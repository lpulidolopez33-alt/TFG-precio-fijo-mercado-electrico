from pathlib import Path
import os
import re
import sys
import subprocess
from datetime import date

import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

import config


PROJECT_DIR = Path(__file__).resolve().parent
CONFIG_PATH = PROJECT_DIR / "config.py"

OUTPUT_DIR = config.OUTPUT_DIR
RUTA_DECISION_COBERTURA = config.RUTA_DECISION_COBERTURA
RUTA_CURVA_CLIENTE = config.RUTA_CURVA_CLIENTE
RUTA_OMIP_MANUAL = config.RUTA_OMIP_MANUAL
MANUAL_DATA_DIR = RUTA_OMIP_MANUAL.parent

RUTA_PREVISION = OUTPUT_DIR / "prevision_commodity_modelo.xlsx"
RUTA_CURVAS_CLIENTES_SIMULADOS = getattr(
    config,
    "RUTA_CURVAS_CLIENTES_SIMULADOS",
    MANUAL_DATA_DIR / "curvas_clientes_simulados.xlsx",
)


st.set_page_config(
    page_title="Calculadora Commodity y Cobertura",
    page_icon="⚡",
    layout="wide",
)


# ============================================================
# ESTILO VISUAL
# ============================================================

st.markdown(
    """
    <style>
    .main-title {
        font-size: 2.6rem;
        font-weight: 800;
        margin-bottom: 0.1rem;
    }

    .subtitle {
        color: #a8a8a8;
        font-size: 1rem;
        margin-bottom: 1.5rem;
    }

    .section-card {
        background: rgba(255,255,255,0.035);
        border: 1px solid rgba(255,255,255,0.08);
        border-radius: 18px;
        padding: 1.2rem;
        margin-bottom: 1rem;
    }

    .metric-card {
        background: linear-gradient(135deg, rgba(255,255,255,0.07), rgba(255,255,255,0.025));
        border: 1px solid rgba(255,255,255,0.09);
        border-radius: 18px;
        padding: 1rem 1.1rem;
        min-height: 118px;
    }

    .metric-label {
        color: #b8b8b8;
        font-size: 0.85rem;
        margin-bottom: 0.45rem;
    }

    .metric-value {
        font-size: 1.85rem;
        font-weight: 750;
        line-height: 1.05;
        white-space: nowrap;
    }

    .metric-unit {
        color: #9d9d9d;
        font-size: 0.85rem;
        margin-top: 0.35rem;
    }

    .ok-pill {
        display: inline-block;
        padding: 0.25rem 0.65rem;
        border-radius: 999px;
        background: rgba(55, 180, 120, 0.18);
        border: 1px solid rgba(55, 180, 120, 0.35);
        color: #74d99f;
        font-size: 0.8rem;
        font-weight: 650;
    }

    .warn-pill {
        display: inline-block;
        padding: 0.25rem 0.65rem;
        border-radius: 999px;
        background: rgba(230, 165, 45, 0.16);
        border: 1px solid rgba(230, 165, 45, 0.35);
        color: #e7bd65;
        font-size: 0.8rem;
        font-weight: 650;
    }

    div[data-testid="stMetric"] {
        overflow: visible;
    }

    .stDataFrame {
        border-radius: 14px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# UTILIDADES
# ============================================================

def formato_numero(valor, decimales=2):
    if valor is None or pd.isna(valor):
        return "—"
    return f"{valor:,.{decimales}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def metric_card(titulo: str, valor: str, unidad: str = ""):
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">{titulo}</div>
            <div class="metric-value">{valor}</div>
            <div class="metric-unit">{unidad}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def leer_excel_seguro(ruta: Path, sheet_name: str | None = None) -> pd.DataFrame:
    if not ruta.exists():
        return pd.DataFrame()

    try:
        if sheet_name is None:
            return pd.read_excel(ruta)

        return pd.read_excel(ruta, sheet_name=sheet_name)

    except Exception as error:
        st.warning(f"No se ha podido leer {ruta.name}: {error}")
        return pd.DataFrame()


def reemplazar_o_anadir_variable_config(
    texto: str,
    nombre_variable: str,
    valor_literal: str,
) -> str:
    patron = rf"^{nombre_variable}\s*=.*$"
    nueva_linea = f"{nombre_variable} = {valor_literal}"

    if re.search(patron, texto, flags=re.MULTILINE):
        texto = re.sub(
            patron,
            nueva_linea,
            texto,
            flags=re.MULTILINE,
        )
    else:
        texto = texto.rstrip() + "\n\n" + nueva_linea + "\n"

    return texto


def aplicar_parametros_config(
    usar_curva_cliente: bool,
    perfil_cliente: str,
    energia_total_mwh: float,
):
    if not CONFIG_PATH.exists():
        st.error(f"No se encuentra config.py en {CONFIG_PATH}")
        return

    texto = CONFIG_PATH.read_text(encoding="utf-8")

    texto = reemplazar_o_anadir_variable_config(
        texto,
        "USAR_CURVA_CLIENTE",
        "True" if usar_curva_cliente else "False",
    )

    texto = reemplazar_o_anadir_variable_config(
        texto,
        "PERFIL_REE_CLIENTE",
        repr(perfil_cliente),
    )

    texto = reemplazar_o_anadir_variable_config(
        texto,
        "ENERGIA_TOTAL_CLIENTE_MWH",
        str(float(energia_total_mwh)),
    )

    CONFIG_PATH.write_text(texto, encoding="utf-8")


def ejecutar_main(
    opcion: str,
    fecha_inicio: str | None = None,
    fecha_fin: str | None = None,
    timeout: int = 1800,
) -> tuple[bool, str]:
    entradas = [opcion]

    if opcion in ["1", "4"]:
        if fecha_inicio is not None:
            entradas.append(fecha_inicio)

        if fecha_fin is not None:
            entradas.append(fecha_fin)

    entrada_texto = "\n".join(entradas) + "\n"

    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"

    try:
        proceso = subprocess.run(
            [sys.executable, "main.py"],
            cwd=PROJECT_DIR,
            input=entrada_texto,
            text=True,
            capture_output=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
            env=env,
        )

        salida = ""
        if proceso.stdout:
            salida += proceso.stdout

        if proceso.stderr:
            salida += "\n\nERRORES / AVISOS:\n" + proceso.stderr

        return proceso.returncode == 0, salida

    except subprocess.TimeoutExpired:
        return False, "El proceso ha superado el tiempo máximo de ejecución."

    except Exception as error:
        return False, f"Error ejecutando main.py: {error}"


def ejecutar_calculo_oferta(
    fecha_inicio: str,
    fecha_fin: str,
    modo_cliente: str,
    perfil_cliente: str,
    energia_total_mwh: float,
) -> tuple[bool, str]:
    logs = []

    if modo_cliente == "Perfil plano":
        aplicar_parametros_config(
            usar_curva_cliente=False,
            perfil_cliente=perfil_cliente,
            energia_total_mwh=energia_total_mwh,
        )

        ok, log = ejecutar_main(
            opcion="4",
            fecha_inicio=fecha_inicio,
            fecha_fin=fecha_fin,
        )

        logs.append("EJECUCIÓN OPCIÓN 4 - PERFIL PLANO")
        logs.append(log)

        return ok, "\n\n".join(logs)

    # Perfil no plano:
    # 1) Generamos previsión base sin curva.
    # 2) Generamos curva horaria.
    # 3) Activamos curva.
    # 4) Recalculamos previsión/decisión con curva.
    aplicar_parametros_config(
        usar_curva_cliente=False,
        perfil_cliente=perfil_cliente,
        energia_total_mwh=energia_total_mwh,
    )

    ok_base, log_base = ejecutar_main(
        opcion="4",
        fecha_inicio=fecha_inicio,
        fecha_fin=fecha_fin,
    )

    logs.append("PASO 1 - PREVISIÓN BASE SIN CURVA")
    logs.append(log_base)

    if not ok_base:
        return False, "\n\n".join(logs)

    ok_curva, log_curva = ejecutar_main(opcion="6")

    logs.append("PASO 2 - GENERACIÓN DE CURVA HORARIA")
    logs.append(log_curva)

    if not ok_curva:
        return False, "\n\n".join(logs)

    aplicar_parametros_config(
        usar_curva_cliente=True,
        perfil_cliente=perfil_cliente,
        energia_total_mwh=energia_total_mwh,
    )

    ok_final, log_final = ejecutar_main(
        opcion="4",
        fecha_inicio=fecha_inicio,
        fecha_fin=fecha_fin,
    )

    logs.append("PASO 3 - CÁLCULO FINAL CON CURVA HORARIA")
    logs.append(log_final)

    return ok_final, "\n\n".join(logs)


def formatear_tabla(df: pd.DataFrame, decimales: int = 3) -> pd.DataFrame:
    if df.empty:
        return df

    df = df.copy()

    for columna in df.columns:
        if pd.api.types.is_float_dtype(df[columna]):
            df[columna] = df[columna].round(decimales)

    return df


# ============================================================
# LECTURA DE RESULTADOS
# ============================================================

resumen_decision = leer_excel_seguro(
    RUTA_DECISION_COBERTURA,
    sheet_name="resumen_decision",
)

plan_cobertura = leer_excel_seguro(
    RUTA_DECISION_COBERTURA,
    sheet_name="plan_cobertura",
)

productos_omip = leer_excel_seguro(
    RUTA_DECISION_COBERTURA,
    sheet_name="productos_omip_usados",
)

curva_horaria = leer_excel_seguro(
    RUTA_CURVA_CLIENTE,
    sheet_name="curva_horaria",
)

resumen_mensual_curva = leer_excel_seguro(
    RUTA_CURVA_CLIENTE,
    sheet_name="resumen_mensual",
)

perfil_horario = leer_excel_seguro(
    RUTA_CURVA_CLIENTE,
    sheet_name="perfil_horario",
)

prevision = leer_excel_seguro(
    RUTA_PREVISION,
    sheet_name="prevision_horaria",
)

omip_manual = leer_excel_seguro(RUTA_OMIP_MANUAL)


# ============================================================
# CABECERA
# ============================================================

st.markdown(
    """
    <div class="main-title">⚡ Calculadora de Commodity y Cobertura</div>
    <div class="subtitle">
    Aplicación local para simular ofertas, comparar modelo vs OMIP, generar curvas de cliente y definir un plan operativo de cobertura.
    </div>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:
    st.header("Panel de control")

    st.markdown("### Datos de la oferta")

    fecha_inicio_input = st.date_input(
        "Fecha inicio oferta",
        value=date(2026, 6, 1),
    )

    fecha_fin_input = st.date_input(
        "Fecha fin oferta",
        value=date(2027, 5, 31),
    )

    energia_total_input = st.number_input(
        "Consumo total cliente [MWh]",
        min_value=1.0,
        value=12000.0,
        step=100.0,
    )

    modo_cliente_input = st.radio(
        "Tipo de cliente",
        options=[
            "Perfil plano",
            "Perfil no plano",
        ],
        index=1,
    )

    perfil_cliente_input = st.selectbox(
        "Tarifa / perfil de curva",
        options=[
            "2.0TD",
            "3.0TD",
            "3.0TDVE",
            "6.XTD",
        ],
        index=1,
    )

    fecha_inicio_str = fecha_inicio_input.strftime("%Y-%m-%d")
    fecha_fin_str = fecha_fin_input.strftime("%Y-%m-%d")

    st.markdown("---")

    st.markdown("### Ejecutar cálculo")

    calcular = st.button(
        "Calcular oferta con estos parámetros",
        use_container_width=True,
        type="primary",
    )

    st.markdown("---")

    st.markdown("### Órdenes del programa")

    col_a, col_b = st.columns(2)

    boton_1 = col_a.button("1 · Proceso completo", use_container_width=True)
    boton_2 = col_b.button("2 · Actualizar datos", use_container_width=True)

    boton_3 = col_a.button("3 · Entrenar modelo", use_container_width=True)
    boton_4 = col_b.button("4 · Previsión", use_container_width=True)

    boton_5 = col_a.button("5 · Backtesting", use_container_width=True)
    boton_6 = col_b.button("6 · Generar curvas", use_container_width=True)

    st.markdown("---")

    if resumen_decision.empty:
        st.markdown('<span class="warn-pill">Sin resultados cargados</span>', unsafe_allow_html=True)
    else:
        st.markdown('<span class="ok-pill">Resultados cargados</span>', unsafe_allow_html=True)


# ============================================================
# EJECUCIONES
# ============================================================

if "ultimo_log" not in st.session_state:
    st.session_state["ultimo_log"] = ""

if calcular:
    with st.spinner("Calculando oferta con los parámetros seleccionados..."):
        ok, log = ejecutar_calculo_oferta(
            fecha_inicio=fecha_inicio_str,
            fecha_fin=fecha_fin_str,
            modo_cliente=modo_cliente_input,
            perfil_cliente=perfil_cliente_input,
            energia_total_mwh=energia_total_input,
        )

        st.session_state["ultimo_log"] = log

    if ok:
        st.success("Cálculo completado correctamente. Actualiza la vista si no ves los nuevos resultados.")
    else:
        st.error("El cálculo ha terminado con errores. Revisa el log.")

if boton_1:
    aplicar_parametros_config(
        usar_curva_cliente=(modo_cliente_input == "Perfil no plano"),
        perfil_cliente=perfil_cliente_input,
        energia_total_mwh=energia_total_input,
    )

    with st.spinner("Ejecutando opción 1..."):
        ok, log = ejecutar_main(
            opcion="1",
            fecha_inicio=fecha_inicio_str,
            fecha_fin=fecha_fin_str,
        )
        st.session_state["ultimo_log"] = log

    st.success("Opción 1 completada.") if ok else st.error("Opción 1 con errores.")

if boton_2:
    aplicar_parametros_config(
        usar_curva_cliente=(modo_cliente_input == "Perfil no plano"),
        perfil_cliente=perfil_cliente_input,
        energia_total_mwh=energia_total_input,
    )

    with st.spinner("Ejecutando opción 2..."):
        ok, log = ejecutar_main(opcion="2")
        st.session_state["ultimo_log"] = log

    st.success("Opción 2 completada.") if ok else st.error("Opción 2 con errores.")

if boton_3:
    aplicar_parametros_config(
        usar_curva_cliente=(modo_cliente_input == "Perfil no plano"),
        perfil_cliente=perfil_cliente_input,
        energia_total_mwh=energia_total_input,
    )

    with st.spinner("Ejecutando opción 3..."):
        ok, log = ejecutar_main(opcion="3")
        st.session_state["ultimo_log"] = log

    st.success("Opción 3 completada.") if ok else st.error("Opción 3 con errores.")

if boton_4:
    with st.spinner("Ejecutando opción 4 con los parámetros de la oferta..."):
        ok, log = ejecutar_calculo_oferta(
            fecha_inicio=fecha_inicio_str,
            fecha_fin=fecha_fin_str,
            modo_cliente=modo_cliente_input,
            perfil_cliente=perfil_cliente_input,
            energia_total_mwh=energia_total_input,
        )

        st.session_state["ultimo_log"] = log

    st.success("Opción 4 completada.") if ok else st.error("Opción 4 con errores.")

if boton_5:
    aplicar_parametros_config(
        usar_curva_cliente=(modo_cliente_input == "Perfil no plano"),
        perfil_cliente=perfil_cliente_input,
        energia_total_mwh=energia_total_input,
    )

    with st.spinner("Ejecutando opción 5..."):
        ok, log = ejecutar_main(opcion="5")
        st.session_state["ultimo_log"] = log

    st.success("Opción 5 completada.") if ok else st.error("Opción 5 con errores.")

if boton_6:
    aplicar_parametros_config(
        usar_curva_cliente=False,
        perfil_cliente=perfil_cliente_input,
        energia_total_mwh=energia_total_input,
    )

    with st.spinner("Generando previsión base y curvas de cliente..."):
        ok_base, log_base = ejecutar_main(
            opcion="4",
            fecha_inicio=fecha_inicio_str,
            fecha_fin=fecha_fin_str,
        )

        ok_curva, log_curva = ejecutar_main(opcion="6")

        st.session_state["ultimo_log"] = (
            "PASO 1 - PREVISIÓN BASE\n"
            + log_base
            + "\n\nPASO 2 - GENERACIÓN DE CURVAS\n"
            + log_curva
        )

    st.success("Curvas generadas.") if ok_base and ok_curva else st.error("Error generando curvas.")


# ============================================================
# TABS
# ============================================================

tabs = st.tabs(
    [
        "Resumen",
        "Plan de cobertura",
        "Curva cliente",
        "Previsión modelo",
        "OMIP",
        "Archivos",
        "Log",
    ]
)


# ============================================================
# TAB RESUMEN
# ============================================================

with tabs[0]:
    st.subheader("Resumen de la oferta")

    if resumen_decision.empty:
        st.info("Todavía no hay resumen disponible. Ejecuta el cálculo desde el panel de control.")
    else:
        fila = resumen_decision.iloc[0]

        col1, col2, col3 = st.columns(3)

        with col1:
            metric_card(
                "Commodity modelo",
                formato_numero(fila.get("commodity_modelo_total_EUR_MWh"), 2),
                "€/MWh",
            )

        with col2:
            metric_card(
                "OMIP equivalente",
                formato_numero(fila.get("omip_equivalente_EUR_MWh"), 2),
                "€/MWh",
            )

        with col3:
            metric_card(
                "Commodity recomendada",
                formato_numero(fila.get("commodity_recomendada_EUR_MWh"), 2),
                "€/MWh",
            )

        col4, col5, col6 = st.columns(3)

        with col4:
            metric_card(
                "Energía cliente",
                formato_numero(fila.get("energia_total_cliente_MWh"), 0),
                "MWh",
            )

        with col5:
            metric_card(
                "Cobertura total",
                formato_numero(100 * fila.get("porcentaje_cobertura_total", 0), 1),
                "%",
            )

        with col6:
            metric_card(
                "Parte abierta",
                formato_numero(100 * fila.get("porcentaje_abierto_total", 0), 1),
                "%",
            )

        col7, col8, col9 = st.columns(3)

        with col7:
            metric_card(
                "Energía cubierta",
                formato_numero(fila.get("energia_cubierta_total_MWh"), 0),
                "MWh",
            )

        with col8:
            metric_card(
                "Energía abierta",
                formato_numero(fila.get("energia_abierta_total_MWh"), 0),
                "MWh",
            )

        with col9:
            metric_card(
                "Importe recomendado",
                formato_numero(fila.get("importe_recomendado_total_EUR"), 0),
                "€",
            )

        st.markdown("### Detalle numérico")
        st.dataframe(
            formatear_tabla(resumen_decision),
            use_container_width=True,
        )


# ============================================================
# TAB PLAN COBERTURA
# ============================================================

with tabs[1]:
    st.subheader("Plan operativo de cobertura")

    if plan_cobertura.empty:
        st.info("No se ha encontrado la hoja plan_cobertura. Ejecuta de nuevo el cálculo.")
    else:
        columnas_principales = [
            "producto_cobertura",
            "fecha_inicio",
            "fecha_fin",
            "decision_operativa",
            "energia_producto_MWh",
            "energia_a_cubrir_MWh",
            "energia_abierta_MWh",
            "cobertura_recomendada_pct",
            "exposicion_abierta_pct",
            "precio_modelo_producto_EUR_MWh",
            "precio_omip_producto_EUR_MWh",
            "commodity_producto_recomendada_EUR_MWh",
            "importe_producto_recomendado_EUR",
            "señal",
        ]

        columnas_existentes = [
            columna
            for columna in columnas_principales
            if columna in plan_cobertura.columns
        ]

        st.dataframe(
            formatear_tabla(plan_cobertura[columnas_existentes]),
            use_container_width=True,
        )

        st.markdown("### Energía por producto de cobertura")

        if "producto_cobertura" in plan_cobertura.columns:
            fig, ax = plt.subplots(figsize=(10, 4))

            df_plot = plan_cobertura.copy()

            if "energia_a_cubrir_MWh" in df_plot.columns and "energia_abierta_MWh" in df_plot.columns:
                ax.bar(
                    df_plot["producto_cobertura"],
                    df_plot["energia_a_cubrir_MWh"],
                    label="Energía a cubrir",
                )

                ax.bar(
                    df_plot["producto_cobertura"],
                    df_plot["energia_abierta_MWh"],
                    bottom=df_plot["energia_a_cubrir_MWh"],
                    label="Energía abierta",
                )

                ax.set_xlabel("Producto")
                ax.set_ylabel("Energía [MWh]")
                ax.set_title("Energía cubierta y abierta por producto")
                ax.tick_params(axis="x", rotation=45)
                ax.legend()
                fig.tight_layout()
                st.pyplot(fig)


# ============================================================
# TAB CURVA CLIENTE
# ============================================================

with tabs[2]:
    st.subheader("Curva horaria del cliente")

    if curva_horaria.empty:
        st.info("No hay curva de cliente cargada. Ejecuta la opción 6.")
    else:
        col1, col2 = st.columns(2)

        with col1:
            st.markdown("### Energía mensual")

            if not resumen_mensual_curva.empty and "año_mes" in resumen_mensual_curva.columns:
                fig, ax = plt.subplots(figsize=(9, 4))
                ax.bar(
                    resumen_mensual_curva["año_mes"],
                    resumen_mensual_curva["energia_MWh"],
                )
                ax.set_xlabel("Mes")
                ax.set_ylabel("Energía [MWh]")
                ax.set_title("Consumo mensual del cliente")
                ax.tick_params(axis="x", rotation=45)
                fig.tight_layout()
                st.pyplot(fig)

        with col2:
            st.markdown("### Perfil horario medio")

            if not perfil_horario.empty and "hora" in perfil_horario.columns:
                fig, ax = plt.subplots(figsize=(9, 4))
                ax.plot(
                    perfil_horario["hora"],
                    perfil_horario["consumo_medio_MWh"],
                    marker="o",
                )
                ax.set_xlabel("Hora")
                ax.set_ylabel("Consumo medio [MWh]")
                ax.set_title("Perfil horario medio")
                ax.set_xticks(range(0, 24))
                fig.tight_layout()
                st.pyplot(fig)

        st.markdown("### Datos horarios")
        st.dataframe(
            formatear_tabla(curva_horaria.head(1000)),
            use_container_width=True,
        )
        st.caption("Se muestran las primeras 1.000 filas.")


# ============================================================
# TAB PREVISIÓN
# ============================================================

with tabs[3]:
    st.subheader("Previsión del modelo")

    if prevision.empty:
        st.info("No hay previsión generada.")
    else:
        df_prev = prevision.copy()

        if "fecha" in df_prev.columns:
            df_prev["fecha"] = pd.to_datetime(df_prev["fecha"], errors="coerce")

        if "precio_omie_previsto" in df_prev.columns:
            df_diario = (
                df_prev.groupby("fecha", as_index=False)
                .agg(precio_medio=("precio_omie_previsto", "mean"))
            )

            fig, ax = plt.subplots(figsize=(11, 4))
            ax.plot(df_diario["fecha"], df_diario["precio_medio"])
            ax.set_xlabel("Fecha")
            ax.set_ylabel("Precio previsto [€/MWh]")
            ax.set_title("Precio OMIE previsto medio diario")
            fig.tight_layout()
            st.pyplot(fig)

        columnas = [
            "fecha_hora_utc",
            "fecha",
            "precio_omie_previsto",
        ]

        columnas_existentes = [
            columna
            for columna in columnas
            if columna in df_prev.columns
        ]

        st.dataframe(
            formatear_tabla(df_prev[columnas_existentes].head(1000)),
            use_container_width=True,
        )


# ============================================================
# TAB OMIP
# ============================================================

with tabs[4]:
    st.subheader("Precios OMIP manuales")

    if omip_manual.empty:
        st.info("No se ha podido cargar omip_manual.xlsx.")
    else:
        st.dataframe(
            formatear_tabla(omip_manual),
            use_container_width=True,
        )


# ============================================================
# TAB ARCHIVOS
# ============================================================

with tabs[5]:
    st.subheader("Archivos generados")

    archivos = [
        RUTA_PREVISION,
        RUTA_DECISION_COBERTURA,
        OUTPUT_DIR / "comparacion_modelo_omip.xlsx",
        RUTA_CURVA_CLIENTE,
        RUTA_CURVAS_CLIENTES_SIMULADOS,
        RUTA_OMIP_MANUAL,
    ]

    tabla_archivos = []

    for ruta in archivos:
        tabla_archivos.append(
            {
                "archivo": ruta.name,
                "existe": ruta.exists(),
                "ruta": str(ruta),
            }
        )

    st.dataframe(
        pd.DataFrame(tabla_archivos),
        use_container_width=True,
    )


# ============================================================
# TAB LOG
# ============================================================

with tabs[6]:
    st.subheader("Log de ejecución")

    if st.session_state["ultimo_log"]:
        st.code(st.session_state["ultimo_log"], language="text")
    else:
        st.info("Todavía no se ha ejecutado ningún proceso desde la app.")