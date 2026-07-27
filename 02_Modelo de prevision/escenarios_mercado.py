from __future__ import annotations

from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
from typing import Iterator

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

import ajuste_omip
import commodity_periodos
import comparacion_omip
import decision_cobertura
import graficas_prevision
from config import (
    OUTPUT_DIR,
    USAR_AJUSTE_OMIP,
    USAR_COMPARACION_OMIP,
    USAR_DECISION_COBERTURA,
)
from prevision_futura import (
    cargar_base_entrenamiento,
    cargar_modelo_entrenado,
    construir_base_futura_modelo,
    crear_resumen_prevision,
    crear_variables_mix_futuro,
    generar_prediccion_precio,
)


FACTOR_VARIACION = 1.20
ESCENARIOS_OUTPUT_DIR = OUTPUT_DIR / "escenarios_mercado"


@dataclass(frozen=True)
class EscenarioMercado:
    codigo: str
    nombre: str
    nombre_corto: str
    nombre_carpeta: str
    variables_modificadas: tuple[str, ...]


ESCENARIOS: dict[str, EscenarioMercado] = {
    "1": EscenarioMercado(
        codigo="1",
        nombre="Incremento del precio del gas en un 20 %",
        nombre_corto="Gas +20 %",
        nombre_carpeta="caso_1_gas_mas_20",
        variables_modificadas=("precio_gas_EUR_MWh",),
    ),
    "2": EscenarioMercado(
        codigo="2",
        nombre="Incremento de la demanda en un 20 %",
        nombre_corto="Demanda +20 %",
        nombre_carpeta="caso_2_demanda_mas_20",
        variables_modificadas=("demanda",),
    ),
    "3": EscenarioMercado(
        codigo="3",
        nombre="Incremento de la generación renovable en un 20 %",
        nombre_corto="Renovables +20 %",
        nombre_carpeta="caso_3_renovables_mas_20",
        variables_modificadas=(
            "generacion_solar",
            "generacion_eolica",
            "generacion_hidraulica",
        ),
    ),
}


def mostrar_menu_escenarios() -> str:
    """Muestra el submenú de sensibilidad y devuelve la opción seleccionada."""

    print("\nPOSIBLES VARIACIONES DEL MERCADO")
    print("-" * 60)
    print("1 - Incremento del precio del gas en un 20 %")
    print("2 - Incremento de la demanda en un 20 %")
    print("3 - Incremento de la generación renovable en un 20 %")
    print("4 - Ejecutar los tres escenarios")
    print("-" * 60)

    while True:
        opcion = input("Selecciona un caso [4]: ").strip()
        if opcion == "":
            opcion = "4"

        if opcion in {"1", "2", "3", "4"}:
            return opcion

        print("Opción no válida. Selecciona 1, 2, 3 o 4.")


def aplicar_escenario_mercado(
    df_base_futura: pd.DataFrame,
    escenario: EscenarioMercado,
) -> pd.DataFrame:
    """
    Aplica la variación sobre una copia independiente de la base futura.

    Las entradas originales, la base histórica y el modelo entrenado no se
    modifican. Después de variar las entradas se recalculan todas las variables
    derivadas utilizadas por el modelo.
    """

    df = df_base_futura.copy(deep=True)

    columnas_faltantes = [
        columna
        for columna in escenario.variables_modificadas
        if columna not in df.columns
    ]
    if columnas_faltantes:
        raise ValueError(
            "No se puede aplicar el escenario porque faltan las columnas: "
            f"{columnas_faltantes}"
        )

    for columna in escenario.variables_modificadas:
        df[columna] = (
            pd.to_numeric(df[columna], errors="coerce") * FACTOR_VARIACION
        )

    # Recalcular hueco térmico, demanda neta, porcentajes renovables y variables
    # vinculadas al precio del gas después de modificar las entradas.
    df = crear_variables_mix_futuro(df)

    columnas_numericas = df.select_dtypes(include=[np.number]).columns
    if np.isinf(df[columnas_numericas].to_numpy()).any():
        raise ValueError(
            "El escenario ha generado valores infinitos. Revisa los datos de "
            "demanda y generación utilizados."
        )

    df["escenario_mercado"] = escenario.nombre
    df["factor_variacion"] = FACTOR_VARIACION

    return df


def _guardar_prevision_escenario(
    df_prevision: pd.DataFrame,
    directorio_salida: Path,
    nombre_escenario: str,
) -> Path:
    """Guarda la previsión con la misma estructura empleada por la opción 2."""

    directorio_salida.mkdir(parents=True, exist_ok=True)
    ruta_salida = directorio_salida / "prevision_commodity_modelo.xlsx"

    resumen_diario, resumen_mensual, resumen_global = crear_resumen_prevision(
        df_prevision
    )
    resumen_global = resumen_global.copy()
    resumen_global["escenario_mercado"] = nombre_escenario

    try:
        with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
            resumen_global.to_excel(writer, sheet_name="resumen_global", index=False)
            resumen_mensual.to_excel(writer, sheet_name="resumen_mensual", index=False)
            resumen_diario.to_excel(writer, sheet_name="resumen_diario", index=False)
            df_prevision.to_excel(writer, sheet_name="prevision_horaria", index=False)
    except PermissionError as exc:
        raise PermissionError(
            f"No se puede guardar {ruta_salida}. Probablemente está abierto "
            "en Excel. Ciérralo y vuelve a ejecutar."
        ) from exc

    print(f"Previsión del escenario guardada en: {ruta_salida}")
    return ruta_salida


@contextmanager
def _rutas_aisladas_escenario(
    directorio_salida: Path,
    directorio_graficas: Path,
) -> Iterator[None]:
    """
    Redirige temporalmente las salidas de la opción 2 a la carpeta del caso.

    Al terminar se restauran todas las rutas originales, por lo que ni la
    opción 2 ni un escenario posterior reutilizan los resultados modificados.
    """

    directorio_salida.mkdir(parents=True, exist_ok=True)
    directorio_graficas.mkdir(parents=True, exist_ok=True)

    cambios = [
        (commodity_periodos, "OUTPUT_DIR", directorio_salida),
        (graficas_prevision, "OUTPUT_DIR", directorio_salida),
        (graficas_prevision, "GRAPH_DIR", directorio_graficas),
        (comparacion_omip, "OUTPUT_DIR", directorio_salida),
        (
            comparacion_omip,
            "RUTA_COMPARACION_OMIP",
            directorio_salida / "comparacion_modelo_omip.xlsx",
        ),
        (decision_cobertura, "OUTPUT_DIR", directorio_salida),
        (
            decision_cobertura,
            "RUTA_DECISION_COBERTURA",
            directorio_salida / "decision_commodity_cobertura.xlsx",
        ),
        (ajuste_omip, "OUTPUT_DIR", directorio_salida),
        (
            ajuste_omip,
            "RUTA_COMPARACION_OMIP",
            directorio_salida / "comparacion_modelo_omip.xlsx",
        ),
        (
            ajuste_omip,
            "RUTA_PREVISION_HIBRIDA_OMIP",
            directorio_salida / "prevision_commodity_hibrida_omip.xlsx",
        ),
        (
            ajuste_omip,
            "RUTA_RESUMEN_HIBRIDO_OMIP",
            directorio_salida / "resumen_commodity_hibrida_omip.xlsx",
        ),
    ]

    valores_originales = [
        (modulo, atributo, getattr(modulo, atributo))
        for modulo, atributo, _ in cambios
    ]
    exportacion_original = decision_cobertura.exportar_resultados_calculadora

    for modulo, atributo, valor_nuevo in cambios:
        setattr(modulo, atributo, valor_nuevo)

    # La sensibilidad debe calcular previsión, OMIP y cobertura, pero no debe
    # escribir resultados provisionales en la calculadora principal del usuario.
    decision_cobertura.exportar_resultados_calculadora = (
        lambda *args, **kwargs: print(
            "Exportación a la calculadora omitida para mantener aislado el escenario."
        )
    )

    try:
        yield
    finally:
        for modulo, atributo, valor_original in valores_originales:
            setattr(modulo, atributo, valor_original)
        decision_cobertura.exportar_resultados_calculadora = exportacion_original


def _ejecutar_herramienta_completa_para_prevision(
    df_prevision: pd.DataFrame,
    directorio_caso: Path,
    nombre_escenario: str,
) -> None:
    """
    Ejecuta el mismo flujo posterior que la opción 2 sobre una previsión dada.

    Se generan la previsión, la commodity por periodos, las gráficas, la
    comparación OMIP y la decisión de cobertura en una carpeta independiente.
    """

    directorio_salida = directorio_caso / "salidas"
    directorio_graficas = directorio_caso / "graficas"

    _guardar_prevision_escenario(
        df_prevision=df_prevision,
        directorio_salida=directorio_salida,
        nombre_escenario=nombre_escenario,
    )

    with _rutas_aisladas_escenario(
        directorio_salida=directorio_salida,
        directorio_graficas=directorio_graficas,
    ):
        commodity_periodos.generar_commodity_por_periodos()
        graficas_prevision.generar_graficas_prevision()

        if USAR_COMPARACION_OMIP:
            comparacion_omip.generar_comparacion_modelo_omip()
        else:
            print("\nComparación con OMIP desactivada.")

        if USAR_DECISION_COBERTURA:
            decision_cobertura.generar_decision_cobertura()
        else:
            print("\nDecisión de cobertura desactivada.")

        if USAR_AJUSTE_OMIP:
            ajuste_omip.generar_ajuste_omip()
        else:
            print("\nAjuste híbrido OMIP-modelo desactivado.")


def _crear_tablas_comparacion(
    previsiones: dict[str, pd.DataFrame],
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Crea comparaciones global, mensual y horaria entre base y escenarios."""

    resumenes_globales: list[dict[str, float | str]] = []
    tablas_mensuales: list[pd.DataFrame] = []
    tabla_horaria: pd.DataFrame | None = None

    precio_base = float(previsiones["Caso base"]["precio_omie_previsto"].mean())

    for nombre, df_original in previsiones.items():
        df = df_original.copy()
        df["fecha"] = pd.to_datetime(df["fecha"])
        precio_medio = float(df["precio_omie_previsto"].mean())
        diferencia = precio_medio - precio_base
        variacion_pct = diferencia / precio_base * 100 if precio_base != 0 else np.nan

        resumenes_globales.append(
            {
                "escenario": nombre,
                "commodity_media_EUR_MWh": precio_medio,
                "diferencia_frente_base_EUR_MWh": diferencia,
                "variacion_frente_base_pct": variacion_pct,
            }
        )

        df["periodo"] = df["fecha"].dt.to_period("M").astype(str)
        mensual = (
            df.groupby("periodo", as_index=False)["precio_omie_previsto"]
            .mean()
            .rename(columns={"precio_omie_previsto": nombre})
        )
        tablas_mensuales.append(mensual)

        columna_horaria = df[["fecha_hora_utc", "precio_omie_previsto"]].copy()
        columna_horaria = columna_horaria.rename(
            columns={"precio_omie_previsto": nombre}
        )
        if tabla_horaria is None:
            tabla_horaria = columna_horaria
        else:
            tabla_horaria = tabla_horaria.merge(
                columna_horaria,
                on="fecha_hora_utc",
                how="outer",
            )

    comparacion_mensual = tablas_mensuales[0]
    for tabla in tablas_mensuales[1:]:
        comparacion_mensual = comparacion_mensual.merge(
            tabla,
            on="periodo",
            how="outer",
        )

    resumen_global = pd.DataFrame(resumenes_globales)
    assert tabla_horaria is not None

    for nombre in previsiones:
        if nombre == "Caso base":
            continue
        tabla_horaria[f"Diferencia {nombre} - base"] = (
            tabla_horaria[nombre] - tabla_horaria["Caso base"]
        )

    return resumen_global, comparacion_mensual, tabla_horaria


def _guardar_comparacion_escenarios(
    previsiones: dict[str, pd.DataFrame],
) -> None:
    """
    Guarda las tablas y las gráficas comparativas de los escenarios.

    El orden de las gráficas es:
    1. Caso base frente al escenario de gas.
    2. Caso base frente al escenario de demanda.
    3. Caso base frente al escenario de renovables.
    4. Comparación conjunta del caso base y todos los escenarios ejecutados.
    5. Peso relativo de cada variable en la sensibilidad del precio previsto.

    Cuando se ejecuta un único escenario, solo se generan las comparaciones
    correspondientes a los casos disponibles.
    """

    ESCENARIOS_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    resumen_global, comparacion_mensual, comparacion_horaria = (
        _crear_tablas_comparacion(previsiones)
    )

    nombres_variables = {
        "Gas +20 %": "Precio del gas",
        "Demanda +20 %": "Demanda",
        "Renovables +20 %": "Generación renovable",
    }
    orden_escenarios = [
        "Gas +20 %",
        "Demanda +20 %",
        "Renovables +20 %",
    ]
    prefijos_archivo = {
        "Gas +20 %": "01_base_vs_caso_1_gas",
        "Demanda +20 %": "02_base_vs_caso_2_demanda",
        "Renovables +20 %": "03_base_vs_caso_3_renovables",
    }

    impacto_variables = resumen_global.loc[
        resumen_global["escenario"] != "Caso base"
    ].copy()
    impacto_variables["variable_analizada"] = impacto_variables[
        "escenario"
    ].map(nombres_variables).fillna(impacto_variables["escenario"])
    impacto_variables["impacto_absoluto_EUR_MWh"] = impacto_variables[
        "diferencia_frente_base_EUR_MWh"
    ].abs()
    suma_impactos = float(impacto_variables["impacto_absoluto_EUR_MWh"].sum())
    if suma_impactos > 0:
        impacto_variables["peso_relativo_pct"] = (
            impacto_variables["impacto_absoluto_EUR_MWh"] / suma_impactos * 100
        )
    else:
        impacto_variables["peso_relativo_pct"] = 0.0
    impacto_variables["peso_relativo_firmado_pct"] = np.where(
        impacto_variables["diferencia_frente_base_EUR_MWh"] >= 0,
        impacto_variables["peso_relativo_pct"],
        -impacto_variables["peso_relativo_pct"],
    )
    impacto_variables["orden_impacto"] = (
        impacto_variables["impacto_absoluto_EUR_MWh"]
        .rank(method="dense", ascending=False)
        .astype(int)
    )
    impacto_variables = impacto_variables.sort_values(
        ["orden_impacto", "variable_analizada"]
    ).reset_index(drop=True)

    ruta_excel = ESCENARIOS_OUTPUT_DIR / "comparacion_escenarios_mercado.xlsx"
    try:
        with pd.ExcelWriter(ruta_excel, engine="openpyxl") as writer:
            resumen_global.to_excel(writer, sheet_name="resumen_global", index=False)
            comparacion_mensual.to_excel(
                writer, sheet_name="comparacion_mensual", index=False
            )
            comparacion_horaria.to_excel(
                writer, sheet_name="comparacion_horaria", index=False
            )
            impacto_variables.to_excel(
                writer, sheet_name="impacto_variables", index=False
            )
    except PermissionError as exc:
        raise PermissionError(
            f"No se puede guardar {ruta_excel}. Probablemente está abierto en Excel."
        ) from exc

    rutas_generadas: list[Path] = []

    # 1-3. Comparaciones individuales: caso base frente a cada escenario.
    for escenario in orden_escenarios:
        if escenario not in comparacion_mensual.columns:
            continue

        plt.figure(figsize=(12, 6))
        plt.plot(
            comparacion_mensual["periodo"],
            comparacion_mensual["Caso base"],
            marker="o",
            label="Caso base",
        )
        plt.plot(
            comparacion_mensual["periodo"],
            comparacion_mensual[escenario],
            marker="o",
            label=escenario,
        )
        plt.title(f"Caso base frente a {escenario}")
        plt.xlabel("Mes")
        plt.ylabel("Precio previsto [€/MWh]")
        plt.xticks(rotation=45, ha="right")
        plt.legend()
        plt.grid(axis="y", alpha=0.25)
        plt.tight_layout()

        ruta_individual = (
            ESCENARIOS_OUTPUT_DIR / f"{prefijos_archivo[escenario]}.png"
        )
        plt.savefig(ruta_individual, dpi=300, bbox_inches="tight")
        plt.close()
        rutas_generadas.append(ruta_individual)

    # 4. Comparación conjunta del caso base y todos los escenarios ejecutados.
    plt.figure(figsize=(12, 6))
    columnas_conjuntas = ["Caso base"] + [
        escenario
        for escenario in orden_escenarios
        if escenario in comparacion_mensual.columns
    ]
    for columna in columnas_conjuntas:
        plt.plot(
            comparacion_mensual["periodo"],
            comparacion_mensual[columna],
            marker="o",
            label=columna,
        )
    plt.title("Comparación conjunta del caso base y los escenarios de mercado")
    plt.xlabel("Mes")
    plt.ylabel("Precio previsto [€/MWh]")
    plt.xticks(rotation=45, ha="right")
    plt.legend()
    plt.grid(axis="y", alpha=0.25)
    plt.tight_layout()
    ruta_conjunta = ESCENARIOS_OUTPUT_DIR / "04_comparacion_todos_los_casos.png"
    plt.savefig(ruta_conjunta, dpi=300, bbox_inches="tight")
    plt.close()
    rutas_generadas.append(ruta_conjunta)

    # 5. Peso relativo de cada variable en la sensibilidad del precio.
    if not impacto_variables.empty:
        impacto_grafica = impacto_variables.sort_values(
            "impacto_absoluto_EUR_MWh", ascending=False
        ).copy()
        color_barras = "#562843"
        etiquetas = impacto_grafica["variable_analizada"].tolist()
        pesos_firmados = impacto_grafica["peso_relativo_firmado_pct"].tolist()
        pesos_absolutos = impacto_grafica["peso_relativo_pct"].tolist()
        diferencias = impacto_grafica[
            "diferencia_frente_base_EUR_MWh"
        ].tolist()

        plt.figure(figsize=(10.5, 5.8))
        barras = plt.barh(
            etiquetas,
            pesos_firmados,
            color=color_barras,
        )
        plt.axvline(0, color="black", linewidth=1)
        plt.title(
            "Peso relativo de las variables en la sensibilidad del precio previsto"
        )
        plt.xlabel("Peso relativo sobre la variación total del precio [%]")
        plt.ylabel("Variable modificada")
        plt.grid(axis="x", alpha=0.25)
        plt.gca().invert_yaxis()

        limite = max(abs(valor) for valor in pesos_firmados) if pesos_firmados else 0.0
        margen = max(limite * 0.12, 3.0)
        for barra, peso_firmado, peso_absoluto, diferencia in zip(
            barras, pesos_firmados, pesos_absolutos, diferencias
        ):
            plt.text(
                peso_firmado / 2,
                barra.get_y() + barra.get_height() / 2,
                f"{peso_absoluto:.2f} % | {diferencia:+.2f} €/MWh",
                va="center",
                ha="center",
                color="white",
                fontweight="bold",
            )

        if limite > 0:
            plt.xlim(
                min(min(pesos_firmados) - margen, -margen),
                max(max(pesos_firmados) + margen, margen),
            )
        plt.tight_layout()
        ruta_impacto = (
            ESCENARIOS_OUTPUT_DIR / "05_peso_relativo_variables_precio.png"
        )
        plt.savefig(ruta_impacto, dpi=300, bbox_inches="tight")
        plt.close()
        rutas_generadas.append(ruta_impacto)

    print(f"Comparación de escenarios guardada en: {ruta_excel}")
    for ruta in rutas_generadas:
        print(f"Gráfica comparativa guardada en: {ruta}")

def generar_escenarios_mercado(
    fecha_inicio_prevision: str,
    fecha_fin_prevision: str,
    opcion_escenario: str,
) -> None:
    """
    Ejecuta el caso base y uno o todos los escenarios con el modelo entrenado.

    Cada escenario vuelve a aplicar el modelo sobre entradas modificadas y,
    después, ejecuta el mismo flujo de la opción 2. Todos los resultados se
    guardan de manera independiente dentro de ``salidas/escenarios_mercado``.
    """

    if opcion_escenario not in {"1", "2", "3", "4"}:
        raise ValueError("La opción de escenario debe ser 1, 2, 3 o 4.")

    print("\nANÁLISIS DE POSIBLES VARIACIONES DEL MERCADO")
    print("-" * 60)
    print("Cargando el modelo y construyendo el caso base...")

    paquete_modelo = cargar_modelo_entrenado()
    df_historico = cargar_base_entrenamiento()

    df_futuro_base = construir_base_futura_modelo(
        df_historico=df_historico,
        fecha_inicio_prevision=fecha_inicio_prevision,
        fecha_fin_prevision=fecha_fin_prevision,
    )
    df_prevision_base = generar_prediccion_precio(
        df_futuro=df_futuro_base.copy(deep=True),
        paquete_modelo=paquete_modelo,
    )
    df_prevision_base["escenario_mercado"] = "Caso base"

    previsiones: dict[str, pd.DataFrame] = {"Caso base": df_prevision_base}

    print("\nEjecutando nuevamente la herramienta para el caso base...")
    _ejecutar_herramienta_completa_para_prevision(
        df_prevision=df_prevision_base,
        directorio_caso=ESCENARIOS_OUTPUT_DIR / "caso_base",
        nombre_escenario="Caso base",
    )

    codigos = ["1", "2", "3"] if opcion_escenario == "4" else [opcion_escenario]

    for codigo in codigos:
        escenario = ESCENARIOS[codigo]
        print(f"\nCASO {codigo}: {escenario.nombre}")
        print("Aplicando la variación y ejecutando nuevamente la previsión...")

        # Cada escenario parte siempre de la misma base original.
        df_futuro_escenario = aplicar_escenario_mercado(
            df_base_futura=df_futuro_base,
            escenario=escenario,
        )
        df_prevision_escenario = generar_prediccion_precio(
            df_futuro=df_futuro_escenario,
            paquete_modelo=paquete_modelo,
        )

        previsiones[escenario.nombre_corto] = df_prevision_escenario

        _ejecutar_herramienta_completa_para_prevision(
            df_prevision=df_prevision_escenario,
            directorio_caso=ESCENARIOS_OUTPUT_DIR / escenario.nombre_carpeta,
            nombre_escenario=escenario.nombre,
        )

        precio_base = float(df_prevision_base["precio_omie_previsto"].mean())
        precio_escenario = float(
            df_prevision_escenario["precio_omie_previsto"].mean()
        )
        diferencia = precio_escenario - precio_base
        variacion = diferencia / precio_base * 100 if precio_base != 0 else np.nan

        print(f"Commodity media del caso base: {precio_base:.2f} €/MWh")
        print(f"Commodity media del escenario: {precio_escenario:.2f} €/MWh")
        print(f"Diferencia: {diferencia:+.2f} €/MWh ({variacion:+.2f} %)")

    _guardar_comparacion_escenarios(previsiones)

    print("-" * 60)
    print("Escenarios de mercado completados correctamente.")