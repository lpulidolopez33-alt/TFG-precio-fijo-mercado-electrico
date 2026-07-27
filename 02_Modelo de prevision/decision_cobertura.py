import pandas as pd
import numpy as np

from config import (
    OUTPUT_DIR,
    RUTA_OMIP_MANUAL,
    RUTA_DECISION_COBERTURA,
    PORCENTAJE_COBERTURA_MODELO_MAYOR_OMIP,
    PORCENTAJE_COBERTURA_MODELO_MENOR_OMIP,
    PRIMA_RIESGO_PARTE_ABIERTA_EUR_MWH,
    USAR_CURVA_CLIENTE,
    RUTA_CURVA_CLIENTE,
    COLUMNA_CONSUMO_CLIENTE,
    ENERGIA_TOTAL_CLIENTE_MWH,
)

from generar_curva_cliente import (
    cargar_perfil_ree_2026,
    crear_tabla_coeficientes_perfil,
    generar_curva_horaria,
    calcular_resumen_curva,
    guardar_curva_cliente_seleccionada,
)

from exportacion_calculadora import exportar_resultados_calculadora


def cargar_prevision_modelo() -> pd.DataFrame:
    """
    Carga la previsión horaria generada por el modelo.
    """

    ruta_prevision = OUTPUT_DIR / "prevision_commodity_modelo.xlsx"

    if not ruta_prevision.exists():
        raise FileNotFoundError(
            f"No existe {ruta_prevision}. Ejecuta primero la opción 4."
        )

    df = pd.read_excel(
        ruta_prevision,
        sheet_name="prevision_horaria",
    )

    return df


def cargar_omip_manual() -> pd.DataFrame:
    """
    Carga los precios OMIP introducidos manualmente.
    """

    if not RUTA_OMIP_MANUAL.exists():
        raise FileNotFoundError(
            f"No existe {RUTA_OMIP_MANUAL}. Crea o actualiza omip_manual.xlsx."
        )

    df = pd.read_excel(RUTA_OMIP_MANUAL)

    return df


def preparar_prevision(df: pd.DataFrame) -> pd.DataFrame:
    """
    Prepara la previsión del modelo.
    """

    df = df.copy()

    columnas_obligatorias = [
        "fecha",
        "precio_omie_previsto",
    ]

    for columna in columnas_obligatorias:
        if columna not in df.columns:
            raise ValueError(
                f"Falta la columna '{columna}' en prevision_commodity_modelo.xlsx"
            )

    df["fecha"] = pd.to_datetime(
        df["fecha"],
        errors="coerce",
    )

    df["precio_omie_previsto"] = pd.to_numeric(
        df["precio_omie_previsto"],
        errors="coerce",
    )

    if "fecha_hora_utc" in df.columns:
        df["fecha_hora_utc"] = pd.to_datetime(
            df["fecha_hora_utc"],
            utc=True,
            errors="coerce",
        )

    df = df.dropna(
        subset=[
            "fecha",
            "precio_omie_previsto",
        ]
    ).copy()

    return df


def preparar_omip(df: pd.DataFrame) -> pd.DataFrame:
    """
    Prepara la tabla manual de OMIP.

    Acepta tanto 'tipo_producto' como 'tipo_prod' por si Excel muestra
    el encabezado abreviado.
    """

    df = df.copy()

    if "tipo_producto" not in df.columns and "tipo_prod" in df.columns:
        df = df.rename(columns={"tipo_prod": "tipo_producto"})

    columnas_obligatorias = [
        "producto",
        "tipo_producto",
        "fecha_inicio",
        "fecha_fin",
        "precio_omip_EUR_MWh",
    ]

    for columna in columnas_obligatorias:
        if columna not in df.columns:
            raise ValueError(
                f"Falta la columna '{columna}' en omip_manual.xlsx"
            )

    df["fecha_inicio"] = pd.to_datetime(
        df["fecha_inicio"],
        errors="coerce",
        dayfirst=False,
    )

    df["fecha_fin"] = pd.to_datetime(
        df["fecha_fin"],
        errors="coerce",
        dayfirst=False,
    )

    df["precio_omip_EUR_MWh"] = (
        df["precio_omip_EUR_MWh"]
        .astype(str)
        .str.replace(",", ".", regex=False)
    )

    df["precio_omip_EUR_MWh"] = pd.to_numeric(
        df["precio_omip_EUR_MWh"],
        errors="coerce",
    )

    df["tipo_producto"] = (
        df["tipo_producto"]
        .astype(str)
        .str.strip()
    )

    df["producto"] = (
        df["producto"]
        .astype(str)
        .str.strip()
    )

    df = df.dropna(
        subset=[
            "producto",
            "tipo_producto",
            "fecha_inicio",
            "fecha_fin",
        ]
    ).copy()

    df["dias_producto"] = (
        df["fecha_fin"].dt.normalize()
        - df["fecha_inicio"].dt.normalize()
    ).dt.days + 1

    return df


def buscar_producto_que_cubre(
    df_omip: pd.DataFrame,
    tipo_producto: str,
    fecha_inicio: pd.Timestamp,
    fecha_fin: pd.Timestamp,
) -> pd.Series | None:
    """
    Busca un producto OMIP de un tipo concreto que cubra completamente
    el tramo indicado.
    """

    df = df_omip.copy()

    df["tipo_producto_normalizado"] = (
        df["tipo_producto"]
        .astype(str)
        .str.lower()
        .str.strip()
    )

    mascara = (
        df["tipo_producto_normalizado"].str.contains(tipo_producto.lower())
        & (df["fecha_inicio"] <= fecha_inicio)
        & (df["fecha_fin"] >= fecha_fin)
        & (df["precio_omip_EUR_MWh"].notna())
    )

    candidatos = df.loc[mascara].copy()

    if candidatos.empty:
        return None

    candidatos = candidatos.sort_values(
        by=[
            "dias_producto",
            "fecha_inicio",
        ],
        ascending=[
            True,
            True,
        ],
    )

    return candidatos.iloc[0]


def obtener_inicio_fin_trimestre(
    año: int,
    trimestre: int,
) -> tuple[pd.Timestamp, pd.Timestamp]:
    """
    Devuelve el inicio y fin natural de un trimestre.
    """

    mes_inicio = (trimestre - 1) * 3 + 1

    fecha_inicio = pd.Timestamp(
        year=año,
        month=mes_inicio,
        day=1,
    )

    fecha_fin = fecha_inicio + pd.offsets.QuarterEnd(0)

    return fecha_inicio, fecha_fin


def crear_bloques_trimestrales_oferta(
    fecha_inicio_oferta: pd.Timestamp,
    fecha_fin_oferta: pd.Timestamp,
) -> list[dict]:
    """
    Divide la oferta en bloques por trimestre natural.
    """

    fecha_inicio_oferta = fecha_inicio_oferta.normalize()
    fecha_fin_oferta = fecha_fin_oferta.normalize()

    meses = pd.date_range(
        start=fecha_inicio_oferta.replace(day=1),
        end=fecha_fin_oferta.replace(day=1),
        freq="MS",
    )

    bloques = {}

    for mes in meses:
        año = mes.year
        trimestre = ((mes.month - 1) // 3) + 1

        inicio_trimestre, fin_trimestre = obtener_inicio_fin_trimestre(
            año=año,
            trimestre=trimestre,
        )

        inicio_bloque = max(
            inicio_trimestre,
            fecha_inicio_oferta,
        )

        fin_bloque = min(
            fin_trimestre,
            fecha_fin_oferta,
        )

        clave = (
            año,
            trimestre,
        )

        bloques[clave] = {
            "año": año,
            "trimestre": trimestre,
            "inicio_bloque": inicio_bloque,
            "fin_bloque": fin_bloque,
            "inicio_trimestre": inicio_trimestre,
            "fin_trimestre": fin_trimestre,
            "es_trimestre_completo": (
                inicio_bloque == inicio_trimestre
                and fin_bloque == fin_trimestre
            ),
        }

    return list(bloques.values())


def crear_tramos_mensuales(
    fecha_inicio: pd.Timestamp,
    fecha_fin: pd.Timestamp,
) -> list[dict]:
    """
    Divide un tramo parcial de trimestre en meses o partes de meses.
    """

    meses = pd.date_range(
        start=fecha_inicio.replace(day=1),
        end=fecha_fin.replace(day=1),
        freq="MS",
    )

    tramos = []

    for mes in meses:
        inicio_mes = mes
        fin_mes = mes + pd.offsets.MonthEnd(0)

        inicio_tramo = max(
            inicio_mes,
            fecha_inicio,
        )

        fin_tramo = min(
            fin_mes,
            fecha_fin,
        )

        tramos.append(
            {
                "inicio": inicio_tramo,
                "fin": fin_tramo,
            }
        )

    return tramos


def asignar_producto_a_tramo(
    df: pd.DataFrame,
    producto: pd.Series,
    fecha_inicio: pd.Timestamp,
    fecha_fin: pd.Timestamp,
    criterio_asignacion: str,
) -> pd.DataFrame:
    """
    Asigna un producto OMIP a un tramo de fechas.
    """

    fechas = df["fecha"].dt.normalize()

    mascara = (
        (fechas >= fecha_inicio)
        & (fechas <= fecha_fin)
    )

    df.loc[mascara, "producto_omip"] = producto["producto"]
    df.loc[mascara, "precio_omip_EUR_MWh"] = producto["precio_omip_EUR_MWh"]
    df.loc[mascara, "tipo_producto_omip"] = producto["tipo_producto"]
    df.loc[mascara, "criterio_producto_omip"] = criterio_asignacion

    return df


def asignar_productos_omip_a_prevision(
    df_prevision: pd.DataFrame,
    df_omip: pd.DataFrame,
) -> pd.DataFrame:
    """
    Asigna productos OMIP a la previsión.

    Regla:
    - Si la oferta contiene un trimestre completo, se usa el producto trimestral.
    - Si la oferta contiene solo parte de un trimestre, se intentan usar meses.
    - Si no existe precio mensual para ese tramo parcial, se usa el trimestre como aproximación.
    - Si tampoco existe trimestre, se intenta usar el anual.
    - Si no existe ningún producto, se usa el modelo para ese tramo.
    """

    df = df_prevision.copy()

    df["producto_omip"] = pd.NA
    df["precio_omip_EUR_MWh"] = pd.NA
    df["tipo_producto_omip"] = pd.NA
    df["criterio_producto_omip"] = pd.NA

    df["fecha"] = pd.to_datetime(
        df["fecha"],
        errors="coerce",
    )

    fecha_inicio_oferta = df["fecha"].min().normalize()
    fecha_fin_oferta = df["fecha"].max().normalize()

    bloques = crear_bloques_trimestrales_oferta(
        fecha_inicio_oferta=fecha_inicio_oferta,
        fecha_fin_oferta=fecha_fin_oferta,
    )

    for bloque in bloques:
        inicio_bloque = bloque["inicio_bloque"]
        fin_bloque = bloque["fin_bloque"]

        # ============================================================
        # CASO 1: TRIMESTRE COMPLETO DENTRO DE LA OFERTA
        # ============================================================
        if bloque["es_trimestre_completo"]:
            producto_trimestral = buscar_producto_que_cubre(
                df_omip=df_omip,
                tipo_producto="trimestral",
                fecha_inicio=bloque["inicio_trimestre"],
                fecha_fin=bloque["fin_trimestre"],
            )

            if producto_trimestral is not None:
                df = asignar_producto_a_tramo(
                    df=df,
                    producto=producto_trimestral,
                    fecha_inicio=inicio_bloque,
                    fecha_fin=fin_bloque,
                    criterio_asignacion=(
                        "Trimestre completo: se usa producto trimestral."
                    ),
                )

            else:
                producto_anual = buscar_producto_que_cubre(
                    df_omip=df_omip,
                    tipo_producto="anual",
                    fecha_inicio=inicio_bloque,
                    fecha_fin=fin_bloque,
                )

                if producto_anual is not None:
                    df = asignar_producto_a_tramo(
                        df=df,
                        producto=producto_anual,
                        fecha_inicio=inicio_bloque,
                        fecha_fin=fin_bloque,
                        criterio_asignacion=(
                            "Trimestre completo sin producto trimestral: "
                            "se usa producto anual."
                        ),
                    )

        # ============================================================
        # CASO 2: TRIMESTRE PARCIAL DENTRO DE LA OFERTA
        # ============================================================
        else:
            tramos_mensuales = crear_tramos_mensuales(
                fecha_inicio=inicio_bloque,
                fecha_fin=fin_bloque,
            )

            for tramo in tramos_mensuales:
                # Primero intentamos usar el producto mensual.
                producto_mensual = buscar_producto_que_cubre(
                    df_omip=df_omip,
                    tipo_producto="mensual",
                    fecha_inicio=tramo["inicio"],
                    fecha_fin=tramo["fin"],
                )

                if producto_mensual is not None:
                    df = asignar_producto_a_tramo(
                        df=df,
                        producto=producto_mensual,
                        fecha_inicio=tramo["inicio"],
                        fecha_fin=tramo["fin"],
                        criterio_asignacion=(
                            "Trimestre parcial: se usa producto mensual."
                        ),
                    )

                else:
                    # Si no existe mensual, usamos el producto trimestral
                    # que cubra el trimestre natural completo.
                    producto_trimestral = buscar_producto_que_cubre(
                        df_omip=df_omip,
                        tipo_producto="trimestral",
                        fecha_inicio=bloque["inicio_trimestre"],
                        fecha_fin=bloque["fin_trimestre"],
                    )

                    if producto_trimestral is not None:
                        df = asignar_producto_a_tramo(
                            df=df,
                            producto=producto_trimestral,
                            fecha_inicio=tramo["inicio"],
                            fecha_fin=tramo["fin"],
                            criterio_asignacion=(
                                "Trimestre parcial sin precio mensual: "
                                "se usa el producto trimestral como referencia."
                            ),
                        )

                    else:
                        # Si tampoco existe trimestre, intentamos usar anual.
                        producto_anual = buscar_producto_que_cubre(
                            df_omip=df_omip,
                            tipo_producto="anual",
                            fecha_inicio=tramo["inicio"],
                            fecha_fin=tramo["fin"],
                        )

                        if producto_anual is not None:
                            df = asignar_producto_a_tramo(
                                df=df,
                                producto=producto_anual,
                                fecha_inicio=tramo["inicio"],
                                fecha_fin=tramo["fin"],
                                criterio_asignacion=(
                                    "Trimestre parcial sin mensual ni trimestral: "
                                    "se usa producto anual."
                                ),
                            )

    df["precio_omip_EUR_MWh"] = pd.to_numeric(
        df["precio_omip_EUR_MWh"],
        errors="coerce",
    )

    df["referencia_commodity_EUR_MWh"] = df["precio_omip_EUR_MWh"].fillna(
        df["precio_omie_previsto"]
    )

    df["criterio_producto_omip"] = df["criterio_producto_omip"].fillna(
        "Sin producto OMIP aplicable: se usa modelo para este tramo."
    )

    return df

def cargar_curva_cliente() -> pd.DataFrame:
    """
    Carga la curva horaria de consumo del cliente.
    """

    if not RUTA_CURVA_CLIENTE.exists():
        raise FileNotFoundError(
            f"No existe {RUTA_CURVA_CLIENTE}. "
            f"Crea el archivo curva_cliente.xlsx o desactiva USAR_CURVA_CLIENTE."
        )

    df = pd.read_excel(RUTA_CURVA_CLIENTE)

    return df


def preparar_curva_cliente(df: pd.DataFrame) -> pd.DataFrame:
    """
    Prepara la curva de consumo del cliente.
    """

    df = df.copy()

    if "fecha_hora_utc" not in df.columns:
        raise ValueError(
            "Falta la columna 'fecha_hora_utc' en curva_cliente.xlsx. "
            "La curva debe tener una fila por hora."
        )

    if COLUMNA_CONSUMO_CLIENTE not in df.columns:
        raise ValueError(
            f"Falta la columna '{COLUMNA_CONSUMO_CLIENTE}' en curva_cliente.xlsx."
        )

    df["fecha_hora_utc"] = pd.to_datetime(
        df["fecha_hora_utc"],
        utc=True,
        errors="coerce",
    )

    df[COLUMNA_CONSUMO_CLIENTE] = pd.to_numeric(
        df[COLUMNA_CONSUMO_CLIENTE],
        errors="coerce",
    )

    df = df.dropna(
        subset=[
            "fecha_hora_utc",
            COLUMNA_CONSUMO_CLIENTE,
        ]
    ).copy()

    df = df[
        [
            "fecha_hora_utc",
            COLUMNA_CONSUMO_CLIENTE,
        ]
    ].copy()

    df = df.rename(
        columns={
            COLUMNA_CONSUMO_CLIENTE: "energia_cliente_MWh",
        }
    )

    return df




def normalizar_tarifa_para_perfil_ree(tarifa: str | None) -> str:
    """
    Convierte la tarifa elegida por el usuario al perfil disponible para construir
    la curva de carga.
    """

    if tarifa is None:
        return "3.0TD"

    tarifa = str(tarifa).upper().replace(" ", "").strip()

    if tarifa.startswith("2.0"):
        return "2.0TD"

    if tarifa.startswith("3.0TDVE"):
        return "3.0TDVE"

    if tarifa.startswith("3.0"):
        return "3.0TD"

    if tarifa.startswith("6."):
        return "6.XTD"

    return "3.0TD"


def preparar_horas_para_curva_cliente(df_horario: pd.DataFrame) -> pd.DataFrame:
    """
    Prepara las fechas de la previsión para poder generar una curva de carga
    exactamente sobre el periodo de oferta elegido.
    """

    df = df_horario.copy()

    if "fecha_hora_utc" in df.columns:
        fecha_utc = pd.to_datetime(df["fecha_hora_utc"], utc=True, errors="coerce")
        fecha_local = fecha_utc.dt.tz_convert("Europe/Madrid").dt.tz_localize(None)

    elif "fecha_hora_local" in df.columns:
        fecha_utc = pd.to_datetime(df["fecha_hora_local"], utc=True, errors="coerce")
        fecha_local = fecha_utc.dt.tz_convert("Europe/Madrid").dt.tz_localize(None)

    elif "fecha_hora" in df.columns:
        fecha_local = pd.to_datetime(df["fecha_hora"], errors="coerce")
        fecha_utc = fecha_local.dt.tz_localize("Europe/Madrid").dt.tz_convert("UTC")

    elif "fecha" in df.columns and "hora" in df.columns:
        fecha_base = pd.to_datetime(df["fecha"], errors="coerce")
        hora = pd.to_numeric(df["hora"], errors="coerce").fillna(0).astype(int)
        fecha_local = fecha_base + pd.to_timedelta(hora, unit="h")
        fecha_utc = fecha_local.dt.tz_localize("Europe/Madrid").dt.tz_convert("UTC")

    else:
        raise ValueError(
            "No se han encontrado columnas de fecha suficientes para crear la curva. "
            "Se necesita 'fecha_hora_utc', 'fecha_hora_local', 'fecha_hora' "
            "o bien 'fecha' + 'hora'."
        )

    if fecha_local.isna().any() or fecha_utc.isna().any():
        raise ValueError("Hay fechas no válidas en la previsión horaria.")

    df_horas = pd.DataFrame(
        {
            "fecha_hora_utc": fecha_utc,
            "fecha_hora_local": fecha_local,
        }
    )

    return df_horas


def crear_clave_local_sin_tz(df: pd.DataFrame) -> pd.Series:
    """
    Devuelve una fecha local sin zona horaria para cruzar previsión y curva.
    """

    if "fecha_hora_utc" in df.columns:
        return (
            pd.to_datetime(df["fecha_hora_utc"], utc=True, errors="coerce")
            .dt.tz_convert("Europe/Madrid")
            .dt.tz_localize(None)
        )

    if "fecha_hora_local" in df.columns:
        fechas = pd.to_datetime(df["fecha_hora_local"], utc=True, errors="coerce")
        return fechas.dt.tz_convert("Europe/Madrid").dt.tz_localize(None)

    if "fecha_hora" in df.columns:
        return pd.to_datetime(df["fecha_hora"], errors="coerce")

    if "fecha" in df.columns and "hora" in df.columns:
        fecha_base = pd.to_datetime(df["fecha"], errors="coerce")
        hora = pd.to_numeric(df["hora"], errors="coerce").fillna(0).astype(int)
        return fecha_base + pd.to_timedelta(hora, unit="h")

    raise ValueError("No se puede crear la clave horaria local.")



def crear_clave_utc_sin_tz(df: pd.DataFrame) -> pd.Series:
    """
    Devuelve una clave horaria UTC sin zona horaria. Se usa para hacer merges
    sin duplicar la hora repetida del cambio horario de octubre.
    """

    if "fecha_hora_utc" in df.columns:
        return (
            pd.to_datetime(df["fecha_hora_utc"], utc=True, errors="coerce")
            .dt.tz_localize(None)
        )

    if "fecha_hora_local" in df.columns:
        fechas = pd.to_datetime(df["fecha_hora_local"], utc=True, errors="coerce")
        return fechas.dt.tz_localize(None)

    if "fecha_hora" in df.columns:
        fecha_local = pd.to_datetime(df["fecha_hora"], errors="coerce")
        return (
            fecha_local.dt.tz_localize("Europe/Madrid", ambiguous="infer", nonexistent="shift_forward")
            .dt.tz_convert("UTC")
            .dt.tz_localize(None)
        )

    if "fecha" in df.columns and "hora" in df.columns:
        fecha_base = pd.to_datetime(df["fecha"], errors="coerce")
        hora = pd.to_numeric(df["hora"], errors="coerce").fillna(0).astype(int)
        fecha_local = fecha_base + pd.to_timedelta(hora, unit="h")
        return (
            fecha_local.dt.tz_localize("Europe/Madrid", ambiguous="infer", nonexistent="shift_forward")
            .dt.tz_convert("UTC")
            .dt.tz_localize(None)
        )

    raise ValueError("No se puede crear la clave horaria UTC.")


def factor_tipo_cliente(row: pd.Series, tipo_cliente: str | None) -> float:
    """
    Ajuste moderado sobre el perfil base REE para que la curva sea coherente
    con el tipo de cliente seleccionado.
    """

    tipo = str(tipo_cliente or "industrial").lower().strip()
    hora = int(row["hora"])
    dia_semana = int(row.get("dia_semana", 0))
    mes = int(row["mes"])
    es_fin_semana = dia_semana >= 5

    factor = 1.0

    if tipo in ["industrial", "industria"]:
        if 8 <= hora <= 18:
            factor *= 1.16
        elif 0 <= hora <= 5:
            factor *= 0.82
        if es_fin_semana:
            factor *= 0.72
        if mes == 8:
            factor *= 0.78

    elif tipo in ["oficinas", "oficina"]:
        if 8 <= hora <= 18 and not es_fin_semana:
            factor *= 1.28
        elif hora <= 6 or hora >= 22:
            factor *= 0.55
        if es_fin_semana:
            factor *= 0.45
        if mes == 8:
            factor *= 0.85

    elif tipo in ["comercio", "retail"]:
        if 10 <= hora <= 21:
            factor *= 1.18
        elif hora <= 7:
            factor *= 0.65
        if dia_semana == 5:
            factor *= 1.08
        if dia_semana == 6:
            factor *= 0.72

    elif tipo in ["residencial", "domestico", "doméstico"]:
        if 7 <= hora <= 9 or 18 <= hora <= 23:
            factor *= 1.20
        if 10 <= hora <= 16:
            factor *= 0.90
        if es_fin_semana:
            factor *= 1.04
        if mes in [1, 2, 7, 8, 12]:
            factor *= 1.08

    return factor


def ajustar_coeficientes_por_tipo_cliente(
    df_coeficientes: pd.DataFrame,
    tipo_cliente: str | None,
) -> pd.DataFrame:
    """
    Parte de un perfil real de REE y lo modula suavemente según el tipo de
    cliente. La curva sigue manteniendo como base los coeficientes oficiales.
    """

    df = df_coeficientes.copy()

    df["fecha_ficticia"] = pd.to_datetime(
        {
            "year": 2026,
            "month": df["mes"],
            "day": df["dia"],
        },
        errors="coerce",
    )

    df["dia_semana"] = df["fecha_ficticia"].dt.dayofweek.fillna(0).astype(int)
    df["factor_tipo_cliente"] = df.apply(
        lambda row: factor_tipo_cliente(row, tipo_cliente),
        axis=1,
    )

    df["coeficiente_perfil"] = (
        pd.to_numeric(df["coeficiente_perfil"], errors="coerce")
        * df["factor_tipo_cliente"]
    )

    df["coeficiente_perfil"] = df["coeficiente_perfil"].clip(lower=0)

    origen = df.get("origen_perfil", "Perfil REE")
    df["origen_perfil"] = (
        origen.astype(str)
        + f" + ajuste tipo cliente {tipo_cliente or 'industrial'}"
    )

    return df.drop(columns=["fecha_ficticia"], errors="ignore")


def generar_curva_cliente_desde_perfil_ree(
    df_horario: pd.DataFrame,
    energia_total_mwh: float,
    tarifa_cliente: str | None,
    tipo_cliente: str | None,
) -> pd.DataFrame:
    """
    Genera una curva de carga para las mismas fechas de la oferta.

    La base es el fichero de perfiles iniciales de REE. Para tarifas 6.XTD se
    usa el perfil industrial sintético ya definido en el proyecto, construido
    a partir del perfil 3.0TD de REE.
    """

    perfil_ree = normalizar_tarifa_para_perfil_ree(tarifa_cliente)

    df_horas = preparar_horas_para_curva_cliente(df_horario)
    df_perfil_ree = cargar_perfil_ree_2026()

    df_coeficientes = crear_tabla_coeficientes_perfil(
        df_perfil_ree=df_perfil_ree,
        perfil=perfil_ree,
    )

    df_coeficientes = ajustar_coeficientes_por_tipo_cliente(
        df_coeficientes=df_coeficientes,
        tipo_cliente=tipo_cliente,
    )

    df_curva = generar_curva_horaria(
        df_horas=df_horas,
        df_coeficientes=df_coeficientes,
        energia_total_mwh=energia_total_mwh,
    )

    resumen_mensual, resumen_horario, resumen_dia_semana = calcular_resumen_curva(
        df_curva
    )

    guardar_curva_cliente_seleccionada(
        df_curva=df_curva,
        resumen_mensual=resumen_mensual,
        resumen_horario=resumen_horario,
        resumen_dia_semana=resumen_dia_semana,
    )

    print("\nCurva de carga generada con base en perfiles REE.")
    print(f"Perfil REE utilizado: {perfil_ree}")
    print(f"Tarifa indicada: {tarifa_cliente}")
    print(f"Tipo de cliente: {tipo_cliente}")
    print(f"Fecha inicio curva: {df_curva['fecha_hora_local'].min()}")
    print(f"Fecha fin curva: {df_curva['fecha_hora_local'].max()}")
    print(f"Energía total curva: {df_curva['consumo_MWh'].sum():,.3f} MWh")

    df_curva["_fecha_hora_utc_key"] = pd.to_datetime(
        df_curva["fecha_hora_utc"],
        errors="coerce",
    )

    df_curva = df_curva.rename(
        columns={
            "consumo_MWh": "energia_cliente_MWh",
        }
    )

    return df_curva[
        [
            "_fecha_hora_utc_key",
            "energia_cliente_MWh",
            "coeficiente_perfil",
            "criterio_coeficiente_perfil",
            "origen_perfil",
        ]
    ].copy()


def incorporar_curva_cliente(
    df_horario: pd.DataFrame,
    configuracion_consumo: dict | None = None,
) -> pd.DataFrame:
    """
    Incorpora al dataframe horario el consumo del cliente.

    - Si el usuario elige perfil plano, reparte la energía total entre todas
      las horas de la oferta.
    - Si elige curva de carga, genera una curva sobre las fechas exactas de la
      oferta usando perfiles REE como base.
    """

    df = df_horario.copy()

    if configuracion_consumo is None:
        configuracion_consumo = {
            "modo_consumo": "plano" if not USAR_CURVA_CLIENTE else "curva",
            "energia_total_cliente_mwh": ENERGIA_TOTAL_CLIENTE_MWH,
            "tarifa_cliente": "3.0TD",
            "tipo_cliente": "industrial",
        }

    modo_consumo = str(configuracion_consumo.get("modo_consumo", "plano")).lower()
    energia_total_mwh = float(
        configuracion_consumo.get(
            "energia_total_cliente_mwh",
            ENERGIA_TOTAL_CLIENTE_MWH,
        )
    )

    if energia_total_mwh <= 0:
        raise ValueError("La energía total del cliente debe ser mayor que cero.")

    df["_fecha_hora_local_key"] = crear_clave_local_sin_tz(df)
    df["_fecha_hora_utc_key"] = crear_clave_utc_sin_tz(df)

    if df["_fecha_hora_local_key"].isna().any() or df["_fecha_hora_utc_key"].isna().any():
        raise ValueError("Hay horas no válidas en el periodo de oferta.")

    print("\nCONSUMO DEL CLIENTE")
    print("-" * 50)

    if modo_consumo == "curva":
        curva = generar_curva_cliente_desde_perfil_ree(
            df_horario=df,
            energia_total_mwh=energia_total_mwh,
            tarifa_cliente=configuracion_consumo.get("tarifa_cliente"),
            tipo_cliente=configuracion_consumo.get("tipo_cliente"),
        )

        df = df.drop(columns=["energia_cliente_MWh"], errors="ignore")
        df = df.merge(
            curva,
            on="_fecha_hora_utc_key",
            how="left",
        )

        df["modo_ponderacion"] = "Curva de carga con perfil REE ajustado"
        df["tarifa_cliente"] = configuracion_consumo.get("tarifa_cliente")
        df["tipo_cliente"] = configuracion_consumo.get("tipo_cliente")

    else:
        n_horas = len(df)
        if n_horas == 0:
            raise ValueError("El periodo de oferta no contiene horas.")

        df["energia_cliente_MWh"] = energia_total_mwh / n_horas
        df["coeficiente_perfil"] = 1.0
        df["criterio_coeficiente_perfil"] = "Perfil plano"
        df["origen_perfil"] = "Energía repartida uniformemente por hora"
        df["modo_ponderacion"] = "Perfil plano"
        df["tarifa_cliente"] = configuracion_consumo.get("tarifa_cliente")
        df["tipo_cliente"] = configuracion_consumo.get("tipo_cliente")

    horas_sin_consumo = df["energia_cliente_MWh"].isna().sum()

    if horas_sin_consumo > 0:
        raise ValueError(
            "No se ha podido incorporar correctamente el consumo del cliente. "
            f"Horas sin consumo: {horas_sin_consumo}."
        )

    energia_final = df["energia_cliente_MWh"].sum()

    print(f"Modo de consumo: {df['modo_ponderacion'].iloc[0]}")
    print(f"Fecha inicio: {df['_fecha_hora_local_key'].min()}")
    print(f"Fecha fin: {df['_fecha_hora_local_key'].max()}")
    print(f"Número de horas: {len(df)}")
    print(f"Energía total incorporada: {energia_final:,.3f} MWh")
    print(f"Consumo medio horario: {df['energia_cliente_MWh'].mean():.6f} MWh")
    print("-" * 50)

    return df

def media_ponderada(
    valores: pd.Series,
    pesos: pd.Series,
) -> float:
    """
    Calcula una media ponderada.
    """

    datos = pd.DataFrame(
        {
            "valor": valores,
            "peso": pesos,
        }
    ).dropna()

    if datos.empty:
        return float("nan")

    suma_pesos = datos["peso"].sum()

    if suma_pesos == 0:
        return float("nan")

    return (datos["valor"] * datos["peso"]).sum() / suma_pesos


def calcular_decision_cobertura(
    df: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Calcula la decisión de cobertura y la commodity recomendada mediante
    una ponderación temporal uniforme.

    Cada hora tiene el mismo peso. No se utiliza consumo, energía, tarifa,
    perfil ni curva de carga de ningún cliente.
    """

    df = df.copy()

    # Peso neutro de una unidad por hora. La denominación de la columna se
    # conserva internamente para mantener la compatibilidad con el resto de
    # funciones, pero no representa energía real de un cliente.
    df["energia_cliente_MWh"] = 1.0
    df["modo_ponderacion"] = (
        "Ponderación temporal uniforme, sin datos del cliente"
    )

    energia_total = df["energia_cliente_MWh"].sum()

    if energia_total <= 0:
        raise ValueError("El periodo de previsión no contiene horas válidas.")

    df["producto_decision"] = df["producto_omip"].fillna("Sin OMIP - modelo")
    df["criterio_decision_producto"] = df["criterio_producto_omip"].fillna(
        "Sin producto OMIP aplicable: se usa modelo para este tramo."
    )

    df["porcentaje_cobertura_sobre_tramo_omip"] = 0.0
    df["porcentaje_abierto_sobre_tramo_omip"] = 1.0
    df["precio_parte_cubierta_EUR_MWh"] = df["precio_omip_EUR_MWh"]
    df["precio_parte_abierta_EUR_MWh"] = (
        df["precio_omie_previsto"]
        + PRIMA_RIESGO_PARTE_ABIERTA_EUR_MWH
    )

    df["commodity_recomendada_hora_EUR_MWh"] = (
        df["precio_omie_previsto"]
        + PRIMA_RIESGO_PARTE_ABIERTA_EUR_MWH
    )

    registros_productos = []

    for (producto, criterio), grupo in df.groupby(
        [
            "producto_decision",
            "criterio_decision_producto",
        ],
        dropna=False,
    ):
        indices = grupo.index

        energia_producto = grupo["energia_cliente_MWh"].sum()

        precio_modelo_producto = media_ponderada(
            valores=grupo["precio_omie_previsto"],
            pesos=grupo["energia_cliente_MWh"],
        )

        tiene_omip = grupo["precio_omip_EUR_MWh"].notna().any()

        if tiene_omip:
            precio_omip_producto = media_ponderada(
                valores=grupo["precio_omip_EUR_MWh"],
                pesos=grupo["energia_cliente_MWh"],
            )

            diferencia_modelo_menos_omip = (
                precio_modelo_producto
                - precio_omip_producto
            )

            if precio_modelo_producto >= precio_omip_producto:
                porcentaje_cobertura = (
                    PORCENTAJE_COBERTURA_MODELO_MAYOR_OMIP
                )

                señal = "OMIP igual o inferior al modelo"

            else:
                porcentaje_cobertura = (
                    PORCENTAJE_COBERTURA_MODELO_MENOR_OMIP
                )

                señal = "OMIP superior al modelo"

            porcentaje_abierto = 1 - porcentaje_cobertura

            df.loc[
                indices,
                "porcentaje_cobertura_sobre_tramo_omip",
            ] = porcentaje_cobertura

            df.loc[
                indices,
                "porcentaje_abierto_sobre_tramo_omip",
            ] = porcentaje_abierto

            df.loc[
                indices,
                "commodity_recomendada_hora_EUR_MWh",
            ] = (
                porcentaje_cobertura
                * df.loc[indices, "precio_omip_EUR_MWh"]
                + porcentaje_abierto
                * df.loc[indices, "precio_parte_abierta_EUR_MWh"]
            )

        else:
            precio_omip_producto = None
            diferencia_modelo_menos_omip = None
            porcentaje_cobertura = 0.0
            porcentaje_abierto = 1.0
            señal = "Sin OMIP disponible"

            df.loc[
                indices,
                "commodity_recomendada_hora_EUR_MWh",
            ] = df.loc[indices, "precio_parte_abierta_EUR_MWh"]

        commodity_recomendada_producto = media_ponderada(
            valores=df.loc[indices, "commodity_recomendada_hora_EUR_MWh"],
            pesos=df.loc[indices, "energia_cliente_MWh"],
        )

        energia_cubierta_producto_MWh = (
            energia_producto
            * porcentaje_cobertura
        )

        energia_abierta_producto_MWh = (
            energia_producto
            * porcentaje_abierto
        )

        importe_modelo_producto_EUR = (
            energia_producto
            * precio_modelo_producto
        )

        if tiene_omip:
            importe_omip_producto_EUR = (
                energia_producto
                * precio_omip_producto
            )

            importe_cobertura_producto_EUR = (
                energia_cubierta_producto_MWh
                * precio_omip_producto
            )
        else:
            importe_omip_producto_EUR = None
            importe_cobertura_producto_EUR = 0.0

        importe_abierto_producto_EUR = (
            energia_abierta_producto_MWh
            * (
                precio_modelo_producto
                + PRIMA_RIESGO_PARTE_ABIERTA_EUR_MWH
            )
        )

        importe_recomendado_producto_EUR = (
            energia_producto
            * commodity_recomendada_producto
        )

        registros_productos.append(
            {
                "producto_omip": producto,
                "criterio_producto_omip": criterio,
                "fecha_inicio": grupo["fecha"].min().date(),
                "fecha_fin": grupo["fecha"].max().date(),
                "horas": len(grupo),

                "energia_cliente_MWh": energia_producto,
                "peso_energia": energia_producto / energia_total,

                "precio_modelo_EUR_MWh": precio_modelo_producto,
                "precio_omip_EUR_MWh": precio_omip_producto,
                "diferencia_modelo_menos_omip_EUR_MWh": diferencia_modelo_menos_omip,

                "porcentaje_cobertura": porcentaje_cobertura,
                "porcentaje_abierto": porcentaje_abierto,

                "energia_cubierta_MWh": energia_cubierta_producto_MWh,
                "energia_abierta_MWh": energia_abierta_producto_MWh,

                "commodity_recomendada_producto_EUR_MWh": commodity_recomendada_producto,

                "importe_modelo_producto_EUR": importe_modelo_producto_EUR,
                "importe_omip_producto_EUR": importe_omip_producto_EUR,
                "importe_cobertura_producto_EUR": importe_cobertura_producto_EUR,
                "importe_abierto_producto_EUR": importe_abierto_producto_EUR,
                "importe_recomendado_producto_EUR": importe_recomendado_producto_EUR,

                "señal": señal,
            }
        )

    resumen_productos = pd.DataFrame(registros_productos)

    commodity_modelo_total = media_ponderada(
        valores=df["precio_omie_previsto"],
        pesos=df["energia_cliente_MWh"],
    )

    commodity_recomendada = media_ponderada(
        valores=df["commodity_recomendada_hora_EUR_MWh"],
        pesos=df["energia_cliente_MWh"],
    )

    df_con_omip = df.dropna(subset=["precio_omip_EUR_MWh"]).copy()

    energia_con_omip = df_con_omip["energia_cliente_MWh"].sum()
    energia_sin_omip = energia_total - energia_con_omip

    if energia_con_omip > 0:
        omip_equivalente = media_ponderada(
            valores=df_con_omip["precio_omip_EUR_MWh"],
            pesos=df_con_omip["energia_cliente_MWh"],
        )

        modelo_horas_omip = media_ponderada(
            valores=df_con_omip["precio_omie_previsto"],
            pesos=df_con_omip["energia_cliente_MWh"],
        )

        diferencia_global = modelo_horas_omip - omip_equivalente

    else:
        omip_equivalente = None
        modelo_horas_omip = None
        diferencia_global = None

    porcentaje_cobertura_total = (
        df["energia_cliente_MWh"]
        * df["porcentaje_cobertura_sobre_tramo_omip"]
    ).sum() / energia_total

    porcentaje_abierto_total = 1 - porcentaje_cobertura_total

    energia_cubierta_total_MWh = (
        df["energia_cliente_MWh"]
        * df["porcentaje_cobertura_sobre_tramo_omip"]
    ).sum()

    energia_abierta_total_MWh = (
        energia_total
        - energia_cubierta_total_MWh
    )

    importe_modelo_total_EUR = (
        energia_total
        * commodity_modelo_total
    )

    importe_recomendado_total_EUR = (
        energia_total
        * commodity_recomendada
    )

    if omip_equivalente is not None:
        importe_omip_equivalente_EUR = (
            energia_con_omip
            * omip_equivalente
        )
    else:
        importe_omip_equivalente_EUR = None

    resumen = pd.DataFrame(
        [
            {
                "fecha_inicio_oferta": df["fecha"].min().date(),
                "fecha_fin_oferta": df["fecha"].max().date(),
                "modo_ponderacion": df["modo_ponderacion"].iloc[0],
                "horas_totales_oferta": len(df),
                "energia_total_cliente_MWh": energia_total,
                "energia_con_omip_MWh": energia_con_omip,
                "energia_sin_omip_MWh": energia_sin_omip,
                "porcentaje_energia_con_omip": energia_con_omip / energia_total,
                "porcentaje_energia_sin_omip": energia_sin_omip / energia_total,
                "commodity_modelo_total_EUR_MWh": commodity_modelo_total,
                "commodity_modelo_horas_omip_EUR_MWh": modelo_horas_omip,
                "omip_equivalente_EUR_MWh": omip_equivalente,
                "diferencia_modelo_menos_omip_EUR_MWh": diferencia_global,
                "porcentaje_cobertura_total": porcentaje_cobertura_total,
                "porcentaje_abierto_total": porcentaje_abierto_total,

                "energia_cubierta_total_MWh": energia_cubierta_total_MWh,
                "energia_abierta_total_MWh": energia_abierta_total_MWh,

                "prima_riesgo_parte_abierta_EUR_MWh": PRIMA_RIESGO_PARTE_ABIERTA_EUR_MWH,

                "commodity_recomendada_EUR_MWh": commodity_recomendada,

                "importe_modelo_total_EUR": importe_modelo_total_EUR,
                "importe_omip_equivalente_EUR": importe_omip_equivalente_EUR,
                "importe_recomendado_total_EUR": importe_recomendado_total_EUR,
            }
        ]
    )

    df.attrs["resumen_productos_decision"] = resumen_productos

    return df, resumen


def calcular_resumen_productos_omip(df: pd.DataFrame) -> pd.DataFrame:
    """
    Devuelve el resumen de decisión por producto OMIP.
    """

    resumen_productos = df.attrs.get("resumen_productos_decision")

    if resumen_productos is None:
        return pd.DataFrame()

    return resumen_productos

def crear_plan_cobertura_operativo(
    resumen_productos: pd.DataFrame,
) -> pd.DataFrame:
    """
    Crea una tabla operativa de cobertura por producto OMIP.

    Esta tabla indica, para cada mes, trimestre o año usado:
    - qué precio OMIP se toma,
    - qué precio estima el modelo,
    - qué porcentaje se recomienda cubrir,
    - cuánta energía se cubre,
    - cuánta energía queda abierta,
    - y qué commodity resulta para ese tramo.
    """

    if resumen_productos.empty:
        return pd.DataFrame()

    df = resumen_productos.copy()

    columnas_necesarias = [
        "producto_omip",
        "fecha_inicio",
        "fecha_fin",
        "energia_cliente_MWh",
        "precio_modelo_EUR_MWh",
        "precio_omip_EUR_MWh",
        "diferencia_modelo_menos_omip_EUR_MWh",
        "porcentaje_cobertura",
        "porcentaje_abierto",
        "energia_cubierta_MWh",
        "energia_abierta_MWh",
        "commodity_recomendada_producto_EUR_MWh",
        "importe_recomendado_producto_EUR",
        "señal",
        "criterio_producto_omip",
    ]

    columnas_existentes = [
        columna
        for columna in columnas_necesarias
        if columna in df.columns
    ]

    df = df[columnas_existentes].copy()

    def clasificar_decision(row):
        if pd.isna(row.get("precio_omip_EUR_MWh")):
            return "Sin OMIP: se usa modelo"

        if row["porcentaje_cobertura"] >= 0.999:
            return "Cubrir 100 %"

        if row["porcentaje_cobertura"] > 0:
            return "Cobertura parcial"

        return "No cubrir"

    df["decision_operativa"] = df.apply(
        clasificar_decision,
        axis=1,
    )

    df["porcentaje_cobertura"] = df["porcentaje_cobertura"] * 100
    df["porcentaje_abierto"] = df["porcentaje_abierto"] * 100

    df = df.rename(
        columns={
            "producto_omip": "producto_cobertura",
            "energia_cliente_MWh": "horas_producto",
            "precio_modelo_EUR_MWh": "precio_modelo_producto_EUR_MWh",
            "precio_omip_EUR_MWh": "precio_omip_producto_EUR_MWh",
            "porcentaje_cobertura": "cobertura_recomendada_pct",
            "porcentaje_abierto": "exposicion_abierta_pct",
            "energia_cubierta_MWh": "horas_equivalentes_cubiertas",
            "energia_abierta_MWh": "horas_equivalentes_abiertas",
            "commodity_recomendada_producto_EUR_MWh": "commodity_producto_recomendada_EUR_MWh",
            "importe_recomendado_producto_EUR": "importe_producto_recomendado_EUR",
        }
    )

    columnas_ordenadas = [
        "producto_cobertura",
        "fecha_inicio",
        "fecha_fin",
        "decision_operativa",
        "horas_producto",
        "horas_equivalentes_cubiertas",
        "horas_equivalentes_abiertas",
        "cobertura_recomendada_pct",
        "exposicion_abierta_pct",
        "precio_modelo_producto_EUR_MWh",
        "precio_omip_producto_EUR_MWh",
        "diferencia_modelo_menos_omip_EUR_MWh",
        "commodity_producto_recomendada_EUR_MWh",
        "importe_producto_recomendado_EUR",
        "señal",
        "criterio_producto_omip",
    ]

    columnas_ordenadas = [
        columna
        for columna in columnas_ordenadas
        if columna in df.columns
    ]

    df = df[columnas_ordenadas].copy()

    return df




def crear_tramos_por_estrategia(
    fecha_inicio_oferta: pd.Timestamp,
    fecha_fin_oferta: pd.Timestamp,
    estrategia: str,
) -> list[dict]:
    """
    Crea los tramos sobre los que se simula la cobertura mensual,
    trimestral o anual.
    """

    fecha_inicio_oferta = pd.Timestamp(fecha_inicio_oferta).normalize()
    fecha_fin_oferta = pd.Timestamp(fecha_fin_oferta).normalize()
    estrategia = estrategia.lower().strip()

    tramos = []

    if estrategia == "mensual":
        periodos = pd.date_range(
            start=fecha_inicio_oferta.replace(day=1),
            end=fecha_fin_oferta.replace(day=1),
            freq="MS",
        )

        for periodo in periodos:
            inicio_natural = periodo
            fin_natural = periodo + pd.offsets.MonthEnd(0)
            tramos.append(
                {
                    "estrategia": estrategia,
                    "tipo_producto_buscado": "mensual",
                    "fecha_inicio": max(inicio_natural, fecha_inicio_oferta),
                    "fecha_fin": min(fin_natural, fecha_fin_oferta),
                    "periodo_natural_inicio": inicio_natural,
                    "periodo_natural_fin": fin_natural,
                }
            )

    elif estrategia == "trimestral":
        bloques = crear_bloques_trimestrales_oferta(
            fecha_inicio_oferta=fecha_inicio_oferta,
            fecha_fin_oferta=fecha_fin_oferta,
        )

        for bloque in bloques:
            tramos.append(
                {
                    "estrategia": estrategia,
                    "tipo_producto_buscado": "trimestral",
                    "fecha_inicio": bloque["inicio_bloque"],
                    "fecha_fin": bloque["fin_bloque"],
                    "periodo_natural_inicio": bloque["inicio_trimestre"],
                    "periodo_natural_fin": bloque["fin_trimestre"],
                }
            )

    elif estrategia == "anual":
        años = range(fecha_inicio_oferta.year, fecha_fin_oferta.year + 1)

        for año in años:
            inicio_natural = pd.Timestamp(year=año, month=1, day=1)
            fin_natural = pd.Timestamp(year=año, month=12, day=31)
            inicio = max(inicio_natural, fecha_inicio_oferta)
            fin = min(fin_natural, fecha_fin_oferta)

            if inicio <= fin:
                tramos.append(
                    {
                        "estrategia": estrategia,
                        "tipo_producto_buscado": "anual",
                        "fecha_inicio": inicio,
                        "fecha_fin": fin,
                        "periodo_natural_inicio": inicio_natural,
                        "periodo_natural_fin": fin_natural,
                    }
                )

    else:
        raise ValueError(f"Estrategia no reconocida: {estrategia}")

    return tramos


def calcular_detalle_estrategia_cobertura(
    df_horario: pd.DataFrame,
    df_omip: pd.DataFrame,
    estrategia: str,
) -> pd.DataFrame:
    """
    Simula una estrategia de cobertura concreta:
    - mensual: intenta cubrir cada mes con productos mensuales;
    - trimestral: intenta cubrir cada trimestre con productos trimestrales;
    - anual: intenta cubrir cada año/calendario con productos anuales.
    """

    df = df_horario.copy()
    df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce")

    fecha_inicio_oferta = df["fecha"].min().normalize()
    fecha_fin_oferta = df["fecha"].max().normalize()

    tramos = crear_tramos_por_estrategia(
        fecha_inicio_oferta=fecha_inicio_oferta,
        fecha_fin_oferta=fecha_fin_oferta,
        estrategia=estrategia,
    )

    registros = []

    for tramo in tramos:
        inicio = tramo["fecha_inicio"]
        fin = tramo["fecha_fin"]

        mascara = (
            (df["fecha"].dt.normalize() >= inicio)
            & (df["fecha"].dt.normalize() <= fin)
        )

        grupo = df.loc[mascara].copy()

        if grupo.empty:
            continue

        energia_tramo = grupo["energia_cliente_MWh"].sum()
        precio_modelo = media_ponderada(
            valores=grupo["precio_omie_previsto"],
            pesos=grupo["energia_cliente_MWh"],
        )

        producto = buscar_producto_que_cubre(
            df_omip=df_omip,
            tipo_producto=tramo["tipo_producto_buscado"],
            fecha_inicio=tramo["periodo_natural_inicio"],
            fecha_fin=tramo["periodo_natural_fin"],
        )

        if producto is None:
            producto_nombre = "Sin producto OMIP"
            precio_omip = np.nan
            diferencia = np.nan
            porcentaje_cobertura = 0.0
            señal = "Sin OMIP disponible: se queda abierto con referencia modelo"
        else:
            producto_nombre = producto["producto"]
            precio_omip = float(producto["precio_omip_EUR_MWh"])
            diferencia = precio_modelo - precio_omip

            if precio_modelo >= precio_omip:
                porcentaje_cobertura = PORCENTAJE_COBERTURA_MODELO_MAYOR_OMIP
                señal = "Cubrir: OMIP igual o inferior al modelo"
            else:
                porcentaje_cobertura = PORCENTAJE_COBERTURA_MODELO_MENOR_OMIP
                señal = "Cobertura parcial: OMIP superior al modelo"

        porcentaje_abierto = 1 - porcentaje_cobertura
        precio_abierto = precio_modelo + PRIMA_RIESGO_PARTE_ABIERTA_EUR_MWH

        if pd.isna(precio_omip):
            commodity_tramo = precio_abierto
        else:
            commodity_tramo = (
                porcentaje_cobertura * precio_omip
                + porcentaje_abierto * precio_abierto
            )

        energia_cubierta = energia_tramo * porcentaje_cobertura
        energia_abierta = energia_tramo * porcentaje_abierto

        registros.append(
            {
                "estrategia": estrategia,
                "tipo_producto_buscado": tramo["tipo_producto_buscado"],
                "producto_omip": producto_nombre,
                "fecha_inicio": inicio.date(),
                "fecha_fin": fin.date(),
                "periodo_natural_inicio": tramo["periodo_natural_inicio"].date(),
                "periodo_natural_fin": tramo["periodo_natural_fin"].date(),
                "energia_cliente_MWh": energia_tramo,
                "precio_modelo_EUR_MWh": precio_modelo,
                "precio_omip_EUR_MWh": precio_omip,
                "diferencia_modelo_menos_omip_EUR_MWh": diferencia,
                "porcentaje_cobertura": porcentaje_cobertura,
                "porcentaje_abierto": porcentaje_abierto,
                "energia_cubierta_MWh": energia_cubierta,
                "energia_abierta_MWh": energia_abierta,
                "commodity_tramo_recomendada_EUR_MWh": commodity_tramo,
                "importe_tramo_recomendado_EUR": energia_tramo * commodity_tramo,
                "señal": señal,
            }
        )

    return pd.DataFrame(registros)


def calcular_comparativa_estrategias_cobertura(
    df_horario: pd.DataFrame,
    df_omip: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Compara si interesa cubrir por meses, por trimestres o por producto anual.
    """

    detalles = []
    resumenes = []

    for estrategia in ["mensual", "trimestral", "anual"]:
        detalle = calcular_detalle_estrategia_cobertura(
            df_horario=df_horario,
            df_omip=df_omip,
            estrategia=estrategia,
        )

        if detalle.empty:
            continue

        detalles.append(detalle)

        energia_total = detalle["energia_cliente_MWh"].sum()
        importe_recomendado = detalle["importe_tramo_recomendado_EUR"].sum()
        commodity_recomendada = importe_recomendado / energia_total

        commodity_modelo = media_ponderada(
            valores=detalle["precio_modelo_EUR_MWh"],
            pesos=detalle["energia_cliente_MWh"],
        )

        detalle_con_omip = detalle.dropna(subset=["precio_omip_EUR_MWh"]).copy()

        if detalle_con_omip.empty:
            omip_equivalente = np.nan
        else:
            omip_equivalente = media_ponderada(
                valores=detalle_con_omip["precio_omip_EUR_MWh"],
                pesos=detalle_con_omip["energia_cliente_MWh"],
            )

        energia_cubierta = detalle["energia_cubierta_MWh"].sum()
        energia_abierta = detalle["energia_abierta_MWh"].sum()
        tramos_sin_omip = int(detalle["precio_omip_EUR_MWh"].isna().sum())

        resumenes.append(
            {
                "estrategia_cobertura": estrategia,
                "energia_total_MWh": energia_total,
                "commodity_modelo_EUR_MWh": commodity_modelo,
                "omip_equivalente_EUR_MWh": omip_equivalente,
                "commodity_recomendada_EUR_MWh": commodity_recomendada,
                "importe_recomendado_EUR": importe_recomendado,
                "energia_cubierta_MWh": energia_cubierta,
                "energia_abierta_MWh": energia_abierta,
                "cobertura_total_pct": energia_cubierta / energia_total * 100,
                "exposicion_abierta_pct": energia_abierta / energia_total * 100,
                "numero_tramos": len(detalle),
                "tramos_sin_omip": tramos_sin_omip,
            }
        )

    resumen = pd.DataFrame(resumenes)
    detalle_total = pd.concat(detalles, ignore_index=True) if detalles else pd.DataFrame()

    if not resumen.empty:
        mejor_idx = resumen["commodity_recomendada_EUR_MWh"].idxmin()
        resumen["recomendacion_estrategia"] = ""
        resumen.loc[mejor_idx, "recomendacion_estrategia"] = (
            "Estrategia con menor commodity recomendada"
        )

    return resumen, detalle_total


def mostrar_comparativa_estrategias(comparativa: pd.DataFrame) -> None:
    """
    Muestra en terminal la comparación mensual vs trimestral vs anual.
    """

    print("\nCOMPARATIVA DE ESTRATEGIAS DE COBERTURA")
    print("-" * 50)

    if comparativa.empty:
        print("No se ha podido construir la comparativa de estrategias.")
        return

    columnas = [
        "estrategia_cobertura",
        "commodity_modelo_EUR_MWh",
        "omip_equivalente_EUR_MWh",
        "commodity_recomendada_EUR_MWh",
        "cobertura_total_pct",
        "exposicion_abierta_pct",
        "tramos_sin_omip",
        "recomendacion_estrategia",
    ]

    columnas = [col for col in columnas if col in comparativa.columns]
    print(comparativa[columnas].round(3).to_string(index=False))


def guardar_decision_cobertura(
    df_horario: pd.DataFrame,
    resumen_decision: pd.DataFrame,
    resumen_productos: pd.DataFrame,
    comparativa_estrategias: pd.DataFrame | None = None,
    detalle_estrategias: pd.DataFrame | None = None,
) -> None:
    """
    Guarda la decisión de cobertura en Excel.
    """

    df_excel = df_horario.copy()

    if "fecha_hora_utc" in df_excel.columns:
        df_excel["fecha_hora_utc"] = pd.to_datetime(
            df_excel["fecha_hora_utc"],
            utc=True,
            errors="coerce",
        ).dt.tz_localize(None)

    if "fecha_hora_local" in df_excel.columns:
        try:
            df_excel["fecha_hora_local"] = (
                pd.to_datetime(df_excel["fecha_hora_local"], utc=True, errors="coerce")
                .dt.tz_convert("Europe/Madrid")
                .dt.tz_localize(None)
            )
        except Exception:
            df_excel["fecha_hora_local"] = pd.to_datetime(
                df_excel["fecha_hora_local"],
                errors="coerce",
            )

    if "fecha" in df_excel.columns:
        df_excel["fecha"] = pd.to_datetime(
            df_excel["fecha"],
            errors="coerce",
        ).dt.tz_localize(None)

    for columna_auxiliar in ["_fecha_dia", "_fecha_hora_local_key", "_fecha_hora_utc_key"]:
        if columna_auxiliar in df_excel.columns:
            df_excel = df_excel.drop(columns=[columna_auxiliar])


    # Preparar tablas de salida sin referencias a energía o consumo del cliente.
    resumen_excel = resumen_decision.copy().rename(
        columns={
            "energia_total_cliente_MWh": "peso_horario_total",
            "energia_con_omip_MWh": "horas_con_omip",
            "energia_sin_omip_MWh": "horas_sin_omip",
            "porcentaje_energia_con_omip": "porcentaje_horas_con_omip",
            "porcentaje_energia_sin_omip": "porcentaje_horas_sin_omip",
            "energia_cubierta_total_MWh": "horas_equivalentes_cubiertas",
            "energia_abierta_total_MWh": "horas_equivalentes_abiertas",
        }
    )
    resumen_excel = resumen_excel.drop(
        columns=[
            "importe_modelo_total_EUR",
            "importe_omip_equivalente_EUR",
            "importe_recomendado_total_EUR",
        ],
        errors="ignore",
    )

    productos_excel = resumen_productos.copy().rename(
        columns={
            "energia_cliente_MWh": "horas_producto",
            "peso_energia": "peso_temporal",
            "energia_cubierta_MWh": "horas_equivalentes_cubiertas",
            "energia_abierta_MWh": "horas_equivalentes_abiertas",
        }
    )
    productos_excel = productos_excel.drop(
        columns=[
            "importe_modelo_producto_EUR",
            "importe_omip_producto_EUR",
            "importe_cobertura_producto_EUR",
            "importe_abierto_producto_EUR",
            "importe_recomendado_producto_EUR",
        ],
        errors="ignore",
    )

    comparativa_excel = (
        comparativa_estrategias.copy()
        if comparativa_estrategias is not None
        else None
    )
    if comparativa_excel is not None:
        comparativa_excel = comparativa_excel.rename(
            columns={
                "energia_total_MWh": "peso_horario_total",
                "energia_cubierta_MWh": "horas_equivalentes_cubiertas",
                "energia_abierta_MWh": "horas_equivalentes_abiertas",
            }
        ).drop(columns=["importe_recomendado_EUR"], errors="ignore")

    detalle_estrategias_excel = (
        detalle_estrategias.copy()
        if detalle_estrategias is not None
        else None
    )
    if detalle_estrategias_excel is not None:
        detalle_estrategias_excel = detalle_estrategias_excel.rename(
            columns={
                "energia_cliente_MWh": "horas_tramo",
                "energia_cubierta_MWh": "horas_equivalentes_cubiertas",
                "energia_abierta_MWh": "horas_equivalentes_abiertas",
            }
        ).drop(columns=["importe_tramo_recomendado_EUR"], errors="ignore")

    df_excel = df_excel.rename(
        columns={"energia_cliente_MWh": "peso_horario"}
    )

    try:
        with pd.ExcelWriter(RUTA_DECISION_COBERTURA, engine="openpyxl") as writer:
            resumen_excel.to_excel(
                writer,
                sheet_name="resumen_decision",
                index=False,
            )

            plan_cobertura = crear_plan_cobertura_operativo(
                resumen_productos=resumen_productos,
            )

            plan_cobertura.to_excel(
                writer,
                sheet_name="plan_cobertura",
                index=False,
            )

            if comparativa_excel is not None:
                comparativa_excel.to_excel(
                    writer,
                    sheet_name="comparativa_estrategias",
                    index=False,
                )

            if detalle_estrategias_excel is not None:
                detalle_estrategias_excel.to_excel(
                    writer,
                    sheet_name="detalle_estrategias",
                    index=False,
                )

            productos_excel.to_excel(
                writer,
                sheet_name="productos_omip_usados",
                index=False,
            )

            df_excel.to_excel(
                writer,
                sheet_name="detalle_horario",
                index=False,
            )

    except PermissionError:
        raise PermissionError(
            f"No se puede guardar {RUTA_DECISION_COBERTURA}. "
            f"Probablemente está abierto en Excel. Ciérralo y vuelve a ejecutar."
        )

    print(f"Decisión de cobertura guardada en: {RUTA_DECISION_COBERTURA}")

def mostrar_resumen_decision(resumen_decision: pd.DataFrame) -> None:
    """
    Muestra el resumen de la comparación y la cobertura sin datos de cliente.
    """

    print("\nRESUMEN DECISIÓN DE COBERTURA")

    columnas = [
        "modo_ponderacion",
        "horas_totales_oferta",
        "commodity_modelo_total_EUR_MWh",
        "omip_equivalente_EUR_MWh",
        "diferencia_modelo_menos_omip_EUR_MWh",
        "porcentaje_cobertura_total",
        "porcentaje_abierto_total",
        "commodity_recomendada_EUR_MWh",
        "porcentaje_energia_con_omip",
        "porcentaje_energia_sin_omip",
    ]

    columnas_existentes = [
        columna
        for columna in columnas
        if columna in resumen_decision.columns
    ]

    df = resumen_decision[columnas_existentes].copy()

    df = df.rename(
        columns={
            "porcentaje_energia_con_omip": "porcentaje_horas_con_omip",
            "porcentaje_energia_sin_omip": "porcentaje_horas_sin_omip",
        }
    )

    columnas_numericas = [
        "horas_totales_oferta",
        "commodity_modelo_total_EUR_MWh",
        "omip_equivalente_EUR_MWh",
        "diferencia_modelo_menos_omip_EUR_MWh",
        "porcentaje_cobertura_total",
        "porcentaje_abierto_total",
        "commodity_recomendada_EUR_MWh",
        "porcentaje_horas_con_omip",
        "porcentaje_horas_sin_omip",
    ]

    for columna in columnas_numericas:
        if columna in df.columns:
            df[columna] = pd.to_numeric(df[columna], errors="coerce").round(3)

    print(df.to_string(index=False))



def generar_decision_cobertura() -> None:
    """
    Genera la decisión de commodity y cobertura.
    """

    print("\nDECISIÓN DE COMMODITY Y COBERTURA")
    print("-" * 50)

    df_prevision = cargar_prevision_modelo()
    df_prevision = preparar_prevision(df_prevision)

    df_omip = cargar_omip_manual()
    df_omip = preparar_omip(df_omip)

    df_horario = asignar_productos_omip_a_prevision(
        df_prevision=df_prevision,
        df_omip=df_omip,
    )

    # La decisión se calcula directamente sobre la previsión y los productos
    # OMIP. No se incorpora información de consumo ni curvas de cliente.
    df_horario, resumen_decision = calcular_decision_cobertura(df_horario)

    resumen_productos = calcular_resumen_productos_omip(df_horario)

    print("\nPLAN OPERATIVO DE COBERTURA")

    plan_cobertura = crear_plan_cobertura_operativo(
        resumen_productos=resumen_productos,
    )

    if plan_cobertura.empty:
        print("No se ha usado ningún producto OMIP.")
    else:
        columnas_terminal = [
            "producto_cobertura",
            "fecha_inicio",
            "fecha_fin",
            "decision_operativa",
            "horas_producto",
            "horas_equivalentes_cubiertas",
            "horas_equivalentes_abiertas",
            "cobertura_recomendada_pct",
            "precio_modelo_producto_EUR_MWh",
            "precio_omip_producto_EUR_MWh",
            "commodity_producto_recomendada_EUR_MWh",
        ]

        columnas_terminal = [
            columna
            for columna in columnas_terminal
            if columna in plan_cobertura.columns
        ]

        print(
            plan_cobertura[columnas_terminal]
            .round(3)
            .to_string(index=False)
        )

    comparativa_estrategias, detalle_estrategias = calcular_comparativa_estrategias_cobertura(
        df_horario=df_horario,
        df_omip=df_omip,
    )

    mostrar_comparativa_estrategias(comparativa_estrategias)

    guardar_decision_cobertura(
        df_horario=df_horario,
        resumen_decision=resumen_decision,
        resumen_productos=resumen_productos,
        comparativa_estrategias=comparativa_estrategias,
        detalle_estrategias=detalle_estrategias,
    )

    exportar_resultados_calculadora(
        resumen_decision=resumen_decision,
    )

    mostrar_resumen_decision(resumen_decision)

    print("Decisión de commodity y cobertura completada correctamente.")
    print("-" * 50)