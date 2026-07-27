import joblib
import numpy as np
import pandas as pd
import holidays

from config import (
    PROCESSED_DATA_DIR,
    OUTPUT_DIR,
    MODEL_DIR,
    FECHA_INICIO_PREVISION,
    FECHA_FIN_PREVISION,
    MAX_DIAS_PREVISION,
    METODO_ESCENARIO_FUTURO,
)

from datos_base import crear_indice_horario


COLUMNAS_EXOGENAS_HISTORICAS = [
    "demanda",
    "generacion_solar",
    "generacion_eolica",
    "generacion_hidraulica",
    "generacion_nuclear",
    "generacion_ciclo_combinado",
    "generacion_carbon",
    "precio_gas_EUR_MWh",
    #"precio_co2_EUR_tCO2",
]


def cargar_modelo_entrenado() -> dict:
    """
    Carga el modelo entrenado y las variables que necesita.
    """

    ruta_modelo = MODEL_DIR / "modelo_precio_omie.joblib"

    if not ruta_modelo.exists():
        raise FileNotFoundError(f"No existe el modelo entrenado: {ruta_modelo}")

    paquete_modelo = joblib.load(ruta_modelo)

    return paquete_modelo


def cargar_base_entrenamiento() -> pd.DataFrame:
    """
    Carga la base histórica usada para entrenar el modelo.
    """

    ruta_base = PROCESSED_DATA_DIR / "base_entrenamiento.xlsx"

    if not ruta_base.exists():
        raise FileNotFoundError(f"No existe la base de entrenamiento: {ruta_base}")

    df = pd.read_excel(ruta_base)

    return df


def validar_fechas_prevision(
    fecha_inicio_prevision: str | None = None,
    fecha_fin_prevision: str | None = None,
) -> tuple[pd.Timestamp, pd.Timestamp]:
    """
    Comprueba que las fechas de previsión sean correctas.

    Si no se pasan fechas, usa las fechas definidas en config.py.
    """

    if fecha_inicio_prevision is None:
        fecha_inicio_prevision = FECHA_INICIO_PREVISION

    if fecha_fin_prevision is None:
        fecha_fin_prevision = FECHA_FIN_PREVISION

    fecha_inicio = pd.Timestamp(fecha_inicio_prevision)
    fecha_fin = pd.Timestamp(fecha_fin_prevision)

    if fecha_fin < fecha_inicio:
        raise ValueError("La fecha fin de previsión no puede ser anterior a la fecha inicio.")

    dias_prevision = (fecha_fin - fecha_inicio).days + 1

    if dias_prevision > MAX_DIAS_PREVISION:
        raise ValueError(
            f"La previsión solicitada tiene {dias_prevision} días. "
            f"El máximo permitido es {MAX_DIAS_PREVISION} días."
        )

    return fecha_inicio, fecha_fin


def crear_base_futura_calendario(
    fecha_inicio: pd.Timestamp,
    fecha_fin: pd.Timestamp,
) -> pd.DataFrame:
    """
    Crea la base horaria futura con variables calendario.
    """

    indice_horario = crear_indice_horario(
        fecha_inicio.strftime("%Y-%m-%d"),
        fecha_fin.strftime("%Y-%m-%d"),
    )

    df = pd.DataFrame()

    df["fecha_hora_local"] = indice_horario.astype(str)
    df["fecha_hora_utc"] = indice_horario.tz_convert("UTC").astype(str)

    df["fecha"] = indice_horario.date
    df["fecha"] = pd.to_datetime(df["fecha"])

    df["año"] = indice_horario.year
    df["mes"] = indice_horario.month
    df["dia_mes"] = indice_horario.day
    df["hora"] = indice_horario.hour

    df["dia_semana"] = indice_horario.dayofweek
    df["es_fin_semana"] = df["dia_semana"].isin([5, 6]).astype(int)

    festivos_españa = holidays.Spain(years=df["año"].unique())

    df["es_festivo"] = df["fecha"].dt.date.apply(
        lambda fecha: 1 if fecha in festivos_españa else 0
    )

    return df


def crear_perfiles_historicos(df_historico: pd.DataFrame) -> dict:
    """
    Crea perfiles históricos para estimar variables futuras.

    Se usa la mediana histórica por mes, día de la semana y hora.
    """

    df = df_historico.copy()

    perfiles = {}

    perfiles["mes_dia_hora"] = (
        df.groupby(["mes", "dia_semana", "hora"], as_index=False)[
            COLUMNAS_EXOGENAS_HISTORICAS
        ]
        .median()
    )

    perfiles["mes_hora"] = (
        df.groupby(["mes", "hora"], as_index=False)[
            COLUMNAS_EXOGENAS_HISTORICAS
        ]
        .median()
    )

    perfiles["mes"] = (
        df.groupby(["mes"], as_index=False)[
            COLUMNAS_EXOGENAS_HISTORICAS
        ]
        .median()
    )

    perfiles["global"] = df[COLUMNAS_EXOGENAS_HISTORICAS].median()

    return perfiles


def asignar_escenario_base_futuro(
    df_futuro: pd.DataFrame,
    perfiles: dict,
) -> pd.DataFrame:
    """
    Asigna valores futuros de demanda y generación usando perfiles históricos.
    """

    df = df_futuro.copy()

    perfil_principal = perfiles["mes_dia_hora"]
    perfil_secundario = perfiles["mes_hora"]
    perfil_mensual = perfiles["mes"]
    perfil_global = perfiles["global"]

    df = df.merge(
        perfil_principal,
        on=["mes", "dia_semana", "hora"],
        how="left",
    )

    columnas_a_rellenar = COLUMNAS_EXOGENAS_HISTORICAS

    df = df.merge(
        perfil_secundario,
        on=["mes", "hora"],
        how="left",
        suffixes=("", "_mes_hora"),
    )

    for columna in columnas_a_rellenar:
        df[columna] = df[columna].fillna(df[f"{columna}_mes_hora"])

    columnas_auxiliares = [
        f"{columna}_mes_hora" for columna in columnas_a_rellenar
    ]

    df = df.drop(columns=columnas_auxiliares, errors="ignore")

    df = df.merge(
        perfil_mensual,
        on=["mes"],
        how="left",
        suffixes=("", "_mes"),
    )

    for columna in columnas_a_rellenar:
        df[columna] = df[columna].fillna(df[f"{columna}_mes"])

    columnas_auxiliares = [
        f"{columna}_mes" for columna in columnas_a_rellenar
    ]

    df = df.drop(columns=columnas_auxiliares, errors="ignore")

    for columna in columnas_a_rellenar:
        df[columna] = df[columna].fillna(perfil_global[columna])

    return df


def crear_variables_calendario_futuro(df: pd.DataFrame) -> pd.DataFrame:
    """
    Crea variables cíclicas de calendario para la previsión futura.
    """

    df = df.copy()

    df["hora_sin"] = np.sin(2 * np.pi * df["hora"] / 24)
    df["hora_cos"] = np.cos(2 * np.pi * df["hora"] / 24)

    df["mes_sin"] = np.sin(2 * np.pi * df["mes"] / 12)
    df["mes_cos"] = np.cos(2 * np.pi * df["mes"] / 12)

    df["dia_semana_sin"] = np.sin(2 * np.pi * df["dia_semana"] / 7)
    df["dia_semana_cos"] = np.cos(2 * np.pi * df["dia_semana"] / 7)

    return df


def crear_variables_mix_futuro(df: pd.DataFrame) -> pd.DataFrame:
    """
    Crea variables derivadas del mix de generación para el futuro.
    """

    df = df.copy()

    df["generacion_renovable"] = (
        df["generacion_solar"].fillna(0)
        + df["generacion_eolica"].fillna(0)
        + df["generacion_hidraulica"].fillna(0)
    )

    df["generacion_termica_convencional"] = (
        df["generacion_nuclear"].fillna(0)
        + df["generacion_ciclo_combinado"].fillna(0)
        + df["generacion_carbon"].fillna(0)
    )

    df["demanda_neta_solar_eolica"] = (
        df["demanda"]
        - df["generacion_solar"].fillna(0)
        - df["generacion_eolica"].fillna(0)
    )

    df["hueco_termico_aproximado"] = (
        df["demanda"]
        - df["generacion_solar"].fillna(0)
        - df["generacion_eolica"].fillna(0)
        - df["generacion_hidraulica"].fillna(0)
        - df["generacion_nuclear"].fillna(0)
    )

    df["porcentaje_renovable_demanda"] = (
        df["generacion_renovable"] / df["demanda"] * 100
    )

    df["porcentaje_solar_demanda"] = (
        df["generacion_solar"] / df["demanda"] * 100
    )

    df["porcentaje_eolica_demanda"] = (
        df["generacion_eolica"] / df["demanda"] * 100
    )

    df["porcentaje_ciclo_combinado_demanda"] = (
        df["generacion_ciclo_combinado"] / df["demanda"] * 100
    )

    df["coste_gas_ciclo_aproximado"] = (
        df["precio_gas_EUR_MWh"] * df["generacion_ciclo_combinado"] / df["demanda"]
    )

    df["gas_por_hueco_termico"] = (
        df["precio_gas_EUR_MWh"] * df["hueco_termico_aproximado"]
    )

    #df["co2_por_hueco_termico"] = (
        #df["precio_co2_EUR_tCO2"] * df["hueco_termico_aproximado"]
    #)

    #df["co2_ciclo_relativo"] = (
        #df["precio_co2_EUR_tCO2"]
        #* df["generacion_ciclo_combinado"]
        #/ df["demanda"]
    #)

    #df["coste_combustible_y_co2_aproximado"] = (
        #df["gas_por_hueco_termico"]
        #+ df["co2_por_hueco_termico"]
    #01-07)

    return df


def construir_base_futura_modelo(
    df_historico: pd.DataFrame,
    fecha_inicio_prevision: str | None = None,
    fecha_fin_prevision: str | None = None,
) -> pd.DataFrame:
    """
    Construye la base futura completa para introducirla en el modelo.
    """

    fecha_inicio, fecha_fin = validar_fechas_prevision(
        fecha_inicio_prevision=fecha_inicio_prevision,
        fecha_fin_prevision=fecha_fin_prevision,
    )

    df_futuro = crear_base_futura_calendario(
        fecha_inicio=fecha_inicio,
        fecha_fin=fecha_fin,
    )

    perfiles = crear_perfiles_historicos(df_historico)

    df_futuro = asignar_escenario_base_futuro(
        df_futuro=df_futuro,
        perfiles=perfiles,
    )

    df_futuro = crear_variables_calendario_futuro(df_futuro)
    df_futuro = crear_variables_mix_futuro(df_futuro)

    return df_futuro


def generar_prediccion_precio(
    df_futuro: pd.DataFrame,
    paquete_modelo: dict,
) -> pd.DataFrame:
    """
    Aplica el modelo entrenado sobre la base futura.
    """

    modelo = paquete_modelo["modelo"]
    columnas_variables = paquete_modelo["columnas_variables"]

    columnas_faltantes = [
        columna for columna in columnas_variables
        if columna not in df_futuro.columns
    ]

    if columnas_faltantes:
        raise ValueError(
            f"Faltan columnas necesarias para la previsión: {columnas_faltantes}"
        )

    df = df_futuro.copy()

    X_futuro = df[columnas_variables]

    df["precio_omie_previsto"] = modelo.predict(X_futuro)

    return df


def crear_resumen_prevision(df_prevision: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    Crea resúmenes diario, mensual y global de la previsión.
    """

    df = df_prevision.copy()

    df["año_mes"] = df["fecha"].dt.to_period("M").astype(str)

    resumen_diario = (
        df.groupby("fecha", as_index=False)
        .agg(
            precio_medio_previsto=("precio_omie_previsto", "mean"),
            precio_minimo_previsto=("precio_omie_previsto", "min"),
            precio_maximo_previsto=("precio_omie_previsto", "max"),
            horas=("precio_omie_previsto", "count"),
        )
    )

    resumen_mensual = (
        df.groupby("año_mes", as_index=False)
        .agg(
            precio_medio_previsto=("precio_omie_previsto", "mean"),
            precio_minimo_previsto=("precio_omie_previsto", "min"),
            precio_maximo_previsto=("precio_omie_previsto", "max"),
            horas=("precio_omie_previsto", "count"),
        )
    )

    resumen_global = pd.DataFrame(
        [
            {
                "fecha_inicio": df["fecha"].min(),
                "fecha_fin": df["fecha"].max(),
                "horas_prevision": len(df),
                "commodity_media_prevista_EUR_MWh": df["precio_omie_previsto"].mean(),
                "precio_minimo_previsto_EUR_MWh": df["precio_omie_previsto"].min(),
                "precio_maximo_previsto_EUR_MWh": df["precio_omie_previsto"].max(),
                "metodo_escenario_futuro": METODO_ESCENARIO_FUTURO,
            }
        ]
    )

    return resumen_diario, resumen_mensual, resumen_global


def guardar_prevision(
    df_prevision: pd.DataFrame,
    resumen_diario: pd.DataFrame,
    resumen_mensual: pd.DataFrame,
    resumen_global: pd.DataFrame,
) -> None:
    """
    Guarda la previsión futura en un Excel de salida.
    """

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    ruta_salida = OUTPUT_DIR / "prevision_commodity_modelo.xlsx"

    try:
        with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
            resumen_global.to_excel(writer, sheet_name="resumen_global", index=False)
            resumen_mensual.to_excel(writer, sheet_name="resumen_mensual", index=False)
            resumen_diario.to_excel(writer, sheet_name="resumen_diario", index=False)
            df_prevision.to_excel(writer, sheet_name="prevision_horaria", index=False)

    except PermissionError:
        raise PermissionError(
            f"No se puede guardar el archivo {ruta_salida}. "
            f"Probablemente está abierto en Excel. Ciérralo y vuelve a ejecutar."
        )

    print(f"Previsión futura guardada en: {ruta_salida}")


def generar_prevision_futura(
    fecha_inicio_prevision: str | None = None,
    fecha_fin_prevision: str | None = None,
) -> None:
    """
    Genera la previsión futura de commodity con el modelo entrenado.
    """

    print("\nPREVISIÓN FUTURA DE COMMODITY")
    print("-" * 50)

    paquete_modelo = cargar_modelo_entrenado()
    df_historico = cargar_base_entrenamiento()

    df_futuro = construir_base_futura_modelo(
        df_historico=df_historico,
        fecha_inicio_prevision=fecha_inicio_prevision,
        fecha_fin_prevision=fecha_fin_prevision,
    )

    df_prevision = generar_prediccion_precio(
        df_futuro=df_futuro,
        paquete_modelo=paquete_modelo,
    )

    resumen_diario, resumen_mensual, resumen_global = crear_resumen_prevision(
        df_prevision
    )

    guardar_prevision(
        df_prevision=df_prevision,
        resumen_diario=resumen_diario,
        resumen_mensual=resumen_mensual,
        resumen_global=resumen_global,
    )

    commodity_media = resumen_global.loc[
        0,
        "commodity_media_prevista_EUR_MWh",
    ]

    fecha_inicio_real = resumen_global.loc[0, "fecha_inicio"]
    fecha_fin_real = resumen_global.loc[0, "fecha_fin"]

    print(f"Fecha inicio previsión: {fecha_inicio_real}")
    print(f"Fecha fin previsión: {fecha_fin_real}")
    print(f"Commodity media prevista: {commodity_media:.2f} €/MWh")
    print("Previsión futura completada correctamente.")
    print("-" * 50)