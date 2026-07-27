import numpy as np
import pandas as pd

from config import PROCESSED_DATA_DIR


COLUMNAS_NUMERICAS_BASE = [
    "precio_omie",
    "demanda",
    "generacion_solar",
    "generacion_eolica",
    "generacion_hidraulica",
    "generacion_nuclear",
    "generacion_ciclo_combinado",
    "generacion_carbon",
    "precio_gas_EUR_MWh",
    "precio_co2_EUR_tCO2",
]


COLUMNAS_OBLIGATORIAS_ENTRENAMIENTO = [
    "precio_omie",
    "demanda",
    "generacion_solar",
    "generacion_eolica",
    "generacion_hidraulica",
    "generacion_nuclear",
    "generacion_ciclo_combinado",
    "precio_gas_EUR_MWh",
    #"precio_co2_EUR_tCO2",
]


def cargar_base_modelo() -> pd.DataFrame:
    """
    Carga la base principal del modelo.
    """

    ruta_base = PROCESSED_DATA_DIR / "base_modelo.xlsx"

    if not ruta_base.exists():
        raise FileNotFoundError(f"No existe el archivo: {ruta_base}")

    df = pd.read_excel(ruta_base)

    return df


def preparar_tipos_de_datos(df: pd.DataFrame) -> pd.DataFrame:
    """
    Convierte columnas de fechas y columnas numéricas al formato correcto.
    """

    df = df.copy()

    df["fecha"] = pd.to_datetime(df["fecha"])
    df["fecha_hora_utc_dt"] = pd.to_datetime(df["fecha_hora_utc"], utc=True)

    for columna in COLUMNAS_NUMERICAS_BASE:
        if columna in df.columns:
            df[columna] = pd.to_numeric(df[columna], errors="coerce")
        else:
            df[columna] = np.nan

    return df


def crear_variables_calendario(df: pd.DataFrame) -> pd.DataFrame:
    """
    Crea variables temporales útiles para el modelo.
    """

    df = df.copy()

    df["año"] = df["fecha"].dt.year
    df["mes"] = df["fecha"].dt.month
    df["dia_mes"] = df["fecha"].dt.day

    df["hora"] = pd.to_numeric(df["hora"], errors="coerce")
    df["dia_semana"] = pd.to_numeric(df["dia_semana"], errors="coerce")

    df["hora_sin"] = np.sin(2 * np.pi * df["hora"] / 24)
    df["hora_cos"] = np.cos(2 * np.pi * df["hora"] / 24)

    df["mes_sin"] = np.sin(2 * np.pi * df["mes"] / 12)
    df["mes_cos"] = np.cos(2 * np.pi * df["mes"] / 12)

    df["dia_semana_sin"] = np.sin(2 * np.pi * df["dia_semana"] / 7)
    df["dia_semana_cos"] = np.cos(2 * np.pi * df["dia_semana"] / 7)

    return df


def crear_variables_mix_generacion(df: pd.DataFrame) -> pd.DataFrame:
    """
    Crea variables derivadas del mix de generación.
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

    df["co2_por_hueco_termico"] = (
        df["precio_co2_EUR_tCO2"] * df["hueco_termico_aproximado"]
    )

    df["co2_ciclo_relativo"] = (
        df["precio_co2_EUR_tCO2"]
        * df["generacion_ciclo_combinado"]
        / df["demanda"]
    )

    df["coste_combustible_y_co2_aproximado"] = (
        df["gas_por_hueco_termico"]
        + df["co2_por_hueco_termico"]
    )

    return df


def seleccionar_columnas_entrenamiento(df: pd.DataFrame) -> pd.DataFrame:
    """
    Selecciona las columnas que formarán parte de la base de entrenamiento.
    """

    columnas_entrenamiento = [
        "fecha_hora_local",
        "fecha_hora_utc",
        "fecha",
        "año",
        "mes",
        "dia_mes",
        "hora",
        "dia_semana",
        "es_fin_semana",
        "es_festivo",
        "precio_omie",
        "demanda",
        "generacion_solar",
        "generacion_eolica",
        "generacion_hidraulica",
        "generacion_nuclear",
        "generacion_ciclo_combinado",
        "generacion_carbon",
        "precio_gas_EUR_MWh",
        #"precio_co2_EUR_tCO2",
        "generacion_renovable",
        "generacion_termica_convencional",
        "demanda_neta_solar_eolica",
        "hueco_termico_aproximado",
        "porcentaje_renovable_demanda",
        "porcentaje_solar_demanda",
        "porcentaje_eolica_demanda",
        "porcentaje_ciclo_combinado_demanda",
        "coste_gas_ciclo_aproximado",
        "gas_por_hueco_termico",
        #"co2_por_hueco_termico",
        #"co2_ciclo_relativo",
        #"coste_combustible_y_co2_aproximado",
        "hora_sin",
        "hora_cos",
        "mes_sin",
        "mes_cos",
        "dia_semana_sin",
        "dia_semana_cos",
    ]

    columnas_existentes = [
        columna for columna in columnas_entrenamiento
        if columna in df.columns
    ]

    return df[columnas_existentes].copy()


def filtrar_filas_validas(df: pd.DataFrame) -> pd.DataFrame:
    """
    Elimina filas sin datos esenciales para entrenar el modelo.
    """

    df_filtrado = df.dropna(
        subset=COLUMNAS_OBLIGATORIAS_ENTRENAMIENTO
    ).copy()

    return df_filtrado


def crear_resumen_base_entrenamiento(
    df_original: pd.DataFrame,
    df_entrenamiento: pd.DataFrame,
) -> pd.DataFrame:
    """
    Crea un resumen de control de la base de entrenamiento.
    """

    resumen = {
        "filas_base_modelo": len(df_original),
        "filas_base_entrenamiento": len(df_entrenamiento),
        "filas_eliminadas": len(df_original) - len(df_entrenamiento),
        "fecha_inicio_entrenamiento": df_entrenamiento["fecha_hora_local"].min(),
        "fecha_fin_entrenamiento": df_entrenamiento["fecha_hora_local"].max(),
        "precio_medio_omie": df_entrenamiento["precio_omie"].mean(),
        "demanda_media": df_entrenamiento["demanda"].mean(),
        "generacion_solar_media": df_entrenamiento["generacion_solar"].mean(),
        "generacion_eolica_media": df_entrenamiento["generacion_eolica"].mean(),
        "generacion_ciclo_combinado_media": df_entrenamiento[
            "generacion_ciclo_combinado"
        ].mean(),
    }

    return pd.DataFrame([resumen])


def crear_tabla_correlaciones(df_entrenamiento: pd.DataFrame) -> pd.DataFrame:
    """
    Calcula correlaciones simples entre el precio OMIE y las variables numéricas.
    """

    columnas_numericas = df_entrenamiento.select_dtypes(
        include=["number"]
    ).columns.tolist()

    correlaciones = (
        df_entrenamiento[columnas_numericas]
        .corr(numeric_only=True)["precio_omie"]
        .sort_values(ascending=False)
        .reset_index()
    )

    correlaciones.columns = [
        "variable",
        "correlacion_con_precio_omie",
    ]

    return correlaciones


def guardar_base_entrenamiento() -> None:
    """
    Prepara y guarda la base de entrenamiento del modelo.
    """

    print("\nPREPARACIÓN DE BASE DE ENTRENAMIENTO")
    print("-" * 50)

    df_base = cargar_base_modelo()

    df_preparado = preparar_tipos_de_datos(df_base)
    df_preparado = crear_variables_calendario(df_preparado)
    df_preparado = crear_variables_mix_generacion(df_preparado)

    df_entrenamiento = seleccionar_columnas_entrenamiento(df_preparado)
    df_entrenamiento = filtrar_filas_validas(df_entrenamiento)

    resumen = crear_resumen_base_entrenamiento(
        df_original=df_base,
        df_entrenamiento=df_entrenamiento,
    )

    correlaciones = crear_tabla_correlaciones(df_entrenamiento)

    ruta_base_entrenamiento = PROCESSED_DATA_DIR / "base_entrenamiento.xlsx"
    ruta_control = PROCESSED_DATA_DIR / "control_base_entrenamiento.xlsx"

    try:
        df_entrenamiento.to_excel(ruta_base_entrenamiento, index=False)

        with pd.ExcelWriter(ruta_control, engine="openpyxl") as writer:
            resumen.to_excel(writer, sheet_name="resumen", index=False)
            correlaciones.to_excel(writer, sheet_name="correlaciones", index=False)

    except PermissionError:
        raise PermissionError(
            "No se pueden guardar los archivos de entrenamiento. "
            "Probablemente alguno está abierto en Excel. Ciérralo y vuelve a ejecutar."
        )

    print(f"Base de entrenamiento guardada en: {ruta_base_entrenamiento}")
    print(f"Control de entrenamiento guardado en: {ruta_control}")
    print(f"Filas base modelo: {len(df_base)}")
    print(f"Filas base entrenamiento: {len(df_entrenamiento)}")
    print("-" * 50)