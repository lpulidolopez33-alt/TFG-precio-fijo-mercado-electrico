import pandas as pd

from config import (
    RUTA_CO2_MANUAL,
    PROCESSED_DATA_DIR,
)

from datos_base import crear_indice_horario


def cargar_co2_manual() -> pd.DataFrame:
    """
    Carga el archivo manual de precio del CO2.

    El archivo debe estar en:
    datos/entrada_manual/co2_manual.xlsx
    """

    if not RUTA_CO2_MANUAL.exists():
        raise FileNotFoundError(
            f"No existe el archivo de CO2 manual: {RUTA_CO2_MANUAL}. "
            f"Crea una copia de plantilla_co2_manual.xlsx y llámala co2_manual.xlsx."
        )

    df = pd.read_excel(RUTA_CO2_MANUAL)

    return df


def preparar_co2_diario(df: pd.DataFrame) -> pd.DataFrame:
    """
    Limpia y prepara el precio diario del CO2.
    """

    df = df.copy()

    columnas_obligatorias = [
        "fecha",
        "precio_co2_EUR_tCO2",
    ]

    for columna in columnas_obligatorias:
        if columna not in df.columns:
            raise ValueError(
                f"Falta la columna obligatoria '{columna}' en co2_manual.xlsx"
            )

    df["fecha"] = pd.to_datetime(df["fecha"])
    df["precio_co2_EUR_tCO2"] = pd.to_numeric(
        df["precio_co2_EUR_tCO2"],
        errors="coerce",
    )

    df = df[
        [
            "fecha",
            "precio_co2_EUR_tCO2",
        ]
    ]

    df = df.drop_duplicates(
        subset=["fecha"],
        keep="last",
    )

    df = df.sort_values("fecha")

    return df


def convertir_co2_diario_a_horario(df_co2_diario: pd.DataFrame) -> pd.DataFrame:
    """
    Convierte el precio diario del CO2 a formato horario.

    El precio del CO2 es diario, por lo que se repite para todas las horas
    de cada día.
    """

    fecha_inicio = df_co2_diario["fecha"].min()
    fecha_fin = df_co2_diario["fecha"].max()

    indice_horario = crear_indice_horario(
        fecha_inicio.strftime("%Y-%m-%d"),
        fecha_fin.strftime("%Y-%m-%d"),
    )

    df_horario = pd.DataFrame()

    df_horario["fecha_hora_local"] = indice_horario.astype(str)
    df_horario["fecha_hora_utc"] = indice_horario.tz_convert("UTC").astype(str)
    df_horario["fecha"] = pd.to_datetime(indice_horario.date)

    df_horario = df_horario.merge(
        df_co2_diario,
        on="fecha",
        how="left",
    )

    return df_horario


def guardar_co2_manual_procesado() -> None:
    """
    Lee el archivo co2_manual.xlsx y guarda una versión horaria procesada.
    """

    print("\nIMPORTACIÓN MANUAL DE PRECIO DEL CO2")
    print("-" * 50)

    df_co2 = cargar_co2_manual()
    df_co2_diario = preparar_co2_diario(df_co2)

    dias_con_co2 = df_co2_diario["precio_co2_EUR_tCO2"].notna().sum()

    if dias_con_co2 == 0:
        print(
            "Aviso: co2_manual.xlsx existe, pero todavía no contiene precios de CO2."
        )
        print(
            "Se creará el archivo procesado, pero la columna precio_co2_EUR_tCO2 "
            "quedará vacía."
        )

    df_co2_horario = convertir_co2_diario_a_horario(df_co2_diario)

    ruta_salida = PROCESSED_DATA_DIR / "co2_manual_procesado.xlsx"

    try:
        df_co2_horario.to_excel(ruta_salida, index=False)

    except PermissionError:
        raise PermissionError(
            f"No se puede guardar el archivo {ruta_salida}. "
            f"Probablemente está abierto en Excel. Ciérralo y vuelve a ejecutar."
        )

    horas_con_co2 = df_co2_horario["precio_co2_EUR_tCO2"].notna().sum()

    print(f"CO2 manual procesado guardado en: {ruta_salida}")
    print(f"Número de días en co2_manual.xlsx: {len(df_co2_diario)}")
    print(f"Días con precio de CO2: {dias_con_co2}")
    print(f"Número de filas horarias generadas: {len(df_co2_horario)}")
    print(f"Horas con precio de CO2: {horas_con_co2}")
    print("-" * 50)