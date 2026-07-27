import pandas as pd

from config import (
    RUTA_GAS_MANUAL,
    PROCESSED_DATA_DIR,
)

from datos_base import crear_indice_horario


def cargar_gas_manual() -> pd.DataFrame:
    """
    Carga el archivo manual de precio del gas.

    El archivo debe estar en:
    datos/entrada_manual/gas_manual.xlsx
    """

    if not RUTA_GAS_MANUAL.exists():
        raise FileNotFoundError(
            f"No existe el archivo de gas manual: {RUTA_GAS_MANUAL}. "
            f"Crea una copia de plantilla_gas_manual.xlsx y llámala gas_manual.xlsx."
        )

    df = pd.read_excel(RUTA_GAS_MANUAL)

    return df


def preparar_gas_diario(df: pd.DataFrame) -> pd.DataFrame:
    """
    Limpia y prepara el precio diario del gas.
    """

    df = df.copy()

    columnas_obligatorias = [
        "fecha",
        "precio_gas_EUR_MWh",
    ]

    for columna in columnas_obligatorias:
        if columna not in df.columns:
            raise ValueError(
                f"Falta la columna obligatoria '{columna}' en gas_manual.xlsx"
            )

    df["fecha"] = pd.to_datetime(df["fecha"])
    df["precio_gas_EUR_MWh"] = pd.to_numeric(
        df["precio_gas_EUR_MWh"],
        errors="coerce",
    )

    df = df[
        [
            "fecha",
            "precio_gas_EUR_MWh",
        ]
    ]

    df = df.drop_duplicates(
        subset=["fecha"],
        keep="last",
    )

    df = df.sort_values("fecha")

    return df


def convertir_gas_diario_a_horario(df_gas_diario: pd.DataFrame) -> pd.DataFrame:
    """
    Convierte el precio diario del gas a formato horario.

    El precio del gas es diario, por lo que se repite para todas las horas
    de cada día.
    """

    fecha_inicio = df_gas_diario["fecha"].min()
    fecha_fin = df_gas_diario["fecha"].max()

    indice_horario = crear_indice_horario(
        fecha_inicio.strftime("%Y-%m-%d"),
        fecha_fin.strftime("%Y-%m-%d"),
    )

    df_horario = pd.DataFrame()

    df_horario["fecha_hora_local"] = indice_horario.astype(str)
    df_horario["fecha_hora_utc"] = indice_horario.tz_convert("UTC").astype(str)
    df_horario["fecha"] = pd.to_datetime(indice_horario.date)

    df_horario = df_horario.merge(
        df_gas_diario,
        on="fecha",
        how="left",
    )

    return df_horario


def guardar_gas_manual_procesado() -> None:
    """
    Lee el archivo gas_manual.xlsx y guarda una versión horaria procesada.
    """

    print("\nIMPORTACIÓN MANUAL DE PRECIO DEL GAS")
    print("-" * 50)

    df_gas = cargar_gas_manual()
    df_gas_diario = preparar_gas_diario(df_gas)

    filas_con_gas_diario = df_gas_diario["precio_gas_EUR_MWh"].notna().sum()

    if filas_con_gas_diario == 0:
        print(
            "Aviso: gas_manual.xlsx existe, pero todavía no contiene precios de gas."
        )
        print(
            "Se creará el archivo procesado, pero la columna precio_gas_EUR_MWh "
            "quedará vacía."
        )

    df_gas_horario = convertir_gas_diario_a_horario(df_gas_diario)

    ruta_salida = PROCESSED_DATA_DIR / "gas_manual_procesado.xlsx"

    try:
        df_gas_horario.to_excel(ruta_salida, index=False)

    except PermissionError:
        raise PermissionError(
            f"No se puede guardar el archivo {ruta_salida}. "
            f"Probablemente está abierto en Excel. Ciérralo y vuelve a ejecutar."
        )

    filas_horarias_con_gas = df_gas_horario["precio_gas_EUR_MWh"].notna().sum()

    print(f"Gas manual procesado guardado en: {ruta_salida}")
    print(f"Número de días en gas_manual.xlsx: {len(df_gas_diario)}")
    print(f"Días con precio de gas: {filas_con_gas_diario}")
    print(f"Número de filas horarias generadas: {len(df_gas_horario)}")
    print(f"Horas con precio de gas: {filas_horarias_con_gas}")
    print("-" * 50)