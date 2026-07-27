import pandas as pd

from config import (
    FECHA_INICIO_HISTORICO,
    FECHA_FIN_HISTORICO,
    MANUAL_DATA_DIR,
    RUTA_PLANTILLA_CO2_MANUAL,
)


def crear_plantilla_co2_manual() -> None:
    """
    Crea una plantilla diaria para introducir manualmente el precio del CO2.

    El precio del CO2 se introduce en €/tCO2.
    Después el programa lo repetirá para todas las horas de cada día.
    """

    MANUAL_DATA_DIR.mkdir(parents=True, exist_ok=True)

    if RUTA_PLANTILLA_CO2_MANUAL.exists():
        print(f"Plantilla de CO2 manual ya existe: {RUTA_PLANTILLA_CO2_MANUAL}")
        return

    fechas = pd.date_range(
        start=FECHA_INICIO_HISTORICO,
        end=FECHA_FIN_HISTORICO,
        freq="D",
    )

    df = pd.DataFrame()

    df["fecha"] = fechas.date
    df["año"] = fechas.year
    df["mes"] = fechas.month
    df["dia"] = fechas.day

    df["precio_co2_EUR_tCO2"] = None
    df["fuente"] = None
    df["observaciones"] = None

    try:
        df.to_excel(RUTA_PLANTILLA_CO2_MANUAL, index=False)

    except PermissionError:
        raise PermissionError(
            f"No se puede guardar el archivo {RUTA_PLANTILLA_CO2_MANUAL}. "
            f"Probablemente está abierto en Excel. Ciérralo y vuelve a ejecutar."
        )

    print(f"Plantilla de CO2 manual creada: {RUTA_PLANTILLA_CO2_MANUAL}")
    print(f"Número de días: {len(df)}")