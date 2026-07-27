import pandas as pd

from config import (
    FECHA_INICIO_HISTORICO,
    FECHA_FIN_HISTORICO,
    MANUAL_DATA_DIR,
    RUTA_PLANTILLA_GENERACION_MANUAL,
)

from datos_base import crear_indice_horario


COLUMNAS_GENERACION_MANUAL = [
    "generacion_solar_fotovoltaica",
    "generacion_solar_termica",
    "generacion_solar",
    "generacion_eolica",
    "generacion_hidraulica",
    "generacion_nuclear",
    "generacion_ciclo_combinado",
    "generacion_carbon",
]


def crear_plantilla_generacion_manual() -> None:
    """
    Crea una plantilla Excel para introducir manualmente la generación horaria.

    La plantilla tendrá una fila por cada hora del histórico y columnas para
    las principales tecnologías que usará el modelo.
    """

    MANUAL_DATA_DIR.mkdir(parents=True, exist_ok=True)

    if RUTA_PLANTILLA_GENERACION_MANUAL.exists():
        print(
            f"Plantilla de generación manual ya existe: "
            f"{RUTA_PLANTILLA_GENERACION_MANUAL}"
        )
        return

    indice_horario = crear_indice_horario(
        FECHA_INICIO_HISTORICO,
        FECHA_FIN_HISTORICO,
    )

    df = pd.DataFrame()

    df["fecha_hora_local"] = indice_horario.astype(str)
    df["fecha_hora_utc"] = indice_horario.tz_convert("UTC").astype(str)
    df["fecha"] = indice_horario.date
    df["año"] = indice_horario.year
    df["mes"] = indice_horario.month
    df["dia"] = indice_horario.day
    df["hora"] = indice_horario.hour

    for columna in COLUMNAS_GENERACION_MANUAL:
        df[columna] = None

    try:
        df.to_excel(RUTA_PLANTILLA_GENERACION_MANUAL, index=False)

    except PermissionError:
        raise PermissionError(
            f"No se puede guardar el archivo {RUTA_PLANTILLA_GENERACION_MANUAL}. "
            f"Probablemente está abierto en Excel. Ciérralo y vuelve a ejecutar."
        )

    print(
        f"Plantilla de generación manual creada: "
        f"{RUTA_PLANTILLA_GENERACION_MANUAL}"
    )
    print(f"Número de filas horarias: {len(df)}")
    print(f"Número de columnas: {len(df.columns)}")