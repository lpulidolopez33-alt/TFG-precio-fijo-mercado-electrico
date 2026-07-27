import pandas as pd
import holidays


from config import (
    FECHA_INICIO_HISTORICO,
    FECHA_FIN_HISTORICO,
    PROCESSED_DATA_DIR,
)


ZONA_HORARIA = "Europe/Madrid"           # indica que vamos a trabajar con horario español.


def crear_indice_horario(fecha_inicio: str, fecha_fin: str) -> pd.DatetimeIndex:                    # crea todas las horas entre dos fechas.
    """
    Crea un índice horario entre dos fechas, usando la zona horaria española.

    La fecha final se incluye completa.
    """

    inicio = pd.Timestamp(fecha_inicio).tz_localize(ZONA_HORARIA)                                   # convierte la fecha de inicio en una fecha entendible para Python y le asigna la zona horaria española.
    fin_exclusivo = (pd.Timestamp(fecha_fin) + pd.Timedelta(days=1)).tz_localize(ZONA_HORARIA)

    indice_horario = pd.date_range(
        start=inicio,
        end=fin_exclusivo,
        freq="h",
        inclusive="left",
    )

    return indice_horario


def construir_plantilla_historica() -> pd.DataFrame:
    """
    Construye una tabla horaria con todas las columnas necesarias para el modelo.
    """

    indice_horario = crear_indice_horario(
        FECHA_INICIO_HISTORICO,
        FECHA_FIN_HISTORICO,
    )

    df = pd.DataFrame()                                                                          # estamos creando una tabla vacía, parecida a una hoja de Excel.

    df["fecha_hora_local"] = indice_horario.astype(str)
    df["fecha_hora_utc"] = indice_horario.tz_convert("UTC").astype(str)

    df["fecha"] = indice_horario.date
    df["año"] = indice_horario.year
    df["mes"] = indice_horario.month
    df["dia"] = indice_horario.day
    df["hora"] = indice_horario.hour

    df["dia_semana"] = indice_horario.dayofweek
    df["es_fin_semana"] = df["dia_semana"].isin([5, 6]).astype(int)

    festivos_españa = holidays.Spain(years=df["año"].unique())

    df["es_festivo"] = df["fecha"].apply(
        lambda fecha: 1 if fecha in festivos_españa else 0
    )

    df["es_hora_repetida_cambio_horario"] = df.duplicated(
        subset=["fecha", "hora"],
        keep=False,
    ).astype(int)

    df["precio_omie"] = None
    df["demanda"] = None
    df["generacion_solar"] = None
    df["generacion_eolica"] = None
    df["generacion_hidraulica"] = None
    df["generacion_nuclear"] = None
    df["generacion_ciclo_combinado"] = None
    df["generacion_carbon"] = None
    df["precio_gas"] = None
    df["precio_co2"] = None

    df["demanda_neta"] = None
    df["porcentaje_renovable"] = None

    df["periodo_20td"] = None
    df["periodo_30td"] = None
    df["periodo_6xtd"] = None

    return df


def guardar_plantilla_historica() -> None:
    """
    Guarda la plantilla histórica en un archivo Excel.
    """

    plantilla = construir_plantilla_historica()

    ruta_salida = PROCESSED_DATA_DIR / "plantilla_historica.xlsx"

    plantilla.to_excel(ruta_salida, index=False)

    print(f"Plantilla histórica creada: {ruta_salida}")
    print(f"Número de filas horarias: {len(plantilla)}")
    print(f"Número de columnas: {len(plantilla.columns)}")