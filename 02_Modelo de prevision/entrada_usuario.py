import pandas as pd

from config import (
    FECHA_INICIO_PREVISION,
    FECHA_FIN_PREVISION,
    MAX_DIAS_PREVISION,
)


def pedir_fecha(mensaje: str, fecha_por_defecto: str) -> str:
    """
    Pide una fecha al usuario en formato AAAA-MM-DD.

    Si el usuario no escribe nada, utiliza la fecha configurada por defecto.
    """

    while True:
        texto = input(f"{mensaje} [{fecha_por_defecto}]: ").strip()

        if texto == "":
            texto = fecha_por_defecto

        try:
            fecha = pd.to_datetime(
                texto,
                format="%Y-%m-%d",
                errors="raise",
            )

            return fecha.strftime("%Y-%m-%d")

        except (ValueError, TypeError):
            print(
                "Fecha no válida. Usa el formato AAAA-MM-DD. "
                "Ejemplo: 2027-01-01"
            )


def pedir_fechas_prevision() -> tuple[str, str]:
    """
    Pide al usuario las fechas inicial y final de la previsión.
    """

    print("\nCONFIGURACIÓN DE FECHAS DE PREVISIÓN")
    print("-" * 50)

    fecha_inicio = pedir_fecha(
        mensaje="Introduce fecha inicio de previsión",
        fecha_por_defecto=FECHA_INICIO_PREVISION,
    )

    fecha_fin = pedir_fecha(
        mensaje="Introduce fecha fin de previsión",
        fecha_por_defecto=FECHA_FIN_PREVISION,
    )

    fecha_inicio_ts = pd.to_datetime(
        fecha_inicio,
        format="%Y-%m-%d",
    )

    fecha_fin_ts = pd.to_datetime(
        fecha_fin,
        format="%Y-%m-%d",
    )

    if fecha_fin_ts < fecha_inicio_ts:
        raise ValueError(
            "La fecha fin de previsión no puede ser anterior "
            "a la fecha de inicio."
        )

    dias_prevision = (fecha_fin_ts - fecha_inicio_ts).days + 1

    if dias_prevision > MAX_DIAS_PREVISION:
        raise ValueError(
            f"La previsión solicitada tiene {dias_prevision} días. "
            f"El máximo permitido es {MAX_DIAS_PREVISION} días."
        )

    print(f"Fecha inicio seleccionada: {fecha_inicio}")
    print(f"Fecha fin seleccionada: {fecha_fin}")
    print(f"Número de días de previsión: {dias_prevision}")
    print("-" * 50)

    return fecha_inicio, fecha_fin