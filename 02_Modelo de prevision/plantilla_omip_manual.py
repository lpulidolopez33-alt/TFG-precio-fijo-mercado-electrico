import pandas as pd

from config import (
    MANUAL_DATA_DIR,
    RUTA_PLANTILLA_OMIP_MANUAL,
)


def crear_plantilla_omip_manual() -> None:
    """
    Crea una plantilla para introducir manualmente precios OMIP.

    Cada fila representa un producto de mercado a plazo:
    mensual, trimestral, anual, etc.
    """

    MANUAL_DATA_DIR.mkdir(parents=True, exist_ok=True)

    if RUTA_PLANTILLA_OMIP_MANUAL.exists():
        print(f"Plantilla OMIP ya existe: {RUTA_PLANTILLA_OMIP_MANUAL}")
        return

    datos_ejemplo = [
        {
            "producto": "Jun-26",
            "tipo_producto": "Mensual",
            "fecha_inicio": "2026-06-01",
            "fecha_fin": "2026-06-30",
            "precio_omip_EUR_MWh": 64.00,
            "fecha_cotizacion": "2026-05-05",
            "fuente": "OMIP",
            "observaciones": "Valor introducido manualmente",
        },
        {
            "producto": "Q3-26",
            "tipo_producto": "Trimestral",
            "fecha_inicio": "2026-07-01",
            "fecha_fin": "2026-09-30",
            "precio_omip_EUR_MWh": 84.80,
            "fecha_cotizacion": "2026-05-05",
            "fuente": "OMIP",
            "observaciones": "Valor introducido manualmente",
        },
        {
            "producto": "Q4-26",
            "tipo_producto": "Trimestral",
            "fecha_inicio": "2026-10-01",
            "fecha_fin": "2026-12-31",
            "precio_omip_EUR_MWh": 88.88,
            "fecha_cotizacion": "2026-05-05",
            "fuente": "OMIP",
            "observaciones": "Valor introducido manualmente",
        },
        {
            "producto": "YR-27",
            "tipo_producto": "Anual",
            "fecha_inicio": "2027-01-01",
            "fecha_fin": "2027-12-31",
            "precio_omip_EUR_MWh": 60.85,
            "fecha_cotizacion": "2026-05-05",
            "fuente": "OMIP",
            "observaciones": "Valor introducido manualmente",
        },
    ]

    df = pd.DataFrame(datos_ejemplo)

    try:
        df.to_excel(RUTA_PLANTILLA_OMIP_MANUAL, index=False)

    except PermissionError:
        raise PermissionError(
            f"No se puede guardar {RUTA_PLANTILLA_OMIP_MANUAL}. "
            f"Probablemente está abierto en Excel. Ciérralo y vuelve a ejecutar."
        )

    print(f"Plantilla OMIP creada: {RUTA_PLANTILLA_OMIP_MANUAL}")
    print(f"Número de productos de ejemplo: {len(df)}")