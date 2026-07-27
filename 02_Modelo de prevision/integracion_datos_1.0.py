import pandas as pd

from config import PROCESSED_DATA_DIR


def cargar_plantilla_historica() -> pd.DataFrame:
    """
    Carga la plantilla histórica generada previamente.
    """

    ruta_plantilla = PROCESSED_DATA_DIR / "plantilla_historica.xlsx"

    if not ruta_plantilla.exists():
        raise FileNotFoundError(f"No existe el archivo: {ruta_plantilla}")

    df_plantilla = pd.read_excel(ruta_plantilla)

    return df_plantilla


def cargar_precios_omie() -> pd.DataFrame:
    """
    Carga el archivo de precios horarios OMIE.
    """

    ruta_omie = PROCESSED_DATA_DIR / "precios_omie_horarios.xlsx"

    if not ruta_omie.exists():
        raise FileNotFoundError(f"No existe el archivo: {ruta_omie}")

    df_omie = pd.read_excel(ruta_omie)

    return df_omie


def integrar_precio_omie(
    df_plantilla: pd.DataFrame,
    df_omie: pd.DataFrame,
) -> pd.DataFrame:
    """
    Integra el precio OMIE horario dentro de la plantilla histórica.
    """

    columnas_omie = [
        "fecha_hora_utc",
        "precio_omie",
        "numero_valores_originales_omie",
        "fichero_origen_omie",
    ]

    df_omie_reducido = df_omie[columnas_omie].copy()

    df_base = df_plantilla.drop(columns=["precio_omie"], errors="ignore").merge(
        df_omie_reducido,
        on="fecha_hora_utc",
        how="left",
    )

    return df_base


def mostrar_resumen_integracion(df_base: pd.DataFrame) -> None:
    """
    Muestra un resumen de cuántos precios OMIE se han integrado.
    """

    total_filas = len(df_base)
    filas_con_precio = df_base["precio_omie"].notna().sum()
    filas_sin_precio = df_base["precio_omie"].isna().sum()

    print("\nINTEGRACIÓN DE DATOS OMIE")
    print("-" * 50)
    print(f"Filas totales de la base: {total_filas}")
    print(f"Filas con precio OMIE: {filas_con_precio}")
    print(f"Filas sin precio OMIE: {filas_sin_precio}")

    if filas_con_precio > 0:
        fecha_min = df_base.loc[df_base["precio_omie"].notna(), "fecha_hora_local"].min()
        fecha_max = df_base.loc[df_base["precio_omie"].notna(), "fecha_hora_local"].max()

        print(f"Primera hora con OMIE: {fecha_min}")
        print(f"Última hora con OMIE: {fecha_max}")

    print("-" * 50)


def guardar_base_modelo() -> None:
    """
    Crea la base de datos del modelo integrando la plantilla histórica con OMIE.
    """

    df_plantilla = cargar_plantilla_historica()
    df_omie = cargar_precios_omie()

    df_base = integrar_precio_omie(
        df_plantilla=df_plantilla,
        df_omie=df_omie,
    )

    ruta_salida = PROCESSED_DATA_DIR / "base_modelo.xlsx"

    df_base.to_excel(ruta_salida, index=False)

    mostrar_resumen_integracion(df_base)

    print(f"Base del modelo guardada en: {ruta_salida}")