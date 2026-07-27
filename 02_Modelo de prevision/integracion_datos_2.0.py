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


def cargar_demanda_ree() -> pd.DataFrame:
    """
    Carga el archivo de demanda horaria de REE.
    """

    ruta_demanda = PROCESSED_DATA_DIR / "demanda_ree_horaria.xlsx"

    if not ruta_demanda.exists():
        raise FileNotFoundError(f"No existe el archivo: {ruta_demanda}")

    df_demanda = pd.read_excel(ruta_demanda)

    return df_demanda


def integrar_precio_omie(
    df_base: pd.DataFrame,
    df_omie: pd.DataFrame,
) -> pd.DataFrame:
    """
    Integra el precio OMIE horario dentro de la base histórica.
    """

    columnas_omie = [
        "fecha_hora_utc",
        "precio_omie",
        "numero_valores_originales_omie",
        "fichero_origen_omie",
    ]

    df_omie_reducido = df_omie[columnas_omie].copy()

    df_base = df_base.drop(
        columns=[
            "precio_omie",
            "numero_valores_originales_omie",
            "fichero_origen_omie",
        ],
        errors="ignore",
    )

    df_base = df_base.merge(
        df_omie_reducido,
        on="fecha_hora_utc",
        how="left",
    )

    return df_base


def integrar_demanda_ree(
    df_base: pd.DataFrame,
    df_demanda: pd.DataFrame,
) -> pd.DataFrame:
    """
    Integra la demanda eléctrica de REE dentro de la base histórica.
    """

    columnas_demanda = [
        "fecha_hora_utc",
        "demanda",
        "serie_origen_ree",
    ]

    df_demanda_reducido = df_demanda[columnas_demanda].copy()

    df_base = df_base.drop(
        columns=[
            "demanda",
            "serie_origen_ree",
        ],
        errors="ignore",
    )

    df_base = df_base.merge(
        df_demanda_reducido,
        on="fecha_hora_utc",
        how="left",
    )

    return df_base


def mostrar_resumen_integracion(df_base: pd.DataFrame) -> None:
    """
    Muestra un resumen de cuántos datos se han integrado.
    """

    total_filas = len(df_base)

    filas_con_precio = df_base["precio_omie"].notna().sum()
    filas_sin_precio = df_base["precio_omie"].isna().sum()

    filas_con_demanda = df_base["demanda"].notna().sum()
    filas_sin_demanda = df_base["demanda"].isna().sum()

    print("\nINTEGRACIÓN DE DATOS DEL MODELO")
    print("-" * 50)
    print(f"Filas totales de la base: {total_filas}")

    print("\nOMIE")
    print(f"Filas con precio OMIE: {filas_con_precio}")
    print(f"Filas sin precio OMIE: {filas_sin_precio}")

    if filas_con_precio > 0:
        fecha_min_omie = df_base.loc[
            df_base["precio_omie"].notna(),
            "fecha_hora_local",
        ].min()

        fecha_max_omie = df_base.loc[
            df_base["precio_omie"].notna(),
            "fecha_hora_local",
        ].max()

        print(f"Primera hora con OMIE: {fecha_min_omie}")
        print(f"Última hora con OMIE: {fecha_max_omie}")

    print("\nREE - DEMANDA")
    print(f"Filas con demanda REE: {filas_con_demanda}")
    print(f"Filas sin demanda REE: {filas_sin_demanda}")

    if filas_con_demanda > 0:
        fecha_min_demanda = df_base.loc[
            df_base["demanda"].notna(),
            "fecha_hora_local",
        ].min()

        fecha_max_demanda = df_base.loc[
            df_base["demanda"].notna(),
            "fecha_hora_local",
        ].max()

        print(f"Primera hora con demanda: {fecha_min_demanda}")
        print(f"Última hora con demanda: {fecha_max_demanda}")

    print("-" * 50)


def guardar_base_modelo() -> None:
    """
    Crea la base de datos del modelo integrando:
    - plantilla histórica
    - precios OMIE
    - demanda REE
    """

    df_base = cargar_plantilla_historica()

    df_omie = cargar_precios_omie()
    df_base = integrar_precio_omie(
        df_base=df_base,
        df_omie=df_omie,
    )

    df_demanda = cargar_demanda_ree()
    df_base = integrar_demanda_ree(
        df_base=df_base,
        df_demanda=df_demanda,
    )

    ruta_salida = PROCESSED_DATA_DIR / "base_modelo.xlsx"

    try:
        df_base.to_excel(ruta_salida, index=False)

    except PermissionError:
        raise PermissionError(
            f"No se puede guardar el archivo {ruta_salida}. "
            f"Probablemente está abierto en Excel. Ciérralo y vuelve a ejecutar."
        )

    mostrar_resumen_integracion(df_base)

    print(f"Base del modelo guardada en: {ruta_salida}")