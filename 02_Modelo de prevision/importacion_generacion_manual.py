import pandas as pd

from config import RUTA_GENERACION_MANUAL, PROCESSED_DATA_DIR


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


def cargar_generacion_manual() -> pd.DataFrame:
    """
    Carga el archivo manual de generación horaria.

    El archivo debe estar en:
    datos/entrada_manual/generacion_manual.xlsx
    """

    if not RUTA_GENERACION_MANUAL.exists():
        raise FileNotFoundError(
            f"No existe el archivo de generación manual: {RUTA_GENERACION_MANUAL}. "
            f"Crea una copia de plantilla_generacion_manual.xlsx y llámala "
            f"generacion_manual.xlsx."
        )

    df = pd.read_excel(RUTA_GENERACION_MANUAL)

    return df


def validar_columnas_generacion_manual(df: pd.DataFrame) -> None:
    """
    Comprueba que el Excel manual tenga las columnas necesarias.
    """

    columnas_obligatorias = [
        "fecha_hora_utc",
    ]

    columnas_faltantes = []

    for columna in columnas_obligatorias:
        if columna not in df.columns:
            columnas_faltantes.append(columna)

    if columnas_faltantes:
        raise ValueError(
            f"Faltan columnas obligatorias en generación_manual.xlsx: "
            f"{columnas_faltantes}"
        )


def preparar_generacion_manual(df: pd.DataFrame) -> pd.DataFrame:
    """
    Limpia y prepara la generación manual para integrarla en la base del modelo.
    """

    df = df.copy()

    validar_columnas_generacion_manual(df)

    for columna in COLUMNAS_GENERACION_MANUAL:
        if columna not in df.columns:
            df[columna] = None

    for columna in COLUMNAS_GENERACION_MANUAL:
        df[columna] = pd.to_numeric(df[columna], errors="coerce")

    hay_solar_fotovoltaica = df["generacion_solar_fotovoltaica"].notna()
    hay_solar_termica = df["generacion_solar_termica"].notna()
    hay_alguna_solar_desagregada = hay_solar_fotovoltaica | hay_solar_termica

    solar_calculada = (
        df["generacion_solar_fotovoltaica"].fillna(0)
        + df["generacion_solar_termica"].fillna(0)
    )

    df.loc[
        df["generacion_solar"].isna() & hay_alguna_solar_desagregada,
        "generacion_solar",
    ] = solar_calculada

    df = df[
        [
            "fecha_hora_utc",
            "generacion_solar_fotovoltaica",
            "generacion_solar_termica",
            "generacion_solar",
            "generacion_eolica",
            "generacion_hidraulica",
            "generacion_nuclear",
            "generacion_ciclo_combinado",
            "generacion_carbon",
        ]
    ]

    df = df.drop_duplicates(
        subset=["fecha_hora_utc"],
        keep="last",
    )

    return df


def guardar_generacion_manual_procesada() -> None:
    """
    Lee el Excel manual de generación y guarda una versión procesada.
    """

    print("\nIMPORTACIÓN MANUAL DE GENERACIÓN")
    print("-" * 50)

    df = cargar_generacion_manual()
    df = preparar_generacion_manual(df)

    ruta_salida = PROCESSED_DATA_DIR / "generacion_manual_procesada.xlsx"

    try:
        df.to_excel(ruta_salida, index=False)

    except PermissionError:
        raise PermissionError(
            f"No se puede guardar el archivo {ruta_salida}. "
            f"Probablemente está abierto en Excel. Ciérralo y vuelve a ejecutar."
        )

    print(f"Generación manual procesada guardada en: {ruta_salida}")
    print(f"Número de filas de generación manual: {len(df)}")

    for columna in COLUMNAS_GENERACION_MANUAL:
        datos_disponibles = df[columna].notna().sum()
        print(f"{columna}: {datos_disponibles} datos disponibles")

    print("-" * 50)