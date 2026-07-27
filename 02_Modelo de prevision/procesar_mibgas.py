from pathlib import Path

import pandas as pd

from config import (
    FECHA_INICIO_HISTORICO,
    FECHA_FIN_HISTORICO,
    MIBGAS_DATA_DIR,
    RUTA_GAS_MANUAL,
)


def normalizar_nombre_columna(columna: str) -> str:
    """
    Normaliza nombres de columnas para poder identificarlas aunque tengan
    saltos de línea o espacios.
    """

    return (
        str(columna)
        .replace("\n", " ")
        .replace("\r", " ")
        .strip()
        .lower()
    )


def localizar_columna(
    columnas: list[str],
    textos_obligatorios: list[str],
) -> str:
    """
    Busca una columna que contenga todos los textos indicados.
    """

    for columna in columnas:
        columna_normalizada = normalizar_nombre_columna(columna)

        if all(texto.lower() in columna_normalizada for texto in textos_obligatorios):
            return columna

    raise ValueError(
        f"No se ha encontrado ninguna columna con los textos: {textos_obligatorios}"
    )


def leer_mibgas_indices(ruta_excel: Path) -> pd.DataFrame:
    """
    Lee la hoja 'MIBGAS Indexes' de un fichero anual de MIBGAS.

    Se toma como referencia principal:
    MIBGAS PVB Average Price Index Day-Ahead [EUR/MWh]

    Si algún día no tiene precio medio, se usa como respaldo:
    MIBGAS PVB Last Price Index Day-Ahead [EUR/MWh]
    """

    df = pd.read_excel(
        ruta_excel,
        sheet_name="MIBGAS Indexes",
    )

    columnas = df.columns.tolist()

    columna_fecha = localizar_columna(
        columnas=columnas,
        textos_obligatorios=["delivery day"],
    )

    columna_precio_medio = localizar_columna(
        columnas=columnas,
        textos_obligatorios=["pvb", "average price", "day-ahead"],
    )

    columna_precio_last = localizar_columna(
        columnas=columnas,
        textos_obligatorios=["pvb", "last price", "day-ahead"],
    )

    df_salida = pd.DataFrame()

    df_salida["fecha"] = pd.to_datetime(df[columna_fecha], errors="coerce")
    df_salida["precio_gas_EUR_MWh"] = pd.to_numeric(
        df[columna_precio_medio],
        errors="coerce",
    )

    precio_respaldo = pd.to_numeric(
        df[columna_precio_last],
        errors="coerce",
    )

    df_salida["precio_gas_EUR_MWh"] = df_salida[
        "precio_gas_EUR_MWh"
    ].fillna(precio_respaldo)

    df_salida["fuente"] = "MIBGAS"
    df_salida["observaciones"] = "PVB Average Price Index Day-Ahead; respaldo Last Price Index"
    df_salida["archivo_origen"] = ruta_excel.name

    df_salida = df_salida.dropna(subset=["fecha"])

    df_salida = df_salida[
        [
            "fecha",
            "precio_gas_EUR_MWh",
            "fuente",
            "observaciones",
            "archivo_origen",
        ]
    ]

    return df_salida


def crear_base_diaria_gas() -> pd.DataFrame:
    """
    Crea la base diaria completa del histórico.
    """

    fechas = pd.date_range(
        start=FECHA_INICIO_HISTORICO,
        end=FECHA_FIN_HISTORICO,
        freq="D",
    )

    df = pd.DataFrame()

    df["fecha"] = fechas
    df["año"] = fechas.year
    df["mes"] = fechas.month
    df["dia"] = fechas.day

    return df


def procesar_ficheros_mibgas() -> None:
    """
    Procesa todos los ficheros Excel de MIBGAS guardados en:
    datos/entrada_manual/mibgas

    y genera automáticamente gas_manual.xlsx.
    """

    print("\nPROCESADO DE FICHEROS MIBGAS")
    print("-" * 50)

    MIBGAS_DATA_DIR.mkdir(parents=True, exist_ok=True)

    archivos = sorted(MIBGAS_DATA_DIR.glob("*.xlsx"))

    if not archivos:
        raise FileNotFoundError(
            f"No hay ficheros .xlsx en {MIBGAS_DATA_DIR}. "
            f"Descarga los ficheros anuales de MIBGAS y guárdalos ahí."
        )

    tablas = []

    for archivo in archivos:
        try:
            df_archivo = leer_mibgas_indices(archivo)
            tablas.append(df_archivo)

            print(
                f"Archivo leído: {archivo.name} - "
                f"{df_archivo['precio_gas_EUR_MWh'].notna().sum()} días con precio"
            )

        except Exception as error:
            print(f"ERROR leyendo {archivo.name}: {error}")

    if not tablas:
        raise RuntimeError("No se ha podido leer ningún fichero MIBGAS válido.")

    df_mibgas = pd.concat(tablas, ignore_index=True)

    df_mibgas = df_mibgas.drop_duplicates(
        subset=["fecha"],
        keep="last",
    )

    df_mibgas = df_mibgas.sort_values("fecha")

    df_base = crear_base_diaria_gas()

    df_base = df_base.merge(
        df_mibgas,
        on="fecha",
        how="left",
    )

    df_base["fecha"] = df_base["fecha"].dt.date

    try:
        df_base.to_excel(RUTA_GAS_MANUAL, index=False)

    except PermissionError:
        raise PermissionError(
            f"No se puede guardar {RUTA_GAS_MANUAL}. "
            f"Probablemente está abierto en Excel. Ciérralo y vuelve a ejecutar."
        )

    dias_con_precio = df_base["precio_gas_EUR_MWh"].notna().sum()
    dias_sin_precio = df_base["precio_gas_EUR_MWh"].isna().sum()

    print(f"Archivo gas_manual.xlsx generado en: {RUTA_GAS_MANUAL}")
    print(f"Días totales: {len(df_base)}")
    print(f"Días con precio gas: {dias_con_precio}")
    print(f"Días sin precio gas: {dias_sin_precio}")

    if dias_con_precio > 0:
        print(f"Precio gas medio: {df_base['precio_gas_EUR_MWh'].mean():.2f} €/MWh")

    print("-" * 50)