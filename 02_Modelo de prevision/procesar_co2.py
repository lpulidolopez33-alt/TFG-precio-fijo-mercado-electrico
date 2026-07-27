from pathlib import Path

import pandas as pd

from config import (
    FECHA_INICIO_HISTORICO,
    FECHA_FIN_HISTORICO,
    CO2_DATA_DIR,
    RUTA_CO2_MANUAL,
)


def leer_fichero_co2_investing(ruta_csv: Path) -> pd.DataFrame:
    """
    Lee un fichero CSV descargado de Investing con precios históricos de EUA.

    Columnas esperadas:
    Date, Price, Open, High, Low, Vol., Change %
    """

    df = pd.read_csv(ruta_csv)

    columnas_obligatorias = ["Date", "Price"]

    for columna in columnas_obligatorias:
        if columna not in df.columns:
            raise ValueError(
                f"Falta la columna obligatoria '{columna}' en {ruta_csv.name}"
            )

    df_salida = pd.DataFrame()

    df_salida["fecha"] = pd.to_datetime(
        df["Date"],
        format="%m/%d/%Y",
        errors="coerce",
    )

    df_salida["precio_co2_EUR_tCO2"] = pd.to_numeric(
        df["Price"],
        errors="coerce",
    )

    df_salida["fuente"] = "Investing"
    df_salida["observaciones"] = "European Union Allowance EUA Yearly Futures Historical Data"
    df_salida["archivo_origen"] = ruta_csv.name

    df_salida = df_salida.dropna(subset=["fecha"])
    df_salida = df_salida.sort_values("fecha")

    return df_salida


def crear_base_diaria_co2() -> pd.DataFrame:
    """
    Crea una base diaria completa entre la fecha inicial y final del histórico.
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


def procesar_ficheros_co2() -> None:
    """
    Procesa los CSV de CO2 guardados en:
    datos/entrada_manual/co2

    y genera automáticamente co2_manual.xlsx.
    """

    print("\nPROCESADO DE FICHEROS CO2")
    print("-" * 50)

    CO2_DATA_DIR.mkdir(parents=True, exist_ok=True)

    archivos = sorted(CO2_DATA_DIR.glob("*.csv"))

    if not archivos:
        raise FileNotFoundError(
            f"No hay ficheros .csv en {CO2_DATA_DIR}. "
            f"Descarga el histórico de CO2 y guárdalo ahí."
        )

    tablas = []

    for archivo in archivos:
        try:
            df_archivo = leer_fichero_co2_investing(archivo)
            tablas.append(df_archivo)

            print(
                f"Archivo leído: {archivo.name} - "
                f"{df_archivo['precio_co2_EUR_tCO2'].notna().sum()} días con precio"
            )

        except Exception as error:
            print(f"ERROR leyendo {archivo.name}: {error}")

    if not tablas:
        raise RuntimeError("No se ha podido leer ningún fichero CO2 válido.")

    df_co2 = pd.concat(tablas, ignore_index=True)

    df_co2 = df_co2.drop_duplicates(
        subset=["fecha"],
        keep="last",
    )

    df_co2 = df_co2.sort_values("fecha")

    df_base = crear_base_diaria_co2()

    df_base = df_base.merge(
        df_co2,
        on="fecha",
        how="left",
    )

    # Los mercados financieros no cotizan todos los días.
    # Para fines de semana y festivos se usa el último precio disponible.
    df_base["precio_co2_EUR_tCO2"] = df_base["precio_co2_EUR_tCO2"].ffill()

    # Si el primer día del histórico no tuviera cotización, se rellena hacia atrás.
    df_base["precio_co2_EUR_tCO2"] = df_base["precio_co2_EUR_tCO2"].bfill()

    df_base["fuente"] = df_base["fuente"].fillna("Investing")
    df_base["observaciones"] = df_base["observaciones"].fillna(
        "Precio rellenado con última cotización disponible"
    )
    df_base["archivo_origen"] = df_base["archivo_origen"].fillna("relleno")

    df_base["fecha"] = df_base["fecha"].dt.date

    try:
        df_base.to_excel(RUTA_CO2_MANUAL, index=False)

    except PermissionError:
        raise PermissionError(
            f"No se puede guardar {RUTA_CO2_MANUAL}. "
            f"Probablemente está abierto en Excel. Ciérralo y vuelve a ejecutar."
        )

    dias_con_precio = df_base["precio_co2_EUR_tCO2"].notna().sum()
    dias_sin_precio = df_base["precio_co2_EUR_tCO2"].isna().sum()

    print(f"Archivo co2_manual.xlsx generado en: {RUTA_CO2_MANUAL}")
    print(f"Días totales: {len(df_base)}")
    print(f"Días con precio CO2: {dias_con_precio}")
    print(f"Días sin precio CO2: {dias_sin_precio}")

    if dias_con_precio > 0:
        print(f"Precio CO2 medio: {df_base['precio_co2_EUR_tCO2'].mean():.2f} €/tCO2")

    print("-" * 50)