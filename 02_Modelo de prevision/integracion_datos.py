import pandas as pd

from config import (
    PROCESSED_DATA_DIR,
    USAR_IMPORTACION_GENERACION_EXCEL,
    USAR_PRECIO_GAS_MANUAL,
    USAR_PRECIO_CO2_MANUAL,
)


COLUMNAS_GENERACION = [
    "generacion_solar_fotovoltaica",
    "generacion_solar_termica",
    "generacion_solar",
    "generacion_eolica",
    "generacion_hidraulica",
    "generacion_nuclear",
    "generacion_ciclo_combinado",
    "generacion_carbon",
]


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


def cargar_generacion_manual_procesada() -> pd.DataFrame:
    """
    Carga el archivo de generación manual procesada.
    """

    ruta_generacion = PROCESSED_DATA_DIR / "generacion_manual_procesada.xlsx"

    if not ruta_generacion.exists():
        raise FileNotFoundError(
            f"No existe el archivo de generación manual procesada: {ruta_generacion}"
        )

    df_generacion = pd.read_excel(ruta_generacion)

    return df_generacion

def cargar_gas_manual_procesado() -> pd.DataFrame:
    """
    Carga el archivo de precio del gas procesado en formato horario.
    """

    ruta_gas = PROCESSED_DATA_DIR / "gas_manual_procesado.xlsx"

    if not ruta_gas.exists():
        raise FileNotFoundError(
            f"No existe el archivo de gas manual procesado: {ruta_gas}"
        )

    df_gas = pd.read_excel(ruta_gas)

    return df_gas

def cargar_co2_manual_procesado() -> pd.DataFrame:
    """
    Carga el archivo de precio del CO2 procesado en formato horario.
    """

    ruta_co2 = PROCESSED_DATA_DIR / "co2_manual_procesado.xlsx"

    if not ruta_co2.exists():
        raise FileNotFoundError(
            f"No existe el archivo de CO2 manual procesado: {ruta_co2}"
        )

    df_co2 = pd.read_excel(ruta_co2)

    return df_co2


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


def integrar_generacion_manual(
    df_base: pd.DataFrame,
    df_generacion: pd.DataFrame,
) -> pd.DataFrame:
    """
    Integra las columnas de generación manual dentro de la base histórica.
    """

    columnas_generacion = ["fecha_hora_utc"] + COLUMNAS_GENERACION

    df_generacion_reducido = df_generacion[columnas_generacion].copy()

    df_base = df_base.drop(
        columns=COLUMNAS_GENERACION,
        errors="ignore",
    )

    df_base = df_base.merge(
        df_generacion_reducido,
        on="fecha_hora_utc",
        how="left",
    )

    return df_base

def integrar_gas_manual(
    df_base: pd.DataFrame,
    df_gas: pd.DataFrame,
) -> pd.DataFrame:
    """
    Integra el precio del gas diario convertido a horario dentro de la base histórica.
    """

    columnas_gas = [
        "fecha_hora_utc",
        "precio_gas_EUR_MWh",
    ]

    df_gas_reducido = df_gas[columnas_gas].copy()

    df_base = df_base.drop(
        columns=["precio_gas_EUR_MWh"],
        errors="ignore",
    )

    df_base = df_base.merge(
        df_gas_reducido,
        on="fecha_hora_utc",
        how="left",
    )

    return df_base

def integrar_co2_manual(
    df_base: pd.DataFrame,
    df_co2: pd.DataFrame,
) -> pd.DataFrame:
    """
    Integra el precio del CO2 diario convertido a horario dentro de la base histórica.
    """

    columnas_co2 = [
        "fecha_hora_utc",
        "precio_co2_EUR_tCO2",
    ]

    df_co2_reducido = df_co2[columnas_co2].copy()

    df_base = df_base.drop(
        columns=["precio_co2_EUR_tCO2"],
        errors="ignore",
    )

    df_base = df_base.merge(
        df_co2_reducido,
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

    print("\nGENERACIÓN MANUAL")

    for columna in COLUMNAS_GENERACION:
        if columna in df_base.columns:
            datos_disponibles = df_base[columna].notna().sum()
            print(f"{columna}: {datos_disponibles} datos disponibles")
        else:
            print(f"{columna}: columna no integrada")

            print("\nPRECIO DEL GAS")

    if "precio_gas_EUR_MWh" in df_base.columns:
        datos_gas = df_base["precio_gas_EUR_MWh"].notna().sum()
        datos_gas_faltantes = df_base["precio_gas_EUR_MWh"].isna().sum()

        print(f"Horas con precio de gas: {datos_gas}")
        print(f"Horas sin precio de gas: {datos_gas_faltantes}")

    if datos_gas > 0:
        precio_gas_medio = df_base["precio_gas_EUR_MWh"].mean()
        print(f"Precio medio gas: {precio_gas_medio:.2f} €/MWh")
    else:
        print("Columna precio_gas_EUR_MWh no integrada")

    print("\nPRECIO DEL CO2")

    if "precio_co2_EUR_tCO2" in df_base.columns:
        datos_co2 = df_base["precio_co2_EUR_tCO2"].notna().sum()
        datos_co2_faltantes = df_base["precio_co2_EUR_tCO2"].isna().sum()

        print(f"Horas con precio de CO2: {datos_co2}")
        print(f"Horas sin precio de CO2: {datos_co2_faltantes}")

        if datos_co2 > 0:
            precio_co2_medio = df_base["precio_co2_EUR_tCO2"].mean()
            print(f"Precio medio CO2: {precio_co2_medio:.2f} €/tCO2")
    else:
        print("Columna precio_co2_EUR_tCO2 no integrada")


    print("-" * 50)


def guardar_base_modelo() -> None:
    """
    Crea la base de datos del modelo integrando:
    - plantilla histórica
    - precios OMIE
    - demanda REE
    - generación manual, si está activada
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

    if USAR_IMPORTACION_GENERACION_EXCEL:
        df_generacion = cargar_generacion_manual_procesada()
        df_base = integrar_generacion_manual(
            df_base=df_base,
            df_generacion=df_generacion,
        )
    else:
        print("\nGeneración manual no integrada porque está desactivada.")

    if USAR_PRECIO_GAS_MANUAL:
        df_gas = cargar_gas_manual_procesado()
        df_base = integrar_gas_manual(
            df_base=df_base,
            df_gas=df_gas,
        )
    else:
        print("\nPrecio del gas manual no integrado porque está desactivado.")

    if USAR_PRECIO_CO2_MANUAL:
        df_co2 = cargar_co2_manual_procesado()
        df_base = integrar_co2_manual(
            df_base=df_base,
            df_co2=df_co2,
        )
    else:
        print("\nPrecio del CO2 manual no integrado porque está desactivado.")

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