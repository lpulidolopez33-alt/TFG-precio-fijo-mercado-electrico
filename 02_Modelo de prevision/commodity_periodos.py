import pandas as pd
import holidays

from config import OUTPUT_DIR


def cargar_prevision_horaria() -> pd.DataFrame:
    """
    Carga la previsión horaria generada por el modelo.
    """

    ruta_prevision = OUTPUT_DIR / "prevision_commodity_modelo.xlsx"

    if not ruta_prevision.exists():
        raise FileNotFoundError(f"No existe el archivo de previsión: {ruta_prevision}")

    df = pd.read_excel(
        ruta_prevision,
        sheet_name="prevision_horaria",
    )

    return df


def preparar_fechas(df: pd.DataFrame) -> pd.DataFrame:
    """
    Prepara las columnas de fecha para poder calcular periodos tarifarios.
    """

    df = df.copy()

    df["fecha"] = pd.to_datetime(df["fecha"])
    df["año"] = df["fecha"].dt.year
    df["mes"] = df["fecha"].dt.month
    df["dia_semana"] = df["fecha"].dt.dayofweek
    df["hora"] = pd.to_numeric(df["hora"], errors="coerce").astype(int)

    festivos_espana = holidays.Spain(years=df["año"].unique())

    df["es_festivo_nacional"] = df["fecha"].dt.date.apply(
        lambda fecha: 1 if fecha in festivos_espana else 0
    )

    df["es_fin_semana"] = df["dia_semana"].isin([5, 6]).astype(int)

    return df


def calcular_periodo_20td(fila: pd.Series) -> str:
    """
    Calcula el periodo energético de la tarifa 2.0TD.

    P1: punta
    P2: llano
    P3: valle
    """

    hora = fila["hora"]

    if fila["es_fin_semana"] == 1 or fila["es_festivo_nacional"] == 1:
        return "P3"

    if 0 <= hora < 8:
        return "P3"

    if 10 <= hora < 14 or 18 <= hora < 22:
        return "P1"

    return "P2"


def obtener_bloques_periodos_seis_periodos(mes: int) -> tuple[str, str]:
    """
    Devuelve los periodos de las horas punta y llano para tarifas de seis periodos.

    Para días laborables:
    - bloque caro: 09-14 y 18-22
    - bloque intermedio: 08-09, 14-18 y 22-24
    - 00-08 siempre P6
    """

    temporada_alta = [1, 2, 7, 12]
    temporada_media_alta = [3, 11]
    temporada_media = [6, 8, 9]
    temporada_baja = [4, 5, 10]

    if mes in temporada_alta:
        return "P1", "P2"

    if mes in temporada_media_alta:
        return "P2", "P3"

    if mes in temporada_media:
        return "P3", "P4"

    if mes in temporada_baja:
        return "P4", "P5"

    raise ValueError(f"Mes no reconocido: {mes}")


def calcular_periodo_seis_periodos_peninsula(fila: pd.Series) -> str:
    """
    Calcula el periodo energético para tarifas 3.0TD y 6.XTD en península.
    """

    hora = fila["hora"]
    mes = fila["mes"]

    if fila["es_fin_semana"] == 1 or fila["es_festivo_nacional"] == 1:
        return "P6"

    if 0 <= hora < 8:
        return "P6"

    periodo_caro, periodo_intermedio = obtener_bloques_periodos_seis_periodos(mes)

    if 9 <= hora < 14 or 18 <= hora < 22:
        return periodo_caro

    return periodo_intermedio


def asignar_periodos_tarifarios(df: pd.DataFrame) -> pd.DataFrame:
    """
    Añade columnas de periodo tarifario a la previsión horaria.
    """

    df = df.copy()

    df["periodo_20td"] = df.apply(calcular_periodo_20td, axis=1)

    df["periodo_30td_peninsula"] = df.apply(
        calcular_periodo_seis_periodos_peninsula,
        axis=1,
    )

    df["periodo_6xtd_peninsula"] = df["periodo_30td_peninsula"]

    return df


def crear_resumen_periodos(df: pd.DataFrame) -> pd.DataFrame:
    """
    Calcula el precio medio previsto por tarifa y periodo.
    """

    configuraciones = [
        {
            "tarifa": "2.0TD",
            "columna_periodo": "periodo_20td",
        },
        {
            "tarifa": "3.0TD_peninsula",
            "columna_periodo": "periodo_30td_peninsula",
        },
        {
            "tarifa": "6.XTD_peninsula",
            "columna_periodo": "periodo_6xtd_peninsula",
        },
    ]

    tablas = []

    for configuracion in configuraciones:
        tarifa = configuracion["tarifa"]
        columna_periodo = configuracion["columna_periodo"]

        resumen = (
            df.groupby(columna_periodo, as_index=False)
            .agg(
                commodity_modelo_EUR_MWh=("precio_omie_previsto", "mean"),
                precio_minimo_previsto=("precio_omie_previsto", "min"),
                precio_maximo_previsto=("precio_omie_previsto", "max"),
                horas=("precio_omie_previsto", "count"),
            )
            .rename(columns={columna_periodo: "periodo"})
        )

        resumen.insert(0, "tarifa", tarifa)

        tablas.append(resumen)

    df_resumen = pd.concat(tablas, ignore_index=True)

    df_resumen["periodo_num"] = df_resumen["periodo"].str.replace("P", "").astype(int)

    df_resumen = df_resumen.sort_values(
        ["tarifa", "periodo_num"]
    ).drop(columns=["periodo_num"])

    return df_resumen


def crear_tabla_calculadora(df_resumen: pd.DataFrame) -> pd.DataFrame:
    """
    Crea una tabla en formato ancho, más cómoda para copiar al Excel.
    """

    tabla = df_resumen.pivot_table(
        index="tarifa",
        columns="periodo",
        values="commodity_modelo_EUR_MWh",
        aggfunc="mean",
    ).reset_index()

    columnas_periodos = ["P1", "P2", "P3", "P4", "P5", "P6"]

    for columna in columnas_periodos:
        if columna not in tabla.columns:
            tabla[columna] = None

    tabla = tabla[["tarifa"] + columnas_periodos]

    return tabla


def guardar_commodity_por_periodos(
    df_horaria: pd.DataFrame,
    df_resumen: pd.DataFrame,
    tabla_calculadora: pd.DataFrame,
) -> None:
    """
    Guarda la commodity por periodos en un Excel.
    """

    ruta_salida = OUTPUT_DIR / "commodity_por_periodos.xlsx"

    try:
        with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
            tabla_calculadora.to_excel(
                writer,
                sheet_name="tabla_calculadora",
                index=False,
            )

            df_resumen.to_excel(
                writer,
                sheet_name="resumen_periodos",
                index=False,
            )

            df_horaria.to_excel(
                writer,
                sheet_name="prevision_horaria_periodos",
                index=False,
            )

    except PermissionError:
        raise PermissionError(
            f"No se puede guardar el archivo {ruta_salida}. "
            f"Probablemente está abierto en Excel. Ciérralo y vuelve a ejecutar."
        )

    print(f"Commodity por periodos guardada en: {ruta_salida}")


def generar_commodity_por_periodos() -> None:
    """
    Genera la commodity prevista por periodos tarifarios.
    """

    print("\nCOMMODITY POR PERIODOS TARIFARIOS")
    print("-" * 50)

    df = cargar_prevision_horaria()
    df = preparar_fechas(df)
    df = asignar_periodos_tarifarios(df)

    df_resumen = crear_resumen_periodos(df)
    tabla_calculadora = crear_tabla_calculadora(df_resumen)

    guardar_commodity_por_periodos(
        df_horaria=df,
        df_resumen=df_resumen,
        tabla_calculadora=tabla_calculadora,
    )

    print(tabla_calculadora)
    print("Commodity por periodos generada correctamente.")
    print("-" * 50)