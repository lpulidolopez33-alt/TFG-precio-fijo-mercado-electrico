import unicodedata

import pandas as pd
import requests

from config import (
    FECHA_INICIO_HISTORICO,
    FECHA_FIN_HISTORICO,
    RAW_DATA_DIR,
    PROCESSED_DATA_DIR,
    REE_BASE_URL,
    REE_GEO_TRUNC,
    REE_GEO_LIMIT,
    REE_GEO_IDS,
    MODO_PRUEBA_GENERACION_REE,
    DIAS_PRUEBA_GENERACION_REE,
)


REE_GENERACION_RAW_DIR = RAW_DATA_DIR / "ree_generacion"


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


def normalizar_texto(texto: str) -> str:
    """
    Normaliza un texto para compararlo sin problemas de mayúsculas o acentos.
    """

    texto = str(texto).strip().lower()

    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(
        caracter for caracter in texto
        if not unicodedata.combining(caracter)
    )

    return texto


def mapear_tecnologia_a_columna(titulo_serie: str) -> str | None:
    """
    Convierte el nombre de una tecnología de REE en el nombre de columna del modelo.
    """

    titulo = normalizar_texto(titulo_serie)

    if "solar fotovoltaica" in titulo:
        return "generacion_solar_fotovoltaica"

    if "solar termica" in titulo:
        return "generacion_solar_termica"

    if "eolica" in titulo:
        return "generacion_eolica"

    if titulo == "hidraulica":
        return "generacion_hidraulica"

    if "nuclear" in titulo:
        return "generacion_nuclear"

    if "ciclo combinado" in titulo:
        return "generacion_ciclo_combinado"

    if "carbon" in titulo:
        return "generacion_carbon"

    return None


def obtener_rango_fechas_generacion_ree() -> tuple[pd.Timestamp, pd.Timestamp]:
    """
    Define el rango de fechas para descargar generación de REE.
    """

    fecha_fin = pd.Timestamp(FECHA_FIN_HISTORICO)

    if MODO_PRUEBA_GENERACION_REE:
        fecha_inicio = fecha_fin - pd.Timedelta(days=DIAS_PRUEBA_GENERACION_REE - 1)
    else:
        fecha_inicio = pd.Timestamp(FECHA_INICIO_HISTORICO)

    return fecha_inicio, fecha_fin


def crear_periodos_mensuales(
    fecha_inicio: pd.Timestamp,
    fecha_fin: pd.Timestamp,
) -> list[tuple[pd.Timestamp, pd.Timestamp]]:
    """
    Divide un rango de fechas en periodos mensuales.
    """

    periodos = []
    inicio_periodo = fecha_inicio

    while inicio_periodo <= fecha_fin:
        fin_mes = inicio_periodo + pd.offsets.MonthEnd(0)
        fin_periodo = min(fin_mes, fecha_fin)

        periodos.append((inicio_periodo, fin_periodo))

        inicio_periodo = fin_periodo + pd.Timedelta(days=1)

    return periodos


def construir_url_generacion_ree() -> str:
    """
    Construye la URL del endpoint de balance eléctrico de REE.

    Usamos balance/balance-electrico porque permite obtener la generación
    horaria desagregada por tecnologías.
    """

    return f"{REE_BASE_URL}/balance/balance-electrico"


def descargar_generacion_ree_periodo(
    fecha_inicio: pd.Timestamp,
    fecha_fin: pd.Timestamp,
) -> dict:
    """
    Descarga la generación por tecnología para un periodo concreto.
    """

    REE_GENERACION_RAW_DIR.mkdir(parents=True, exist_ok=True)

    url = construir_url_generacion_ree()

    parametros = {
        "start_date": fecha_inicio.strftime("%Y-%m-%dT00:00"),
        "end_date": fecha_fin.strftime("%Y-%m-%dT23:59"),
        "time_trunc": "hour",
        "geo_trunc": REE_GEO_TRUNC,
        "geo_limit": REE_GEO_LIMIT,
        "geo_ids": REE_GEO_IDS,
    }

    respuesta = requests.get(
        url,
        params=parametros,
        timeout=60,
        headers={
            "Accept": "application/json",
            "Content-Type": "application/json",
            "User-Agent": "Mozilla/5.0",
        },
    )

    if respuesta.status_code != 200:
        raise RuntimeError(
            f"Error REE {respuesta.status_code}: {respuesta.text[:500]}"
        )

    return respuesta.json()


def convertir_serie_generacion_a_dataframe(
    serie: dict,
) -> pd.DataFrame | None:
    """
    Convierte una serie de generación de REE en una tabla con una sola tecnología.
    """

    atributos = serie.get("attributes", {})

    titulo_serie = atributos.get("title", "")
    columna_modelo = mapear_tecnologia_a_columna(titulo_serie)

    if columna_modelo is None:
        return None

    valores = atributos.get("values", [])

    if not valores:
        return None

    df = pd.DataFrame(valores)

    df = df.rename(
        columns={
            "value": columna_modelo,
            "datetime": "fecha_hora_ree",
        }
    )

    df[columna_modelo] = pd.to_numeric(
        df[columna_modelo],
        errors="coerce",
    )

    df["fecha_hora_utc_dt"] = pd.to_datetime(
        df["fecha_hora_ree"],
        utc=True,
    )

    df["fecha_hora_utc"] = df["fecha_hora_utc_dt"].astype(str)

    df = df[
        [
            "fecha_hora_utc",
            columna_modelo,
        ]
    ]

    return df


def combinar_series_generacion(tablas_series: list[pd.DataFrame]) -> pd.DataFrame:
    """
    Combina varias tecnologías de generación en una única tabla horaria.
    """

    if not tablas_series:
        raise ValueError("No se ha encontrado ninguna serie de generación útil.")

    df_generacion = tablas_series[0]

    for tabla in tablas_series[1:]:
        df_generacion = df_generacion.merge(
            tabla,
            on="fecha_hora_utc",
            how="outer",
        )

    return df_generacion


def completar_columnas_generacion(df_generacion: pd.DataFrame) -> pd.DataFrame:
    """
    Asegura que existan todas las columnas de generación necesarias.
    """

    df_generacion = df_generacion.copy()

    for columna in COLUMNAS_GENERACION:
        if columna not in df_generacion.columns:
            df_generacion[columna] = 0.0

    df_generacion["generacion_solar"] = (
        df_generacion["generacion_solar_fotovoltaica"].fillna(0)
        + df_generacion["generacion_solar_termica"].fillna(0)
    )

    df_generacion = df_generacion[
        ["fecha_hora_utc"] + COLUMNAS_GENERACION
    ]

    df_generacion = df_generacion.sort_values("fecha_hora_utc")

    return df_generacion

def extraer_series_generacion_utiles(datos_json: dict) -> list[dict]:
    """
    Extrae todas las series de generación útiles desde la respuesta de REE.

    En el endpoint balance/balance-electrico, las tecnologías pueden venir
    dentro de bloques anidados como Renovable y No renovable.
    """

    series_utiles = []

    def recorrer_serie(serie: dict) -> None:
        atributos = serie.get("attributes", {})

        titulo = atributos.get("title", "")
        valores = atributos.get("values", [])

        columna_modelo = mapear_tecnologia_a_columna(titulo)

        if columna_modelo is not None and valores:
            series_utiles.append(serie)

        contenido = atributos.get("content", [])

        for subserie in contenido:
            recorrer_serie(subserie)

    for serie in datos_json.get("included", []):
        recorrer_serie(serie)

    return series_utiles

def convertir_json_generacion_a_dataframe(datos_json: dict) -> pd.DataFrame:
    """
    Convierte la respuesta JSON de REE en una tabla horaria de generación.
    """

    series = extraer_series_generacion_utiles(datos_json)

    if not series:
        raise ValueError(
            "La respuesta de REE no contiene series de generación útiles."
        )

    tablas_series = []

    for serie in series:
        df_serie = convertir_serie_generacion_a_dataframe(serie)

        if df_serie is not None:
            tablas_series.append(df_serie)

    df_generacion = combinar_series_generacion(tablas_series)
    df_generacion = completar_columnas_generacion(df_generacion)

    return df_generacion


def descargar_historico_generacion_ree() -> None:
    """
    Descarga la generación horaria por tecnología desde REE.
    """

    fecha_inicio, fecha_fin = obtener_rango_fechas_generacion_ree()

    periodos = crear_periodos_mensuales(
        fecha_inicio=fecha_inicio,
        fecha_fin=fecha_fin,
    )

    tablas = []
    errores = []

    print("\nDESCARGA DE GENERACIÓN REE")
    print("-" * 50)
    print(f"Fecha inicial generación REE: {fecha_inicio.date()}")
    print(f"Fecha final generación REE: {fecha_fin.date()}")
    print(f"Número de periodos a descargar: {len(periodos)}")
    print("-" * 50)

    for inicio_periodo, fin_periodo in periodos:
        try:
            datos_json = descargar_generacion_ree_periodo(
                fecha_inicio=inicio_periodo,
                fecha_fin=fin_periodo,
            )

            df_periodo = convertir_json_generacion_a_dataframe(datos_json)
            tablas.append(df_periodo)

            print(
                f"OK {inicio_periodo.date()} a {fin_periodo.date()} "
                f"- {len(df_periodo)} registros"
            )

        except Exception as error:
            errores.append(
                {
                    "fecha_inicio": inicio_periodo.strftime("%Y-%m-%d"),
                    "fecha_fin": fin_periodo.strftime("%Y-%m-%d"),
                    "error": str(error),
                }
            )

            print(
                f"ERROR {inicio_periodo.date()} a {fin_periodo.date()} "
                f"- {error}"
            )

    if not tablas:
        raise RuntimeError("No se ha podido descargar ningún dato de generación REE.")

    df_generacion = pd.concat(tablas, ignore_index=True)

    df_generacion = df_generacion.drop_duplicates(
        subset=["fecha_hora_utc"],
        keep="last",
    )

    df_generacion = df_generacion.sort_values("fecha_hora_utc")

    ruta_salida = PROCESSED_DATA_DIR / "generacion_ree_horaria.xlsx"

    try:
        df_generacion.to_excel(ruta_salida, index=False)

    except PermissionError:
        raise PermissionError(
            f"No se puede guardar el archivo {ruta_salida}. "
            f"Probablemente está abierto en Excel. Ciérralo y vuelve a ejecutar."
        )

    print("-" * 50)
    print(f"Archivo generación REE guardado: {ruta_salida}")
    print(f"Número de filas horarias generación REE: {len(df_generacion)}")

    if errores:
        ruta_errores = PROCESSED_DATA_DIR / "errores_descarga_generacion_ree.xlsx"
        pd.DataFrame(errores).to_excel(ruta_errores, index=False)

        print(f"Periodos con error guardados en: {ruta_errores}")