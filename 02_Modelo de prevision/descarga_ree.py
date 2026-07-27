import pandas as pd
import requests

from config import (
    FECHA_INICIO_HISTORICO,
    FECHA_FIN_HISTORICO,
    RAW_DATA_DIR,
    PROCESSED_DATA_DIR,
    MODO_PRUEBA_REE,
    DIAS_PRUEBA_REE,
    REE_BASE_URL,
    REE_GEO_TRUNC,
    REE_GEO_LIMIT,
    REE_GEO_IDS,
)


REE_RAW_DIR = RAW_DATA_DIR / "ree"


def obtener_rango_fechas_ree() -> tuple[pd.Timestamp, pd.Timestamp]:
    """
    Define el rango de fechas para descargar datos de REE.
    """

    fecha_fin = pd.Timestamp(FECHA_FIN_HISTORICO)

    if MODO_PRUEBA_REE:
        fecha_inicio = fecha_fin - pd.Timedelta(days=DIAS_PRUEBA_REE - 1)
    else:
        fecha_inicio = pd.Timestamp(FECHA_INICIO_HISTORICO)

    return fecha_inicio, fecha_fin


def crear_periodos_mensuales(
    fecha_inicio: pd.Timestamp,
    fecha_fin: pd.Timestamp,
) -> list[tuple[pd.Timestamp, pd.Timestamp]]:
    """
    Divide un rango largo de fechas en tramos mensuales.

    Esto evita pedir demasiados datos a la API en una única llamada.
    """

    periodos = []

    inicio_periodo = fecha_inicio

    while inicio_periodo <= fecha_fin:
        fin_mes = inicio_periodo + pd.offsets.MonthEnd(0)
        fin_periodo = min(fin_mes, fecha_fin)

        periodos.append((inicio_periodo, fin_periodo))

        inicio_periodo = fin_periodo + pd.Timedelta(days=1)

    return periodos


def construir_url_demanda_ree() -> str:
    """
    Construye la URL base para la demanda eléctrica de REE.
    """

    return f"{REE_BASE_URL}/demanda/evolucion"


def descargar_demanda_ree_periodo(
    fecha_inicio: pd.Timestamp,
    fecha_fin: pd.Timestamp,
) -> dict:
    """
    Descarga la demanda eléctrica para un periodo concreto.
    """

    REE_RAW_DIR.mkdir(parents=True, exist_ok=True)

    url = construir_url_demanda_ree()

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


def seleccionar_serie_demanda_real(datos_json: dict) -> dict:
    """
    Busca dentro de la respuesta de REE la serie de demanda real.

    El widget puede traer varias series: demanda real, prevista o programada.
    Para el modelo nos interesa la demanda real histórica.
    """

    series = datos_json.get("included", [])

    if not series:
        raise ValueError("La respuesta de REE no contiene series de datos.")

    for serie in series:
        atributos = serie.get("attributes", {})
        titulo = atributos.get("title", "").lower()

        if "real" in titulo:
            return serie

    return series[0]


def convertir_json_demanda_a_dataframe(datos_json: dict) -> pd.DataFrame:
    """
    Convierte la respuesta JSON de REE en una tabla de pandas.
    """

    serie_demanda = seleccionar_serie_demanda_real(datos_json)

    atributos = serie_demanda.get("attributes", {})
    titulo_serie = atributos.get("title", "demanda")
    valores = atributos.get("values", [])

    if not valores:
        raise ValueError("La serie de demanda no contiene valores.")

    df = pd.DataFrame(valores)

    df = df.rename(
        columns={
            "value": "demanda",
            "datetime": "fecha_hora_ree",
        }
    )

    df["demanda"] = pd.to_numeric(df["demanda"], errors="coerce")

    df["fecha_hora_utc_dt"] = pd.to_datetime(
        df["fecha_hora_ree"],
        utc=True,
    )

    df["fecha_hora_utc"] = df["fecha_hora_utc_dt"].astype(str)

    df["serie_origen_ree"] = titulo_serie

    df = df[
        [
            "fecha_hora_utc",
            "demanda",
            "serie_origen_ree",
        ]
    ]

    return df


def descargar_historico_demanda_ree() -> None:
    """
    Descarga la demanda eléctrica peninsular desde REE y la guarda en Excel.
    """

    fecha_inicio, fecha_fin = obtener_rango_fechas_ree()

    periodos = crear_periodos_mensuales(
        fecha_inicio=fecha_inicio,
        fecha_fin=fecha_fin,
    )

    tablas = []
    errores = []

    print("\nDESCARGA DE DEMANDA REE")
    print("-" * 50)
    print(f"Fecha inicial REE: {fecha_inicio.date()}")
    print(f"Fecha final REE: {fecha_fin.date()}")
    print(f"Número de periodos a descargar: {len(periodos)}")
    print("-" * 50)

    for inicio_periodo, fin_periodo in periodos:
        try:
            datos_json = descargar_demanda_ree_periodo(                      # se conecta a la API de REE y pide la demanda horaria para un periodo concreto.
                fecha_inicio=inicio_periodo,
                fecha_fin=fin_periodo,
            )

            df_periodo = convertir_json_demanda_a_dataframe(datos_json)
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
        raise RuntimeError("No se ha podido descargar ningún dato de demanda REE.")

    df_demanda = pd.concat(tablas, ignore_index=True)

    df_demanda = df_demanda.drop_duplicates(
        subset=["fecha_hora_utc"],
        keep="last",
    )

    df_demanda = df_demanda.sort_values("fecha_hora_utc")

    ruta_salida = PROCESSED_DATA_DIR / "demanda_ree_horaria.xlsx"

    try:
        df_demanda.to_excel(ruta_salida, index=False)

    except PermissionError:
        raise PermissionError(
            f"No se puede guardar el archivo {ruta_salida}. "
            f"Probablemente está abierto en Excel. Ciérralo y vuelve a ejecutar."
        )

    print("-" * 50)
    print(f"Archivo demanda REE guardado: {ruta_salida}")
    print(f"Número de filas horarias REE: {len(df_demanda)}")

    if errores:
        ruta_errores = PROCESSED_DATA_DIR / "errores_descarga_ree.xlsx"
        pd.DataFrame(errores).to_excel(ruta_errores, index=False)

        print(f"Periodos con error guardados en: {ruta_errores}")