import pandas as pd
import matplotlib.pyplot as plt

from config import OUTPUT_DIR, GRAPH_DIR


def cargar_prevision_horaria() -> pd.DataFrame:
    """
    Carga la previsión horaria futura generada por el modelo.
    """

    ruta_prevision = OUTPUT_DIR / "prevision_commodity_modelo.xlsx"

    if not ruta_prevision.exists():
        raise FileNotFoundError(f"No existe el archivo: {ruta_prevision}")

    df = pd.read_excel(
        ruta_prevision,
        sheet_name="prevision_horaria",
    )

    return df


def cargar_commodity_periodos() -> pd.DataFrame:
    """
    Carga la tabla de commodity por periodos tarifarios.
    """

    ruta_periodos = OUTPUT_DIR / "commodity_por_periodos.xlsx"

    if not ruta_periodos.exists():
        raise FileNotFoundError(f"No existe el archivo: {ruta_periodos}")

    df = pd.read_excel(
        ruta_periodos,
        sheet_name="tabla_calculadora",
    )

    return df


def preparar_prevision(df: pd.DataFrame) -> pd.DataFrame:
    """
    Prepara fechas y columnas numéricas de la previsión.
    """

    df = df.copy()

    df["fecha"] = pd.to_datetime(df["fecha"])
    df["precio_omie_previsto"] = pd.to_numeric(
        df["precio_omie_previsto"],
        errors="coerce",
    )

    df["hora"] = pd.to_numeric(df["hora"], errors="coerce")

    return df


def graficar_evolucion_mensual_prevision(df: pd.DataFrame) -> None:
    """
    Genera una gráfica mensual del precio OMIE previsto.
    """

    GRAPH_DIR.mkdir(parents=True, exist_ok=True)

    df = df.copy()
    df["año_mes"] = df["fecha"].dt.to_period("M").astype(str)

    resumen_mensual = (
        df.groupby("año_mes", as_index=False)
        .agg(
            precio_medio_previsto=("precio_omie_previsto", "mean"),
            precio_minimo_previsto=("precio_omie_previsto", "min"),
            precio_maximo_previsto=("precio_omie_previsto", "max"),
            horas=("precio_omie_previsto", "count"),
        )
    )

    plt.figure(figsize=(12, 5))
    plt.plot(
        resumen_mensual["año_mes"],
        resumen_mensual["precio_medio_previsto"],
        marker="o",
    )
    plt.xticks(rotation=90)
    plt.title("Evolución mensual de la commodity prevista")
    plt.xlabel("Mes")
    plt.ylabel("Precio previsto [€/MWh]")
    plt.tight_layout()

    ruta_grafica = GRAPH_DIR / "evolucion_mensual_prevision_commodity.png"
    plt.savefig(ruta_grafica, dpi=300)
    plt.close()

    print(f"Gráfica guardada: {ruta_grafica}")


def graficar_perfil_horario_prevision(df: pd.DataFrame) -> None:
    """
    Genera el perfil horario medio del precio previsto.
    """

    resumen_horario = (
        df.groupby("hora", as_index=False)
        .agg(
            precio_medio_previsto=("precio_omie_previsto", "mean"),
            precio_minimo_previsto=("precio_omie_previsto", "min"),
            precio_maximo_previsto=("precio_omie_previsto", "max"),
            horas=("precio_omie_previsto", "count"),
        )
    )

    plt.figure(figsize=(10, 5))
    plt.plot(
        resumen_horario["hora"],
        resumen_horario["precio_medio_previsto"],
        marker="o",
    )
    plt.xticks(range(0, 24))
    plt.title("Perfil horario medio de la commodity prevista")
    plt.xlabel("Hora del día")
    plt.ylabel("Precio previsto [€/MWh]")
    plt.tight_layout()

    ruta_grafica = GRAPH_DIR / "perfil_horario_prevision_commodity.png"
    plt.savefig(ruta_grafica, dpi=300)
    plt.close()

    print(f"Gráfica guardada: {ruta_grafica}")


def graficar_commodity_por_periodos(df_periodos: pd.DataFrame) -> None:
    """
    Genera una gráfica comparativa de commodity por periodos tarifarios.
    """

    df = df_periodos.copy()

    columnas_periodos = ["P1", "P2", "P3", "P4", "P5", "P6"]

    df_largo = df.melt(
        id_vars="tarifa",
        value_vars=columnas_periodos,
        var_name="periodo",
        value_name="commodity_EUR_MWh",
    )

    df_largo = df_largo.dropna(subset=["commodity_EUR_MWh"])

    etiquetas = (
        df_largo["tarifa"].astype(str)
        + " - "
        + df_largo["periodo"].astype(str)
    )

    plt.figure(figsize=(12, 5))
    plt.bar(
        etiquetas,
        df_largo["commodity_EUR_MWh"],
    )
    plt.xticks(rotation=90)
    plt.title("Commodity prevista por tarifa y periodo")
    plt.xlabel("Tarifa y periodo")
    plt.ylabel("Commodity prevista [€/MWh]")
    plt.tight_layout()

    ruta_grafica = GRAPH_DIR / "commodity_por_periodos.png"
    plt.savefig(ruta_grafica, dpi=300)
    plt.close()

    print(f"Gráfica guardada: {ruta_grafica}")


def generar_graficas_prevision() -> None:
    """
    Genera todas las gráficas asociadas a la previsión futura.
    """

    print("\nGRÁFICAS DE PREVISIÓN FUTURA")
    print("-" * 50)

    df_prevision = cargar_prevision_horaria()
    df_prevision = preparar_prevision(df_prevision)

    df_periodos = cargar_commodity_periodos()

    graficar_evolucion_mensual_prevision(df_prevision)
    graficar_perfil_horario_prevision(df_prevision)
    graficar_commodity_por_periodos(df_periodos)

    print("Gráficas de previsión generadas correctamente.")
    print("-" * 50)