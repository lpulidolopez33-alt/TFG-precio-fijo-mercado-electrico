import pandas as pd
import matplotlib.pyplot as plt

from config import PROCESSED_DATA_DIR, GRAPH_DIR


COLUMNAS_GENERACION = [
    "generacion_solar",
    "generacion_eolica",
    "generacion_hidraulica",
    "generacion_nuclear",
    "generacion_ciclo_combinado",
    "generacion_carbon",
]


def cargar_base_modelo() -> pd.DataFrame:
    """
    Carga la base del modelo con la generación ya integrada.
    """

    ruta_base = PROCESSED_DATA_DIR / "base_modelo.xlsx"

    if not ruta_base.exists():
        raise FileNotFoundError(f"No existe el archivo: {ruta_base}")

    df = pd.read_excel(ruta_base)

    return df


def preparar_base_generacion(df: pd.DataFrame) -> pd.DataFrame:
    """
    Prepara la base para analizar las columnas de generación.
    """

    df = df.copy()

    df["fecha"] = pd.to_datetime(df["fecha"])

    for columna in COLUMNAS_GENERACION:
        if columna in df.columns:
            df[columna] = pd.to_numeric(df[columna], errors="coerce")
        else:
            df[columna] = None

    return df


def crear_resumen_general_generacion(df: pd.DataFrame) -> pd.DataFrame:
    """
    Crea un resumen general de calidad para cada tecnología.
    """

    registros = []

    for columna in COLUMNAS_GENERACION:
        registros.append(
            {
                "variable": columna,
                "datos_disponibles": df[columna].notna().sum(),
                "datos_faltantes": df[columna].isna().sum(),
                "valor_minimo_MW": df[columna].min(),
                "valor_maximo_MW": df[columna].max(),
                "valor_medio_MW": df[columna].mean(),
                "valor_mediana_MW": df[columna].median(),
            }
        )

    return pd.DataFrame(registros)


def crear_resumen_diario_generacion(df: pd.DataFrame) -> pd.DataFrame:
    """
    Crea una tabla diaria de generación.

    Como la generación procede de datos diarios convertidos a MW medios,
    esta tabla evita repetir artificialmente el mismo dato 24 veces.
    """

    resumen_diario = (
        df.groupby("fecha", as_index=False)
        .agg(
            generacion_solar=("generacion_solar", "mean"),
            generacion_eolica=("generacion_eolica", "mean"),
            generacion_hidraulica=("generacion_hidraulica", "mean"),
            generacion_nuclear=("generacion_nuclear", "mean"),
            generacion_ciclo_combinado=("generacion_ciclo_combinado", "mean"),
            generacion_carbon=("generacion_carbon", "mean"),
        )
    )

    return resumen_diario


def crear_resumen_mensual_generacion(df_diario: pd.DataFrame) -> pd.DataFrame:
    """
    Calcula la generación media mensual por tecnología.
    """

    df = df_diario.copy()

    df["año_mes"] = df["fecha"].dt.to_period("M").astype(str)

    resumen_mensual = (
        df.groupby("año_mes", as_index=False)
        .agg(
            generacion_solar=("generacion_solar", "mean"),
            generacion_eolica=("generacion_eolica", "mean"),
            generacion_hidraulica=("generacion_hidraulica", "mean"),
            generacion_nuclear=("generacion_nuclear", "mean"),
            generacion_ciclo_combinado=("generacion_ciclo_combinado", "mean"),
            generacion_carbon=("generacion_carbon", "mean"),
            dias_con_dato=("fecha", "count"),
            primera_fecha=("fecha", "min"),
            ultima_fecha=("fecha", "max"),
        )
    )

    resumen_mensual["dias_mes"] = pd.to_datetime(
        resumen_mensual["año_mes"] + "-01"
    ).dt.days_in_month

    resumen_mensual["mes_completo"] = (
        resumen_mensual["dias_con_dato"] >= resumen_mensual["dias_mes"]
    ).astype(int)

    return resumen_mensual


def guardar_control_calidad_generacion(
    resumen_general: pd.DataFrame,
    resumen_diario: pd.DataFrame,
    resumen_mensual: pd.DataFrame,
) -> None:
    """
    Guarda los resúmenes de generación en un archivo Excel.
    """

    ruta_salida = PROCESSED_DATA_DIR / "control_calidad_generacion.xlsx"

    try:
        with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
            resumen_general.to_excel(writer, sheet_name="resumen_general", index=False)
            resumen_diario.to_excel(writer, sheet_name="resumen_diario", index=False)
            resumen_mensual.to_excel(writer, sheet_name="resumen_mensual", index=False)

    except PermissionError:
        raise PermissionError(
            f"No se puede guardar el archivo {ruta_salida}. "
            f"Probablemente está abierto en Excel. Ciérralo y vuelve a ejecutar."
        )

    print(f"Control de calidad de generación guardado en: {ruta_salida}")


def graficar_generacion_renovable(resumen_mensual: pd.DataFrame) -> None:
    """
    Genera una gráfica mensual de solar, eólica e hidráulica.
    """

    GRAPH_DIR.mkdir(parents=True, exist_ok=True)

    df_grafica = resumen_mensual[
        resumen_mensual["mes_completo"] == 1
    ].copy()

    plt.figure(figsize=(12, 5))
    plt.plot(df_grafica["año_mes"], df_grafica["generacion_solar"], marker="o", label="Solar")
    plt.plot(df_grafica["año_mes"], df_grafica["generacion_eolica"], marker="o", label="Eólica")
    plt.plot(df_grafica["año_mes"], df_grafica["generacion_hidraulica"], marker="o", label="Hidráulica")
    plt.xticks(rotation=90)
    plt.title("Evolución mensual de generación renovable")
    plt.xlabel("Mes")
    plt.ylabel("Generación media [MW]")
    plt.legend()
    plt.tight_layout()

    ruta_grafica = GRAPH_DIR / "evolucion_mensual_generacion_renovable.png"
    plt.savefig(ruta_grafica, dpi=300)
    plt.close()

    print(f"Gráfica guardada: {ruta_grafica}")


def graficar_generacion_termica(resumen_mensual: pd.DataFrame) -> None:
    """
    Genera una gráfica mensual de nuclear, ciclo combinado y carbón.
    """

    GRAPH_DIR.mkdir(parents=True, exist_ok=True)

    df_grafica = resumen_mensual[
        resumen_mensual["mes_completo"] == 1
    ].copy()

    plt.figure(figsize=(12, 5))
    plt.plot(df_grafica["año_mes"], df_grafica["generacion_nuclear"], marker="o", label="Nuclear")
    plt.plot(df_grafica["año_mes"], df_grafica["generacion_ciclo_combinado"], marker="o", label="Ciclo combinado")
    plt.plot(df_grafica["año_mes"], df_grafica["generacion_carbon"], marker="o", label="Carbón")
    plt.xticks(rotation=90)
    plt.title("Evolución mensual de generación térmica")
    plt.xlabel("Mes")
    plt.ylabel("Generación media [MW]")
    plt.legend()
    plt.tight_layout()

    ruta_grafica = GRAPH_DIR / "evolucion_mensual_generacion_termica.png"
    plt.savefig(ruta_grafica, dpi=300)
    plt.close()

    print(f"Gráfica guardada: {ruta_grafica}")

def graficar_generacion_todas(resumen_mensual: pd.DataFrame) -> None:
    """
    Genera una gráfica mensual con todas las tecnologías principales
    representadas en una única figura.
    """

    GRAPH_DIR.mkdir(parents=True, exist_ok=True)

    df_grafica = resumen_mensual[
        resumen_mensual["mes_completo"] == 1
    ].copy()

    plt.figure(figsize=(14, 6))

    plt.plot(
        df_grafica["año_mes"],
        df_grafica["generacion_solar"],
        marker="o",
        label="Solar",
    )

    plt.plot(
        df_grafica["año_mes"],
        df_grafica["generacion_eolica"],
        marker="o",
        label="Eólica",
    )

    plt.plot(
        df_grafica["año_mes"],
        df_grafica["generacion_hidraulica"],
        marker="o",
        label="Hidráulica",
    )

    plt.plot(
        df_grafica["año_mes"],
        df_grafica["generacion_nuclear"],
        marker="o",
        label="Nuclear",
    )

    plt.plot(
        df_grafica["año_mes"],
        df_grafica["generacion_ciclo_combinado"],
        marker="o",
        label="Ciclo combinado",
    )

    plt.plot(
        df_grafica["año_mes"],
        df_grafica["generacion_carbon"],
        marker="o",
        label="Carbón",
    )

    plt.xticks(rotation=90)
    plt.title("Evolución mensual de generación por tecnología")
    plt.xlabel("Mes")
    plt.ylabel("Generación media [MW]")
    plt.legend()
    plt.tight_layout()

    ruta_grafica = GRAPH_DIR / "evolucion_mensual_generacion_todas.png"
    plt.savefig(ruta_grafica, dpi=300)
    plt.close()

    print(f"Gráfica guardada: {ruta_grafica}")

def ejecutar_validacion_generacion() -> None:
    """
    Ejecuta la validación completa de generación.
    """

    print("\nVALIDACIÓN DE GENERACIÓN")
    print("-" * 50)

    df = cargar_base_modelo()
    df = preparar_base_generacion(df)

    resumen_general = crear_resumen_general_generacion(df)
    resumen_diario = crear_resumen_diario_generacion(df)
    resumen_mensual = crear_resumen_mensual_generacion(resumen_diario)

    guardar_control_calidad_generacion(
        resumen_general=resumen_general,
        resumen_diario=resumen_diario,
        resumen_mensual=resumen_mensual,
    )

    graficar_generacion_renovable(resumen_mensual)
    graficar_generacion_termica(resumen_mensual)
    graficar_generacion_todas(resumen_mensual)

    print("Validación de generación completada correctamente.")
    print("-" * 50)