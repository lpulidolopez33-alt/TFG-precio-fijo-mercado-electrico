import pandas as pd
import matplotlib.pyplot as plt

from config import PROCESSED_DATA_DIR, GRAPH_DIR
from entrenamiento_modelo import COLUMNA_OBJETIVO_MODELO


def cargar_predicciones_test() -> pd.DataFrame:
    """
    Carga las predicciones del periodo de prueba del modelo.
    """

    ruta_resultados = PROCESSED_DATA_DIR / "resultados_entrenamiento_modelos.xlsx"

    if not ruta_resultados.exists():
        raise FileNotFoundError(f"No existe el archivo: {ruta_resultados}")

    df = pd.read_excel(
        ruta_resultados,
        sheet_name="predicciones_test",
    )

    return df


def preparar_errores(df: pd.DataFrame) -> pd.DataFrame:
    """
    Prepara las columnas de error del modelo.

    error_estimacion = precio_estimado - precio_real

    Si error_estimacion > 0:
        el modelo ha previsto por encima del precio real.

    Si error_estimacion < 0:
        el modelo se ha quedado corto.
        Este caso es el riesgo más relevante para una oferta a precio fijo.
    """

    df = df.copy()

    df["fecha"] = pd.to_datetime(df["fecha"])

    df[COLUMNA_OBJETIVO_MODELO] = pd.to_numeric(
        df[COLUMNA_OBJETIVO_MODELO],
        errors="coerce",
    )

    df["precio_estimado_modelo"] = pd.to_numeric(
        df["precio_estimado_modelo"],
        errors="coerce",
    )

    df["error_estimacion"] = (
        df["precio_estimado_modelo"]
        - df[COLUMNA_OBJETIVO_MODELO]
    )

    df["error_absoluto"] = df["error_estimacion"].abs()

    df["error_adverso"] = (
        df[COLUMNA_OBJETIVO_MODELO]
        - df["precio_estimado_modelo"]
    )

    df.loc[df["error_adverso"] < 0, "error_adverso"] = 0

    return df


def crear_resumen_errores(df: pd.DataFrame) -> pd.DataFrame:
    """
    Crea un resumen general de errores del modelo.
    """

    resumen = {
        "numero_horas_test": len(df),
        "precio_real_medio": df[COLUMNA_OBJETIVO_MODELO].mean(),
        "precio_estimado_medio": df["precio_estimado_modelo"].mean(),
        "error_medio": df["error_estimacion"].mean(),
        "error_absoluto_medio_MAE": df["error_absoluto"].mean(),
        "error_absoluto_mediana": df["error_absoluto"].median(),
        "error_maximo_positivo_modelo_sobreestima": df["error_estimacion"].max(),
        "error_maximo_negativo_modelo_subestima": df["error_estimacion"].min(),
        "porcentaje_horas_modelo_subestima": (
            (df["error_estimacion"] < 0).mean() * 100
        ),
        "error_adverso_medio": df["error_adverso"].mean(),
        "error_adverso_maximo": df["error_adverso"].max(),
    }

    return pd.DataFrame([resumen])


def crear_percentiles_prima(df: pd.DataFrame) -> pd.DataFrame:
    """
    Calcula posibles primas de incertidumbre a partir del error adverso.

    El error adverso mide cuánto se ha quedado corto el modelo:
    precio_real - precio_estimado, solo cuando es positivo.
    """

    percentiles = [50, 60, 70, 75, 80, 85, 90, 95]

    registros = []

    errores_adversos = df["error_adverso"]

    for percentil in percentiles:
        prima = errores_adversos.quantile(percentil / 100)

        registros.append(
            {
                "percentil": percentil,
                "prima_incertidumbre_EUR_MWh": prima,
            }
        )

    return pd.DataFrame(registros)


def crear_error_mensual(df: pd.DataFrame) -> pd.DataFrame:
    """
    Calcula errores medios del modelo por mes.
    """

    df = df.copy()

    df["año_mes"] = df["fecha"].dt.to_period("M").astype(str)

    resumen_mensual = (
        df.groupby("año_mes", as_index=False)
        .agg(
            precio_real_medio=(COLUMNA_OBJETIVO_MODELO, "mean"),
            precio_estimado_medio=("precio_estimado_modelo", "mean"),
            error_medio=("error_estimacion", "mean"),
            error_absoluto_medio=("error_absoluto", "mean"),
            error_adverso_medio=("error_adverso", "mean"),
            error_adverso_p90=("error_adverso", lambda x: x.quantile(0.90)),
            horas=("error_estimacion", "count"),
        )
    )

    return resumen_mensual


def guardar_analisis_riesgo(
    resumen_errores: pd.DataFrame,
    percentiles_prima: pd.DataFrame,
    error_mensual: pd.DataFrame,
    df_errores: pd.DataFrame,
) -> None:
    """
    Guarda el análisis de errores y primas en Excel.
    """

    ruta_salida = PROCESSED_DATA_DIR / "analisis_riesgo_modelo.xlsx"

    try:
        with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
            resumen_errores.to_excel(writer, sheet_name="resumen_errores", index=False)
            percentiles_prima.to_excel(writer, sheet_name="primas_percentiles", index=False)
            error_mensual.to_excel(writer, sheet_name="error_mensual", index=False)
            df_errores.to_excel(writer, sheet_name="errores_horarios", index=False)

    except PermissionError:
        raise PermissionError(
            f"No se puede guardar el archivo {ruta_salida}. "
            f"Probablemente está abierto en Excel. Ciérralo y vuelve a ejecutar."
        )

    print(f"Análisis de riesgo del modelo guardado en: {ruta_salida}")


def graficar_distribucion_errores(df: pd.DataFrame) -> None:
    """
    Genera un histograma con la distribución de errores del modelo.
    """

    GRAPH_DIR.mkdir(parents=True, exist_ok=True)

    plt.figure(figsize=(10, 5))
    plt.hist(df["error_estimacion"], bins=50)
    plt.axvline(0)
    plt.title("Distribución de errores del modelo")
    plt.xlabel("Error de estimación [€/MWh]")
    plt.ylabel("Número de horas")
    plt.tight_layout()

    ruta_grafica = GRAPH_DIR / "modelo_distribucion_errores.png"
    plt.savefig(ruta_grafica, dpi=300)
    plt.close()

    print(f"Gráfica guardada: {ruta_grafica}")


def graficar_prima_incertidumbre(percentiles_prima: pd.DataFrame) -> None:
    """
    Genera una gráfica con la prima de incertidumbre según percentil.
    """

    GRAPH_DIR.mkdir(parents=True, exist_ok=True)

    plt.figure(figsize=(8, 5))
    plt.plot(
        percentiles_prima["percentil"],
        percentiles_prima["prima_incertidumbre_EUR_MWh"],
        marker="o",
    )
    plt.title("Prima de incertidumbre según percentil de error adverso")
    plt.xlabel("Percentil del error adverso")
    plt.ylabel("Prima de incertidumbre [€/MWh]")
    plt.tight_layout()

    ruta_grafica = GRAPH_DIR / "modelo_prima_incertidumbre.png"
    plt.savefig(ruta_grafica, dpi=300)
    plt.close()

    print(f"Gráfica guardada: {ruta_grafica}")


def ejecutar_analisis_riesgo_modelo() -> None:
    """
    Ejecuta el análisis completo de errores y prima de incertidumbre.
    """

    print("\nANÁLISIS DE RIESGO DEL MODELO")
    print("-" * 50)

    df = cargar_predicciones_test()
    df = preparar_errores(df)

    resumen_errores = crear_resumen_errores(df)
    percentiles_prima = crear_percentiles_prima(df)
    error_mensual = crear_error_mensual(df)

    guardar_analisis_riesgo(
        resumen_errores=resumen_errores,
        percentiles_prima=percentiles_prima,
        error_mensual=error_mensual,
        df_errores=df,
    )

    graficar_distribucion_errores(df)
    graficar_prima_incertidumbre(percentiles_prima)

    print("Análisis de riesgo del modelo completado correctamente.")
    print("-" * 50)