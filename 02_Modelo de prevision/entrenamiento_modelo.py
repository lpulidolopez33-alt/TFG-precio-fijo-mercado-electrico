import joblib
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

from sklearn.ensemble import RandomForestRegressor, ExtraTreesRegressor
from sklearn.ensemble import HistGradientBoostingRegressor
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
from sklearn.inspection import permutation_importance
from sklearn.model_selection import ParameterGrid

from xgboost import XGBRegressor

from config import (
    PROCESSED_DATA_DIR,
    GRAPH_DIR,
    MODEL_DIR,
    COLUMNA_OBJETIVO_MODELO,
    PORCENTAJE_TEST,
    RANDOM_STATE,
)


COLUMNAS_NO_MODELO = [
    "fecha_hora_local",
    "fecha_hora_utc",
    "fecha",
    COLUMNA_OBJETIVO_MODELO,
]


def cargar_base_entrenamiento() -> pd.DataFrame:
    """
    Carga la base preparada para el entrenamiento del modelo.
    """

    ruta_base = PROCESSED_DATA_DIR / "base_entrenamiento.xlsx"

    if not ruta_base.exists():
        raise FileNotFoundError(f"No existe el archivo: {ruta_base}")

    df = pd.read_excel(ruta_base)

    return df


def preparar_base_para_entrenamiento(df: pd.DataFrame) -> pd.DataFrame:
    """
    Ordena la base, convierte fechas y elimina posibles filas incompletas.
    """

    df = df.copy()

    df["fecha_hora_utc_dt"] = pd.to_datetime(df["fecha_hora_utc"], utc=True)
    df[COLUMNA_OBJETIVO_MODELO] = pd.to_numeric(
        df[COLUMNA_OBJETIVO_MODELO],
        errors="coerce",
    )

    df = df.sort_values("fecha_hora_utc_dt")

    df = df.dropna(subset=[COLUMNA_OBJETIVO_MODELO])

    return df


def obtener_columnas_variables(df: pd.DataFrame) -> list[str]:
    """
    Selecciona las variables que usará el modelo.

    Se excluyen fechas, textos y la variable objetivo.
    """

    columnas_candidatas = []

    for columna in df.columns:
        if columna in COLUMNAS_NO_MODELO:
            continue

        if columna == "fecha_hora_utc_dt":
            continue

        if pd.api.types.is_numeric_dtype(df[columna]):
            columnas_candidatas.append(columna)

    return columnas_candidatas


def dividir_train_test_temporal(
    df: pd.DataFrame,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Divide la base en entrenamiento y prueba respetando el orden temporal.

    No se mezclan horas aleatoriamente, porque en series temporales sería
    incorrecto entrenar con datos futuros y probar con datos pasados.
    """

    numero_filas = len(df)

    numero_filas_test = int(numero_filas * PORCENTAJE_TEST)
    punto_corte = numero_filas - numero_filas_test

    df_train = df.iloc[:punto_corte].copy()
    df_test = df.iloc[punto_corte:].copy()

    return df_train, df_test


def crear_modelos() -> dict:
    """
    Define los modelos que se van a entrenar y comparar.
    """

    modelos = {
        "random_forest": RandomForestRegressor(
            n_estimators=300,
            max_depth=18,
            min_samples_leaf=3,
            random_state=RANDOM_STATE,
            n_jobs=-1,
        ),

        "extra_trees": ExtraTreesRegressor(
            n_estimators=300,
            max_depth=18,
            min_samples_leaf=3,
            random_state=RANDOM_STATE,
            n_jobs=-1,
        ),

        "hist_gradient_boosting": HistGradientBoostingRegressor(
            max_iter=400,
            learning_rate=0.05,
            max_leaf_nodes=31,
            random_state=RANDOM_STATE,
        ),

        "xgboost": XGBRegressor(
            n_estimators=700,
            max_depth=5,
            learning_rate=0.03,
            subsample=0.85,
            colsample_bytree=0.85,
            objective="reg:squarederror",
            random_state=RANDOM_STATE,
            n_jobs=-1,
        ),
    }

    return modelos

def optimizar_hist_gradient_boosting(
    df_train: pd.DataFrame,
    df_test: pd.DataFrame,
    columnas_variables: list[str],
) -> tuple[HistGradientBoostingRegressor, dict, pd.DataFrame]:
    """
    Prueba distintas configuraciones del modelo HistGradientBoostingRegressor
    y devuelve la mejor según el MAE en el periodo de prueba.
    """

    print("\nOPTIMIZACIÓN DE HIST GRADIENT BOOSTING")
    print("-" * 50)

    X_train = df_train[columnas_variables]
    y_train = df_train[COLUMNA_OBJETIVO_MODELO]

    X_test = df_test[columnas_variables]
    y_test = df_test[COLUMNA_OBJETIVO_MODELO]

    parametros = {
        "max_iter": [400, 700, 1000],
        "learning_rate": [0.03, 0.05, 0.07],
        "max_leaf_nodes": [15, 31, 45],
        "min_samples_leaf": [20, 40, 80],
        "l2_regularization": [0.0, 0.1, 1.0],
    }

    registros = []

    mejor_modelo = None
    mejor_parametros = None
    mejor_mae = float("inf")

    for combinacion in ParameterGrid(parametros):
        modelo = HistGradientBoostingRegressor(
            **combinacion,
            random_state=RANDOM_STATE,
        )

        modelo.fit(X_train, y_train)

        predicciones = modelo.predict(X_test)

        metricas = calcular_metricas(
            y_real=y_test,
            y_predicho=predicciones,
        )

        registro = {
            **combinacion,
            **metricas,
        }

        registros.append(registro)

        if metricas["MAE_EUR_MWh"] < mejor_mae:
            mejor_mae = metricas["MAE_EUR_MWh"]
            mejor_modelo = modelo
            mejor_parametros = combinacion

    df_resultados = pd.DataFrame(registros)

    df_resultados = df_resultados.sort_values("MAE_EUR_MWh")

    print("Mejores parámetros encontrados:")
    print(mejor_parametros)
    print(f"Mejor MAE optimizado: {mejor_mae:.2f} €/MWh")
    print("-" * 50)

    return mejor_modelo, mejor_parametros, df_resultados


def calcular_metricas(
    y_real: pd.Series,
    y_predicho: np.ndarray,
) -> dict:
    """
    Calcula las métricas de error del modelo.
    """

    mae = mean_absolute_error(y_real, y_predicho)
    rmse = np.sqrt(mean_squared_error(y_real, y_predicho))
    r2 = r2_score(y_real, y_predicho)

    error = y_predicho - y_real
    sesgo_medio = np.mean(error)

    return {
        "MAE_EUR_MWh": mae,
        "RMSE_EUR_MWh": rmse,
        "R2": r2,
        "sesgo_medio_EUR_MWh": sesgo_medio,
    }


def entrenar_y_evaluar_modelos(
    df_train: pd.DataFrame,
    df_test: pd.DataFrame,
    columnas_variables: list[str],
) -> tuple[pd.DataFrame, dict, str]:
    """
    Entrena todos los modelos y devuelve:
    - tabla de métricas
    - modelos entrenados
    - nombre del mejor modelo
    """

    X_train = df_train[columnas_variables]
    y_train = df_train[COLUMNA_OBJETIVO_MODELO]

    X_test = df_test[columnas_variables]
    y_test = df_test[COLUMNA_OBJETIVO_MODELO]

    modelos = crear_modelos()

    registros_metricas = []
    modelos_entrenados = {}

    for nombre_modelo, modelo in modelos.items():
        print(f"Entrenando modelo: {nombre_modelo}")

        modelo.fit(X_train, y_train)

        predicciones = modelo.predict(X_test)

        metricas = calcular_metricas(
            y_real=y_test,
            y_predicho=predicciones,
        )

        metricas["modelo"] = nombre_modelo
        registros_metricas.append(metricas)

        modelos_entrenados[nombre_modelo] = modelo

        print(
            f"Modelo {nombre_modelo} - "
            f"MAE: {metricas['MAE_EUR_MWh']:.2f} €/MWh - "
            f"RMSE: {metricas['RMSE_EUR_MWh']:.2f} €/MWh - "
            f"R2: {metricas['R2']:.3f}"
        )

    df_metricas = pd.DataFrame(registros_metricas)

    df_metricas = df_metricas[
        [
            "modelo",
            "MAE_EUR_MWh",
            "RMSE_EUR_MWh",
            "R2",
            "sesgo_medio_EUR_MWh",
        ]
    ]

    df_metricas = df_metricas.sort_values("MAE_EUR_MWh")

    mejor_modelo = df_metricas.iloc[0]["modelo"]

    return df_metricas, modelos_entrenados, mejor_modelo


def crear_tabla_predicciones(
    df_test: pd.DataFrame,
    modelo,
    columnas_variables: list[str],
) -> pd.DataFrame:
    """
    Crea una tabla con precio real, precio estimado y error.
    """

    df_predicciones = df_test[
        [
            "fecha_hora_local",
            "fecha_hora_utc",
            "fecha",
            "año",
            "mes",
            "hora",
            COLUMNA_OBJETIVO_MODELO,
        ]
    ].copy()

    predicciones = modelo.predict(df_test[columnas_variables])

    df_predicciones["precio_estimado_modelo"] = predicciones
    df_predicciones["error_estimacion"] = (
        df_predicciones["precio_estimado_modelo"]
        - df_predicciones[COLUMNA_OBJETIVO_MODELO]
    )

    df_predicciones["error_absoluto"] = df_predicciones[
        "error_estimacion"
    ].abs()

    return df_predicciones


def crear_importancia_variables(
    modelo,
    X_test: pd.DataFrame,
    y_test: pd.Series,
    columnas_variables: list[str],
) -> pd.DataFrame:
    """
    Calcula la importancia de variables del modelo.

    Si el modelo proporciona feature_importances_, se utiliza directamente.
    Si no, se calcula importancia por permutación.
    """

    if hasattr(modelo, "feature_importances_"):
        df_importancia = pd.DataFrame(
            {
                "variable": columnas_variables,
                "importancia": modelo.feature_importances_,
                "metodo": "feature_importances",
            }
        )

    else:
        resultado_permutacion = permutation_importance(
            modelo,
            X_test,
            y_test,
            n_repeats=10,
            random_state=42,
            scoring="neg_mean_absolute_error",
            n_jobs=-1,
        )

        df_importancia = pd.DataFrame(
            {
                "variable": columnas_variables,
                "importancia": resultado_permutacion.importances_mean,
                "desviacion_importancia": resultado_permutacion.importances_std,
                "metodo": "permutation_importance",
            }
        )

    df_importancia = df_importancia.sort_values(
        "importancia",
        ascending=False,
    )

    return df_importancia


def graficar_precio_real_vs_estimado(df_predicciones: pd.DataFrame) -> None:
    """
    Genera una gráfica temporal de precio real frente a precio estimado.
    """

    GRAPH_DIR.mkdir(parents=True, exist_ok=True)

    df_grafica = df_predicciones.copy()
    df_grafica["fecha_hora"] = pd.to_datetime(df_grafica["fecha_hora_utc"], utc=True)

    plt.figure(figsize=(14, 6))
    plt.plot(
        df_grafica["fecha_hora"],
        df_grafica[COLUMNA_OBJETIVO_MODELO],
        label="Precio real OMIE",
    )
    plt.plot(
        df_grafica["fecha_hora"],
        df_grafica["precio_estimado_modelo"],
        label="Precio estimado modelo",
    )
    plt.title("Precio real vs precio estimado en periodo de prueba")
    plt.xlabel("Fecha")
    plt.ylabel("Precio [€/MWh]")
    plt.legend()
    plt.tight_layout()

    ruta_grafica = GRAPH_DIR / "modelo_precio_real_vs_estimado.png"
    plt.savefig(ruta_grafica, dpi=300)
    plt.close()

    print(f"Gráfica guardada: {ruta_grafica}")


def graficar_dispersion_real_estimado(df_predicciones: pd.DataFrame) -> None:
    """
    Genera una gráfica de dispersión entre precio real y estimado.
    """

    GRAPH_DIR.mkdir(parents=True, exist_ok=True)

    y_real = df_predicciones[COLUMNA_OBJETIVO_MODELO]
    y_pred = df_predicciones["precio_estimado_modelo"]

    minimo = min(y_real.min(), y_pred.min())
    maximo = max(y_real.max(), y_pred.max())

    plt.figure(figsize=(6, 6))
    plt.scatter(y_real, y_pred, alpha=0.3)
    plt.plot([minimo, maximo], [minimo, maximo])
    plt.title("Dispersión precio real vs precio estimado")
    plt.xlabel("Precio real [€/MWh]")
    plt.ylabel("Precio estimado [€/MWh]")
    plt.tight_layout()

    ruta_grafica = GRAPH_DIR / "modelo_dispersion_real_estimado.png"
    plt.savefig(ruta_grafica, dpi=300)
    plt.close()

    print(f"Gráfica guardada: {ruta_grafica}")


def graficar_importancia_variables(df_importancia: pd.DataFrame) -> None:
    """
    Genera una gráfica de importancia de variables.
    """

    if df_importancia["importancia"].isna().all():
        print("El mejor modelo no proporciona importancia de variables.")
        return

    GRAPH_DIR.mkdir(parents=True, exist_ok=True)

    df_grafica = df_importancia.head(15).sort_values("importancia")

    plt.figure(figsize=(10, 6))
    plt.barh(
        df_grafica["variable"],
        df_grafica["importancia"],
    )
    plt.title("Importancia de variables del modelo")
    plt.xlabel("Importancia relativa")
    plt.ylabel("Variable")
    plt.tight_layout()

    ruta_grafica = GRAPH_DIR / "modelo_importancia_variables.png"
    plt.savefig(ruta_grafica, dpi=300)
    plt.close()

    print(f"Gráfica guardada: {ruta_grafica}")


def guardar_resultados_entrenamiento(
    df_metricas: pd.DataFrame,
    df_predicciones: pd.DataFrame,
    df_importancia: pd.DataFrame,
    mejor_modelo: str,
    modelo_entrenado,
    columnas_variables: list[str],
    df_optimizacion: pd.DataFrame | None = None,
) -> None:
    """
    Guarda métricas, predicciones, importancia de variables y modelo final.
    """

    ruta_resultados = PROCESSED_DATA_DIR / "resultados_entrenamiento_modelos.xlsx"
    ruta_modelo = MODEL_DIR / "modelo_precio_omie.joblib"
    ruta_variables = MODEL_DIR / "variables_modelo.xlsx"

    MODEL_DIR.mkdir(parents=True, exist_ok=True)

    try:
        with pd.ExcelWriter(ruta_resultados, engine="openpyxl") as writer:
            df_metricas.to_excel(writer, sheet_name="metricas_modelos", index=False)
            df_predicciones.to_excel(writer, sheet_name="predicciones_test", index=False)
            df_importancia.to_excel(writer, sheet_name="importancia_variables", index=False)

            if df_optimizacion is not None and not df_optimizacion.empty:
               df_optimizacion.to_excel(writer, sheet_name="optimizacion_hgb", index=False)

        pd.DataFrame(
            {
                "variable": columnas_variables,
            }
        ).to_excel(ruta_variables, index=False)

        joblib.dump(
            {
                "modelo": modelo_entrenado,
                "mejor_modelo": mejor_modelo,
                "columnas_variables": columnas_variables,
            },
            ruta_modelo,
        )

    except PermissionError:
        raise PermissionError(
            "No se pueden guardar los resultados del entrenamiento. "
            "Probablemente algún Excel está abierto."
        )

    print(f"Resultados de entrenamiento guardados en: {ruta_resultados}")
    print(f"Modelo final guardado en: {ruta_modelo}")
    print(f"Variables del modelo guardadas en: {ruta_variables}")


def entrenar_modelos_precio() -> None:
    """
    Ejecuta el entrenamiento completo de los modelos de precio OMIE.
    """

    print("\nENTRENAMIENTO DE MODELOS DE PRECIO OMIE")
    print("-" * 50)

    df = cargar_base_entrenamiento()
    df = preparar_base_para_entrenamiento(df)

    columnas_variables = obtener_columnas_variables(df)

    df_train, df_test = dividir_train_test_temporal(df)

    print(f"Filas entrenamiento: {len(df_train)}")
    print(f"Filas prueba: {len(df_test)}")
    print(f"Número de variables utilizadas: {len(columnas_variables)}")
    print(
        f"Periodo entrenamiento: "
        f"{df_train['fecha_hora_local'].min()} a {df_train['fecha_hora_local'].max()}"
    )
    print(
        f"Periodo prueba: "
        f"{df_test['fecha_hora_local'].min()} a {df_test['fecha_hora_local'].max()}"
    )

    df_metricas, modelos_entrenados, mejor_modelo = entrenar_y_evaluar_modelos(
        df_train=df_train,
        df_test=df_test,
        columnas_variables=columnas_variables,
    )

    print(f"Mejor modelo seleccionado inicialmente: {mejor_modelo}")

    modelo_final = modelos_entrenados[mejor_modelo]
    df_optimizacion = pd.DataFrame()

    if mejor_modelo == "hist_gradient_boosting":
        modelo_optimizado, mejores_parametros, df_optimizacion = optimizar_hist_gradient_boosting(
            df_train=df_train,
            df_test=df_test,
            columnas_variables=columnas_variables,
        )

        predicciones_optimizadas = modelo_optimizado.predict(
            df_test[columnas_variables]
        )

        metricas_optimizadas = calcular_metricas(
            y_real=df_test[COLUMNA_OBJETIVO_MODELO],
            y_predicho=predicciones_optimizadas,
        )

        metricas_optimizadas["modelo"] = "hist_gradient_boosting_optimizado"

        df_metricas = pd.concat(
            [
                df_metricas,
                pd.DataFrame([metricas_optimizadas])[
                    [
                        "modelo",
                        "MAE_EUR_MWh",
                        "RMSE_EUR_MWh",
                        "R2",
                        "sesgo_medio_EUR_MWh",
                    ]
                ],
            ],
            ignore_index=True,
        )

        df_metricas = df_metricas.sort_values("MAE_EUR_MWh")

        mejor_modelo = df_metricas.iloc[0]["modelo"]

        if mejor_modelo == "hist_gradient_boosting_optimizado":
            modelo_final = modelo_optimizado
            print("Se selecciona el modelo HistGradientBoosting optimizado.")
        else:
            print("La optimización no mejora al modelo inicial.")

    print(f"Mejor modelo final seleccionado: {mejor_modelo}")

    df_predicciones = crear_tabla_predicciones(
        df_test=df_test,
        modelo=modelo_final,
        columnas_variables=columnas_variables,
    )

    df_importancia = crear_importancia_variables(
        modelo=modelo_final,
        X_test=df_test[columnas_variables],
        y_test=df_test[COLUMNA_OBJETIVO_MODELO],
        columnas_variables=columnas_variables,
    )

    guardar_resultados_entrenamiento(
        df_metricas=df_metricas,
        df_predicciones=df_predicciones,
        df_importancia=df_importancia,
        mejor_modelo=mejor_modelo,
        modelo_entrenado=modelo_final,
        columnas_variables=columnas_variables,
        df_optimizacion=df_optimizacion,
    )

    graficar_precio_real_vs_estimado(df_predicciones)
    graficar_dispersion_real_estimado(df_predicciones)
    graficar_importancia_variables(df_importancia)

    print("Entrenamiento de modelos completado correctamente.")
    print("-" * 50)