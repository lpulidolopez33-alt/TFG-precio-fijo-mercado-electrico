from pathlib import Path
from datetime import date, timedelta


# ============================================================
# RUTAS PRINCIPALES DEL PROYECTO
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

DATA_DIR = BASE_DIR / "datos"
RAW_DATA_DIR = DATA_DIR / "historicos"
PROCESSED_DATA_DIR = DATA_DIR / "procesados"
MANUAL_DATA_DIR = DATA_DIR / "entrada_manual"

OUTPUT_DIR = BASE_DIR / "salidas"
GRAPH_DIR = BASE_DIR / "graficas"
MODEL_DIR = BASE_DIR / "modelos"


# ============================================================
# CONFIGURACIÓN GENERAL DEL MODELO
# ============================================================

NOMBRE_PROYECTO = "Previsión de commodity eléctrica"

ZONA_ELECTRICA = "peninsula"

COLUMNA_OBJETIVO = "precio_omie"


# ============================================================
# RANGO HISTÓRICO DE ENTRENAMIENTO
# ============================================================

FECHA_INICIO_HISTORICO = "2023-01-01"

FECHA_EJECUCION = date.today()

FECHA_FIN_HISTORICO = (FECHA_EJECUCION - timedelta(days=1)).strftime("%Y-%m-%d")


# ============================================================
# LÍMITE DE PREVISIÓN
# ============================================================

MAX_DIAS_PREVISION = 366


# ============================================================
# VARIABLES EXPLICATIVAS DEL MODELO
# ============================================================

VARIABLES_MODELO = [
    "demanda",
    "generacion_solar",
    "generacion_eolica",
    "generacion_hidraulica",
    "generacion_nuclear",
    "generacion_ciclo_combinado",
    "precio_gas",
    "mes",
    "dia_semana",
    "hora",
    "es_festivo",
]

# ============================================================
# CONFIGURACIÓN DE DESCARGA OMIE
# ============================================================

MODO_PRUEBA_OMIE = False
# MODO_PRUEBA_OMIE = True

DIAS_PRUEBA_OMIE = 5

VERSIONES_OMIE = [1, 2, 3, 4, 5]

# ============================================================
# CONFIGURACIÓN DE DESCARGA REE / REData
# ============================================================

MODO_PRUEBA_REE = False
# MODO_PRUEBA_REE = True

DIAS_PRUEBA_REE = 5

REE_BASE_URL = "https://apidatos.ree.es/es/datos"

REE_GEO_TRUNC = "electric_system"
REE_GEO_LIMIT = "peninsular"
REE_GEO_IDS = "8741"

# ============================================================
# CONFIGURACIÓN DE DESCARGA GENERACIÓN REE
# ============================================================

MODO_PRUEBA_GENERACION_REE = True

DIAS_PRUEBA_GENERACION_REE = 5

# ============================================================
# CONTROL DE MÓDULOS DEL PROGRAMA
# ============================================================

USAR_DESCARGA_GENERACION_REE = False

# ============================================================
# CONFIGURACIÓN DE IMPORTACIÓN MANUAL DE GENERACIÓN
# ============================================================

USAR_IMPORTACION_GENERACION_EXCEL = True

RUTA_PLANTILLA_GENERACION_MANUAL = MANUAL_DATA_DIR / "plantilla_generacion_manual.xlsx"

RUTA_GENERACION_MANUAL = MANUAL_DATA_DIR / "generacion_manual.xlsx"

# ============================================================
# CONFIGURACIÓN DE BALANCE REE DIARIO MANUAL
# ============================================================

USAR_BALANCE_REE_DIARIO = True

REE_BALANCE_DIARIO_DIR = MANUAL_DATA_DIR / "ree_balance"

# ============================================================
# CONFIGURACIÓN DE ENTRENAMIENTO DEL MODELO
# ============================================================

COLUMNA_OBJETIVO_MODELO = "precio_omie"

PORCENTAJE_TEST = 0.20

RANDOM_STATE = 42

# ============================================================
# CONFIGURACIÓN DE PREVISIÓN FUTURA
# ============================================================

USAR_PREVISION_FUTURA = True

FECHA_INICIO_PREVISION = "2027-01-01"
FECHA_FIN_PREVISION = "2027-12-31"

METODO_ESCENARIO_FUTURO = "mediana_historica_mes_dia_hora"

# ============================================================
# CONFIGURACIÓN DE PRECIO DEL GAS
# ============================================================

USAR_PRECIO_GAS_MANUAL = True

RUTA_PLANTILLA_GAS_MANUAL = MANUAL_DATA_DIR / "plantilla_gas_manual.xlsx"

RUTA_GAS_MANUAL = MANUAL_DATA_DIR / "gas_manual.xlsx"

# ============================================================
# CONFIGURACIÓN DE FICHEROS MIBGAS
# ============================================================

USAR_PROCESADO_MIBGAS = True

MIBGAS_DATA_DIR = MANUAL_DATA_DIR / "mibgas"

# ============================================================
# CONFIGURACIÓN DE PRECIO DEL CO2
# ============================================================

USAR_PRECIO_CO2_MANUAL = True

RUTA_PLANTILLA_CO2_MANUAL = MANUAL_DATA_DIR / "plantilla_co2_manual.xlsx"

RUTA_CO2_MANUAL = MANUAL_DATA_DIR / "co2_manual.xlsx"

# ============================================================
# CONFIGURACIÓN DE FICHEROS CO2
# ============================================================

USAR_PROCESADO_CO2 = True

CO2_DATA_DIR = MANUAL_DATA_DIR / "co2"

# ============================================================
# CONFIGURACIÓN OMIP
# ============================================================

USAR_COMPARACION_OMIP = True

RUTA_PLANTILLA_OMIP_MANUAL = MANUAL_DATA_DIR / "plantilla_omip_manual.xlsx"

RUTA_OMIP_MANUAL = MANUAL_DATA_DIR / "omip_manual.xlsx"

RUTA_COMPARACION_OMIP = OUTPUT_DIR / "comparacion_modelo_omip.xlsx"


# ============================================================
# CONFIGURACIÓN AJUSTE HÍBRIDO OMIP
# ============================================================

USAR_AJUSTE_OMIP = False

RUTA_PREVISION_HIBRIDA_OMIP = OUTPUT_DIR / "prevision_commodity_hibrida_omip.xlsx"

RUTA_RESUMEN_HIBRIDO_OMIP = OUTPUT_DIR / "resumen_commodity_hibrida_omip.xlsx"

UMBRAL_DIFERENCIA_MODELO_OMIP_BAJO = 3.0

UMBRAL_DIFERENCIA_MODELO_OMIP_ALTO = 10.0

FACTOR_AJUSTE_OMIP_BAJO = 0.0

FACTOR_AJUSTE_OMIP_MEDIO = 0.35

FACTOR_AJUSTE_OMIP_ALTO = 0.65

# ============================================================
# CONFIGURACIÓN BACKTESTING TEMPORAL
# ============================================================

USAR_BACKTESTING_TEMPORAL = True

RUTA_BACKTESTING_TEMPORAL = OUTPUT_DIR / "backtesting_temporal_modelo.xlsx"

# ============================================================
# CONFIGURACIÓN DECISIÓN DE COBERTURA
# ============================================================

USAR_DECISION_COBERTURA = True

RUTA_DECISION_COBERTURA = OUTPUT_DIR / "decision_commodity_cobertura.xlsx"

PORCENTAJE_COBERTURA_MODELO_MAYOR_OMIP = 1.00

PORCENTAJE_COBERTURA_MODELO_MENOR_OMIP = 0.70

PRIMA_RIESGO_PARTE_ABIERTA_EUR_MWH = 0.0

# ============================================================
# CONFIGURACIÓN CURVA DE CONSUMO DEL CLIENTE
# ============================================================

USAR_CURVA_CLIENTE = True

RUTA_CURVA_CLIENTE = RUTA_OMIP_MANUAL.parent / "curva_cliente.xlsx"

COLUMNA_CONSUMO_CLIENTE = "consumo_MWh"

# Energía total del cliente para la oferta, en MWh.
# Se usa cuando USAR_CURVA_CLIENTE = False.
ENERGIA_TOTAL_CLIENTE_MWH = 12000.0

# ============================================================
# CONFIGURACIÓN PERFIL OFICIAL REE PARA CURVA CLIENTE
# ============================================================

RUTA_PERFILES_REE_2026 = RUTA_OMIP_MANUAL.parent / "perfiles_iniciales_2026.xlsx"

# Perfil elegido para generar curva_cliente.xlsx.
# Opciones:
# "2.0TD"
# "3.0TD"
# "3.0TDVE"
# "6.XTD"
PERFIL_REE_CLIENTE = "3.0TD"

RUTA_CURVAS_CLIENTES_SIMULADOS = RUTA_OMIP_MANUAL.parent / "curvas_clientes_simulados.xlsx"

# ============================================================
# INTEGRACIÓN CON LA CALCULADORA DE PRECIOS
# ============================================================

RUTA_CALCULADORA_PRECIOS = Path(
    r"C:\Users\mjlop\OneDrive\Desktop\TFG\Calculadora Precios.xlsx"
)

HOJA_CALCULADORA_COMMODITY = "COMMODITY"

CELDA_COMMODITY_FINAL = "B2"
CELDA_COBERTURA_FINAL = "B3"

CELDA_ETIQUETA_COMMODITY_FINAL = "A2"
CELDA_ETIQUETA_COBERTURA_FINAL = "A3"