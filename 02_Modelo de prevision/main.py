from config import (
    DATA_DIR,
    RAW_DATA_DIR,
    PROCESSED_DATA_DIR,
    OUTPUT_DIR,
    GRAPH_DIR,
    NOMBRE_PROYECTO,
    FECHA_INICIO_HISTORICO,
    FECHA_FIN_HISTORICO,
    MAX_DIAS_PREVISION,
    VARIABLES_MODELO,
    USAR_DESCARGA_GENERACION_REE,
    MANUAL_DATA_DIR,
    USAR_IMPORTACION_GENERACION_EXCEL,
    USAR_BALANCE_REE_DIARIO,
    REE_BALANCE_DIARIO_DIR,
    MODEL_DIR,
    USAR_PRECIO_GAS_MANUAL,
    USAR_PROCESADO_MIBGAS,
    MIBGAS_DATA_DIR,
    USAR_PRECIO_CO2_MANUAL,
    USAR_PROCESADO_CO2,
    CO2_DATA_DIR,
    USAR_COMPARACION_OMIP,
    USAR_AJUSTE_OMIP,
    USAR_BACKTESTING_TEMPORAL,
    USAR_DECISION_COBERTURA,
    RUTA_GENERACION_MANUAL,
    RUTA_GAS_MANUAL,
    RUTA_CO2_MANUAL,
    RUTA_OMIP_MANUAL,
)

from datos_base import guardar_plantilla_historica
from descarga_omie import descargar_historico_omie
from descarga_ree import descargar_historico_demanda_ree
from integracion_datos import guardar_base_modelo
from validacion_omie import ejecutar_validacion_omie
from validacion_demanda import ejecutar_validacion_demanda
from descarga_generacion_ree import descargar_historico_generacion_ree
from plantilla_generacion_manual import crear_plantilla_generacion_manual
from importacion_generacion_manual import guardar_generacion_manual_procesada
from procesar_balance_ree_diario import procesar_balance_ree_diario
from validacion_generacion import ejecutar_validacion_generacion
from preparacion_modelo import guardar_base_entrenamiento
from entrenamiento_modelo import entrenar_modelos_precio
from analisis_riesgo_modelo import ejecutar_analisis_riesgo_modelo
from prevision_futura import generar_prevision_futura
from commodity_periodos import generar_commodity_por_periodos
from entrada_usuario import pedir_fechas_prevision
from plantilla_gas_manual import crear_plantilla_gas_manual
from importacion_gas_manual import guardar_gas_manual_procesado
from procesar_mibgas import procesar_ficheros_mibgas
from graficas_prevision import generar_graficas_prevision
from plantilla_co2_manual import crear_plantilla_co2_manual
from importacion_co2_manual import guardar_co2_manual_procesado
from procesar_co2 import procesar_ficheros_co2
from plantilla_omip_manual import crear_plantilla_omip_manual
from comparacion_omip import generar_comparacion_modelo_omip
from ajuste_omip import generar_ajuste_omip
from backtesting_temporal import ejecutar_backtesting_temporal
from decision_cobertura import generar_decision_cobertura
from escenarios_mercado import generar_escenarios_mercado, mostrar_menu_escenarios




def crear_estructura_carpetas() -> None:
    """
    Crea las carpetas principales del proyecto si todavía no existen.
    """

    carpetas = [
        DATA_DIR,
        RAW_DATA_DIR,
        PROCESSED_DATA_DIR,
        OUTPUT_DIR,
        GRAPH_DIR,
        MANUAL_DATA_DIR,
        REE_BALANCE_DIARIO_DIR,
        MODEL_DIR,
        MIBGAS_DATA_DIR,
        CO2_DATA_DIR,
    ]

    for carpeta in carpetas:
        carpeta.mkdir(parents=True, exist_ok=True)
        print(f"Carpeta preparada: {carpeta}")


def mostrar_configuracion() -> None:
    """
    Muestra por pantalla la configuración principal del proyecto.
    """

    print("\nCONFIGURACIÓN DEL PROYECTO")
    print("-" * 50)
    print(f"Proyecto: {NOMBRE_PROYECTO}")
    print(f"Histórico desde: {FECHA_INICIO_HISTORICO}")
    print(f"Histórico hasta: {FECHA_FIN_HISTORICO}")
    print(f"Máximo días de previsión: {MAX_DIAS_PREVISION}")
    print(f"Número de variables del modelo: {len(VARIABLES_MODELO)}")
    print("-" * 50)


def mostrar_menu() -> str:
    """
    Muestra el menú principal simplificado.
    """

    print("\nMENÚ PRINCIPAL")
    print("-" * 50)
    print("1 - Actualizar todos los datos y entrenar el modelo")
    print("2 - Ejecutar previsión, comparación OMIP y decisión de cobertura")
    print("3 - Posibles variaciones del mercado")
    print("-" * 50)

    opcion = input("Selecciona una opción [2]: ").strip()

    if opcion == "":
        opcion = "2"

    return opcion


def pedir_si_no(mensaje: str, defecto: bool | None = None) -> bool:
    """
    Pregunta de tipo sí/no.
    """

    while True:
        if defecto is True:
            sufijo = " [S]"
        elif defecto is False:
            sufijo = " [N]"
        else:
            sufijo = ""

        respuesta = input(f"{mensaje}{sufijo}: ").strip().lower()

        if respuesta == "" and defecto is not None:
            return defecto

        if respuesta in ["s", "si", "sí", "y", "yes"]:
            return True

        if respuesta in ["n", "no"]:
            return False

        print("Respuesta no válida. Escribe S para sí o N para no.")



def mostrar_recordatorio_datos_manuales() -> None:
    """
    Recuerda al usuario qué entradas manuales debe revisar antes de actualizar.
    """

    print("\nRECORDATORIO DE DATOS MANUALES")
    print("-" * 50)
    print("Antes de actualizar, revisa los ficheros o carpetas manuales que apliquen:")

    elementos = []

    if USAR_IMPORTACION_GENERACION_EXCEL:
        elementos.append(("Generación manual", RUTA_GENERACION_MANUAL))

    if USAR_BALANCE_REE_DIARIO:
        elementos.append(("Balance diario REE", REE_BALANCE_DIARIO_DIR))

    if USAR_PROCESADO_MIBGAS:
        elementos.append(("Ficheros MIBGAS", MIBGAS_DATA_DIR))

    if USAR_PRECIO_GAS_MANUAL:
        elementos.append(("Gas manual", RUTA_GAS_MANUAL))

    if USAR_PROCESADO_CO2:
        elementos.append(("Ficheros CO2", CO2_DATA_DIR))

    if USAR_PRECIO_CO2_MANUAL:
        elementos.append(("CO2 manual", RUTA_CO2_MANUAL))

    elementos.append(("OMIP manual para la comparación/cobertura", RUTA_OMIP_MANUAL))

    for nombre, ruta in elementos:
        estado = "OK" if ruta.exists() else "NO ENCONTRADO"
        print(f"- {nombre}: {ruta} [{estado}]")

    print("-" * 50)
    continuar = pedir_si_no("¿Quieres continuar con la actualización?", defecto=True)
    if not continuar:
        raise SystemExit("Actualización cancelada por el usuario.")


def ejecutar_actualizacion_datos() -> None:
    """
    Actualiza y valida la base histórica del modelo.
    """

    mostrar_recordatorio_datos_manuales()

    guardar_plantilla_historica()
    crear_plantilla_generacion_manual()
    crear_plantilla_gas_manual()
    crear_plantilla_co2_manual()
    crear_plantilla_omip_manual()

    descargar_historico_omie()
    descargar_historico_demanda_ree()

    if USAR_DESCARGA_GENERACION_REE:
        descargar_historico_generacion_ree()
    else:
        print("\nDescarga de generación REE desactivada temporalmente.")

    if USAR_BALANCE_REE_DIARIO:
        procesar_balance_ree_diario()
    else:
        print("\nProcesado de balance REE diario desactivado.")

    if USAR_IMPORTACION_GENERACION_EXCEL:
        guardar_generacion_manual_procesada()
    else:
        print("\nImportación manual de generación desactivada.")

    if USAR_PROCESADO_MIBGAS:
        procesar_ficheros_mibgas()
    else:
        print("\nProcesado de ficheros MIBGAS desactivado.")

    if USAR_PRECIO_GAS_MANUAL:
        guardar_gas_manual_procesado()
    else:
        print("\nImportación manual de gas desactivada.")

    if USAR_PROCESADO_CO2:
        procesar_ficheros_co2()
    else:
        print("\nProcesado de ficheros CO2 desactivado.")

    if USAR_PRECIO_CO2_MANUAL:
        guardar_co2_manual_procesado()
    else:
        print("\nImportación manual de CO2 desactivada.")

    guardar_base_modelo()

    ejecutar_validacion_omie()
    ejecutar_validacion_demanda()
    ejecutar_validacion_generacion()

    guardar_base_entrenamiento()


def ejecutar_entrenamiento() -> None:
    """
    Entrena los modelos y actualiza el análisis de riesgo.
    """

    entrenar_modelos_precio()
    ejecutar_analisis_riesgo_modelo()


def ejecutar_prevision() -> None:
    """
    Genera la previsión futura, la comparación con OMIP y la decisión de
    cobertura utilizando únicamente las fechas solicitadas.

    No se solicita ni se utiliza información sobre el consumo, la energía,
    la tarifa o la curva de carga de ningún cliente.
    """

    fecha_inicio_prevision, fecha_fin_prevision = pedir_fechas_prevision()

    generar_prevision_futura(
        fecha_inicio_prevision=fecha_inicio_prevision,
        fecha_fin_prevision=fecha_fin_prevision,
    )

    generar_commodity_por_periodos()
    generar_graficas_prevision()

    if USAR_COMPARACION_OMIP:
        generar_comparacion_modelo_omip()
    else:
        print("\nComparación con OMIP desactivada.")

    if USAR_DECISION_COBERTURA:
        generar_decision_cobertura()
    else:
        print("\nDecisión de cobertura desactivada.")

    if USAR_AJUSTE_OMIP:
        generar_ajuste_omip()
    else:
        print("\nAjuste híbrido OMIP-modelo desactivado.")



def ejecutar_variaciones_mercado() -> None:
    """
    Ejecuta los escenarios de sensibilidad del precio previsto.

    Los cambios se aplican únicamente sobre copias en memoria de la base futura.
    No se modifican los datos históricos, el modelo entrenado ni las salidas de
    la previsión normal.
    """

    fecha_inicio_prevision, fecha_fin_prevision = pedir_fechas_prevision()
    opcion_escenario = mostrar_menu_escenarios()

    generar_escenarios_mercado(
        fecha_inicio_prevision=fecha_inicio_prevision,
        fecha_fin_prevision=fecha_fin_prevision,
        opcion_escenario=opcion_escenario,
    )

def ejecutar_backtesting() -> None:
    """
    Ejecuta la validación retrospectiva del modelo.
    """

    if USAR_BACKTESTING_TEMPORAL:
        ejecutar_backtesting_temporal()
    else:
        print("\nBacktesting temporal desactivado.")


def main() -> None:
    """
    Función principal del programa.
    """

    print("Iniciando programa de previsión de commodity...")

    crear_estructura_carpetas()
    mostrar_configuracion()

    opcion = mostrar_menu()

    if opcion == "1":
        print("\nActualizando datos y reentrenando modelo...")
        ejecutar_actualizacion_datos()
        ejecutar_entrenamiento()

    elif opcion == "2":
        print("\nEjecutando previsión, comparación y decisión de cobertura...")
        ejecutar_prevision()

    elif opcion == "3":
        print("\nEjecutando posibles variaciones del mercado...")
        ejecutar_variaciones_mercado()

    else:
        print("\nOpción no válida. No se ha ejecutado ningún proceso.")

    print("\nPrograma finalizado correctamente.")


if __name__ == "__main__":
    main()