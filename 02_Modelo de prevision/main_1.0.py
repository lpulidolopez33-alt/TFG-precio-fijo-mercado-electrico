

from pathlib import Path                                # con from traemos desde la libreria x la herramienta con la funcion import x
                                                        # Sirve para representar una ruta de archivo o carpeta como un objeto de Python.

BASE_DIR = Path(__file__).resolve().parent              # Esto localiza automáticamente la carpeta donde está guardado main.py.           

DATA_DIR = BASE_DIR / "datos"                           # Crea la ruta de la carpeta datos.
RAW_DATA_DIR = DATA_DIR / "historicos"                  # Esta será la carpeta donde guardaremos los datos originales descargados o importados.
PROCESSED_DATA_DIR = DATA_DIR / "procesados"            # Esta carpeta guardará los datos ya limpiados y preparados para el modelo.

OUTPUT_DIR = BASE_DIR / "salidas"                       # Aquí guardaremos los resultados finales.
GRAPH_DIR = BASE_DIR / "graficas"                       # Aquí guardaremos las gráficas del modelo.


def crear_estructura_carpetas() -> None:                # def xxxxx() sirve para ejecutar una funcion, y el none es para saber que eso que ejecutamos no devuelve ningun resultado solo ejecuta acciones.
    """
    Crea las carpetas principales del proyecto si todavía no existen.
    """

    carpetas = [
        DATA_DIR,
        RAW_DATA_DIR,
        PROCESSED_DATA_DIR,
        OUTPUT_DIR,
        GRAPH_DIR,
    ]

    for carpeta in carpetas:                             # Esto recorre la lista de carpetas una por una.
        carpeta.mkdir(parents=True, exist_ok=True)       # carpeta.mkdir --> lo que hace es crear carpetas 
        print(f"Carpeta preparada: {carpeta}")


def main() -> None:
    """
    Función principal del programa.
    """

    print("Iniciando programa de previsión de commodity...")
    crear_estructura_carpetas()
    print("Estructura inicial creada correctamente.")


if __name__ == "__main__":
    main()