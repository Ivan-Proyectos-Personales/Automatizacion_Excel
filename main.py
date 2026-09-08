from pathlib import Path
import xlwings as xw

from funciones_excel import (
      cargar_excel,
      llamadas_a_todo_lo_de_comparativos,
      letra_valida
)

def main():
    while True:
            try:
                texto = input("¿Qué rango quieres procesar:?")
                inicio, fin = letra_valida(texto)
                break
            except ValueError as error:
                 print(error)

    #Construimos las rutas desde la ubicación de este archivo
    carpeta_proyecto = Path(__file__).resolve().parent

    ruta_entrada = (carpeta_proyecto / "entrada" / "_Control de obra_7185.xls")
    ruta_salida = (carpeta_proyecto / "salida" / "_Control de obra_7185.xls_MODIFICADO.xls")

    if not ruta_entrada.is_file():
          raise FileNotFoundError(f"No se encuentra el archivo: {ruta_entrada}")

    ruta_salida.parent.mkdir(parents=True, exist_ok=True)

    # Abrimos una instancia de Excel sin mostrar su ventana.
    with xw.App(visible=False, add_book=False) as app:
        wb = cargar_excel(app, ruta_entrada)
        try:
            llamadas_a_todo_lo_de_comparativos(wb, ruta_salida, inicio, fin)
        finally:
            wb.close()
    print(f"Archivo generado: {ruta_salida}")

if __name__ == "__main__":
    main()