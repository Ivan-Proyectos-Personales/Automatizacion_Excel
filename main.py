from pathlib import Path
import xlwings as xw

from funciones_excel import (
      cargar_excel,
      llamadas_a_todo_lo_de_comparativos,
      letra_valida,
      obtener_carpeta_aplicacion
)

def main():
    carpeta_proyecto = obtener_carpeta_aplicacion()
    (carpeta_proyecto / "entrada").mkdir(exist_ok=True)
    (carpeta_proyecto / "salida").mkdir(exist_ok=True)
    while True:
        nombre_excel = input("Nombre del Excel (con extensión): ").strip()
        ruta_entrada = carpeta_proyecto / "entrada" / nombre_excel
        if ruta_entrada.is_file():
             break
        
        print(f"No se encuentra el archivo: {ruta_entrada}")
    
    while True:
            try:
                texto = input("¿Qué rango quieres procesar:?")
                inicio, fin = letra_valida(texto)
                break
            except ValueError as error:
                 print(error)

    ruta_salida = (carpeta_proyecto / "salida" / nombre_excel)

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
    try:
        main()
    except Exception:
        import traceback
        traceback.print_exc()
    finally:
        input("\nPulsa Enter para cerrar...")