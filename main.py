import xlwings as xw

from funciones_excel import (
      cargar_excel,
      llamadas_a_todo_lo_de_comparativos,
      letra_valida,
      obtener_carpeta_aplicacion,
      total_venta_lote
)

def main():
    print("Automatizar Excel Ingeco - Version 1.2")
    carpeta_proyecto = obtener_carpeta_aplicacion()
    (carpeta_proyecto / "entrada").mkdir(exist_ok=True)
    (carpeta_proyecto / "salida").mkdir(exist_ok=True)
    archivos_excel = sorted(
        archivo
        for archivo in (carpeta_proyecto / "entrada").iterdir()
        if archivo.is_file()
        and archivo.suffix.lower() in {".xlsx", ".xlsm", ".xls", ".xlsb"}
        and not archivo.name.startswith("~$")
    )

    if not archivos_excel:
        print("No hay ningún Excel en la carpeta entrada. Coloca uno y vuelve a ejecutar el programa.")
        return

    if len(archivos_excel) > 1:
        print("Hay varios Excel en la carpeta entrada. Deja solo el que quieras procesar:")
        for archivo in archivos_excel:
            print(f"  - {archivo.name}")
        return

    ruta_entrada = archivos_excel[0]
    nombre_excel = ruta_entrada.name
    print(f"Excel seleccionado: {nombre_excel}")

    ruta_salida = (carpeta_proyecto / "salida" / nombre_excel)

    ruta_salida.parent.mkdir(parents=True, exist_ok=True)

    # Abrimos una instancia de Excel sin mostrar su ventana.
    with xw.App(visible=False, add_book=False) as app:
        wb = cargar_excel(app, ruta_entrada)

        try:
            total_venta_lote(wb, ruta_salida)

            while True:
                try:
                    texto = input(
                        "¿Qué rango quieres procesar?: "
                    ).strip().upper()

                    inicio, fin = letra_valida(texto)
                    break

                except ValueError as error:
                    print(error)

            llamadas_a_todo_lo_de_comparativos(
                wb, ruta_salida, inicio, fin
            )

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
