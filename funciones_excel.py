# load_workbook sirve para abrir un archivo Excel existente
from openpyxl import load_workbook

def cargar_excel(ruta_entrada):
    wb = load_workbook(ruta_entrada)
    return wb

#Hay que pasarle la hoja COMPRAS para que funcione
def obtener_comparativos(hoja):
    comparativos = []

    #Recorremos todas las filas de esa columna I
    for fila in range(1, hoja.max_row + 1):

        #Metemos el valor de la variable fila dentro del texto I para así después poder acceder a ella y conocer su valor
        celda = hoja[f"{fila}"]
        valor = celda.value

        #isinstance sirve para comprobar si el valor es una variable del mismo tipo que queremos, en este caso string
        if isinstance(valor, str):
            #Con Strip se eliminan los espacios al inicio y al final
            valor = valor.strip()

            if len(valor) == 1 and valor in "ABCDEFGHIJKLMNOPQRSTUVWXYZ":
                comparativos.append({
                    "letra": valor,
                    "trabajo": hoja[f"D{fila}"].value
                }) 
    return comparativos

def llamadas_a_todo_lo_de_comparativos(ruta_entrada, ruta_salida):
    wb = ruta_entrada

    hoja_compras = wb["COMPRAS"]

    #Obtenemos todos los comparativos para despues poder crear cada una de las pestañas
    comparativos = obtener_comparativos(hoja_compras)

    for comparativo in comparativos:  
        crear_pestaña_comparativo(ruta_entrada, comparativo)

    wb.save(ruta_salida)

def crear_pestaña_comparativo(wb, comparativo):
    nombre_hoja = f"{comparativo['letra']}_{comparativo['nombre']}"

    #sheetnames son los nombres de las pestañas
    if nombre_hoja not in wb.sheetnames:
        #Copiamos la plantilla
        nueva_hoja = wb.copy_worksheet(wb["Plantilla Comparativos"])
        #Ponemos título a la plantilla con el nombre correspondiente
        nueva_hoja.title = nombre_hoja





            