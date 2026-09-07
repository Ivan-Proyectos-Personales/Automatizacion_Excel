from funciones_excel import cargar_excel
from funciones_excel import llamadas_a_todo_lo_de_comparativos

ruta_entrada = "_Control de obra_7185.xls"
ruta_salida = "_Control de obra_7185_MODIFICADO.xls"

wb = cargar_excel(ruta_entrada)

llamadas_a_todo_lo_de_comparativos(wb, ruta_salida)
