def cargar_excel(app, ruta_entrada):
    return app.books.open(str(ruta_entrada), update_links=False)


def letra_a_numero(letra):
    letra = letra.strip().upper()
    if not 1 <= len(letra) <= 2:
        raise ValueError("Introduce una o dos letras, desde A hasta ZZ.")
    if any(caracter not in "ABCDEFGHIJKLMNOPQRSTUVWXYZ" for caracter in letra):
        raise ValueError("Solo se permiten letras de la A a la Z.")

    # Orden de Excel: A = 1, Z = 26, AA = 27, ..., ZZ = 702.
    numero = 0
    for caracter in letra:
        numero = numero * 26 + ord(caracter) - ord("A") + 1
    return numero


def letra_valida(texto):
    partes = texto.split("-")
    if len(partes) != 2:
        raise ValueError("Introduce un rango como A-G, A-AA o AA-AG.")
    inicio = letra_a_numero(partes[0])
    fin = letra_a_numero(partes[1])
    if inicio > fin:
        raise ValueError("La letra inicial no puede ir después de la final.")
    return inicio, fin


def obtener_comparativos(hoja, inicio, fin):
    comparativos = []
    ultima_fila = hoja.used_range.last_cell.row
    for fila in range(1, ultima_fila + 1):
        valor = hoja.range(f"I{fila}").value
        if not isinstance(valor, str):
            continue
        valor = valor.strip().upper()
        try:
            numero = letra_a_numero(valor)
        except ValueError:
            # Ignoramos celdas que no sean letras de comparativos.
            continue
        if inicio <= numero <= fin:
            comparativos.append({
                "letra": valor,
                "trabajo": hoja.range(f"D{fila}").value,
            })
    return comparativos


def crear_pestaña_comparativo(wb, comparativo):
    nombre_hoja = f"{comparativo['letra']}_{comparativo['trabajo']}"[:31]

    letras_existentes = {hoja.name.split("_", 1)[0]
                         for hoja in wb.sheets
                         if "_" in hoja.name}

    if comparativo["letra"] not in letras_existentes:
        plantilla = wb.sheets["Plantilla Comparativos"]
        nueva_hoja = plantilla.copy(after=wb.sheets[-1], name=nombre_hoja)
        nueva_hoja.range("D2").value = comparativo["trabajo"]


def llamadas_a_todo_lo_de_comparativos(wb, ruta_salida, inicio, fin):
    hoja_compras = wb.sheets["COMPRAS"]
    comparativos = obtener_comparativos(hoja_compras, inicio, fin)
    for comparativo in comparativos:
        crear_pestaña_comparativo(wb, comparativo)
    wb.save(str(ruta_salida))
