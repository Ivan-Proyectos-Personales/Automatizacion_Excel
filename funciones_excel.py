import sys
from pathlib import Path

def obtener_carpeta_aplicacion():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent

    return Path(__file__).resolve().parent


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
            partidas = hoja.range(f"E{fila}").value

            comparativos.append({
                "letra": valor,
                "trabajo": hoja.range(f"D{fila}").value,
                "partidas": partidas
            })
    return comparativos


def crear_pestaña_comparativo(wb, comparativo):
    nombre_hoja = f"{comparativo['letra']}_{comparativo['trabajo']}"[:20]

    letras_existentes = {hoja.name.split("_", 1)[0]
                         for hoja in wb.sheets
                         if "_" in hoja.name}

    if comparativo["letra"] not in letras_existentes:
        plantilla = wb.sheets["Plantilla Comparativos"]
        nueva_hoja = plantilla.copy(after=wb.sheets[-1], name=nombre_hoja)
        nueva_hoja.range("D2").value = comparativo["trabajo"]


def llamadas_a_todo_lo_de_comparativos(wb, ruta_salida, inicio, fin):
    hoja_compras = wb.sheets["COMPRAS"]
    hoja_presupuesto = wb.sheets["PPTO INIC (E)"]

    comparativos = obtener_comparativos(hoja_compras, inicio, fin)
    partidas_por_letra = {}

    for comparativo in comparativos:
        letra = comparativo["letra"]


        if letra not in partidas_por_letra:
            crear_pestaña_comparativo(wb, comparativo)
            partidas_por_letra[letra] = []

        intervalos = obtener_intervalos_partidas(comparativo["partidas"])
        partidas = buscar_partidas(hoja_presupuesto, intervalos)

        partidas_por_letra[letra].extend(partidas)

    for letra, partidas in partidas_por_letra.items():
        escribir_partidas(wb, letra, partidas)

    rellenar_importes_compras(hoja_compras, hoja_presupuesto)

    wb.save(str(ruta_salida))

def obtener_intervalos_partidas(texto):
    if texto is None or str(texto).strip() == "":
        return []

    intervalos = []

    for bloque in str(texto).split(","):
        partes = bloque.strip().split("-")

        if len(partes) == 1:
            inicio = partes[0].strip()
            fin = inicio
        elif len(partes) == 2:
            inicio = partes[0].strip()
            fin = partes[1].strip()
        else:
            raise ValueError(f"Formato de partidas incorrecto: {bloque}")

        if not inicio or not fin:
            raise ValueError(f"Falta una partida en: {texto}")

        intervalos.append((inicio, fin))

    return intervalos

def buscar_partidas(hoja_presupuesto, intervalos):
    ultima_fila = hoja_presupuesto.used_range.last_cell.row
    filas = hoja_presupuesto.range(f"A5:I{ultima_fila}").value

    posiciones = {}

    for posicion, fila in enumerate(filas):
        codigo = fila[0]

        if codigo is not None:
            codigo = str(codigo).strip()
            posiciones[codigo] = posicion

    partidas = []

    for inicio, fin in intervalos:
        if inicio not in posiciones:
            raise ValueError(f"No se encuentra la partida: {inicio}")

        if fin not in posiciones:
            raise ValueError(f"No se encuentra la partida: {fin}")

        primera = posiciones[inicio]
        ultima = posiciones[fin]

        if primera > ultima:
            raise ValueError(f"El intervalo está invertido: {inicio}-{fin}")

        for posicion in range(primera, ultima + 1):
            partidas.append(filas[posicion])

    return partidas

def escribir_partidas(wb, letra, partidas):
    if not partidas:
        return

    for hoja in wb.sheets:
        if hoja.name.startswith(f"{letra}_"):
            # A, B, C y D del presupuesto
            hoja.range("A9").value = [partida[:4] for partida in partidas]

            # H del presupuesto -> L
            hoja.range("L9").value = [[partida[7]] for partida in partidas]

            # E del presupuesto -> AF
            hoja.range("AF9").value = [[partida[4]] for partida in partidas]
            return

    raise ValueError(f"No se encuentra la pestaña del comparativo{letra}")

def rellenar_importes_compras(hoja_compras, hoja_presupuesto):
    ultima_fila = hoja_compras.used_range.last_cell.row

    for fila in range(6, ultima_fila + 1):
        texto = hoja_compras.range(f"E{fila}").value
        intervalos = obtener_intervalos_partidas(texto)

        if not intervalos:
            continue

        partidas = buscar_partidas(hoja_presupuesto, intervalos)

        total_objetivo = sum(partida[8] or 0 for partida in partidas)

        total_presupuesto = sum(partida[5] or 0 for partida in partidas)

        hoja_compras.range(f"R{fila}").value = total_objetivo
        hoja_compras.range(f"T{fila}").value = total_presupuesto