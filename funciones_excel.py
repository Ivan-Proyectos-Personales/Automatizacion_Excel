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

def total_venta_lote(wb, ruta_salida):
    hoja_compras = wb.sheets["COMPRAS"]
    hoja_presupuesto = None
    for hoja in wb.sheets:
        if hoja.name.strip().upper().startswith("PPTO"):
            hoja_presupuesto = hoja
            break

    if hoja_presupuesto is None:
        raise ValueError("No se encuentra ninguna pestaña que empiece por PPTO")

    while True:
            texto = input(
                "¿Rellenar los importes de cada lote? (S/N): "
            ).strip().upper()
            if(texto == "S"):
                rellenar_importes_compras(hoja_compras, hoja_presupuesto)
                break
            elif texto == "N":
                break
            else:
                print("Introduce S o N")


def llamadas_a_todo_lo_de_comparativos(wb, ruta_salida, inicio, fin):
    hoja_compras = wb.sheets["COMPRAS"]
    hoja_presupuesto = None
    for hoja in wb.sheets:
        if hoja.name.strip().upper().startswith("PPTO"):
            hoja_presupuesto = hoja
            break

    if hoja_presupuesto is None:
        raise ValueError("No se encuentra ninguna pestaña que empiece por PPTO")

    
    comparativos = obtener_comparativos(hoja_compras, inicio, fin)
    nombres_por_letra = {}
    for comparativo in comparativos:
        nombres = nombres_por_letra.setdefault(comparativo["letra"], [])
        nombres.append(comparativo["trabajo"])
        if len(nombres) > 8:
            raise ValueError(
                f"El comparativo {comparativo['letra']} tiene más de 8 lotes. "
                "Solo hay espacio para sus nombres en D2:K2."
            )

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
        escribir_nombres_lotes(wb, letra, nombres_por_letra[letra])
        escribir_partidas(wb, letra, partidas)

    actualizar_comparativos_en_presupuesto(wb, hoja_presupuesto)
    ordenar_pestanas_comparativos(wb)
    # Excel recuerda la pestaña activa al guardar el libro.
    hoja_compras.activate()
    wb.save(str(ruta_salida))

import re 
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

              if not inicio or not fin:
                  raise ValueError(f"Falta una partida en: {bloque}")

              if re.fullmatch(r"[a-zA-Z]+", fin):
                  # 2.11a-e → 2.11a-2.11e
                  coincidencia = re.fullmatch(
                      r"([a-zA-Z]?\d+(?:\.\d+)*)([a-zA-Z]+)",
                      inicio,
                  )

                  if coincidencia is None:
                      raise ValueError(
                          f"El inicio debe terminar en letra: {bloque}"
                      )

                  fin = coincidencia.group(1) + fin

              elif re.fullmatch(r"\d+[a-zA-Z]*", fin):
                  # 1.02-06 → 1.02-1.06
                  # 7.01.02-04 → 7.01.02-7.01.04
                  # 2.11a-11e → 2.11a-2.11e
                  prefijo, separador, _ = inicio.rpartition(".")

                  if not separador:
                      raise ValueError(
                          f"Escribe los códigos completos: {bloque}"
                      )

                  fin = f"{prefijo}.{fin}"

          else:
              raise ValueError(
                  f"Formato de partidas incorrecto: {bloque}"
              )

          patron = r"[a-zA-Z]?\d+(?:\.\d+)+[a-zA-Z]*"

          if not (
              re.fullmatch(patron, inicio)
              and re.fullmatch(patron, fin)
          ):
              raise ValueError(
                  f"Formato de partidas incorrecto: {bloque}"
              )

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

def escribir_nombres_lotes(wb, letra, nombres):
    """Actualiza la cabecera por filas de COMPRAS, también en hojas reutilizadas."""
    destinos = ("D2", "E2", "F2", "G2", "H2", "I2", "J2", "K2")
    if len(nombres) > len(destinos):
        raise ValueError(f"El comparativo {letra} tiene más de 8 lotes.")

    for hoja in wb.sheets:
        if hoja.name.startswith(f"{letra}_"):
            referencia = hoja.range("D2")
            # Retira los nombres colocados por la versión anterior.
            hoja.range("L2:R2").clear_contents()
            for posicion, destino in enumerate(destinos):
                celda = hoja.range(destino)
                if posicion < len(nombres):
                    if posicion:
                        referencia.copy()
                        celda.paste(paste="formats")
                    celda.api.WrapText = True
                # Limpia los nombres sobrantes si ahora se incluyen menos lotes.
                celda.value = (
                    nombres[posicion] if posicion < len(nombres) else None
                )
            hoja.range("D2:K2").rows.autofit()
            return

    raise ValueError(f"No se encuentra la pestaña del comparativo {letra}")


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

def actualizar_comparativos_en_presupuesto(wb, hoja_presupuesto):
    """Anota en K los prefijos de las hojas que contienen cada código de PPTO."""
    letras_por_codigo = {}
    for hoja in wb.sheets:
        prefijo, separador, _ = hoja.name.partition("_")
        if not separador or not re.fullmatch(r"[A-Z]{1,2}", prefijo):
            continue

        ultima_fila = hoja.used_range.last_cell.row
        if ultima_fila < 9:
            continue

        codigos = hoja.range(f"A9:A{ultima_fila}").options(ndim=1).value
        for codigo in codigos:
            if codigo is None or not str(codigo).strip():
                continue
            codigo = str(codigo).strip()
            letras_por_codigo.setdefault(codigo, set()).add(prefijo)

    ultima_fila = hoja_presupuesto.used_range.last_cell.row
    if ultima_fila < 5:
        return

    codigos = hoja_presupuesto.range(f"A5:A{ultima_fila}").options(ndim=1).value
    for fila, codigo in enumerate(codigos, start=5):
        if codigo is None or not str(codigo).strip():
            continue
        letras = letras_por_codigo.get(str(codigo).strip(), set())
        celda = hoja_presupuesto.range(f"K{fila}")
        if letras:
            # Copiamos solo el formato del ejemplo, conservando las letras calculadas.
            hoja_presupuesto.range("J117").copy()
            celda.paste(paste="formats")
        # Recalculamos para quitar referencias a comparativos ya eliminados.
        celda.value = (
            ", ".join(sorted(letras, key=letra_a_numero)) or None
        )


def rellenar_importes_compras(hoja_compras, hoja_presupuesto):
    ultima_fila = hoja_compras.used_range.last_cell.row

    for fila in range(6, ultima_fila + 1):
        texto = hoja_compras.range(f"E{fila}").value

        if str(texto).strip().casefold() == "posible aumento":
            continue

        intervalos = obtener_intervalos_partidas(texto)

        if not intervalos:
            continue

        partidas = buscar_partidas(hoja_presupuesto, intervalos)

        total_objetivo = sum(partida[8] or 0 for partida in partidas)

        total_presupuesto = sum(partida[5] or 0 for partida in partidas)

        hoja_compras.range(f"U{fila}").value = total_objetivo
        hoja_compras.range(f"S{fila}").value = total_presupuesto

def ordenar_pestanas_comparativos(wb):
    hojas = list(wb.sheets)

    if len(hojas) <= 4:
        return

    # Primero las hojas con guion bajo; en cada grupo, de Z a A
    # por el primer carácter, sin distinguir mayúsculas.
    hojas_ordenadas = sorted(
        hojas[4:],
        key=lambda hoja: ("_" in hoja.name, hoja.name[0].casefold()),
        reverse=True
    )

    # Colocamos las hojas ordenadas después de la cuarta.
    anterior = hojas[3]

    for hoja in hojas_ordenadas:
        hoja.api.Move(After=anterior.api)
        anterior = hoja
