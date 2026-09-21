import unittest
from types import SimpleNamespace
from unittest.mock import MagicMock, Mock, patch

from funciones_excel import (
    actualizar_comparativos_en_presupuesto,
    escribir_nombres_lotes,
    llamadas_a_todo_lo_de_comparativos,
)


def crear_hoja(nombre, primera_fila, codigos):
    hoja = Mock()
    hoja.name = nombre
    ultima_fila = primera_fila + len(codigos) - 1
    hoja.used_range.last_cell.row = ultima_fila
    celdas = {}

    def rango(direccion):
        return celdas.setdefault(direccion, Mock())

    hoja.range.side_effect = rango
    rango(f"A{primera_fila}:A{ultima_fila}").options.return_value.value = codigos
    return hoja, celdas


class ComparativosEnPresupuestoTests(unittest.TestCase):
    def test_reune_todas_las_hojas_sin_duplicados_en_orden_excel(self):
        presupuesto, celdas = crear_hoja("PPTO OBRA", 5, ["1.01", "1.02"])
        hojas = [presupuesto]
        for nombre, codigos in [
            ("AA_Otros", ["1.01", "1.01"]),
            ("Z_Trabajo", ["1.01"]),
            ("A_Tuberias", [" 1.01 ", "1.02"]),
            ("A_Otra", ["1.01"]),
        ]:
            hojas.append(crear_hoja(nombre, 9, codigos)[0])

        actualizar_comparativos_en_presupuesto(SimpleNamespace(sheets=hojas), presupuesto)

        self.assertEqual(celdas["K5"].value, "A, Z, AA")
        self.assertEqual(celdas["K6"].value, "A")

    def test_limpia_referencias_obsoletas_y_respeta_filas_sin_codigo(self):
        presupuesto, celdas = crear_hoja("PPTO", 5, ["1.01", None, "", "1.02"])
        presupuesto.range("K5").value = "B"
        presupuesto.range("K6").value = "Nota"
        presupuesto.range("K7").value = "Otra nota"
        ajena, _ = crear_hoja("Resumen_Obra", 9, ["1.01"])
        plantilla, _ = crear_hoja("Plantilla Comparativos", 9, ["1.02"])

        actualizar_comparativos_en_presupuesto(
            SimpleNamespace(sheets=[presupuesto, ajena, plantilla]), presupuesto
        )

        self.assertIsNone(celdas["K5"].value)
        self.assertIsNone(celdas["K8"].value)
        self.assertEqual(celdas["K6"].value, "Nota")
        self.assertEqual(celdas["K7"].value, "Otra nota")

    def test_distingue_codigos_y_admite_una_sola_fila(self):
        presupuesto, celdas = crear_hoja("PPTO", 5, ["A1.01a"])
        comparativo, _ = crear_hoja("AA_Trabajo", 9, ["A1.01a"])
        otro, _ = crear_hoja("B_Trabajo", 9, ["a1.01a", "A1.1a"])

        actualizar_comparativos_en_presupuesto(
            SimpleNamespace(sheets=[presupuesto, comparativo, otro]), presupuesto
        )

        self.assertEqual(celdas["K5"].value, "AA")


class NombresLotesTests(unittest.TestCase):
    def test_ocho_lotes_y_reutilizacion_con_un_solo_lote(self):
        hoja, celdas = crear_hoja("AA_Trabajo", 9, [])
        wb = SimpleNamespace(sheets=[hoja])
        nombres = [f"Lote {numero}" for numero in range(1, 9)]

        escribir_nombres_lotes(wb, "AA", nombres)

        destinos = ("D2", "E2", "F2", "G2", "H2", "I2", "J2", "K2")
        self.assertEqual([celdas[d].value for d in destinos], nombres)
        celdas["L2:R2"].clear_contents.assert_called_once()
        for destino in destinos[1:]:
            celdas[destino].paste.assert_called_once_with(paste="formats")
            self.assertTrue(celdas[destino].api.WrapText)

        escribir_nombres_lotes(wb, "AA", ["Nuevo lote"])

        self.assertEqual(celdas["D2"].value, "Nuevo lote")
        self.assertTrue(all(celdas[d].value is None for d in destinos[1:]))

    def test_mas_de_ocho_lotes_no_escribe_una_lista_incompleta(self):
        hoja, _ = crear_hoja("A_Trabajo", 9, [])
        hoja.range.reset_mock()
        with self.assertRaisesRegex(ValueError, "más de 8 lotes"):
            escribir_nombres_lotes(SimpleNamespace(sheets=[hoja]), "A", ["Lote"] * 9)
        hoja.range.assert_not_called()

    def test_flujo_agrupa_filas_no_consecutivas_en_orden_de_compras(self):
        presupuesto, _ = crear_hoja("PPTO", 5, [])
        compras, _ = crear_hoja("COMPRAS", 6, [])
        a, celdas_a = crear_hoja("A_Trabajo", 9, [])
        aa, celdas_aa = crear_hoja("AA_Trabajo", 9, [])
        hojas = [compras, presupuesto, a, aa]
        wb = Mock()
        wb.sheets = MagicMock()
        wb.sheets.__iter__.side_effect = lambda: iter(hojas)
        wb.sheets.__getitem__.side_effect = lambda nombre: next(
            h for h in hojas if h.name == nombre
        )
        comparativos = [
            {"letra": "A", "trabajo": "Tuberías", "partidas": None},
            {"letra": "AA", "trabajo": "Electricidad", "partidas": None},
            {"letra": "A", "trabajo": "Accesorios", "partidas": None},
        ]
        with (
            patch("funciones_excel.obtener_comparativos", return_value=comparativos),
            patch("funciones_excel.buscar_partidas", return_value=[]),
            patch("funciones_excel.actualizar_comparativos_en_presupuesto"),
        ):
            llamadas_a_todo_lo_de_comparativos(wb, "resultado.xls", 1, 27)

        self.assertEqual(celdas_a["D2"].value, "Tuberías")
        self.assertEqual(celdas_a["E2"].value, "Accesorios")
        self.assertIsNone(celdas_a["F2"].value)
        self.assertEqual(celdas_aa["D2"].value, "Electricidad")
        self.assertIsNone(celdas_aa["E2"].value)
        wb.save.assert_called_once_with("resultado.xls")


if __name__ == "__main__":
    unittest.main()
