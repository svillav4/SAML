"""Pruebas de las copias actualizadas de los maestros con Excel REAL (se omiten si no hay Excel).

Usa maestros sinteticos; nunca los archivos de GRECO.

Uso (desde la raiz del proyecto):
    .venv/bin/python -m unittest discover -s Transformacion/tests -v
"""
from __future__ import annotations

import hashlib
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import load_workbook

CARPETA = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(CARPETA))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import actualizar_maestros as am  # noqa: E402
import generar_archivos_asesores as gen  # noqa: E402
from fabrica_maestros import crear_maestros  # noqa: E402
from test_asesores import adjudicado  # noqa: E402
from test_limpieza import escribir_csv  # noqa: E402

DISPONIBLE, MOTIVO = am.excel_disponible()


@unittest.skipUnless(DISPONIBLE, f"Sin Excel: {MOTIVO}")
class ActualizarMaestrosTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)
        self.entrada = self.dir / "output"
        self.entrada.mkdir()
        self.maestros_dir = self.dir / "maestros"
        existentes = [{"codigo": "EXISTENTE-1", "empresa": "UNO SAS"}, {"codigo": "EXISTENTE-2", "empresa": "DOS SAS"}]
        rutas = crear_maestros(self.maestros_dir, existentes_alcaldias=existentes)
        # Resumen con rangos fijos sobre las hojas, como en los maestros de GRECO.
        self.ruta_alcaldias = next(r for r in rutas if "Alcaldias" in r.name)
        libro = load_workbook(self.ruta_alcaldias)
        resumen = libro.create_sheet("Resumen", 0)
        resumen["A1"], resumen["B1"] = "# obras", "=COUNTA('Alcaldías Locales'!I3:I4)"
        resumen["A2"], resumen["B2"] = "Valor", "=SUM('Alcaldías Locales'!N3:N4)"
        resumen["A3"], resumen["B3"] = "Municipios", "=COUNTA('Municipios Aledaños'!I3:I3)"
        libro.save(self.ruta_alcaldias)
        self.huella = hashlib.sha256(self.ruta_alcaldias.read_bytes()).hexdigest()
        escribir_csv(
            self.entrada,
            "licitaciones_20261108_080000.csv",
            [
                adjudicado("1", "ALCALDIA LOCAL DE BOSA", "NUEVA UNO SAS"),
                adjudicado("2", "ALCALDIA LOCAL DE SUBA", "NUEVA DOS SAS"),
                adjudicado("3", "ALCALDÍA MUNICIPAL COTA", "NUEVA TRES SAS", TextoDepartamento="Cundinamarca : Cota"),
            ],
        )
        self.g = gen.generar(self.entrada, self.maestros_dir, self.dir / "salida")

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_inserta_las_obras_dentro_de_los_rangos_sin_tocar_el_original(self) -> None:
        self.assertEqual(hashlib.sha256(self.ruta_alcaldias.read_bytes()).hexdigest(), self.huella)
        act = next(a for a in self.g.actualizacion.archivos if "Alcaldias" in a.ruta.name)
        self.assertEqual(act.filas_por_hoja, {"Alcaldías Locales": 2, "Municipios Aledaños": 1})
        libro = load_workbook(act.ruta)
        ws = libro["Alcaldías Locales"]
        codigos = [ws[f"M{r}"].value for r in range(3, 7)]
        self.assertEqual(codigos, ["EXISTENTE-1", "PRUEBA-1", "PRUEBA-2", "EXISTENTE-2"])  # encima de la ultima obra
        self.assertEqual(ws["S4"].value, '=IF(OR(L4="Conservación, mantenimiento y reparcheo",L4="Repetido",E4="Sugeine"),0,+(1-P4)*N4)')
        self.assertEqual(ws["S6"].value, '=IF(OR(L6="Conservación, mantenimiento y reparcheo",L6="Repetido",E6="Sugeine"),0,+(1-P6)*N6)')
        self.assertEqual((ws["E4"].value, ws["B4"].value, ws["N4"].value), ("Mariana", "Bosa", 1500))
        resumen = libro["Resumen"]
        self.assertEqual(resumen["B1"].value, "=COUNTA('Alcaldías Locales'!I3:I6)")  # el rango se amplio solo
        self.assertEqual(resumen["B2"].value, "=SUM('Alcaldías Locales'!N3:N6)")

    def test_hoja_con_una_sola_obra_avisa_si_el_resumen_no_la_cubre(self) -> None:
        act = next(a for a in self.g.actualizacion.archivos if "Alcaldias" in a.ruta.name)
        ws = load_workbook(act.ruta)["Municipios Aledaños"]
        self.assertEqual([ws[f"M{r}"].value for r in (3, 4)], [ws["M3"].value, "PRUEBA-3"])
        self.assertTrue(any("Municipios Aledaños" in aviso and "Resumen!B3" in aviso for aviso in act.avisos), act.avisos)

    def test_desde_una_salida_sin_excel_como_la_de_n8n(self) -> None:
        salida = self.dir / "salida_n8n"
        g = gen.generar(self.entrada, self.maestros_dir, salida, actualizar_maestros=False)  # como en Docker
        self.assertEqual(list(g.carpeta.glob("*_actualizado_*.xlsx")), [])

        resultado = am.desde_carpeta(g.carpeta, self.maestros_dir)

        act = next(a for a in resultado.archivos if "Alcaldias" in a.ruta.name)
        ws = load_workbook(act.ruta)["Alcaldías Locales"]
        self.assertEqual([ws[f"M{r}"].value for r in range(3, 7)], ["EXISTENTE-1", "PRUEBA-1", "PRUEBA-2", "EXISTENTE-2"])
        self.assertEqual(ws["S5"].value, '=IF(OR(L5="Conservación, mantenimiento y reparcheo",L5="Repetido",E5="Sugeine"),0,+(1-P5)*N5)')
        self.assertEqual(hashlib.sha256(self.ruta_alcaldias.read_bytes()).hexdigest(), self.huella)
        # Una segunda vez no duplica.
        otra = am.desde_carpeta(g.carpeta, self.maestros_dir)
        self.assertEqual(otra.archivos, [])
        self.assertIn("ya tiene su copia actualizada", otra.motivo)

    def test_no_actualiza_si_el_maestro_cambio_despues(self) -> None:
        salida = self.dir / "salida_n8n"
        g = gen.generar(self.entrada, self.maestros_dir, salida, actualizar_maestros=False)
        libro = load_workbook(self.ruta_alcaldias)
        libro["Resumen"]["A9"] = "cambio posterior"
        libro.save(self.ruta_alcaldias)

        resultado = am.desde_carpeta(g.carpeta, self.maestros_dir)

        self.assertFalse(any("Alcaldias" in a.ruta.name for a in resultado.archivos))
        self.assertTrue(any("cambió desde que se generó" in e for e in resultado.errores))

    def test_las_instrucciones_apuntan_a_la_copia_actualizada(self) -> None:
        nuevas = next(a for a in self.g.archivos_asesores if "Alcaldias" in a.ruta.name)
        textos = " ".join(str(c.value) for fila in load_workbook(nuevas.ruta)["Instrucciones"].iter_rows() for c in fila if c.value)
        self.assertIn("_actualizado_", textos)
        self.assertIn("No necesita pegar nada", textos)


if __name__ == "__main__":
    unittest.main()
