"""Pruebas del reparto por asesor y de los Excel generados, con datos y maestros SINTETICOS.

Uso (desde la raiz del proyecto):
    .venv/bin/python -m unittest discover -s Transformacion/tests -v
"""
from __future__ import annotations

import hashlib
import sys
import tempfile
import time
import unittest
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

CARPETA = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(CARPETA))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import generar_archivos_asesores as gen  # noqa: E402
import limpieza_licitaciones as lim  # noqa: E402
import maestros as mae  # noqa: E402
import obras_asesores as oa  # noqa: E402
from fabrica_maestros import crear_maestros  # noqa: E402
from test_limpieza import JWT_FALSO, escribir_csv, ganador, proceso  # noqa: E402

BOGOTA = "Cundinamarca : Bogotá D.C."


def adjudicado(id_contrato: str, entidad: str, ganador_nombre: str = "EMPRESA NUEVA SAS", **cambios) -> dict:
    ganadores = [ganador(id_contrato, int(id_contrato), ganador_nombre)]
    base = dict(EntidadContratante=entidad, ganadores=ganadores, proponentes=ganadores, TextoDepartamento=BOGOTA, cronogramas=[])
    base.update(cambios)
    return proceso(id_contrato, **base)


class RepartoTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)
        self.entrada = self.dir / "output"
        self.entrada.mkdir()
        self.maestros_dir = self.dir / "maestros"

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def _repartir(self, procesos: list[dict], **opciones) -> oa.Reparto:
        crear_maestros(self.maestros_dir, **opciones)
        escribir_csv(self.entrada, "licitaciones_20261108_080000.csv", procesos)
        return oa.repartir(lim.procesar(self.entrada), mae.cargar(self.maestros_dir))

    @staticmethod
    def _por_id(reparto: oa.Reparto) -> dict[str, list[oa.FilaObra]]:
        filas: dict[str, list[oa.FilaObra]] = {}
        for fila in reparto.filas:
            filas.setdefault(fila.proceso["idContrato"], []).append(fila)
        return filas

    def test_reparte_por_entidad_y_clasifica_las_entidades_sin_regla(self) -> None:
        reparto = self._repartir(
            [
                adjudicado("1", "INSTITUTO DE DESARROLLO URBANO"),
                adjudicado("2", "TRANSMILENIO S.A."),
                adjudicado("3", "FONDO DE DESARROLLO LOCAL DE BOSA"),
                adjudicado("4", "ALCALDÍA MUNICIPAL COTA", TextoDepartamento="Cundinamarca : Cota"),
                adjudicado("5", "POLICIA METROPOLITANA DE BOGOTA"),
                adjudicado("6", "ALCALDIA MUNICIPIO DE GARZON", TextoDepartamento="Huila : Garzón"),
                adjudicado("7", "SECRETARÍA DISTRITAL DE CULTURA, RECREACIÓN Y DEPORTE"),
                adjudicado("8", "INVIAS", TextoDepartamento="Cundinamarca : Madrid", Objeto="GESTIÓN VIAL INTEGRAL DE LA CARRETERA"),
                adjudicado("9", "DEPARTAMENTO DE CUNDINAMARCA -SECRETARIA DE GOBIERNO", TextoDepartamento="Cundinamarca"),
                adjudicado("10", "FISCALIA GENERAL DE LA NACION", Objeto="MANTENIMIENTO PREVENTIVO Y CORRECTIVO DE LAS SEDES"),
                adjudicado("11", "FISCALIA GENERAL DE LA NACION", Objeto="MANTENIMIENTO, ADECUACIÓN Y REMODELACIÓN DE LA SEDE"),
                adjudicado("12", "UNP", TextoDepartamento="Colombia: Todo el pais"),
            ]
        )
        filas = self._por_id(reparto)
        destino = {i: (f[0].archivo.ruta.name, f[0].hoja.nombre, f[0].valores["E"]) for i, f in filas.items()}
        self.assertEqual(destino["1"], ("Obras_IDU_Prueba.xlsx", "IDU", "Gustavo"))
        self.assertEqual(destino["2"], ("Obras_IDU_Prueba.xlsx", "IDU", "Gustavo"))
        self.assertEqual(destino["3"], ("Obras_Alcaldias_Prueba.xlsx", "Alcaldías Locales", "Mariana"))
        self.assertEqual(destino["4"], ("Obras_Alcaldias_Prueba.xlsx", "Municipios Aledaños", "Mariana"))
        self.assertEqual(reparto.sin_clasificar, [])
        self.assertEqual(sorted(p["idContrato"] for p in reparto.fuera_de_zona), ["12", "6"])
        self.assertEqual([p["idContrato"] for p in reparto.sin_consumo], ["10"])
        sin_asignar = sorted(i for i, f in filas.items() if f[0].archivo is reparto.archivo_sin_asignar)
        self.assertEqual(sin_asignar, ["11", "5", "7"])
        self.assertEqual(filas["5"][0].valores["E"], oa.ASESOR_SIN_ASIGNAR)
        self.assertEqual(filas["5"][0].valores["A"], "POLICIA METROPOLITANA DE BOGOTA")
        self.assertEqual(filas["5"][0].hoja.nombre, oa.HOJA_SIN_ASIGNAR)
        self.assertEqual(destino["8"][1:], ("Municipios Aledaños", "Mariana"))
        self.assertEqual(filas["8"][0].valores["B"], "Madrid")
        self.assertEqual(destino["9"][1], "Municipios Aledaños")
        self.assertEqual(filas["9"][0].valores["B"], "Cundinamarca")
        self.assertEqual(filas["3"][0].valores["B"], "Bosa")
        self.assertEqual(filas["4"][0].valores["B"], "Cota")
        self.assertEqual(filas["1"][0].valores["A"], "IDU")
        self.assertEqual(filas["3"][0].valores["A"], "Alcaldía Local")

    def test_no_repite_obras_que_ya_estan_en_los_maestros(self) -> None:
        existentes = [
            {"codigo": "FDLBOSA-LIC-001-2026", "empresa": "CONSORCIO PARQUES P Y P", "valor": 900},
            {"codigo": "LP-001-2026", "empresa": "OTRA EMPRESA SAS", "valor": 50},
        ]
        reparto = self._repartir(
            [
                adjudicado("1", "FONDO DE DESARROLLO LOCAL DE BOSA", "CONSORCIO PARQUES P Y P", CodigoProceso="FDLBOSA LIC 001-2026-"),
                adjudicado("2", "ALCALDIA LOCAL DE SUBA", "EMPRESA DISTINTA SAS", CodigoProceso="LP-001-2026", Valor="2000000000"),
            ],
            existentes_alcaldias=existentes,
        )
        self.assertEqual([p["idContrato"] for p in reparto.ya_en_maestros], ["1"])
        self.assertIn("Obras_Alcaldias_Prueba.xlsx / Alcaldías Locales (fila 3)", reparto.ya_en_maestros[0]["Ya está en"])
        self.assertEqual([f.proceso["idContrato"] for f in reparto.filas], ["2"])
        self.assertTrue(any(i["Campo"] == "CodigoProceso" and i["idContrato"] == "2" for i in reparto.incidencias))

    def test_cliente_de_sugeine_va_a_los_dos_archivos_y_el_parecido_queda_por_confirmar(self) -> None:
        reparto = self._repartir(
            [
                adjudicado("1", "INSTITUTO DE DESARROLLO URBANO", "Consorcio Vías CB 2024 S.A.S."),
                adjudicado("2", "ALCALDIA LOCAL DE KENNEDY", "INCITECO LTDA"),
                adjudicado("3", "ALCALDIA LOCAL DE USME", "CONSORCIO VIAS CB 2025"),
                adjudicado("4", "ALCALDIA LOCAL DE USME", "EMPRESA SIN RELACION SAS"),
                adjudicado("5", "ALCALDIA LOCAL DE USME", "CONSORCIO VIAS NORTE"),  # solo comparte "CONSORCIO VIAS": no es parecido
            ]
        )
        filas = self._por_id(reparto)
        uno = {(f.archivo.ruta.name, f.valores["E"]) for f in filas["1"]}
        self.assertEqual(uno, {("Obras_IDU_Prueba.xlsx", "Sugeine"), ("Obras_Sugeine_Prueba.xlsx", "Sugeine")})
        self.assertEqual({f.valores["BI"] for f in filas["1"]}, {oa.CONFIRMADO})
        self.assertEqual({f.valores["BM"] for f in filas["1"]}, {"Gustavo"})
        self.assertEqual(len(filas["2"]), 2)
        self.assertEqual([(f.archivo.ruta.name, f.valores["E"]) for f in filas["3"]], [("Obras_Alcaldias_Prueba.xlsx", "Mariana (posible Sugeine)")])
        self.assertEqual(filas["3"][0].valores["BI"], oa.POR_CONFIRMAR)
        self.assertEqual([f.valores["E"] for f in filas["4"]], ["Mariana"])
        self.assertNotIn("BI", filas["4"][0].valores)
        self.assertEqual([f.valores["E"] for f in filas["5"]], ["Mariana"])

    def test_entidades_sin_regla_se_reparten_por_tipo_de_obra(self) -> None:
        reparto = self._repartir(
            [
                adjudicado("1", "FISCALIA GENERAL DE LA NACION", Objeto="ADECUACIÓN Y REMODELACIÓN DE LA SEDE"),
                adjudicado("2", "POLICIA METROPOLITANA DE BOGOTA", Objeto="CONSTRUCCIÓN DE ANDENES Y MALLA VIAL"),
                adjudicado("3", "SECRETARIA DISTRITAL DE INTEGRACION SOCIAL", Objeto="CONSTRUCCIÓN DEL PARQUE INFANTIL"),
                adjudicado("4", "CLUB MILITAR", Objeto="REPOSICIÓN DE REDES DE ALCANTARILLADO"),
                adjudicado("5", "SENA", Objeto="CONSTRUCCIÓN DE AULAS", TextoDepartamento="Cundinamarca : Chía"),
                adjudicado("6", "MINISTERIO DE DEFENSA NACIONAL"),  # "Obra de prueba": sin tipo claro
            ],
            con_idrd=True,
        )
        filas = self._por_id(reparto)
        destino = {i: (f[0].archivo.ruta.name, f[0].hoja.nombre, f[0].valores["E"]) for i, f in filas.items()}
        self.assertEqual(destino["1"], ("Obras_IDRD_Prueba.xlsx", "IDRD", "Asesor 3 (por contratar)"))
        self.assertEqual(filas["1"][0].valores["L"], oa.TIPO_EDIFICACIONES)
        self.assertEqual(filas["1"][0].valores["A"], "FISCALIA GENERAL DE LA NACION")
        self.assertEqual(destino["2"][1:], ("IDU", "Gustavo"))
        self.assertEqual(destino["3"][1:], ("IDRD", "Asesor 3 (por contratar)"))
        self.assertEqual(destino["4"][1], oa.HOJA_SIN_ASIGNAR)  # los maestros de prueba no tienen hoja Acueducto
        self.assertEqual(destino["5"][1], "Municipios Aledaños")
        self.assertEqual(destino["6"][1], oa.HOJA_SIN_ASIGNAR)

    def test_inicio_y_meses_desde_el_cronograma_solo_si_son_claros(self) -> None:
        def hito(nombre: str, fecha: str) -> dict:
            return {"nombre": nombre, "fecha": fecha}

        reparto = self._repartir(
            [
                adjudicado("1", "INSTITUTO DE DESARROLLO URBANO", cronogramas=[
                    hito("Fecha de inicio de ejecución del contrato", "15/11/2026 - 10:00 AM"),
                    hito("Fecha de inicio de ejecución del contrato", "20/11/2026 - 10:00 AM"),  # adenda, mismo mes
                    hito("Plazo de ejecución del contrato", "30/04/2027 - 11:59 PM"),
                ]),
                adjudicado("2", "INSTITUTO DE DESARROLLO URBANO", cronogramas=[
                    hito("Fecha de inicio de ejecución del contrato", "15/11/2026 - 10:00 AM"),
                    hito("Fecha de inicio de ejecución del contrato", "10/01/2027 - 10:00 AM"),  # adenda, otro mes
                    hito("Plazo de ejecución del contrato", "30/04/2027 - 11:59 PM"),
                ]),
                adjudicado("3", "INSTITUTO DE DESARROLLO URBANO", cronogramas=[
                    hito("Fecha de inicio de ejecución del contrato", "01/08/2026 - 10:00 AM"),  # antes de la proyeccion
                    hito("Ejecución del Contrato", "31/12/2026 - 11:59 PM"),
                ]),
                adjudicado("4", "INSTITUTO DE DESARROLLO URBANO", cronogramas=[
                    hito("Inicio de Ejecución del Contrato", "05/12/2026 - 10:00 AM"),
                ]),
                adjudicado("5", "INSTITUTO DE DESARROLLO URBANO"),
            ]
        )
        filas = {i: f[0].valores for i, f in self._por_id(reparto).items()}
        self.assertEqual((filas["1"]["R"], filas["1"]["Q"]), (datetime(2026, 11, 1), 6))
        self.assertEqual((filas["2"]["R"], filas["2"]["Q"]), (datetime(2026, 11, 1), None))  # mes de la extraccion
        self.assertEqual((filas["3"]["R"], filas["3"]["Q"]), (datetime(2026, 10, 1), 3))  # oct a dic 2026
        self.assertEqual((filas["4"]["R"], filas["4"]["Q"]), (datetime(2026, 12, 1), None))
        self.assertEqual((filas["5"]["R"], filas["5"]["Q"]), (datetime(2026, 11, 1), None))
        self.assertTrue(all(v["P"] == 0 and v["BB"] == 0 for v in filas.values()))
        self.assertIn("programados según el cronograma", filas["1"]["J"])
        self.assertIn("completar Meses por Terminar", filas["2"]["J"])

    def test_tipo_de_contrato_propuesto_por_palabras_clave(self) -> None:
        permitidos = oa.TIPO_POR_DEFECTO, "Interventoría", "Parques y escenarios deportivos", "Conservación, mantenimiento y reparcheo", "Construcción de obra"
        casos = {
            "INTERVENTORÍA TÉCNICA A LA CONSTRUCCIÓN DEL PARQUE": "Interventoría",
            "CONSTRUCCIÓN DEL PARQUE VECINAL EL TINTAL": "Parques y escenarios deportivos",
            "MANTENIMIENTO DE LA MALLA VIAL LOCAL": "Conservación, mantenimiento y reparcheo",
            "CONSTRUCCIÓN DE ANDENES": "Construcción de obra",
            "SUMINISTRO DE PAPELERÍA": oa.TIPO_POR_DEFECTO,
        }
        for objeto, esperado in casos.items():
            self.assertEqual(oa.tipo_contrato(objeto, list(permitidos)), esperado, objeto)


class ArchivosGeneradosTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)
        self.entrada = self.dir / "output"
        self.entrada.mkdir()
        self.maestros_dir = self.dir / "maestros"
        self.rutas_maestros = crear_maestros(self.maestros_dir)
        self.salida = self.dir / "salida"
        escribir_csv(
            self.entrada,
            "licitaciones_20261108_080000.csv",
            [
                adjudicado("1", "INSTITUTO DE DESARROLLO URBANO"),
                adjudicado("2", "INSTITUTO DE DESARROLLO URBANO", "INCITECO"),
                adjudicado("3", "FONDO DE DESARROLLO LOCAL DE BOSA"),
                adjudicado("4", "MINISTERIO DE DEFENSA NACIONAL"),
            ],
        )

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_genera_un_archivo_por_asesor_con_la_plantilla_del_maestro(self) -> None:
        huellas = {r.name: hashlib.sha256(r.read_bytes()).hexdigest() for r in self.rutas_maestros}

        g = gen.generar(self.entrada, self.maestros_dir, self.salida, "Prueba", actualizar_maestros=False)

        self.assertEqual({r.name: hashlib.sha256(r.read_bytes()).hexdigest() for r in self.rutas_maestros}, huellas)
        nombres = sorted(a.ruta.name for a in g.archivos_asesores)
        self.assertEqual(len(nombres), 4)
        self.assertTrue(all(n.startswith("Nuevas_") for n in nombres))
        sin_asignar = next(a for a in g.archivos_asesores if a.maestro.asesor == oa.ASESOR_SIN_ASIGNAR)
        self.assertEqual(sin_asignar.ruta.name.split("_2026")[0], "Nuevas_Sin_asignar_GRECO")
        ws_sa = load_workbook(sin_asignar.ruta)[oa.HOJA_SIN_ASIGNAR]
        self.assertEqual((ws_sa["A3"].value, ws_sa["E3"].value), ("MINISTERIO DE DEFENSA NACIONAL", oa.ASESOR_SIN_ASIGNAR))
        self.assertTrue(str(ws_sa["S3"].value).startswith("=IF("))
        idu = next(a for a in g.archivos_asesores if a.maestro.asesor == "Gustavo")
        libro = load_workbook(idu.ruta)
        self.assertEqual(libro.sheetnames, ["Instrucciones", "IDU", "Listas"])
        ws = libro["IDU"]
        self.assertEqual(ws["T1"].value, 0.4)
        self.assertEqual(ws["BC1"].value, "=SUM(BC3:BC4)")
        self.assertEqual(ws["S2"].value, "Valor por Ejecutar")
        self.assertEqual([ws[f"E{r}"].value for r in (3, 4)], ["Gustavo", "Sugeine"])
        self.assertEqual(ws["S4"].value, '=IF(OR(L4="Conservación, mantenimiento y reparcheo",L4="Repetido",E4="Sugeine"),0,+(1-P4)*N4)')
        self.assertEqual(ws["X4"].value, '=IF(AND(X$2>=$BN4,X$2<=$BO4),$W4,"")')
        self.assertEqual(ws["M3"].value, "PRUEBA-1")
        self.assertEqual(ws["N3"].value, 1500.0)
        self.assertEqual(ws["R3"].number_format, "mmm-yy")
        self.assertEqual(ws["BU2"].value, "SAML · Entidad contratante")
        self.assertEqual(ws["BU3"].value, "INSTITUTO DE DESARROLLO URBANO")
        self.assertEqual(ws.freeze_panes, "A3")
        self.assertEqual({str(v.sqref) for v in ws.data_validations.dataValidation}, {"L3:L4", "R3:R4", "E3:E4"})

        sugeine = next(a for a in g.archivos_asesores if a.maestro.es_sugeine)
        self.assertEqual(sugeine.por_hoja, {"IDU": 1})

        revision = load_workbook(g.revision)
        self.assertEqual(
            revision.sheetnames,
            ["Resumen", "Sin clasificar", "Sin consumo de material", "Fuera de zona", "Ya en maestros", "Procesos", "Ganadores", "Incidencias"],
        )
        self.assertEqual(revision["Sin clasificar"].max_row, 1)
        textos = [str(c.value) for archivo in [*g.archivos_asesores] for hoja in load_workbook(archivo.ruta) for fila in hoja.iter_rows() for c in fila if c.value]
        textos += [str(c.value) for hoja in revision for fila in hoja.iter_rows() for c in fila if c.value]
        self.assertFalse(any(JWT_FALSO in t for t in textos))

    def test_cada_ejecucion_crea_una_carpeta_nueva_sin_tocar_la_anterior(self) -> None:
        primera = gen.generar(self.entrada, self.maestros_dir, self.salida, actualizar_maestros=False)
        contenido = {p.name: p.read_bytes() for p in primera.carpeta.iterdir()}
        time.sleep(1.1)
        segunda = gen.generar(self.entrada, self.maestros_dir, self.salida, actualizar_maestros=False)
        self.assertNotEqual(primera.carpeta, segunda.carpeta)
        self.assertEqual({p.name: p.read_bytes() for p in primera.carpeta.iterdir()}, contenido)

    def test_edificaciones_van_a_la_hoja_idrd_existente(self) -> None:
        crear_maestros(self.maestros_dir, con_idrd=True)
        escribir_csv(
            self.entrada,
            "licitaciones_20261108_090000.csv",
            [adjudicado("9", "FISCALIA GENERAL DE LA NACION", Objeto="REMODELACIÓN DE LA SEDE CENTRAL")],
        )
        g = gen.generar(self.entrada, self.maestros_dir, self.salida, actualizar_maestros=False)
        archivo = next(a for a in g.archivos_asesores if a.maestro.asesor == "Asesor 3 (por contratar)")
        libro = load_workbook(archivo.ruta)
        self.assertEqual(libro.sheetnames, ["Instrucciones", "IDRD", "Listas"])
        ws = libro["IDRD"]
        self.assertEqual((ws["T1"].value, ws["U1"].value), (0.4, 0.23))
        self.assertEqual((ws["A3"].value, ws["L3"].value), ("FISCALIA GENERAL DE LA NACION", oa.TIPO_EDIFICACIONES))

    def test_sin_maestros_falla_con_mensaje_claro(self) -> None:
        with self.assertRaises(mae.ErrorMaestros):
            gen.generar(self.entrada, self.dir / "no_existe", self.salida, actualizar_maestros=False)


if __name__ == "__main__":
    unittest.main()
