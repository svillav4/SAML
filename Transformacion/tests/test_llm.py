"""Pruebas de la clasificacion con LLM, SIN llamar a OpenAI (cliente simulado, datos sinteticos).

Uso (desde la raiz del proyecto):
    .venv/bin/python -m unittest discover -s Transformacion/tests -v
"""
from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest import mock

CARPETA = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(CARPETA))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import clasificador_llm as cl  # noqa: E402
import limpieza_licitaciones as lim  # noqa: E402
import maestros as mae  # noqa: E402
import obras_asesores as oa  # noqa: E402
from fabrica_maestros import crear_maestros  # noqa: E402
from test_asesores import adjudicado  # noqa: E402
from test_limpieza import escribir_csv  # noqa: E402


def clasificacion(id_contrato: str, destino: str, **cambios) -> SimpleNamespace:
    base = dict(
        id=id_contrato,
        destino=destino,
        tipo_contrato="Edificaciones y adecuaciones",
        localidad="Teusaquillo",
        consume_material=destino not in (cl.SIN_CONSUMO, cl.FUERA_DE_ZONA),
        cliente_sugeine="NO",
        cliente_sugeine_nombre=None,
        confianza="alta",
        justificacion=f"Prueba: {destino}.",
    )
    base.update(cambios)
    return SimpleNamespace(**base)


class ClienteFalso:
    """Imita client.responses.parse: devuelve las clasificaciones dadas o lanza el error dado."""

    def __init__(self, clasificaciones: list | None = None, error: Exception | None = None) -> None:
        self.clasificaciones = clasificaciones or []
        self.error = error
        self.llamadas: list[dict] = []
        self.responses = SimpleNamespace(parse=self._parse)

    def _parse(self, **kwargs):
        self.llamadas.append(kwargs)
        if self.error is not None:
            raise self.error
        return SimpleNamespace(
            output_parsed=SimpleNamespace(clasificaciones=self.clasificaciones),
            usage=SimpleNamespace(input_tokens=1000, output_tokens=200),
        )


class ClasificadorLLMTest(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)
        self.entrada = self.dir / "output"
        self.entrada.mkdir()
        self.maestros_dir = self.dir / "maestros"
        crear_maestros(self.maestros_dir)
        self.maestros = mae.cargar(self.maestros_dir)
        procesos = [
            adjudicado("1", "INSTITUTO DE DESARROLLO URBANO"),  # tiene regla: no va al LLM
            adjudicado("2", "SERVICIO GEOLOGICO COLOMBIANO", Objeto="ADECUACIÓN DE LA SEDE CENTRAL"),
            adjudicado("3", "FISCALIA GENERAL DE LA NACION", Objeto="ASEO Y CAFETERÍA"),
            adjudicado("4", "MINISTERIO DE DEFENSA NACIONAL", "INCITECO S.A.S", Objeto="CONSTRUCCIÓN DE ALOJAMIENTOS"),
            adjudicado("5", "ARMADA NACIONAL", Objeto="OBRAS EN CARTAGENA"),
        ]
        escribir_csv(self.entrada, "licitaciones_20261108_080000.csv", procesos)
        self.resultado = lim.procesar(self.entrada)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def _clasificador(self, cliente: ClienteFalso) -> cl.ClasificadorOpenAI:
        return cl.ClasificadorOpenAI(cl.Configuracion("sk-prueba-sintetica-123456", "modelo-prueba"), cliente=cliente)

    @staticmethod
    def _fila(reparto: oa.Reparto, id_contrato: str) -> oa.FilaObra:
        return next(f for f in reparto.filas if f.proceso["idContrato"] == id_contrato)

    def test_solo_envia_las_entidades_sin_regla_y_aplica_sus_decisiones(self) -> None:
        cliente = ClienteFalso(
            [
                clasificacion("2", "Alcaldías Locales"),
                clasificacion("3", cl.SIN_CONSUMO, tipo_contrato="Por clasificar"),
                clasificacion("4", "Municipios Aledaños", tipo_contrato="Construcción de obra", localidad="Bogotá D.C.", confianza="baja"),
                clasificacion("5", cl.FUERA_DE_ZONA),
            ]
        )

        reparto = oa.repartir(self.resultado, self.maestros, self._clasificador(cliente))

        self.assertEqual(len(cliente.llamadas), 1)
        enviados = cliente.llamadas[0]["input"]
        self.assertNotIn('"id": "1"', enviados)
        self.assertIn('"id": "2"', enviados)
        self.assertIn("INCITECO", cliente.llamadas[0]["instructions"])  # lista de clientes de Sugeine
        self.assertEqual(cliente.llamadas[0]["model"], "modelo-prueba")

        dos = self._fila(reparto, "2")
        self.assertEqual((dos.archivo.asesor, dos.hoja.nombre), ("Mariana", "Alcaldías Locales"))
        self.assertEqual((dos.valores["L"], dos.valores["B"]), ("Edificaciones y adecuaciones", "Teusaquillo"))
        self.assertTrue(dos.trazabilidad["SAML · Clasificación"].startswith("LLM (modelo-prueba, confianza alta): Prueba"))
        cuatro = self._fila(reparto, "4")
        self.assertEqual(cuatro.hoja.nombre, "Municipios Aledaños")
        self.assertIn("confianza BAJA", cuatro.valores["J"])
        self.assertTrue(any(i["idContrato"] == "4" and "confianza baja" in i["Detalle"] for i in reparto.incidencias))
        self.assertIsNone(reparto.archivo_sin_asignar)  # con el LLM, ninguna obra queda sin asesor
        self.assertIn("propuesto por el LLM", dos.valores["J"])
        self.assertEqual(self._fila(reparto, "1").trazabilidad["SAML · Clasificación"], "Reglas")
        self.assertEqual([p["idContrato"] for p in reparto.sin_consumo], ["3"])
        self.assertTrue(reparto.sin_consumo[0]["Motivo"].startswith("LLM"))
        self.assertEqual([p["idContrato"] for p in reparto.fuera_de_zona], ["5"])
        self.assertEqual(reparto.metodo, {"1": "Reglas", "2": "LLM", "3": "LLM", "4": "LLM", "5": "LLM"})
        self.assertEqual((reparto.llm.tokens_entrada, reparto.llm.tokens_salida), (1000, 200))

    def test_cliente_de_sugeine_detectado_por_el_llm_y_coincidencia_exacta_local(self) -> None:
        # 4: el ganador "INCITECO S.A.S" coincide localmente con el cliente "INCITECO", aunque el LLM diga NO.
        cliente = ClienteFalso(
            [
                clasificacion("2", "Alcaldías Locales", cliente_sugeine="CONFIRMADO", cliente_sugeine_nombre="CONSORCIO VIAS CB 2024"),
                clasificacion("4", "IDU"),
                clasificacion("3", "Alcaldías Locales", cliente_sugeine="CONFIRMADO", cliente_sugeine_nombre="EMPRESA QUE NO ES CLIENTE"),
            ]
        )

        reparto = oa.repartir(self.resultado, self.maestros, self._clasificador(cliente))

        dos = [f for f in reparto.filas if f.proceso["idContrato"] == "2"]
        self.assertEqual(
            sorted((f.archivo.ruta.name, f.valores["E"]) for f in dos),
            [("Obras_Alcaldias_Prueba.xlsx", "Sugeine"), ("Obras_Sugeine_Prueba.xlsx", "Sugeine")],
        )
        self.assertEqual(self._fila(reparto, "4").valores["BI"], oa.CONFIRMADO)
        tres = self._fila(reparto, "3")
        self.assertEqual(tres.valores["E"], "Mariana")  # un "cliente" fuera de la lista no se acepta
        self.assertNotIn("BI", tres.valores)

    def test_si_la_api_falla_se_usan_las_reglas(self) -> None:
        import httpx2
        import openai

        error = openai.APIConnectionError(request=httpx2.Request("POST", "https://api.openai.com/v1/responses"))
        reparto = oa.repartir(self.resultado, self.maestros, self._clasificador(ClienteFalso(error=error)))

        self.assertTrue(all(m == "Reglas" for m in reparto.metodo.values()))
        self.assertTrue(any("APIConnectionError" in e for e in reparto.llm.errores))
        self.assertEqual([p["idContrato"] for p in reparto.fuera_de_zona], [])  # por reglas, Cartagena no se detecta
        self.assertTrue(any(i["Campo"] == "LLM" for i in reparto.incidencias))

    def test_respuestas_invalidas_o_incompletas_van_por_reglas(self) -> None:
        cliente = ClienteFalso(
            [
                clasificacion("2", "Hoja Inventada"),
                clasificacion("3", "Alcaldías Locales", tipo_contrato="Tipo inventado"),
                clasificacion("4", cl.SIN_ASIGNAR),  # ya no es una opcion del LLM
                clasificacion("99", "Alcaldías Locales"),  # id que no se envió
            ]
        )

        reparto = oa.repartir(self.resultado, self.maestros, self._clasificador(cliente))

        self.assertEqual(reparto.llm.decisiones, {})
        self.assertTrue(all(m == "Reglas" for m in reparto.metodo.values()))
        self.assertEqual(len(reparto.llm.errores), 4)  # 3 invalidos + aviso de procesos sin respuesta

    def test_esquema_con_los_valores_de_los_maestros(self) -> None:
        esquema = cl._esquema(cl.destinos_permitidos(self.maestros), self.maestros.tipos_contrato).model_json_schema()
        texto = str(esquema)
        for valor in ("IDU", "Alcaldías Locales", "Municipios Aledaños", cl.SIN_CONSUMO, cl.FUERA_DE_ZONA, "Interventoría", "baja"):
            self.assertIn(valor, texto)
        self.assertNotIn(cl.SIN_ASIGNAR, cl.destinos_permitidos(self.maestros))
        self.assertNotIn("Edificaciones y Mant.", cl.destinos_permitidos(self.maestros))

    def test_configuracion_desde_el_env(self) -> None:
        env = self.dir / ".env"
        with mock.patch.dict("os.environ", {}, clear=True):
            self.assertIsNone(cl.leer_configuracion(env)[0])
            env.write_text("OPENAI_API_KEY=sk-prueba-sintetica-123456\nOPENAI_MODEL=modelo-x\n", encoding="utf-8")
            configuracion, _ = cl.leer_configuracion(env)
            self.assertEqual((configuracion.api_key, configuracion.modelo), ("sk-prueba-sintetica-123456", "modelo-x"))
            env.write_text("OPENAI_API_KEY=sk-prueba-sintetica-123456\nSAML_USAR_LLM=false\n", encoding="utf-8")
            self.assertIsNone(cl.leer_configuracion(env)[0])
            env.write_text("OPENAI_API_KEY=sk-prueba-sintetica-123456\n", encoding="utf-8")
            self.assertEqual(cl.leer_configuracion(env)[0].modelo, cl.MODELO_POR_DEFECTO)


if __name__ == "__main__":
    unittest.main()
