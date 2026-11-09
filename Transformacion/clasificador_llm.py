"""Clasificacion con un LLM (OpenAI) de las licitaciones que las reglas no resuelven.

Las reglas de obras_asesores.py reparten las entidades conocidas (IDU, alcaldias
locales, acueducto, IDRD, municipios). El resto (entidades nacionales, distritales
o departamentales sin regla) se envia al LLM, que analiza el objeto, la entidad y
la ubicacion y decide:
    - destino: la hoja de un asesor (siempre que la obra consuma material y este en la
      zona: no hay opcion "sin asignar"), "Sin consumo de material" o "Fuera de zona";
    - confianza (alta, media o baja) en esa asignacion;
    - tipo de contrato (de la lista del maestro), localidad y si consume material;
    - si un ganador es cliente de Sugeine (lista "Cobertura clientes").

La respuesta llega como JSON validado contra un esquema (salida estructurada) y SAML
vuelve a comprobar cada valor contra las listas del maestro. Si no hay API key, la
API falla o una respuesta no es valida, ese proceso se clasifica con las reglas.

Configuracion en Transformacion/.env (plantilla: Transformacion/.env.example):
    OPENAI_API_KEY   clave de la API (obligatoria para usar el LLM)
    OPENAI_MODEL     modelo (por defecto MODELO_POR_DEFECTO)
    SAML_USAR_LLM    true/false (por defecto true si hay clave)

Datos que se envian al proveedor: datos publicos del proceso (entidad, objeto,
ubicacion, valor, ganadores) y la lista de clientes de Sugeine (decision de GRECO).
"""
from __future__ import annotations

import json
import os
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Literal, Optional

from dotenv import dotenv_values

import maestros as mae

ARCHIVO_ENV = Path(__file__).resolve().parent / ".env"
MODELO_POR_DEFECTO = "gpt-5-mini"
TAMANO_LOTE = 20  # procesos por llamada
MAX_CARACTERES_OBJETO = 600

FUERA_DE_ZONA = "Fuera de zona"
SIN_CONSUMO = "Sin consumo de material"
SIN_ASIGNAR = "Sin asignar (GRECO)"
# Hojas a las que el LLM NO asigna: no aplican a contratacion publica o son de referencia.
HOJAS_EXCLUIDAS = {"Obras Privadas", "Megaproyectos (ref.)"}

# Que significa cada hoja para el LLM (las que no esten aqui se describen solo por su nombre).
# Que significa cada hoja para el LLM: el reparto es por TIPO DE OBRA (las que no esten aqui se describen
# solo por su nombre). Las edificaciones van a la hoja IDRD de Asesor 3 (obras_asesores.HOJA_EDIFICACIONES).
DESCRIPCION_HOJAS = {
    "IDU": "Obras viales y de espacio público en Bogotá: malla vial, andenes, ciclorrutas, puentes, troncales, "
    "terminales y obras de transporte (TransMilenio, terminales de transporte).",
    "Acueducto": "Redes de acueducto, alcantarillado, colectores, drenajes y obras hidráulicas en Bogotá.",
    "Alcaldías Locales": "Obra civil comunitaria en barrios de Bogotá: salones comunales, malla vial local y "
    "espacio público de barrio. NO incluye edificaciones ni sedes de entidades (esas van a \"IDRD\").",
    "IDRD": "Parques, escenarios deportivos y recreativos en Bogotá, Y TAMBIÉN construcción, adecuación, remodelación, "
    "ampliación o reforzamiento de edificaciones y sedes en Bogotá (hospitales y subredes de salud, colegios, sedes de entidades nacionales o distritales, cuarteles, "
    "estaciones de policía o bomberos, bibliotecas, viviendas).",
    "Municipios Aledaños": "Cualquier obra en municipios de Cundinamarca fuera de Bogotá (alcaldías municipales, "
    "Gobernación, ICCU, INVIAS y sedes de entidades en esos municipios).",
}

INSTRUCCIONES = """\
Eres analista comercial de GRECO, empresa de Bogotá que vende agregados y materiales reciclados de construcción \
(producidos a partir de escombros) a los contratistas que ganan obras públicas. GRECO reparte las obras adjudicadas \
entre sus asesores comerciales; cada asesor tiene hojas por tipo de entidad u obra.

Vas a clasificar procesos de contratación ADJUDICADOS que las reglas automáticas no pudieron repartir. Para cada uno:

1. destino:
   - "{fuera}": la obra no se ejecuta en Bogotá ni en Cundinamarca (o la ubicación es todo el país y nada indica \
que la obra sea en la zona).
   - "{sin_consumo}": el contrato no compra materiales de construcción: mantenimiento preventivo o correctivo \
sin obra civil, aseo, pintura, señalización o demarcación, semaforización, cámaras, interventoría, consultoría, \
estudios y diseños sin obra, suministro de bienes o servicios. Si el objeto combina mantenimiento con adecuación, \
remodelación, construcción, mejoramiento, rehabilitación, pavimentación u obra civil, SÍ consume material.
   - En cualquier otro caso (consume material y está en la zona), una de las hojas de asesor listadas abajo, \nelegida por el TIPO DE OBRA (no por el tipo de entidad). TODA obra que consume material en la zona debe quedar con un asesor: si dudas entre varias hojas, elige la más probable según el tipo de obra y la ubicación, y explica la duda en la justificación.
2. tipo_contrato: el valor de la lista que mejor describe la obra.
3. localidad: localidad de Bogotá (p. ej. "Kennedy") si el texto la indica; si no, el municipio de Cundinamarca; \
si no, "Bogotá D.C." para obras en Bogotá, o "" si no se sabe.
4. consume_material: true o false, coherente con el destino.
5. cliente_sugeine: compara los ganadores con la lista de clientes de Sugeine. "CONFIRMADO" solo si es la misma \
empresa o consorcio (puede variar la escritura, tildes, signos o el sufijo S.A.S./LTDA); "POR CONFIRMAR" si es muy \
probable pero no seguro; "NO" en otro caso. cliente_sugeine_nombre: el nombre tal como aparece en la lista, o null.
6. confianza: "alta" si la hoja es evidente, "media" si es razonable, "baja" si es una apuesta entre varias hojas.
7. justificacion: una o dos frases en español con el motivo del destino (y la duda, si la hay).

Devuelve exactamente una clasificación por cada id recibido, sin inventar ids.

Hojas de asesor disponibles:
{hojas}

Tipos de contrato permitidos: {tipos}

Clientes de Sugeine (Cobertura clientes):
{clientes}
"""


@dataclass
class Decision:
    destino: str
    tipo_contrato: str
    localidad: str
    consume_material: bool
    cliente_sugeine: str  # CONFIRMADO, POR CONFIRMAR o NO
    cliente_sugeine_nombre: Optional[str]
    justificacion: str
    confianza: str = "alta"  # alta, media o baja


@dataclass
class ResultadoLLM:
    modelo: str
    decisiones: dict[str, Decision] = field(default_factory=dict)
    errores: list[str] = field(default_factory=list)
    enviados: int = 0
    tokens_entrada: int = 0
    tokens_salida: int = 0


@dataclass
class Configuracion:
    api_key: str
    modelo: str


def leer_configuracion(archivo_env: Path = ARCHIVO_ENV) -> tuple[Optional[Configuracion], str]:
    """(configuracion o None, explicacion). Una variable del entorno prevalece sobre el .env."""
    valores = dotenv_values(archivo_env) if archivo_env.exists() else {}

    def valor(nombre: str) -> str:
        return (os.environ.get(nombre) if nombre in os.environ else valores.get(nombre)) or ""

    if valor("SAML_USAR_LLM").strip().lower() in ("0", "false", "no"):
        return None, "SAML_USAR_LLM=false: se usan solo las reglas."
    clave = valor("OPENAI_API_KEY").strip()
    if not clave:
        return None, "No hay OPENAI_API_KEY en Transformacion/.env: se usan solo las reglas."
    return Configuracion(clave, valor("OPENAI_MODEL").strip() or MODELO_POR_DEFECTO), ""


def _hojas_de_asesor(maestros: mae.Maestros) -> list[tuple[str, str]]:
    """(hoja, asesor) a las que puede asignar el LLM."""
    return [(h, a.asesor) for a in maestros.archivos if not a.es_sugeine for h in a.hojas if h not in HOJAS_EXCLUIDAS]


def destinos_permitidos(maestros: mae.Maestros) -> list[str]:
    hojas = [h for h, _ in _hojas_de_asesor(maestros)]
    # Sin opcion "Sin asignar": toda obra que consume material en la zona queda con un asesor.
    return [*hojas, SIN_CONSUMO, FUERA_DE_ZONA]


def _esquema(destinos: list[str], tipos: list[str]):
    """Modelo Pydantic de la respuesta, con enums tomados de los maestros."""
    from pydantic import BaseModel, ConfigDict

    Destino = Literal[tuple(destinos)]  # type: ignore[valid-type]
    Tipo = Literal[tuple(tipos)]  # type: ignore[valid-type]

    class Clasificacion(BaseModel):
        model_config = ConfigDict(extra="forbid")
        id: str
        destino: Destino
        tipo_contrato: Tipo
        localidad: str
        consume_material: bool
        cliente_sugeine: Literal["CONFIRMADO", "POR CONFIRMAR", "NO"]
        cliente_sugeine_nombre: Optional[str]
        confianza: Literal["alta", "media", "baja"]
        justificacion: str

    class Respuesta(BaseModel):
        model_config = ConfigDict(extra="forbid")
        clasificaciones: list[Clasificacion]

    return Respuesta


def construir_instrucciones(maestros: mae.Maestros) -> str:
    hojas = [f'- "{nombre}" (asesor: {asesor}): {DESCRIPCION_HOJAS.get(nombre, nombre)}' for nombre, asesor in _hojas_de_asesor(maestros)]
    clientes = "\n".join(f"- {c}" for c in sorted(maestros.clientes_sugeine.values())) or "(sin clientes)"
    return INSTRUCCIONES.format(
        fuera=FUERA_DE_ZONA,
        sin_consumo=SIN_CONSUMO,
        hojas="\n".join(hojas),
        tipos=", ".join(f'"{t}"' for t in maestros.tipos_contrato),
        clientes=clientes,
    )


def _entrada(proceso: dict[str, Any]) -> dict[str, Any]:
    """Solo datos publicos del proceso."""
    return {
        "id": proceso["idContrato"],
        "entidad": proceso["Entidad contratante"],
        "objeto": (proceso["Objeto"] or "")[:MAX_CARACTERES_OBJETO],
        "ubicacion": proceso["Ubicación (origen)"],
        "valor_millones": round(proceso["Valor Millones"], 1) if proceso.get("Valor Millones") else None,
        "ganadores": proceso.get("_ganadores") or [],
    }


def _validar(item: Any, destinos: list[str], tipos: list[str], clientes: set[str]) -> tuple[Optional[Decision], str]:
    if item.destino not in destinos:
        return None, f"destino '{item.destino}' no permitido"
    if item.tipo_contrato not in tipos:
        return None, f"tipo '{item.tipo_contrato}' no permitido"
    nombre = item.cliente_sugeine_nombre
    nivel = item.cliente_sugeine
    if nivel != "NO" and (not nombre or mae.normalizar(nombre) not in clientes):
        nivel, nombre = "NO", None  # un cliente que no esta en la lista no se acepta
    return (
        Decision(
            destino=item.destino,
            tipo_contrato=item.tipo_contrato,
            localidad=(item.localidad or "").strip(),
            consume_material=bool(item.consume_material),
            cliente_sugeine=nivel,
            cliente_sugeine_nombre=nombre,
            justificacion=(item.justificacion or "").strip(),
            confianza=item.confianza if item.confianza in ("alta", "media", "baja") else "baja",
        ),
        "",
    )


class ClasificadorOpenAI:
    """Clasifica por lotes con la Responses API de OpenAI y salida estructurada."""

    def __init__(self, configuracion: Configuracion, cliente: Any = None) -> None:
        self.configuracion = configuracion
        if cliente is None:
            from openai import OpenAI

            cliente = OpenAI(api_key=configuracion.api_key, max_retries=3, timeout=120)
        self.cliente = cliente

    @property
    def modelo(self) -> str:
        return self.configuracion.modelo

    def clasificar(self, procesos: list[dict[str, Any]], maestros: mae.Maestros) -> ResultadoLLM:
        import openai

        resultado = ResultadoLLM(modelo=self.modelo)
        if not procesos:
            return resultado
        destinos = destinos_permitidos(maestros)
        tipos = list(maestros.tipos_contrato)
        clientes = set(maestros.clientes_sugeine)
        esquema = _esquema(destinos, tipos)
        instrucciones = construir_instrucciones(maestros)

        for inicio in range(0, len(procesos), TAMANO_LOTE):
            lote = procesos[inicio : inicio + TAMANO_LOTE]
            ids = {p["idContrato"] for p in lote}
            resultado.enviados += len(lote)
            try:
                respuesta = self.cliente.responses.parse(
                    model=self.modelo,
                    instructions=instrucciones,
                    input=json.dumps({"procesos": [_entrada(p) for p in lote]}, ensure_ascii=False),
                    text_format=esquema,
                )
            except openai.AuthenticationError:
                resultado.errores.append("La API rechazó la clave (401). Revisa OPENAI_API_KEY; se usan las reglas.")
                return resultado
            except openai.NotFoundError:
                resultado.errores.append(f"El modelo '{self.modelo}' no existe o no está disponible (404); se usan las reglas.")
                return resultado
            except (openai.RateLimitError, openai.APIConnectionError, openai.APITimeoutError, openai.InternalServerError) as exc:
                resultado.errores.append(f"Lote {inicio // TAMANO_LOTE + 1}: {type(exc).__name__}; esos procesos van por reglas.")
                continue
            except (openai.APIStatusError, openai.LengthFinishReasonError, openai.ContentFilterFinishReasonError) as exc:
                resultado.errores.append(f"Lote {inicio // TAMANO_LOTE + 1}: {type(exc).__name__}: {str(exc)[:200]}; van por reglas.")
                continue

            uso = getattr(respuesta, "usage", None)
            if uso is not None:
                resultado.tokens_entrada += getattr(uso, "input_tokens", 0) or 0
                resultado.tokens_salida += getattr(uso, "output_tokens", 0) or 0
            datos = respuesta.output_parsed
            if datos is None:
                resultado.errores.append(f"Lote {inicio // TAMANO_LOTE + 1}: respuesta sin contenido estructurado; van por reglas.")
                continue
            for item in datos.clasificaciones:
                if item.id not in ids or item.id in resultado.decisiones:
                    continue
                decision, error = _validar(item, destinos, tipos, clientes)
                if decision is None:
                    resultado.errores.append(f"Proceso {item.id}: {error}; va por reglas.")
                    continue
                resultado.decisiones[item.id] = decision
            faltan = ids - set(resultado.decisiones)
            if faltan:
                resultado.errores.append(f"Lote {inicio // TAMANO_LOTE + 1}: {len(faltan)} procesos sin respuesta; van por reglas.")
        return resultado


def crear_clasificador(archivo_env: Path = ARCHIVO_ENV) -> tuple[Optional[ClasificadorOpenAI], str]:
    configuracion, motivo = leer_configuracion(archivo_env)
    if configuracion is None:
        return None, motivo
    try:
        return ClasificadorOpenAI(configuracion), ""
    except ImportError:
        return None, "Falta el paquete 'openai' (pip install -r Transformacion/requirements.txt): se usan solo las reglas."
