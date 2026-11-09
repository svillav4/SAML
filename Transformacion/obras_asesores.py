"""Reparto de las licitaciones adjudicadas entre los archivos de los asesores de GRECO.

Para cada proceso limpio (limpieza_licitaciones.py):
    1. Si su codigo ya esta en algun maestro, no se vuelve a agregar ("Ya en maestros").
    2. Se busca la hoja de destino segun la entidad contratante (REGLAS_HOJA). Si
       ninguna regla aplica, va a "Sin clasificar" para que GRECO decida.
    3. El archivo y el asesor salen del maestro que tiene esa hoja.
    4. Si un ganador es cliente de Sugeine (hoja "Cobertura clientes"), la obra
       tambien va al archivo de Sugeine y en el del otro asesor queda con
       Asesor = Sugeine (su formula la pone en 0 para no contarla dos veces). Si el
       nombre solo se parece, queda como "(posible Sugeine)" para revision.
    5. Se arma la fila con las 71 columnas de la plantilla: SAML llena los datos que
       vienen de la API y deja vacios los que diligencia el asesor.

Inicio de obra y Meses por Terminar salen del cronograma (fechas PROGRAMADAS) solo
cuando hay una unica fecha de inicio y una unica de fin, coherentes; si no, Inicio
de obra = mes de la extraccion y Meses por Terminar queda vacio. Avance = 0.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field, replace
from datetime import datetime
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Optional

import limpieza_licitaciones as lim
import maestros as mae
from maestros import normalizar, normalizar_codigo

# ----------------------------------------------------------------------
# Reglas de reparto (entidad contratante -> hoja). Se evaluan en orden sobre el
# nombre normalizado (mayusculas, sin tildes ni signos). Ajustar aqui si GRECO
# cambia la distribucion.
# ----------------------------------------------------------------------
REGLAS_HOJA: tuple[tuple[str, str, str], ...] = (
    ("IDU", r"\bINSTITUTO DE DESARROLLO URBANO\b|^IDU\b|\bTRANSMILENIO\b", "IDU y TransMilenio"),
    ("Acueducto", r"\bACUEDUCTO\b|\bEAAB\b", "Empresa de Acueducto (EAAB)"),
    ("IDRD", r"\bINSTITUTO DISTRITAL DE (LA )?RECREACION\b|\bIDRD\b", "Instituto Distrital de Recreación y Deporte"),
    ("Alcaldías Locales", r"\bALCALDIA LOCAL\b|\bFONDO DE DESARROLLO LOCAL\b", "Alcaldías locales y FDL de Bogotá"),
)
# Municipios aledanos: entidades municipales o departamentales de Cundinamarca (no de Bogota).
PATRON_MUNICIPAL = re.compile(
    r"\bALCALDIA (MUNICIPAL|DEL MUNICIPIO|DE MUNICIPIO|MUNICIPIO)\b|^MUNICIPIO DE\b|\bICCU\b"
    r"|\bINSTITUTO DE CAMINOS Y CONSTRUCCIONES DE CUNDINAMARCA\b|\bFONDO DE DESARROLLO DE PROYECTOS DE CUNDINAMARCA\b|\bIDUVI\b"
)
HOJA_MUNICIPIOS = "Municipios Aledaños"

# Procesos de entidades sin regla (nacionales, distritales, departamentales). Se evaluan en orden:
FUERA_DE_ZONA = "Fuera de zona"  # 1. ubicacion sin Bogota ni Cundinamarca: se descartan
# 2. Viales nacionales y 3. Gobernacion de Cundinamarca en municipios -> Municipios Aledanos.
PATRON_VIAL_NACIONAL = re.compile(r"\bINVIAS\b|\bINSTITUTO NACIONAL DE VIAS\b|\bANI\b|\bAGENCIA NACIONAL DE INFRAESTRUCTURA\b")
PATRON_GOBERNACION = re.compile(r"\bDEPARTAMENTO DE CUNDINAMARCA\b|\bGOBERNACION DE CUNDINAMARCA\b")
# 4. Solo mantenimiento (sin obra): la formula de GRECO no les da potencial -> hoja "Sin consumo de material".
SIN_CONSUMO = "Sin consumo de material"
PATRON_MANTENIMIENTO = re.compile(r"\bMANTENIMIENTO\b|\bCONSERVACION\b")
PATRON_OBRA = re.compile(
    r"\bADECUACI|\bREMODELACI|\bCONSTRUCCI|\bMEJORAMIENTO\b|\bREHABILITACI|\bAMPLIACI|\bREFORZAMIENTO|\bURBANISMO\b"
    r"|\bPAVIMENT|\bVIAS?\b|\bVIAL"
)
# 5. El resto (con sede en la zona) va a un archivo "Sin asignar (GRECO)" para que GRECO elija el asesor;
#    si el ganador es cliente de Sugeine, va a su hoja "Edificaciones y Mant.".
SIN_ASIGNAR = "Sin asignar (GRECO)"
ASESOR_SIN_ASIGNAR = SIN_ASIGNAR
HOJA_SIN_ASIGNAR = SIN_ASIGNAR
HOJA_PLANTILLA_SIN_ASIGNAR = "Alcaldías Locales"  # todas las hojas comparten la plantilla de 71 columnas
HOJA_SUGEINE_OTRAS = "Edificaciones y Mant."
# Edificaciones y adecuaciones de entidades sin regla (nacionales, distritales) en Bogota: hoja IDRD de Asesor 3
# (ya existe en su maestro; se distinguen de los parques por Tipo de contrato = "Edificaciones y adecuaciones").
HOJA_EDIFICACIONES = "IDRD"
TIPO_EDIFICACIONES = "Edificaciones y adecuaciones"
# Reparto por tipo de obra de las entidades sin regla (respaldo cuando no hay LLM).
TIPO_A_HOJA = {
    TIPO_EDIFICACIONES: HOJA_EDIFICACIONES,
    "Parques y escenarios deportivos": "IDRD",
    "Redes de acueducto y alcantarillado": "Acueducto",
}
PATRON_VIAL = re.compile(r"\bMALLA VIAL\b|\bVIAS?\b|\bVIAL\b|\bANDEN|\bPUENTES?\b|\bPAVIMENT|\bESPACIO PUBLICO\b|\bCICLORRUTA")

LOCALIDADES_BOGOTA = (
    "Usaquén", "Chapinero", "Santa Fe", "San Cristóbal", "Usme", "Tunjuelito", "Bosa", "Kennedy", "Fontibón",
    "Engativá", "Suba", "Barrios Unidos", "Teusaquillo", "Los Mártires", "Antonio Nariño", "Puente Aranda",
    "La Candelaria", "Rafael Uribe Uribe", "Ciudad Bolívar", "Sumapaz",
)
BOGOTA = "Bogotá D.C."

# Tipo de contrato propuesto por palabras clave del objeto (en orden; la primera que coincide gana).
REGLAS_TIPO: tuple[tuple[str, str], ...] = (
    ("Interventoría", r"\bINTERVENTOR"),
    ("Señalización", r"\bSENALIZ|\bDEMARCACI"),
    ("Semaforización", r"\bSEMAFOR"),
    ("Cámaras de seguridad", r"\bCAMARAS?\b|\bCCTV\b|\bVIDEOVIGILANCIA\b"),
    ("Redes de acueducto y alcantarillado", r"\bACUEDUCTO\b|\bALCANTARILLAD|\bREDES HIDR"),
    ("Parques y escenarios deportivos", r"\bPARQUES?\b|\bESCENARIOS? DEPORTIV|\bCANCHAS?\b|\bPOLIDEPORTIV"),
    ("Conservación, mantenimiento y reparcheo", r"\bMANTENIMIENTO\b|\bCONSERVACION\b|\bREPARCHEO\b|\bPARCHEO\b"),
    ("Edificaciones y adecuaciones", r"\bADECUACI|\bEDIFICACI|\bEDIFICIO|\bLOCATIV|\bREMODELACI|\bSEDES?\b"),
    ("Construcción de obra", r"\bCONSTRUCCION\b|\bOBRAS?\b"),
)
TIPO_POR_DEFECTO = "Por clasificar"

UMBRAL_SIMILITUD = 0.88  # "posible Sugeine"
MESES_MAXIMOS = 60
SUFIJOS_SOCIETARIOS = {"SAS", "SA", "S", "A", "LTDA", "LIMITADA", "EU", "ESP", "E", "P", "BIC", "EN", "REORGANIZACION", "SUCURSAL"}
MESES_ABREV = ("ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic")

# Columnas (fuera de las 71 del maestro) con datos de origen para revisar cada fila. Van
# despues de una columna vacia (BT) y NO se pegan en el maestro.
COLUMNAS_TRAZABILIDAD_SAML = (
    "SAML · Entidad contratante",
    "SAML · Ubicación (origen)",
    "SAML · Enlace",
    "SAML · idContrato",
    "SAML · Inicio ejecución (cronograma)",
    "SAML · Fin ejecución (cronograma)",
    "SAML · Origen de Inicio y Meses",
    "SAML · Coincidencia con clientes de Sugeine",
    "SAML · Clasificación",
)

CONFIRMADO = "CONFIRMADO"
POR_CONFIRMAR = "POR CONFIRMAR"


@dataclass
class FilaObra:
    """Una fila nueva para una hoja del archivo de un asesor."""

    archivo: mae.ArchivoMaestro
    hoja: mae.HojaMaestro
    valores: dict[str, Any]  # letra de columna -> valor (solo las que llena SAML)
    trazabilidad: dict[str, Any]
    proceso: dict[str, Any]


@dataclass
class Reparto:
    filas: list[FilaObra] = field(default_factory=list)
    sin_clasificar: list[dict[str, Any]] = field(default_factory=list)
    sin_consumo: list[dict[str, Any]] = field(default_factory=list)
    fuera_de_zona: list[dict[str, Any]] = field(default_factory=list)
    archivo_sin_asignar: Optional[mae.ArchivoMaestro] = None
    metodo: dict[str, str] = field(default_factory=dict)  # idContrato -> "Reglas" o "LLM"
    llm: Any = None  # clasificador_llm.ResultadoLLM de esta ejecucion (o None)
    llm_aviso: str = ""  # por que no se uso el LLM, si no se uso
    ya_en_maestros: list[dict[str, Any]] = field(default_factory=list)
    incidencias: list[dict[str, Any]] = field(default_factory=list)
    destinos: dict[str, str] = field(default_factory=dict)  # idContrato -> resumen del destino

    def filas_de(self, archivo: mae.ArchivoMaestro) -> list[FilaObra]:
        return [f for f in self.filas if f.archivo is archivo]


# ----------------------------------------------------------------------
# Reglas
# ----------------------------------------------------------------------
def hoja_para(entidad: str, ubicacion: str) -> Optional[str]:
    """Hoja de destino segun la entidad contratante, o None si no hay regla."""
    nombre = normalizar(entidad)
    for hoja, patron, _ in REGLAS_HOJA:
        if re.search(patron, nombre):
            return hoja
    if PATRON_MUNICIPAL.search(nombre) and "CUNDINAMARCA" in normalizar(ubicacion):
        return HOJA_MUNICIPIOS
    return None


def clasificar_sin_regla(proceso: dict[str, Any]) -> tuple[str, Optional[str], str]:
    """(categoria, hoja o None, motivo) para un proceso cuya entidad no tiene regla de hoja."""
    entidad = normalizar(proceso["Entidad contratante"])
    ubicacion = proceso["Ubicación (origen)"]
    objeto = normalizar(proceso["Objeto"])
    zona = normalizar(ubicacion)
    if "BOGOTA" not in zona and "CUNDINAMARCA" not in zona:
        return FUERA_DE_ZONA, None, f"Ubicación fuera de Bogotá y Cundinamarca ('{ubicacion}')."
    if PATRON_VIAL_NACIONAL.search(entidad):
        return "Vial nacional", HOJA_MUNICIPIOS, "Obra vial nacional en la zona: va a Municipios Aledaños."
    if PATRON_GOBERNACION.search(entidad) and (_municipios(ubicacion) or "BOGOTA" not in zona):
        return "Gobernación", HOJA_MUNICIPIOS, "Gobernación de Cundinamarca con obra en municipios: va a Municipios Aledaños."
    if PATRON_MANTENIMIENTO.search(objeto) and not PATRON_OBRA.search(objeto):
        return SIN_CONSUMO, None, "Solo mantenimiento (sin adecuación, construcción ni obra): no consume material."
    # Por tipo de obra (solo obras en Bogota; fuera de Bogota, en Cundinamarca, van a Municipios Aledanos).
    if "BOGOTA" not in zona:
        return "Cundinamarca", HOJA_MUNICIPIOS, "Obra en Cundinamarca fuera de Bogotá: va a Municipios Aledaños."
    if PATRON_VIAL.search(objeto):
        return "Vial", "IDU", "Obra vial o de espacio público en Bogotá: va a IDU."
    for tipo, hoja in TIPO_A_HOJA.items():
        if re.search(dict(REGLAS_TIPO)[tipo], objeto):
            return tipo, hoja, f"Tipo de obra '{tipo}' en Bogotá: va a {hoja}."
    return SIN_ASIGNAR, None, "Entidad sin regla de reparto con obra en la zona: GRECO decide el asesor."


def archivo_y_hoja(maestros: mae.Maestros, nombre_hoja: str) -> tuple[Optional[mae.ArchivoMaestro], Optional[mae.HojaMaestro]]:
    archivo = maestros.archivo_de_hoja(nombre_hoja)
    return (archivo, archivo.hojas[nombre_hoja]) if archivo else (None, None)


def _archivo_sin_asignar(maestros: mae.Maestros) -> Optional[mae.ArchivoMaestro]:
    """Archivo virtual 'Sin asignar (GRECO)' con la plantilla de una hoja de obras de los maestros."""
    dueno = maestros.archivo_de_hoja(HOJA_PLANTILLA_SIN_ASIGNAR)
    if dueno is None:
        dueno = next((a for a in maestros.archivos if not a.es_sugeine), None)
        if dueno is None:
            return None
        plantilla = next(iter(dueno.hojas.values()))
    else:
        plantilla = dueno.hojas[HOJA_PLANTILLA_SIN_ASIGNAR]
    hoja = replace(plantilla, nombre=HOJA_SIN_ASIGNAR, archivo="Sin_asignar_GRECO.xlsx", entidad="")
    return mae.ArchivoMaestro(
        ruta=Path("Sin_asignar_GRECO.xlsx"), asesor=ASESOR_SIN_ASIGNAR, es_sugeine=False, sha256="", hojas={HOJA_SIN_ASIGNAR: hoja}
    )


def tipo_contrato(objeto: str, permitidos: list[str]) -> str:
    texto = normalizar(objeto)
    for tipo, patron in REGLAS_TIPO:
        if re.search(patron, texto) and tipo in permitidos:
            return tipo
    return TIPO_POR_DEFECTO if TIPO_POR_DEFECTO in permitidos else ""


def _municipios(ubicacion: str) -> list[str]:
    """'Cundinamarca : Cota, Chía; Tolima : Melgar' -> ['Cota', 'Chía', 'Melgar'] (sin Bogota)."""
    nombres = []
    for tramo in ubicacion.split(";"):
        if ":" not in tramo:
            continue
        for municipio in tramo.split(":", 1)[1].split(","):
            municipio = municipio.strip()
            if municipio and normalizar(municipio) != normalizar(BOGOTA):
                nombres.append(municipio)
    return nombres


def localidad(hoja: str, entidad: str, ubicacion: str, escrituras: dict[str, str]) -> str:
    if hoja == "Alcaldías Locales":
        nombre = normalizar(entidad)
        for candidata in LOCALIDADES_BOGOTA:
            clave = normalizar(candidata)
            corta = re.sub(r"^(LOS|LA) ", "", clave)
            corta = re.sub(r" URIBE URIBE$", " URIBE", corta)
            if re.search(rf"\b({clave}|{corta})\b", nombre):
                return escrituras.get(clave) or escrituras.get(corta) or candidata
    municipios = _municipios(ubicacion)
    if hoja == HOJA_MUNICIPIOS and municipios:
        return municipios[0]
    if normalizar(BOGOTA) in normalizar(ubicacion):
        return BOGOTA
    if municipios:
        return municipios[0]
    return "Cundinamarca" if "CUNDINAMARCA" in normalizar(ubicacion) else ""


def _nombre_empresa(nombre: str) -> str:
    palabras = normalizar(nombre).split()
    while palabras and palabras[-1] in SUFIJOS_SOCIETARIOS:
        palabras.pop()
    return " ".join(palabras)


def _nucleo(nombre: str) -> str:
    return re.sub(r"^(CONSORCIO|UNION TEMPORAL|UT|U T)\s+", "", nombre).strip()


def coincidencia_sugeine(ganadores: list[str], clientes: dict[str, str]) -> tuple[str, str, str]:
    """(nivel, ganador, cliente): nivel CONFIRMADO si el nombre coincide, POR_CONFIRMAR si solo se parece."""
    if not clientes:
        return "", "", ""
    base = {_nombre_empresa(original): original for original in clientes.values()}
    mejor = ("", "", "", 0.0)
    for ganador in ganadores:
        nombre = _nombre_empresa(ganador)
        if not nombre:
            continue
        if nombre in base:
            return CONFIRMADO, ganador, base[nombre]
        # El parecido se mide sin "CONSORCIO" / "UNION TEMPORAL", que harian parecidos a casi todos.
        nucleo = _nucleo(nombre)
        for clave, original in base.items():
            nucleo_cliente = _nucleo(clave)
            if min(len(nucleo), len(nucleo_cliente)) < 6:
                continue
            similitud = SequenceMatcher(None, nucleo, nucleo_cliente).ratio()
            if similitud > mejor[3]:
                mejor = (POR_CONFIRMAR, ganador, original, similitud)
    if mejor[3] >= UMBRAL_SIMILITUD:
        return mejor[0], mejor[1], mejor[2]
    return "", "", ""


def es_misma_obra(proceso: dict[str, Any], existente: mae.ObraExistente) -> bool:
    """Mismo codigo y, ademas, la misma empresa ganadora o el mismo valor (con 2 % de tolerancia)."""
    empresa = _nombre_empresa(existente.empresa)
    for ganador in proceso.get("_ganadores") or []:
        nombre = _nombre_empresa(ganador)
        if nombre and empresa and (nombre == empresa or SequenceMatcher(None, nombre, empresa).ratio() >= UMBRAL_SIMILITUD):
            return True
    valor = proceso.get("Valor Millones")
    if valor and existente.valor_millones:
        return abs(valor - existente.valor_millones) <= 0.02 * max(valor, existente.valor_millones)
    return False


def _primero_de_mes(fecha: datetime) -> datetime:
    return datetime(fecha.year, fecha.month, 1)


def _meses_entre(desde: datetime, hasta: datetime) -> int:
    return (hasta.year * 12 + hasta.month) - (desde.year * 12 + desde.month)


def inicio_y_meses(
    proceso: dict[str, Any], inicio_proyeccion: Optional[datetime]
) -> tuple[datetime, Optional[int], str]:
    """(Inicio de obra, Meses por Terminar, explicacion) segun el cronograma o el mes de extraccion."""
    extraccion = proceso.get("Fecha extracción") or datetime.now()
    piso = _primero_de_mes(inicio_proyeccion) if inicio_proyeccion else None
    por_defecto = _primero_de_mes(extraccion)
    if piso and por_defecto < piso:
        por_defecto = piso
    # El cronograma guarda las versiones anteriores (adendas): se compara por MES. Si todas las
    # fechas caen en el mismo mes, ese es el programado; si caen en meses distintos, es ambiguo.
    inicios = sorted({_primero_de_mes(f) for f in proceso.get("_inicios") or []})
    fines = sorted({_primero_de_mes(f) for f in proceso.get("_fines") or []})

    if len(inicios) != 1:
        motivo = "el cronograma no trae fecha de inicio de ejecución" if not inicios else (
            f"las fechas de inicio de ejecución del cronograma caen en {len(inicios)} meses distintos (adendas)"
        )
        return por_defecto, None, f"Inicio = mes de la extracción y Meses vacío: {motivo}."

    inicio = inicios[0]
    # La lista desplegable y la proyeccion empiezan en X2: una obra iniciada antes empieza a contar en X2.
    inicio_obra = max(inicio, piso) if piso else inicio
    if len(fines) != 1:
        motivo = "no trae fecha de fin" if not fines else f"trae fechas de fin en {len(fines)} meses distintos (adendas)"
        return inicio_obra, None, f"Inicio programado según el cronograma; Meses vacío: el cronograma {motivo}."
    fin = fines[0]
    if fin < inicio:
        return inicio_obra, None, "Inicio programado según el cronograma; Meses vacío: la fecha de fin no es posterior al inicio."
    meses = _meses_entre(inicio_obra, _primero_de_mes(fin)) + 1
    if meses < 1:
        return inicio_obra, None, (
            f"Inicio programado según el cronograma; Meses vacío: el fin programado ({fin:%m/%Y}) es anterior "
            "al inicio de la proyección."
        )
    if meses > MESES_MAXIMOS:
        return inicio_obra, None, f"Inicio programado según el cronograma; Meses vacío: daría {meses} meses (más de {MESES_MAXIMOS})."
    return inicio_obra, meses, (
        f"Inicio y Meses programados según el cronograma (de {inicio:%m/%Y} a {fin:%m/%Y})."
    )


def _fecha_corta(fecha: datetime) -> str:
    return f"{fecha.day:02d}-{MESES_ABREV[fecha.month - 1]}-{fecha:%y}"


# ----------------------------------------------------------------------
# Construccion de las filas
# ----------------------------------------------------------------------
def _valores_fila(
    proceso: dict[str, Any],
    hoja: mae.HojaMaestro,
    asesor: str,
    asesor_dueno: str,
    sugeine: tuple[str, str, str],
    maestros: mae.Maestros,
) -> tuple[dict[str, Any], dict[str, Any], list[str]]:
    extraccion: datetime = proceso.get("Fecha extracción") or datetime.now()
    inicio, meses, origen_fechas = inicio_y_meses(proceso, hoja.inicio_proyeccion)
    tipo = tipo_contrato(proceso["Objeto"], maestros.tipos_contrato)
    fuente = (proceso.get("Fuente") or "SECOP").upper()
    avisos = []
    if meses is None:
        avisos.append("completar Meses por Terminar")
    avisos.append(f"Tipo de contrato propuesto por palabras clave ({tipo or 'sin propuesta'}): confirmar")
    nota = (
        f"[SAML · {_fecha_corta(extraccion)}] ADJUDICADA · VERIFICAR INICIO — {origen_fechas} "
        + "Pendiente: " + "; ".join(avisos) + "."
    )
    nivel, ganador, cliente = sugeine
    valores: dict[str, Any] = {
        "A": hoja.entidad,
        "B": localidad(hoja.nombre, proceso["Entidad contratante"], proceso["Ubicación (origen)"], maestros.localidades),
        "E": asesor,
        "I": proceso["Empresa (ganadores)"] or None,
        "J": nota,
        "K": proceso["Objeto"] or None,
        "L": tipo or None,
        "M": proceso["CodigoProceso"],
        "N": proceso["Valor Millones"],
        "P": 0,
        "Q": meses,
        "R": inicio,
        "AY": f"NUEVA · SAML (Licitaciones.info, extracción {_fecha_corta(extraccion)})",
        "AZ": proceso["CodigoProceso"],
        "BA": datetime(extraccion.year, extraccion.month, extraccion.day),
        "BB": 0,
        "BH": f"{fuente} {proceso['CodigoProceso']} (vía Licitaciones.info)",
    }
    if nivel:
        explicacion = f"SAML: el ganador '{ganador}' {'coincide con' if nivel == CONFIRMADO else 'se parece a'} el cliente '{cliente}' de Cobertura clientes."
        valores.update({"BI": nivel, "BJ": explicacion, "BK": f"Ganador en Licitaciones.info: {ganador}", "BM": asesor_dueno})
    trazabilidad = {
        "SAML · Entidad contratante": proceso["Entidad contratante"],
        "SAML · Ubicación (origen)": proceso["Ubicación (origen)"],
        "SAML · Enlace": proceso["Enlace"],
        "SAML · idContrato": proceso["idContrato"],
        "SAML · Inicio ejecución (cronograma)": proceso["Inicio ejecución (cronograma)"],
        "SAML · Fin ejecución (cronograma)": proceso["Fin ejecución (cronograma)"],
        "SAML · Origen de Inicio y Meses": origen_fechas,
        "SAML · Coincidencia con clientes de Sugeine": (f"{nivel}: {ganador} ~ {cliente}" if nivel else ""),
        "SAML · Clasificación": "Reglas",
    }
    return valores, trazabilidad, avisos


def repartir(resultado: lim.Resultado, maestros: mae.Maestros, clasificador: Any = None) -> Reparto:
    """Reparte los procesos. `clasificador` (opcional, clasificador_llm.ClasificadorOpenAI) decide los que no tienen regla."""
    reparto = Reparto()

    def incidencia(tipo: str, detalle: str, proceso: dict[str, Any], campo: str = "") -> None:
        reparto.incidencias.append(
            {
                "Tipo": tipo,
                "Alcance": "Registro",
                "idContrato": proceso["idContrato"],
                "CodigoProceso": proceso["CodigoProceso"],
                "Hoja": "Reparto",
                "Campo": campo,
                "Detalle": detalle,
            }
        )

    sugeine = maestros.sugeine
    if clasificador is not None:
        candidatos = [
            p
            for p in resultado.datos
            if hoja_para(p["Entidad contratante"], p["Ubicación (origen)"]) is None
            and not any(es_misma_obra(p, e) for e in maestros.codigos.get(normalizar_codigo(p["CodigoProceso"]), []))
        ]
        reparto.llm = clasificador.clasificar(candidatos, maestros)
        for error in reparto.llm.errores:
            reparto.incidencias.append(
                {"Tipo": lim.INC_REVISION, "Alcance": "Ejecución", "idContrato": "", "CodigoProceso": "",
                 "Hoja": "Reparto", "Campo": "LLM", "Detalle": error}
            )
    decisiones = reparto.llm.decisiones if reparto.llm is not None else {}
    vistos: dict[str, str] = {}
    for proceso in resultado.datos:
        codigo = normalizar_codigo(proceso["CodigoProceso"])
        existentes = maestros.codigos.get(codigo, []) if codigo else []
        iguales = [e for e in existentes if es_misma_obra(proceso, e)]
        if iguales:
            donde = "; ".join(e.ubicacion for e in iguales[:3])
            reparto.ya_en_maestros.append({**proceso, "Ya está en": donde})
            reparto.destinos[proceso["idContrato"]] = f"Ya en maestros: {donde}"
            continue
        if existentes:
            # Codigos genericos (p. ej. "LP-001-2026") se repiten entre entidades distintas.
            incidencia(
                lim.INC_REVISION,
                "El código coincide con "
                + "; ".join(e.ubicacion for e in existentes[:3])
                + ", pero la empresa y el valor son distintos: se agrega como obra nueva; verificar que no sea la misma.",
                proceso,
                "CodigoProceso",
            )
        if codigo and codigo in vistos:
            incidencia(
                lim.INC_DUPLICADO,
                f"Mismo código de proceso que el idContrato {vistos[codigo]}; se agrega una sola vez.",
                proceso,
                "CodigoProceso",
            )
            reparto.destinos[proceso["idContrato"]] = f"Omitido: mismo código que {vistos[codigo]}"
            continue

        nombre_hoja = hoja_para(proceso["Entidad contratante"], proceso["Ubicación (origen)"])
        decision = None
        categoria_sin_regla = nombre_hoja is None
        reparto.metodo[proceso["idContrato"]] = "Reglas"
        if nombre_hoja is None:
            decision = decisiones.get(proceso["idContrato"])
            if decision is not None:
                reparto.metodo[proceso["idContrato"]] = "LLM"
                categoria = decision.destino
                nombre_hoja = None if categoria in (FUERA_DE_ZONA, SIN_CONSUMO, SIN_ASIGNAR) else categoria
                motivo = f"LLM ({reparto.llm.modelo}, confianza {decision.confianza}): {decision.justificacion}"
                if decision.confianza == "baja" and categoria not in (FUERA_DE_ZONA, SIN_CONSUMO):
                    incidencia(
                        lim.INC_REVISION,
                        f"Asignación del LLM con confianza baja ({categoria}): revisar. {decision.justificacion}",
                        proceso,
                        "Asesor",
                    )
            else:
                categoria, nombre_hoja, motivo = clasificar_sin_regla(proceso)
            if categoria in (FUERA_DE_ZONA, SIN_CONSUMO):
                lista = reparto.fuera_de_zona if categoria == FUERA_DE_ZONA else reparto.sin_consumo
                lista.append({**proceso, "Motivo": motivo})
                reparto.destinos[proceso["idContrato"]] = categoria
                continue
        archivo, hoja = archivo_y_hoja(maestros, nombre_hoja) if nombre_hoja else (None, None)
        if decision is None and categoria_sin_regla and nombre_hoja and archivo is None:
            nombre_hoja = None  # la hoja por tipo de obra no existe en los maestros: GRECO decide (Sin asignar)
        if nombre_hoja and archivo is None:
            reparto.sin_clasificar.append({**proceso, "Motivo": f"Ningún maestro de asesor tiene la hoja '{nombre_hoja}'."})
            reparto.destinos[proceso["idContrato"]] = "Sin clasificar"
            continue
        if codigo:
            vistos[codigo] = proceso["idContrato"]
        nivel, ganador, cliente = coincidencia_sugeine(proceso["_ganadores"], maestros.clientes_sugeine)
        if decision is not None and nivel != CONFIRMADO:
            # Una coincidencia exacta local siempre vale; si no la hay, decide el LLM.
            nivel = {"CONFIRMADO": CONFIRMADO, "POR CONFIRMAR": POR_CONFIRMAR}.get(decision.cliente_sugeine, "")
            cliente = maestros.clientes_sugeine.get(normalizar(decision.cliente_sugeine_nombre or ""), "") if nivel else ""
            ganador = (proceso["_ganadores"] or [""])[0] if nivel else ""
            nivel = nivel if cliente else ""

        if archivo is None:
            # Entidad con sede en la zona pero sin regla: GRECO decide el asesor.
            hoja_sugeine = sugeine.hojas.get(HOJA_SUGEINE_OTRAS) if sugeine else None
            if nivel == CONFIRMADO and hoja_sugeine is not None:
                destino_archivo, hoja, asesor = sugeine, hoja_sugeine, mae.ASESOR_SUGEINE
            else:
                destino_archivo = reparto.archivo_sin_asignar or _archivo_sin_asignar(maestros)
                if destino_archivo is None:
                    reparto.sin_clasificar.append({**proceso, "Motivo": f"{motivo} No hay plantilla para 'Sin asignar'."})
                    reparto.destinos[proceso["idContrato"]] = "Sin clasificar"
                    continue
                reparto.archivo_sin_asignar = destino_archivo
                hoja, asesor = destino_archivo.hojas[HOJA_SIN_ASIGNAR], ASESOR_SIN_ASIGNAR
                if nivel == POR_CONFIRMAR:
                    incidencia(
                        lim.INC_REVISION,
                        f"El ganador '{ganador}' se parece al cliente de Sugeine '{cliente}'; considerar asignarla a Sugeine.",
                        proceso,
                        "Asesor",
                    )
            valores, trazabilidad, _ = _valores_fila(proceso, hoja, asesor, ASESOR_SIN_ASIGNAR, (nivel, ganador, cliente), maestros)
            valores["A"] = proceso["Entidad contratante"] or None  # como en Megaproyectos y Edificaciones y Mant.
            _aplicar_decision(valores, trazabilidad, decision, motivo)
            reparto.filas.append(FilaObra(destino_archivo, hoja, valores, trazabilidad, proceso))
            reparto.destinos[proceso["idContrato"]] = f"{destino_archivo.ruta.name} / {hoja.nombre} (Asesor: {asesor})"
            _avisos_de_fila(valores, trazabilidad, proceso, incidencia)
            continue

        asesor = archivo.asesor
        if nivel == CONFIRMADO:
            asesor = mae.ASESOR_SUGEINE
        elif nivel == POR_CONFIRMAR:
            posible = f"{archivo.asesor} (posible Sugeine)"
            asesor = posible if posible in maestros.asesores else archivo.asesor
            incidencia(
                lim.INC_REVISION,
                f"El ganador '{ganador}' se parece al cliente de Sugeine '{cliente}'; confirmar si la obra es de Sugeine.",
                proceso,
                "Asesor",
            )

        valores, trazabilidad, avisos = _valores_fila(proceso, hoja, asesor, archivo.asesor, (nivel, ganador, cliente), maestros)
        if categoria_sin_regla:
            valores["A"] = proceso["Entidad contratante"] or None  # entidad sin categoria propia en la hoja
        _aplicar_decision(valores, trazabilidad, decision, motivo if decision is not None else "")
        reparto.filas.append(FilaObra(archivo, hoja, valores, trazabilidad, proceso))
        destino = f"{archivo.ruta.name} / {hoja.nombre} (Asesor: {asesor})"

        if nivel == CONFIRMADO and sugeine is not None:
            # Las edificaciones de clientes de Sugeine van a su hoja "Edificaciones y Mant.".
            es_edificacion = categoria_sin_regla and valores.get("L") == TIPO_EDIFICACIONES
            hoja_sugeine = sugeine.hojas.get(HOJA_SUGEINE_OTRAS if es_edificacion else nombre_hoja) or sugeine.hojas.get(nombre_hoja)
            if hoja_sugeine is None:
                incidencia(
                    lim.INC_REVISION,
                    f"Cliente de Sugeine, pero su archivo no tiene la hoja '{nombre_hoja}': solo se agrega en {archivo.ruta.name}.",
                    proceso,
                    "Asesor",
                )
            else:
                valores_s, traz_s, _ = _valores_fila(
                    proceso, hoja_sugeine, mae.ASESOR_SUGEINE, archivo.asesor, (nivel, ganador, cliente), maestros
                )
                _aplicar_decision(valores_s, traz_s, decision, motivo if decision is not None else "")
                reparto.filas.append(FilaObra(sugeine, hoja_sugeine, valores_s, traz_s, proceso))
                destino += f" y {sugeine.ruta.name} / {hoja_sugeine.nombre}"
        reparto.destinos[proceso["idContrato"]] = destino
        _avisos_de_fila(valores, trazabilidad, proceso, incidencia)
    return reparto


def _aplicar_decision(valores: dict[str, Any], trazabilidad: dict[str, Any], decision: Any, motivo: str) -> None:
    """Usa el tipo de contrato y la localidad que propuso el LLM, y deja su justificacion como trazabilidad."""
    if decision is None:
        return
    valores["L"] = decision.tipo_contrato
    if decision.localidad:
        valores["B"] = decision.localidad
    valores["J"] = valores["J"].replace("propuesto por palabras clave", "propuesto por el LLM")
    if decision.confianza == "baja":
        valores["J"] += " Asignación del LLM con confianza BAJA: confirmar el asesor."
    trazabilidad["SAML · Clasificación"] = motivo


def _avisos_de_fila(valores: dict[str, Any], trazabilidad: dict[str, Any], proceso: dict[str, Any], incidencia) -> None:
    if valores["N"] is None:
        incidencia(lim.INC_SIN_DATO, "Sin valor: 'Valor Millones' queda vacío y la obra no suma potencial.", proceso, "Valor Millones")
    if not valores["I"]:
        incidencia(lim.INC_SIN_DATO, "Sin ganador en origen: 'Empresa' queda vacía.", proceso, "Empresa")
    if not valores["B"]:
        incidencia(lim.INC_REVISION, "No se pudo determinar la Localidad.", proceso, "Localidad")
    if valores["Q"] is None:
        incidencia(lim.INC_REVISION, trazabilidad["SAML · Origen de Inicio y Meses"], proceso, "Meses por Terminar")
