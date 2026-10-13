"""Limpieza intermedia de licitaciones adjudicadas (version local de demostracion).

Lee los CSV generados por Ingesta/extract.py, los consolida por idContrato y
prepara las filas de las hojas del Excel intermedio (Datos, Ganadores e
Incidencias) junto con las cifras de la hoja Resumen. Solo lee los CSV: no los
modifica ni escribe archivos. La escritura del Excel esta en
generar_excel_intermedio.py.

Las columnas de "Obras V6" son provisionales: se tomaron de una captura del
archivo de GRECO y deben confirmarse con el .xlsx original. Esta etapa no
reemplaza la integracion final con Obras V6, Supabase ni n8n.
"""
from __future__ import annotations

import ast
import csv
import hashlib
import re
import statistics
import sys
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

ESTADO_ADJUDICADO = "Adjudicado"
PATRON_ARCHIVO = re.compile(r"^licitaciones_(\d{8}_\d{6})\.csv$")
PATRON_HITO_ADJUDICACION = re.compile(r"adjudic", re.IGNORECASE)
PATRON_DIRECCION = re.compile(
    r"\b(calle|carrera|cra\.?|cl\.?|kr\.?|avenida|av\.|diagonal|transversal)\s*\d", re.IGNORECASE
)
PATRON_POSIBLE_NIT = re.compile(r"\bNIT\b|\d{9}", re.IGNORECASE)
LONGITUD_MAXIMA_OBJETO = 500  # la API parece recortar "Objeto" a 500 caracteres

# Campos del CSV que nunca se llevan al Excel ni a los mensajes: identificadores
# firmados de la plataforma (jwt_contratista_id, pivot) y el estado del usuario
# en licitaciones.info. El CSV original los conserva intactos.
CAMPOS_EXCLUIDOS = (
    "jwt_contratista_id",
    "pivot",
    "tipo",
    "leido",
    "is_fijado",
    "en_seguimiento",
    "en_favorito",
    "Grupo",
    "seguimiento_vigente",
    "fuente_enable_seguimiento_limit",
    "custom_contratista",
    "alias_fuente",
)

# Estado de cada columna; tambien define el color del encabezado en el Excel.
MAPEADO = "Mapeado"
PROVISIONAL = "Provisional"
SIN_FUENTE = "Sin fuente"
MANUAL = "Manual GRECO"
CALCULADO = "Calculado GRECO"
TRAZABILIDAD = "Trazabilidad"


@dataclass(frozen=True)
class Columna:
    nombre: str
    estado: str
    origen: str
    regla: str


_MANUAL = "Lo diligencia el equipo comercial de GRECO; no proviene de fuentes públicas."
_CALCULADO = "Depende de fórmulas o datos del archivo original de GRECO; se deja vacío."

# Orden y nombres provisionales tomados de la captura de "Obras V6".
COLUMNAS_GRECO = (
    Columna("Entidad", MAPEADO, "EntidadContratante", "Texto de origen sin espacios sobrantes."),
    Columna("Localidad", SIN_FUENTE, "-", "La API solo informa departamento y municipio (ver 'Ubicación (origen)'). No se infiere."),
    Columna("Latitud", SIN_FUENTE, "-", "No hay coordenadas en origen. No se usa geocodificación."),
    Columna("Longitud", SIN_FUENTE, "-", "No hay coordenadas en origen. No se usa geocodificación."),
    Columna("Asesor", MANUAL, "-", _MANUAL),
    Columna("Despacho/Disposición", MANUAL, "-", _MANUAL),
    Columna("Contactada", MANUAL, "-", _MANUAL),
    Columna(
        "NIT",
        SIN_FUENTE,
        "-",
        "Ningún campo de la API contiene el NIT. 'id_contratista_verificado' es un ID interno de "
        "licitaciones.info (2 a 6 dígitos), no un NIT.",
    ),
    Columna(
        "Empresa",
        PROVISIONAL,
        "ganadores[].nombreEntidad",
        "Nombre del ganador; si hay varios, separados por ' | '. Detalle en la hoja Ganadores.",
    ),
    Columna("Notas", MANUAL, "-", _MANUAL),
    Columna(
        "Nombre de la Obra",
        SIN_FUENTE,
        "-",
        "Pendiente de confirmar con Obras V6. Candidato: 'Objeto (origen)'. "
        "'Nombre (origen)' es la modalidad más el código del proceso.",
    ),
    Columna(
        "Contrato No",
        SIN_FUENTE,
        "-",
        "La API no trae número de contrato; 'CodigoProceso' es el código del proceso de selección.",
    ),
    Columna(
        "Valor Millones",
        MAPEADO,
        "Valor",
        "Valor (en pesos) / 1.000.000. Un Valor de 0 se deja vacío y se registra como incidencia.",
    ),
    Columna(
        "Fecha Adjudicación",
        SIN_FUENTE,
        "-",
        "El cronograma solo trae fechas programadas, a veces varias por proceso; no se identifica la "
        "fecha real con certeza. Ver 'Adjudicación programada (cronograma)'.",
    ),
    Columna("Avance %", CALCULADO, "-", _CALCULADO),
    Columna("Meses por Terminar", CALCULADO, "-", _CALCULADO),
    Columna("Valor por Ejecutar", CALCULADO, "-", _CALCULADO),
    Columna("Potencial Infraestructura", CALCULADO, "-", _CALCULADO),
    Columna("Meta 25% Material", CALCULADO, "-", _CALCULADO),
    Columna("MS Greco", CALCULADO, "-", _CALCULADO),
)

COLUMNAS_TRAZABILIDAD = (
    Columna("idContrato", TRAZABILIDAD, "idContrato", "Clave del proceso en licitaciones.info; única en esta hoja."),
    Columna("CodigoProceso", TRAZABILIDAD, "CodigoProceso", "Código del proceso en SECOP; se conserva como texto."),
    Columna("Fuente", TRAZABILIDAD, "nombre_fuente", "Fuente original del proceso (SECOP I / SECOP II)."),
    Columna("Enlace", TRAZABILIDAD, "link", "Ficha del proceso en licitaciones.info."),
    Columna("Nombre (origen)", TRAZABILIDAD, "Nombre", "Modalidad y código del proceso, tal como llegan."),
    Columna("Objeto (origen)", TRAZABILIDAD, "Objeto", "Descripción del contrato; espacios sobrantes normalizados."),
    Columna("Valor (origen, pesos)", TRAZABILIDAD, "Valor", "Valor tal como llega de la API."),
    Columna("Fecha publicación (origen)", TRAZABILIDAD, "FechaPublicacion", "No es la fecha de adjudicación."),
    Columna("Fecha vencimiento (origen)", TRAZABILIDAD, "FechaVencimiento", "No es la fecha de adjudicación."),
    Columna("Ubicación (origen)", TRAZABILIDAD, "TextoDepartamento", "Departamento : municipios, sin modificar."),
    Columna("Direccion", TRAZABILIDAD, "-", "La API no trae un campo de dirección; queda vacío."),
    Columna("Num. ganadores", TRAZABILIDAD, "ganadores", "Ganadores distintos del proceso."),
    Columna("Num. proponentes", TRAZABILIDAD, "proponentes", "Proponentes registrados en el proceso."),
    Columna(
        "Adjudicación programada (cronograma)",
        TRAZABILIDAD,
        "cronogramas[] con 'adjudic' en el hito",
        "Fechas programadas, no confirmadas; no se usan como Fecha Adjudicación.",
    ),
    Columna("Estado (origen)", TRAZABILIDAD, "estado_agrupado", "Solo se incluyen procesos 'Adjudicado'."),
    Columna("Archivo origen", TRAZABILIDAD, "-", "CSV de Ingesta/output del que proviene la fila."),
    Columna("Fecha extracción", TRAZABILIDAD, "-", "Fecha y hora de la ingesta, tomada del nombre del CSV."),
)

COLUMNAS_DATOS = COLUMNAS_GRECO + COLUMNAS_TRAZABILIDAD

COLUMNAS_GANADORES = (
    Columna("idContrato", TRAZABILIDAD, "ganadores[].idContrato", "Relaciona con la hoja Datos."),
    Columna("CodigoProceso", TRAZABILIDAD, "CodigoProceso", "Relaciona con la hoja Datos."),
    Columna("Orden", TRAZABILIDAD, "-", "Posición del ganador en la lista de origen."),
    Columna("Empresa", PROVISIONAL, "ganadores[].nombreEntidad", "Nombre del ganador sin espacios sobrantes."),
    Columna("NIT", SIN_FUENTE, "-", "No disponible en la API."),
    Columna("Es consorcio (origen)", MAPEADO, "ganadores[].es_consorcio", "Sí / No / Sin dato."),
    Columna(
        "Integrantes del consorcio",
        SIN_FUENTE,
        "-",
        "La API solo trae el nombre del consorcio, no sus integrantes. No se infieren.",
    ),
    Columna("Está en proponentes", TRAZABILIDAD, "proponentes[].idProponente", "Si el ganador figura entre los proponentes."),
    Columna("ID proponente (licitaciones.info)", TRAZABILIDAD, "ganadores[].idProponente", "Identificador interno de la plataforma."),
    Columna(
        "ID contratista verificado (licitaciones.info)",
        TRAZABILIDAD,
        "ganadores[].id_contratista_verificado",
        "Identificador interno de la plataforma; no es un NIT.",
    ),
    Columna("Repeticiones en origen", TRAZABILIDAD, "-", "Veces que la misma entrada aparece en la lista de ganadores."),
)

COLUMNAS_INCIDENCIAS = ("Tipo", "Alcance", "idContrato", "CodigoProceso", "Hoja", "Campo", "Detalle")

# Tipos de incidencia
INC_SIN_DATO = "Campo sin dato en origen"
INC_SIN_FUENTE = "Columna sin fuente"
INC_PROVISIONAL = "Mapeo provisional"
INC_CALIDAD = "Calidad de datos"
INC_REVISION = "Revisión sugerida"
INC_DUPLICADO = "Duplicado"
INC_EXCLUIDO = "Excluido"
INC_NO_TRANSFORMABLE = "No transformable"


@dataclass
class ArchivoFuente:
    ruta: Path
    fecha_extraccion: Optional[datetime]
    sha256: str
    filas: int = 0


@dataclass
class Resultado:
    archivos: list[ArchivoFuente]
    datos: list[dict[str, Any]]
    ganadores: list[dict[str, Any]]
    incidencias: list[dict[str, Any]]
    metricas: dict[str, Any] = field(default_factory=dict)


# ----------------------------------------------------------------------
# Lectura y consolidacion
# ----------------------------------------------------------------------
def fecha_desde_nombre(ruta: Path) -> Optional[datetime]:
    coincidencia = PATRON_ARCHIVO.match(ruta.name)
    if not coincidencia:
        return None
    return datetime.strptime(coincidencia.group(1), "%Y%m%d_%H%M%S")


def sha256_de(ruta: Path) -> str:
    h = hashlib.sha256()
    with ruta.open("rb") as f:
        for bloque in iter(lambda: f.read(1 << 20), b""):
            h.update(bloque)
    return h.hexdigest()


def listar_csvs(directorio: Path) -> list[ArchivoFuente]:
    """CSV de la ingesta presentes en la carpeta, del mas antiguo al mas reciente."""
    archivos = [
        ArchivoFuente(ruta=ruta, fecha_extraccion=fecha_desde_nombre(ruta), sha256=sha256_de(ruta))
        for ruta in directorio.glob("licitaciones_*.csv")
    ]
    archivos.sort(key=lambda a: (a.fecha_extraccion or datetime.fromtimestamp(a.ruta.stat().st_mtime), a.ruta.name))
    return archivos


def leer_filas(archivo: ArchivoFuente) -> list[dict[str, str]]:
    csv.field_size_limit(sys.maxsize)
    with archivo.ruta.open(newline="", encoding="utf-8") as f:
        lector = csv.DictReader(f)
        if lector.fieldnames == ["Sin resultados"]:
            return []
        return list(lector)


# ----------------------------------------------------------------------
# Normalizacion de valores
# ----------------------------------------------------------------------
def limpiar_texto(valor: Any) -> str:
    """Quita espacios al inicio/final y colapsa espacios y saltos de linea internos."""
    if valor is None:
        return ""
    return re.sub(r"\s+", " ", str(valor)).strip()


def fecha_iso(valor: str) -> Optional[datetime]:
    """'2026-02-12 10:30:00' -> datetime; vacio -> None; formato desconocido -> ValueError."""
    texto = limpiar_texto(valor)
    if not texto:
        return None
    return datetime.strptime(texto, "%Y-%m-%d %H:%M:%S")


def fecha_cronograma(valor: str) -> datetime:
    """'12/02/2026 - 10:30 AM' (dia/mes/anio) -> datetime."""
    return datetime.strptime(limpiar_texto(valor), "%d/%m/%Y - %I:%M %p")


def lista_anidada(valor: str) -> list:
    """Interpreta las listas que extract.py guarda como texto de Python."""
    texto = (valor or "").strip()
    if not texto:
        return []
    resultado = ast.literal_eval(texto)
    if not isinstance(resultado, list):
        raise ValueError("no es una lista")
    return resultado


def si_no(valor: Any) -> str:
    if valor is True:
        return "Sí"
    if valor is False:
        return "No"
    return "Sin dato"


# ----------------------------------------------------------------------
# Transformacion
# ----------------------------------------------------------------------
class _Registro:
    """Acumula las incidencias de una ejecucion."""

    def __init__(self) -> None:
        self.incidencias: list[dict[str, Any]] = []

    def agregar(
        self,
        tipo: str,
        detalle: str,
        *,
        alcance: str = "Registro",
        id_contrato: str = "",
        codigo: str = "",
        hoja: str = "Datos",
        campo: str = "",
    ) -> None:
        self.incidencias.append(
            {
                "Tipo": tipo,
                "Alcance": alcance,
                "idContrato": id_contrato,
                "CodigoProceso": codigo,
                "Hoja": hoja,
                "Campo": campo,
                "Detalle": detalle,
            }
        )


def consolidar(archivos: list[ArchivoFuente], registro: _Registro) -> tuple[list[tuple[dict, ArchivoFuente]], dict]:
    """Une los CSV por idContrato; ante repetidos gana la version del archivo mas reciente."""
    por_id: dict[str, tuple[dict, ArchivoFuente]] = {}
    conteo = Counter()

    for archivo in archivos:
        filas = leer_filas(archivo)
        archivo.filas = len(filas)
        conteo["filas_leidas"] += len(filas)
        vistos_en_archivo: set[str] = set()

        for fila in filas:
            id_contrato = limpiar_texto(fila.get("idContrato"))
            codigo = limpiar_texto(fila.get("CodigoProceso"))
            if not id_contrato:
                conteo["sin_id"] += 1
                registro.agregar(
                    INC_NO_TRANSFORMABLE,
                    f"Fila sin idContrato en {archivo.ruta.name}; no se puede consolidar y se excluye.",
                    codigo=codigo,
                    campo="idContrato",
                )
                continue

            if id_contrato in vistos_en_archivo:
                conteo["duplicados_mismo_archivo"] += 1
                registro.agregar(
                    INC_DUPLICADO,
                    f"idContrato repetido dentro de {archivo.ruta.name}; se conserva la última aparición.",
                    id_contrato=id_contrato,
                    codigo=codigo,
                    campo="idContrato",
                )
            elif id_contrato in por_id:
                conteo["reemplazados_por_version_reciente"] += 1
                anterior = por_id[id_contrato][1].ruta.name
                registro.agregar(
                    INC_DUPLICADO,
                    f"idContrato presente en {anterior} y {archivo.ruta.name}; se usa la versión de {archivo.ruta.name}.",
                    id_contrato=id_contrato,
                    codigo=codigo,
                    campo="idContrato",
                )

            vistos_en_archivo.add(id_contrato)
            por_id[id_contrato] = (fila, archivo)

    return list(por_id.values()), dict(conteo)


def _resumen_adjudicacion(cronogramas: list, registro: _Registro, id_contrato: str, codigo: str) -> str:
    fechas: set[datetime] = set()
    for hito in cronogramas:
        if not isinstance(hito, dict) or not PATRON_HITO_ADJUDICACION.search(str(hito.get("nombre", ""))):
            continue
        try:
            fechas.add(fecha_cronograma(str(hito.get("fecha", ""))))
        except ValueError:
            registro.agregar(
                INC_NO_TRANSFORMABLE,
                "Fecha de un hito de adjudicación con formato no reconocido.",
                id_contrato=id_contrato,
                codigo=codigo,
                campo="cronogramas",
            )
    if not fechas:
        return ""
    ordenadas = sorted(fechas)
    if len(ordenadas) == 1:
        return ordenadas[0].strftime("%d/%m/%Y %H:%M")
    return f"{len(ordenadas)} fechas, de {ordenadas[0]:%d/%m/%Y} a {ordenadas[-1]:%d/%m/%Y}"


def _transformar_ganadores(
    ganadores: list, proponentes: list, registro: _Registro, id_contrato: str, codigo: str
) -> list[dict[str, Any]]:
    ids_proponentes = {p.get("idProponente") for p in proponentes if isinstance(p, dict)}
    filas: list[dict[str, Any]] = []
    por_clave: dict[tuple, dict[str, Any]] = {}

    for entrada in ganadores:
        if not isinstance(entrada, dict):
            registro.agregar(
                INC_NO_TRANSFORMABLE,
                "Entrada de ganador sin estructura reconocible.",
                id_contrato=id_contrato,
                codigo=codigo,
                hoja="Ganadores",
                campo="ganadores",
            )
            continue

        # Solo se usan campos seguros; jwt_contratista_id y pivot se ignoran.
        clave = (
            entrada.get("idProponente"),
            entrada.get("nombreEntidad"),
            entrada.get("es_consorcio"),
            entrada.get("id_contratista_verificado"),
        )
        if clave in por_clave:
            por_clave[clave]["Repeticiones en origen"] += 1
            continue

        nombre = limpiar_texto(entrada.get("nombreEntidad"))
        fila = {
            "idContrato": id_contrato,
            "CodigoProceso": codigo,
            "Orden": len(filas) + 1,
            "Empresa": nombre,
            "NIT": None,
            "Es consorcio (origen)": si_no(entrada.get("es_consorcio")),
            "Integrantes del consorcio": None,
            "Está en proponentes": "Sí" if entrada.get("idProponente") in ids_proponentes else "No",
            "ID proponente (licitaciones.info)": entrada.get("idProponente"),
            "ID contratista verificado (licitaciones.info)": entrada.get("id_contratista_verificado"),
            "Repeticiones en origen": 1,
        }
        por_clave[clave] = fila
        filas.append(fila)

        if not nombre:
            registro.agregar(INC_SIN_DATO, "Ganador sin nombre.", id_contrato=id_contrato, codigo=codigo, hoja="Ganadores", campo="nombreEntidad")
        if entrada.get("es_consorcio") is None:
            registro.agregar(
                INC_SIN_DATO,
                f"Ganador '{nombre}' sin dato de es_consorcio.",
                id_contrato=id_contrato,
                codigo=codigo,
                hoja="Ganadores",
                campo="es_consorcio",
            )
        if fila["Está en proponentes"] == "No":
            registro.agregar(
                INC_CALIDAD,
                f"El ganador '{nombre}' no figura en la lista de proponentes del proceso.",
                id_contrato=id_contrato,
                codigo=codigo,
                hoja="Ganadores",
                campo="proponentes",
            )
        if PATRON_POSIBLE_NIT.search(nombre):
            registro.agregar(
                INC_REVISION,
                "El nombre del ganador parece incluir un NIT; no se extrae automáticamente.",
                id_contrato=id_contrato,
                codigo=codigo,
                hoja="Ganadores",
                campo="nombreEntidad",
            )

    for fila in filas:
        if fila["Repeticiones en origen"] > 1:
            registro.agregar(
                INC_CALIDAD,
                f"El ganador '{fila['Empresa']}' aparece {fila['Repeticiones en origen']} veces con datos idénticos; "
                "se muestra una sola fila.",
                id_contrato=id_contrato,
                codigo=codigo,
                hoja="Ganadores",
                campo="ganadores",
            )
    return filas


def _transformar_proceso(
    fila: dict[str, str], archivo: ArchivoFuente, registro: _Registro
) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    id_contrato = limpiar_texto(fila.get("idContrato"))
    codigo = limpiar_texto(fila.get("CodigoProceso"))

    def incidencia(tipo: str, detalle: str, campo: str) -> None:
        registro.agregar(tipo, detalle, id_contrato=id_contrato, codigo=codigo, campo=campo)

    anidados: dict[str, list] = {}
    for campo in ("ganadores", "proponentes", "cronogramas"):
        try:
            anidados[campo] = lista_anidada(fila.get(campo, ""))
        except (ValueError, SyntaxError):
            anidados[campo] = []
            incidencia(INC_NO_TRANSFORMABLE, f"No se pudo interpretar la lista '{campo}'.", campo)

    ganadores = _transformar_ganadores(anidados["ganadores"], anidados["proponentes"], registro, id_contrato, codigo)
    if not anidados["ganadores"]:
        incidencia(INC_SIN_DATO, "Proceso adjudicado sin ganadores en origen; 'Empresa' queda vacía.", "ganadores")
    if not anidados["proponentes"]:
        incidencia(INC_SIN_DATO, "Proceso sin proponentes en origen.", "proponentes")
    if len(ganadores) > 1:
        incidencia(INC_REVISION, f"Proceso con {len(ganadores)} ganadores; ver hoja Ganadores.", "ganadores")

    valor_origen: Optional[float] = None
    valor_millones: Optional[float] = None
    texto_valor = limpiar_texto(fila.get("Valor"))
    if not texto_valor:
        incidencia(INC_SIN_DATO, "Valor vacío en origen.", "Valor")
    else:
        try:
            valor_origen = float(texto_valor)
        except ValueError:
            incidencia(INC_NO_TRANSFORMABLE, "Valor no numérico en origen.", "Valor")
        else:
            if valor_origen == 0:
                incidencia(INC_SIN_DATO, "Valor igual a 0 en origen; 'Valor Millones' queda vacío.", "Valor")
            else:
                valor_millones = valor_origen / 1_000_000

    fechas: dict[str, Any] = {}
    for campo in ("FechaPublicacion", "FechaVencimiento"):
        try:
            fechas[campo] = fecha_iso(fila.get(campo, ""))
        except ValueError:
            fechas[campo] = limpiar_texto(fila.get(campo))
            incidencia(INC_NO_TRANSFORMABLE, f"{campo} con formato no reconocido; se conserva el texto.", campo)
        else:
            if fechas[campo] is None:
                incidencia(INC_SIN_DATO, f"{campo} vacía en origen; no se completa.", campo)

    objeto = limpiar_texto(fila.get("Objeto"))
    if len(objeto) == LONGITUD_MAXIMA_OBJETO:
        incidencia(INC_CALIDAD, f"Objeto con exactamente {LONGITUD_MAXIMA_OBJETO} caracteres; posiblemente recortado por la API.", "Objeto")
    if re.search(r"localidad", objeto, re.IGNORECASE):
        incidencia(INC_REVISION, "El Objeto menciona una localidad; 'Localidad' no se infiere del texto.", "Objeto")
    if PATRON_DIRECCION.search(objeto):
        incidencia(INC_REVISION, "El Objeto parece mencionar una dirección; 'Direccion' no se extrae del texto.", "Objeto")

    entidad = limpiar_texto(fila.get("EntidadContratante"))
    if not entidad:
        incidencia(INC_SIN_DATO, "EntidadContratante vacía en origen.", "EntidadContratante")

    datos = {columna.nombre: None for columna in COLUMNAS_DATOS}
    datos.update(
        {
            "Entidad": entidad,
            "Empresa": " | ".join(g["Empresa"] for g in ganadores if g["Empresa"]),
            "Valor Millones": valor_millones,
            "idContrato": id_contrato,
            "CodigoProceso": codigo,
            "Fuente": limpiar_texto(fila.get("nombre_fuente")),
            "Enlace": limpiar_texto(fila.get("link")),
            "Nombre (origen)": limpiar_texto(fila.get("Nombre")),
            "Objeto (origen)": objeto,
            "Valor (origen, pesos)": valor_origen if valor_origen is not None else texto_valor,
            "Fecha publicación (origen)": fechas["FechaPublicacion"],
            "Fecha vencimiento (origen)": fechas["FechaVencimiento"],
            "Ubicación (origen)": limpiar_texto(fila.get("TextoDepartamento")),
            "Num. ganadores": len(ganadores),
            "Num. proponentes": len(anidados["proponentes"]),
            "Adjudicación programada (cronograma)": _resumen_adjudicacion(
                anidados["cronogramas"], registro, id_contrato, codigo
            ),
            "Estado (origen)": limpiar_texto(fila.get("estado_agrupado")),
            "Archivo origen": archivo.ruta.name,
            "Fecha extracción": archivo.fecha_extraccion,
        }
    )
    return datos, ganadores


def _incidencias_de_columnas(registro: _Registro) -> None:
    """Una incidencia por columna sin fuente o con mapeo provisional (afecta a todas las filas)."""
    for hoja, columnas in (("Datos", COLUMNAS_GRECO), ("Ganadores", COLUMNAS_GANADORES)):
        for columna in columnas:
            if columna.estado == SIN_FUENTE:
                registro.agregar(INC_SIN_FUENTE, columna.regla, alcance="Columna", hoja=hoja, campo=columna.nombre)
            elif columna.estado == PROVISIONAL:
                registro.agregar(
                    INC_PROVISIONAL,
                    f"Mapeo provisional desde '{columna.origen}': {columna.regla} Confirmar con Obras V6.",
                    alcance="Columna",
                    hoja=hoja,
                    campo=columna.nombre,
                )
    registro.agregar(
        INC_PROVISIONAL,
        "Los nombres y el orden de las columnas GRECO se tomaron de una captura de Obras V6; "
        "algunos encabezados estaban cortados. Confirmar con el .xlsx original.",
        alcance="Columna",
        campo="(encabezados GRECO)",
    )


def procesar(directorio_entrada: Path) -> Resultado:
    """Ejecuta la limpieza completa sobre los CSV presentes en la carpeta de la ingesta."""
    archivos = listar_csvs(directorio_entrada)
    if not archivos:
        raise FileNotFoundError(
            f"No hay archivos licitaciones_*.csv en {directorio_entrada}. Ejecuta primero Ingesta/extract.py."
        )

    registro = _Registro()
    consolidados, conteo = consolidar(archivos, registro)

    datos: list[dict[str, Any]] = []
    ganadores: list[dict[str, Any]] = []
    estados = Counter()
    for fila, archivo in consolidados:
        estado = limpiar_texto(fila.get("estado_agrupado"))
        estados[estado or "(vacío)"] += 1
        if estado != ESTADO_ADJUDICADO:
            registro.agregar(
                INC_EXCLUIDO,
                f"Estado '{estado or '(vacío)'}' distinto de '{ESTADO_ADJUDICADO}'; el proceso no se incluye.",
                id_contrato=limpiar_texto(fila.get("idContrato")),
                codigo=limpiar_texto(fila.get("CodigoProceso")),
                campo="estado_agrupado",
            )
            continue
        fila_datos, filas_ganadores = _transformar_proceso(fila, archivo, registro)
        datos.append(fila_datos)
        ganadores.extend(filas_ganadores)

    codigos = Counter(d["CodigoProceso"] for d in datos if d["CodigoProceso"])
    for d in datos:
        if codigos[d["CodigoProceso"]] > 1:
            registro.agregar(
                INC_DUPLICADO,
                "CodigoProceso compartido por varios idContrato; se conservan todos.",
                id_contrato=d["idContrato"],
                codigo=d["CodigoProceso"],
                campo="CodigoProceso",
            )

    _incidencias_de_columnas(registro)

    valores = [d["Valor (origen, pesos)"] for d in datos if isinstance(d["Valor (origen, pesos)"], float)]
    no_cero = [v for v in valores if v > 0]
    metricas = {
        "filas_leidas": conteo.get("filas_leidas", 0),
        "procesos_unicos": len(consolidados),
        "duplicados_mismo_archivo": conteo.get("duplicados_mismo_archivo", 0),
        "reemplazados_por_version_reciente": conteo.get("reemplazados_por_version_reciente", 0),
        "filas_sin_id": conteo.get("sin_id", 0),
        "estados": dict(estados),
        "incluidos": len(datos),
        "excluidos_por_estado": sum(n for e, n in estados.items() if e != ESTADO_ADJUDICADO),
        "ids_unicos_en_datos": len({d["idContrato"] for d in datos}) == len(datos),
        "codigos_duplicados": sum(1 for n in codigos.values() if n > 1),
        "filas_ganadores": len(ganadores),
        "procesos_varios_ganadores": sum(1 for d in datos if (d["Num. ganadores"] or 0) > 1),
        "procesos_sin_ganador": sum(1 for d in datos if not d["Num. ganadores"]),
        "valor_minimo_no_cero": min(no_cero) if no_cero else None,
        "valor_mediana": statistics.median(no_cero) if no_cero else None,
        "incidencias_por_tipo": dict(Counter(i["Tipo"] for i in registro.incidencias)),
    }
    return Resultado(archivos, datos, ganadores, registro.incidencias, metricas)
