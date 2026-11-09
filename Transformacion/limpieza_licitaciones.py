"""Limpieza de las licitaciones adjudicadas que trae la ingesta.

Lee los CSV generados por Ingesta/extract.py, los consolida por idContrato y
deja un registro limpio por proceso adjudicado (con sus ganadores y las fechas
de ejecucion programadas en el cronograma), mas las incidencias encontradas.
Solo lee los CSV: no los modifica ni escribe archivos.

El reparto de cada proceso al archivo de un asesor y su conversion a las 71
columnas de GRECO estan en obras_asesores.py; la escritura de los Excel, en
generar_archivos_asesores.py.
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
# Hitos del cronograma con la fecha programada de inicio y de fin de la ejecucion del contrato.
PATRON_HITO_INICIO = re.compile(r"inicio\s+de\s+(la\s+)?ejecuci", re.IGNORECASE)
PATRON_HITO_FIN = re.compile(r"^\s*(plazo\s+de\s+ejecuci|ejecuci[oó]n\s+del\s+contrato)", re.IGNORECASE)
PATRON_DIRECCION = re.compile(
    r"\b(calle|carrera|cra\.?|cl\.?|kr\.?|avenida|av\.|diagonal|transversal)\s*\d", re.IGNORECASE
)
PATRON_POSIBLE_NIT = re.compile(r"\bNIT\b|\d{9}", re.IGNORECASE)
LONGITUD_MAXIMA_OBJETO = 500  # la API parece recortar "Objeto" a 500 caracteres

# Campos del CSV que nunca se llevan a los Excel ni a los mensajes: identificadores
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


@dataclass(frozen=True)
class Columna:
    nombre: str
    origen: str
    regla: str


# Columnas del registro limpio de cada proceso (hoja "Procesos" del Excel de revision).
COLUMNAS_PROCESOS = (
    Columna("idContrato", "idContrato", "Clave del proceso en licitaciones.info; única por proceso."),
    Columna("CodigoProceso", "CodigoProceso", "Código del proceso en SECOP; se conserva como texto."),
    Columna("Entidad contratante", "EntidadContratante", "Texto de origen sin espacios sobrantes."),
    Columna("Objeto", "Objeto", "Descripción del contrato; espacios sobrantes normalizados."),
    Columna("Empresa (ganadores)", "ganadores[].nombreEntidad", "Ganadores distintos, separados por ' | '."),
    Columna("Valor Millones", "Valor", "Valor (en pesos) / 1.000.000. Un Valor de 0 queda vacío y se registra."),
    Columna("Valor (origen, pesos)", "Valor", "Valor tal como llega de la API."),
    Columna("Ubicación (origen)", "TextoDepartamento", "Departamento : municipios, sin modificar."),
    Columna("Fuente", "nombre_fuente", "Fuente original del proceso (SECOP I / SECOP II)."),
    Columna("Enlace", "link", "Ficha del proceso en licitaciones.info."),
    Columna("Nombre (origen)", "Nombre", "Modalidad y código del proceso, tal como llegan."),
    Columna("Fecha publicación (origen)", "FechaPublicacion", "No es la fecha de adjudicación."),
    Columna("Fecha vencimiento (origen)", "FechaVencimiento", "No es la fecha de adjudicación."),
    Columna("Num. ganadores", "ganadores", "Ganadores distintos del proceso."),
    Columna("Num. proponentes", "proponentes", "Proponentes registrados en el proceso."),
    Columna(
        "Adjudicación programada (cronograma)",
        "cronogramas[] con 'adjudic' en el hito",
        "Fechas programadas, no confirmadas; no se usan como Fecha Adjudicación.",
    ),
    Columna(
        "Inicio ejecución (cronograma)",
        "cronogramas[] 'inicio de ejecución'",
        "Fecha programada de inicio; solo si el cronograma trae una única fecha.",
    ),
    Columna(
        "Fin ejecución (cronograma)",
        "cronogramas[] 'plazo de ejecución' / 'ejecución del contrato'",
        "Fecha programada de fin; solo si el cronograma trae una única fecha.",
    ),
    Columna("Estado (origen)", "estado_agrupado", "Solo se incluyen procesos 'Adjudicado'."),
    Columna("Archivo origen", "-", "CSV de Ingesta/output del que proviene la fila."),
    Columna("Fecha extracción", "-", "Fecha y hora de la ingesta, tomada del nombre del CSV."),
)

COLUMNAS_GANADORES = (
    Columna("idContrato", "ganadores[].idContrato", "Relaciona con la hoja Procesos."),
    Columna("CodigoProceso", "CodigoProceso", "Relaciona con la hoja Procesos."),
    Columna("Orden", "-", "Posición del ganador en la lista de origen."),
    Columna("Empresa", "ganadores[].nombreEntidad", "Nombre del ganador sin espacios sobrantes."),
    Columna("Es consorcio (origen)", "ganadores[].es_consorcio", "Sí / No / Sin dato."),
    Columna("Está en proponentes", "proponentes[].idProponente", "Si el ganador figura entre los proponentes."),
    Columna("ID proponente (licitaciones.info)", "ganadores[].idProponente", "Identificador interno de la plataforma."),
    Columna(
        "ID contratista verificado (licitaciones.info)",
        "ganadores[].id_contratista_verificado",
        "Identificador interno de la plataforma; no es un NIT.",
    ),
    Columna("Repeticiones en origen", "-", "Veces que la misma entrada aparece en la lista de ganadores."),
)

COLUMNAS_INCIDENCIAS = ("Tipo", "Alcance", "idContrato", "CodigoProceso", "Hoja", "Campo", "Detalle")

# Tipos de incidencia
INC_SIN_DATO = "Campo sin dato en origen"
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
    csv.field_size_limit(min(sys.maxsize, 2**31 - 1))
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
    if valor is True or str(valor).strip().lower() == "true":
        return "Sí"
    if valor is False or str(valor).strip().lower() == "false":
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
        hoja: str = "Procesos",
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


def _fechas_de_hitos(
    cronogramas: list, patron: re.Pattern, registro: _Registro, id_contrato: str, codigo: str
) -> list[datetime]:
    """Fechas distintas (por dia) de los hitos del cronograma cuyo nombre cumple el patron."""
    fechas: set[datetime] = set()
    for hito in cronogramas:
        if not isinstance(hito, dict) or not patron.search(str(hito.get("nombre", ""))):
            continue
        try:
            fechas.add(fecha_cronograma(str(hito.get("fecha", ""))))
        except ValueError:
            registro.agregar(
                INC_NO_TRANSFORMABLE,
                f"Fecha del hito '{limpiar_texto(hito.get('nombre'))[:60]}' con formato no reconocido.",
                id_contrato=id_contrato,
                codigo=codigo,
                campo="cronogramas",
            )
    por_dia = {f.replace(hour=0, minute=0, second=0, microsecond=0) for f in fechas}
    return sorted(por_dia)


def _resumen_fechas(fechas: list[datetime]) -> str:
    if not fechas:
        return ""
    if len(fechas) == 1:
        return fechas[0].strftime("%d/%m/%Y")
    return f"{len(fechas)} fechas, de {fechas[0]:%d/%m/%Y} a {fechas[-1]:%d/%m/%Y}"


def _transformar_ganadores(
    ganadores: list, proponentes: list, registro: _Registro, id_contrato: str, codigo: str
) -> list[dict[str, Any]]:
    ids_proponentes = {str(p.get("idProponente")) for p in proponentes if isinstance(p, dict)}
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
            str(entrada.get("idProponente")),
            entrada.get("nombreEntidad"),
            str(entrada.get("es_consorcio")),
            str(entrada.get("id_contratista_verificado")),
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
            "Es consorcio (origen)": si_no(entrada.get("es_consorcio")),
            "Está en proponentes": "Sí" if str(entrada.get("idProponente")) in ids_proponentes else "No",
            "ID proponente (licitaciones.info)": entrada.get("idProponente"),
            "ID contratista verificado (licitaciones.info)": entrada.get("id_contratista_verificado"),
            "Repeticiones en origen": 1,
        }
        por_clave[clave] = fila
        filas.append(fila)

        if not nombre:
            registro.agregar(INC_SIN_DATO, "Ganador sin nombre.", id_contrato=id_contrato, codigo=codigo, hoja="Ganadores", campo="nombreEntidad")
        if entrada.get("es_consorcio") in (None, "", "None"):
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
    if PATRON_DIRECCION.search(objeto):
        incidencia(INC_REVISION, "El Objeto parece mencionar una dirección; no se extrae del texto.", "Objeto")

    entidad = limpiar_texto(fila.get("EntidadContratante"))
    if not entidad:
        incidencia(INC_SIN_DATO, "EntidadContratante vacía en origen.", "EntidadContratante")

    cronogramas = anidados["cronogramas"]
    adjudicacion = _fechas_de_hitos(cronogramas, PATRON_HITO_ADJUDICACION, registro, id_contrato, codigo)
    inicios = _fechas_de_hitos(cronogramas, PATRON_HITO_INICIO, registro, id_contrato, codigo)
    fines = _fechas_de_hitos(cronogramas, PATRON_HITO_FIN, registro, id_contrato, codigo)

    datos = {
        "idContrato": id_contrato,
        "CodigoProceso": codigo,
        "Entidad contratante": entidad,
        "Objeto": objeto,
        "Empresa (ganadores)": " | ".join(g["Empresa"] for g in ganadores if g["Empresa"]),
        "Valor Millones": valor_millones,
        "Valor (origen, pesos)": valor_origen if valor_origen is not None else texto_valor,
        "Ubicación (origen)": limpiar_texto(fila.get("TextoDepartamento")),
        "Fuente": limpiar_texto(fila.get("nombre_fuente")),
        "Enlace": limpiar_texto(fila.get("link")),
        "Nombre (origen)": limpiar_texto(fila.get("Nombre")),
        "Fecha publicación (origen)": fechas["FechaPublicacion"],
        "Fecha vencimiento (origen)": fechas["FechaVencimiento"],
        "Num. ganadores": len(ganadores),
        "Num. proponentes": len(anidados["proponentes"]),
        "Adjudicación programada (cronograma)": _resumen_fechas(adjudicacion),
        "Inicio ejecución (cronograma)": _resumen_fechas(inicios),
        "Fin ejecución (cronograma)": _resumen_fechas(fines),
        "Estado (origen)": limpiar_texto(fila.get("estado_agrupado")),
        "Archivo origen": archivo.ruta.name,
        "Fecha extracción": archivo.fecha_extraccion,
        # Campos internos (no se escriben como columnas): los usa obras_asesores.py.
        "_ganadores": [g["Empresa"] for g in ganadores if g["Empresa"]],
        "_inicios": inicios,
        "_fines": fines,
    }
    return datos, ganadores


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
    }
    return Resultado(archivos, datos, ganadores, registro.incidencias, metricas)
