"""Genera los Excel de obras nuevas para cada asesor de GRECO a partir de la ingesta.

Lee los CSV licitaciones_*.csv de Ingesta/output (sin modificarlos), los limpia
(limpieza_licitaciones.py), los reparte entre los maestros de datos/maestros
(obras_asesores.py) y escribe, en una carpeta nueva datos/salida/SAML_<fecha_hora>/:

    Nuevas_<maestro>_<fecha_hora>.xlsx   uno por asesor con obras nuevas: una hoja por
                                         cada hoja de destino, con las 71 columnas, las
                                         formulas y los formatos del maestro, listas para
                                         copiar y pegar debajo de la ultima obra.
    SAML_revision_<fecha_hora>.xlsx      Resumen, Sin clasificar, Ya en maestros, Procesos,
                                         Ganadores e Incidencias.

Nunca modifica los maestros ni sobrescribe un archivo existente.

Uso (desde la raiz del proyecto):
    .venv/bin/python Transformacion/generar_archivos_asesores.py
    .venv/bin/python Transformacion/generar_archivos_asesores.py --entrada Ingesta/output --maestros datos/maestros --salida datos/salida
"""
from __future__ import annotations

import argparse
import re
import sys
from collections import Counter
from copy import copy
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE, Cell
from openpyxl.comments import Comment
from openpyxl.formula.translate import Translator
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.worksheet import Worksheet

import actualizar_maestros as am
import limpieza_licitaciones as lim
import maestros as mae
import obras_asesores as oa

RAIZ = Path(__file__).resolve().parent.parent
ENTRADA_POR_DEFECTO = RAIZ / "Ingesta" / "output"
MAESTROS_POR_DEFECTO = RAIZ / "datos" / "maestros"
SALIDA_POR_DEFECTO = RAIZ / "datos" / "salida"

FUENTE_ENCABEZADO = Font(bold=True, color="1F1F1F")
FUENTE_TITULO = Font(bold=True, size=14)
FUENTE_SECCION = Font(bold=True, size=12, color="1F4E78")
BORDE = Border(bottom=Side(style="thin", color="808080"))
ALINEACION_ENCABEZADO = Alignment(wrap_text=True, vertical="center", horizontal="center")
RELLENO_ENCABEZADO = PatternFill("solid", fgColor="DDEBF7")
RELLENO_SAML = PatternFill("solid", fgColor="D9D9D9")
COLUMNA_TRAZABILIDAD = column_index_from_string(mae.ULTIMA_COLUMNA) + 2  # BU: se deja BT vacia
PATRON_RANGO_DATOS = re.compile(r"([A-Z]{1,3})3:(\$?)([A-Z]{1,3})(\$?)(\d+)")

ANCHOS = {
    "Entidad contratante": 40,
    "Objeto": 70,
    "Empresa (ganadores)": 40,
    "Enlace": 30,
    "Nombre (origen)": 40,
    "Ubicación (origen)": 30,
    "Destino": 60,
    "Motivo": 60,
    "Ya está en": 60,
    "Detalle": 90,
    "Empresa": 45,
    "Campo": 22,
    "Tipo": 24,
}
FORMATOS = {
    "Valor Millones": "#,##0.00",
    "Valor (origen, pesos)": "#,##0",
    "Fecha publicación (origen)": "dd/mm/yyyy hh:mm",
    "Fecha vencimiento (origen)": "dd/mm/yyyy hh:mm",
    "Fecha extracción": "dd/mm/yyyy hh:mm",
}

# Lo que SAML llena en las 71 columnas (para la hoja Resumen e Instrucciones).
MAPEO_COLUMNAS = (
    ("A", "Entidad", "Categoría de la hoja tal como la usa el maestro (p. ej. 'IDU', 'Alcaldía Local', 'Municipio')."),
    ("B", "Localidad", "Alcaldías Locales: localidad tomada del nombre de la entidad. Municipios: municipio de la ubicación. Resto: 'Bogotá D.C.'."),
    ("E", "Asesor", "Dueño de la hoja; 'Sugeine' si el ganador es su cliente; '<asesor> (posible Sugeine)' si solo se parece."),
    ("I", "Empresa", "Ganador(es) del proceso, separados por ' | '."),
    ("J", "Notas", "Etiqueta [SAML · fecha] con el origen de Inicio/Meses y lo que falta por completar."),
    ("K", "Nombre de la Obra", "Objeto del contrato (como indica la nota del maestro)."),
    ("L", "Tipo de contrato", "Propuesto por palabras clave del objeto, de la lista del maestro. Confirmar."),
    ("M", "Contrato No", "Código del proceso (CodigoProceso), como lo usa el equipo."),
    ("N", "Valor Millones", "Valor de la API (pesos) / 1.000.000."),
    ("P", "Avance %", "0."),
    ("Q", "Meses por Terminar", "Desde el cronograma (programado) solo si inicio y fin son claros; si no, vacío."),
    ("R", "Inicio de obra", "Mes de inicio programado en el cronograma si es claro; si no, mes de la extracción (mínimo el primer mes de la proyección)."),
    ("AY", "Origen del registro", "'NUEVA · SAML (Licitaciones.info, extracción <fecha>)'."),
    ("AZ", "Proceso de contratación", "Código del proceso."),
    ("BA", "Fecha de cargue", "Fecha de la extracción."),
    ("BB", "Meses de espera", "0."),
    ("BH", "Fuente del valor", "'SECOP II <código> (vía Licitaciones.info)'."),
    ("BI:BM", "Datos de Sugeine", "Solo si el ganador coincide o se parece a un cliente de 'Cobertura clientes'."),
    ("S:AX, BC:BG, BN:BS", "Fórmulas", "Copiadas de la primera fila de la hoja del maestro, ajustadas a cada fila."),
    ("C, D, F, G, H, O, BL", "Manuales", "Quedan vacías: Latitud, Longitud, Despacho, Contactado, NIT, Fecha Adjudicación y ventas."),
)


@dataclass
class ArchivoAsesor:
    ruta: Path
    maestro: mae.ArchivoMaestro
    filas: int
    por_hoja: dict[str, int]


@dataclass
class Generacion:
    carpeta: Path
    archivos_asesores: list[ArchivoAsesor]
    revision: Path
    resultado: lim.Resultado
    reparto: oa.Reparto
    maestros: mae.Maestros
    generado: datetime
    actualizacion: am.ResultadoActualizacion | None = None


def _celda(valor: Any) -> Any:
    """openpyxl rechaza caracteres de control en los textos."""
    if isinstance(valor, str):
        return ILLEGAL_CHARACTERS_RE.sub("", valor)
    return valor


def _copiar_estilo(origen: Cell, destino: Cell) -> None:
    if origen.has_style:
        destino.font = copy(origen.font)
        destino.fill = copy(origen.fill)
        destino.border = copy(origen.border)
        destino.alignment = copy(origen.alignment)
        destino.number_format = origen.number_format
        destino.protection = copy(origen.protection)


def _escribir_tabla(hoja: Worksheet, columnas: list[str], filas: Iterable[dict[str, Any]], congelar: str = "A2") -> None:
    hoja.append(columnas)
    for indice, nombre in enumerate(columnas, start=1):
        celda = hoja.cell(row=1, column=indice)
        celda.font = FUENTE_ENCABEZADO
        celda.alignment = ALINEACION_ENCABEZADO
        celda.border = BORDE
        celda.fill = RELLENO_ENCABEZADO
        hoja.column_dimensions[get_column_letter(indice)].width = ANCHOS.get(nombre, max(12, min(len(nombre) + 4, 24)))
    hoja.row_dimensions[1].height = 45
    for fila in filas:
        hoja.append([_celda(fila.get(nombre)) for nombre in columnas])
    for indice, nombre in enumerate(columnas, start=1):
        formato = FORMATOS.get(nombre)
        for celda in hoja[get_column_letter(indice)][1:]:
            celda.alignment = Alignment(vertical="top")
            if formato:
                celda.number_format = formato
            if nombre == "Enlace" and isinstance(celda.value, str) and celda.value.startswith("https://"):
                celda.hyperlink = celda.value
                celda.style = "Hyperlink"
    hoja.freeze_panes = congelar
    hoja.auto_filter.ref = hoja.dimensions


# ----------------------------------------------------------------------
# Archivo de obras nuevas de un asesor
# ----------------------------------------------------------------------
def _hoja_listas(libro: Workbook, maestros: mae.Maestros) -> None:
    hoja = libro.create_sheet(mae.HOJA_LISTAS)
    for fila in maestros.listas:
        hoja.append([_celda(v) for v in fila])
    for celda in hoja["D"][1:]:
        if isinstance(celda.value, datetime):
            celda.number_format = "mmm-yy"
    for letra in ("A", "B", "F"):
        hoja.column_dimensions[letra].width = 40


def _validaciones(hoja: Worksheet, maestros: mae.Maestros, ultima: int) -> None:
    """Las mismas listas desplegables del maestro (Tipo de contrato, Inicio de obra y Asesor)."""
    for letra, columna_lista, cantidad in (
        ("L", "A", len(maestros.tipos_contrato)),
        ("R", "D", len(maestros.meses_inicio)),
        ("E", "F", len(maestros.asesores)),
    ):
        if not cantidad:
            continue
        validacion = DataValidation(
            type="list",
            formula1=f"={mae.HOJA_LISTAS}!${columna_lista}$2:${columna_lista}${cantidad + 1}",
            allow_blank=True,
            showErrorMessage=True,
            errorStyle="stop",
        )
        validacion.add(f"{letra}3:{letra}{max(ultima, 3)}")
        hoja.add_data_validation(validacion)


def _formula_encabezado(valor: Any, ultima: int) -> Any:
    """Los totales de la fila 1 del maestro (p. ej. =SUM(BC3:BC116)) se ajustan a las filas de este archivo."""
    if isinstance(valor, str) and valor.startswith("="):
        return PATRON_RANGO_DATOS.sub(lambda m: f"{m.group(1)}3:{m.group(2)}{m.group(3)}{m.group(4)}{max(ultima, 3)}", valor)
    return valor


def _hoja_obras(libro: Workbook, hoja: mae.HojaMaestro, filas: list[oa.FilaObra], maestros: mae.Maestros) -> None:
    ws = libro.create_sheet(hoja.nombre[:31])
    ultima = mae.FILA_PLANTILLA + len(filas) - 1

    for numero_fila, celdas in enumerate(hoja.filas_encabezado, start=1):
        for origen in celdas:
            destino = ws.cell(numero_fila, origen.column, _formula_encabezado(origen.value, ultima))
            _copiar_estilo(origen, destino)
            if origen.comment is not None:
                nota = Comment(origen.comment.text, origen.comment.author or "GRECO")
                nota.width, nota.height = 360, 200
                destino.comment = nota
    for letra, ancho in hoja.anchos.items():
        ws.column_dimensions[letra].width = ancho
    ws.row_dimensions[2].height = 60

    for desplazamiento, fila in enumerate(filas):
        numero = mae.FILA_PLANTILLA + desplazamiento
        for plantilla in hoja.plantilla:
            letra = plantilla.column_letter
            valor = plantilla.value
            if isinstance(valor, str) and valor.startswith("="):
                valor = Translator(valor, origin=f"{letra}{mae.FILA_PLANTILLA}").translate_formula(f"{letra}{numero}")
            else:
                valor = _celda(fila.valores.get(letra))
            destino = ws.cell(numero, plantilla.column, valor)
            _copiar_estilo(plantilla, destino)

        for indice, nombre in enumerate(oa.COLUMNAS_TRAZABILIDAD_SAML):
            celda = ws.cell(numero, COLUMNA_TRAZABILIDAD + indice, _celda(fila.trazabilidad.get(nombre)))
            celda.alignment = Alignment(vertical="top")
            if nombre == "SAML · Enlace" and isinstance(celda.value, str) and celda.value.startswith("https://"):
                celda.hyperlink = celda.value
                celda.style = "Hyperlink"

    for indice, nombre in enumerate(oa.COLUMNAS_TRAZABILIDAD_SAML):
        columna = COLUMNA_TRAZABILIDAD + indice
        celda = ws.cell(2, columna, nombre)
        celda.font = FUENTE_ENCABEZADO
        celda.fill = RELLENO_SAML
        celda.alignment = ALINEACION_ENCABEZADO
        nota = Comment("Columna de revisión de SAML: NO se pega en el maestro (solo se copian las columnas A a BS).", "SAML")
        nota.width, nota.height = 300, 80
        celda.comment = nota
        ws.column_dimensions[get_column_letter(columna)].width = 34 if "Enlace" not in nombre else 45
    ws.cell(1, COLUMNA_TRAZABILIDAD, "Datos de origen para revisar (no pegar en el maestro)").font = FUENTE_SECCION

    ws.freeze_panes = "A3"
    ws.auto_filter.ref = f"A2:{mae.ULTIMA_COLUMNA}{max(ultima, 2)}"
    _validaciones(ws, maestros, ultima)


def _hoja_instrucciones(
    libro: Workbook,
    maestro: mae.ArchivoMaestro,
    por_hoja: dict[str, int],
    generado: datetime,
    modo: str | None,
    actualizado: str | None = None,
) -> None:
    ws = libro.active
    ws.title = "Instrucciones"
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 110
    if maestro.asesor == oa.ASESOR_SIN_ASIGNAR:
        filas = [
            ("Obras sin asesor asignado (para que GRECO decida)", None),
            (None, None),
            ("Generado el", generado),
            ("Modo", modo or ""),
            (None, None),
            ("Qué hay aquí", "Obras adjudicadas de entidades sin regla de reparto (nacionales, distritales u otras) con sede en Bogotá o Cundinamarca. "
                             "La columna A trae el nombre de la entidad contratante."),
            (None, None),
            ("Cómo asignarlas", None),
            ("1.", "Decida el asesor de cada obra y cámbielo en la columna E (Asesor), con la lista desplegable."),
            ("2.", "Copie la fila, SOLO columnas A a BS (no copie desde la columna BU)."),
            ("3.", "En el archivo de ese asesor, elija la hoja que corresponda, haga clic en la celda A de su ÚLTIMA obra → clic derecho → "
                   "«Insertar celdas copiadas» → «Desplazar las celdas hacia abajo». Los porcentajes de potencial los toma de esa hoja."),
            ("4.", "Si Meses por Terminar está vacío, la obra no suma en la proyección hasta que lo complete."),
            (None, None),
            ("Obras", sum(por_hoja.values())),
            (None, None),
            ("Importante", "Las fechas de Inicio de obra y Meses por Terminar que vienen del cronograma son PROGRAMADAS por la entidad y pueden cambiar."),
        ]
    else:
        filas = []
    filas = filas or [
        (f"Obras nuevas para {maestro.asesor}", None),
        (None, None),
        ("Generado el", generado),
        ("Modo", modo or ""),
        ("Archivo de destino", maestro.ruta.name),
        (None, None),
        *(
            [
                ("Maestro actualizado", actualizado),
                (None, "SAML ya insertó estas obras en una COPIA de su maestro (archivo de arriba, en esta misma carpeta), "
                       "dentro de los rangos del Resumen, Mercado 2027 y Top 15, sin tocar gráficos ni listas. Revise esa copia, "
                       "complete lo que falte y, si está bien, úsela en lugar de su archivo. No necesita pegar nada."),
                (None, None),
                ("Si prefiere hacerlo a mano", None),
            ]
            if actualizado
            else [("Cómo usar este archivo", None)]
        ),
        ("1.", "Revise cada fila (Tipo de contrato, Localidad, Inicio de obra y Meses por Terminar) y complete lo que falte."),
        ("2.", f"En su archivo {maestro.ruta.name}, vaya a la hoja con el MISMO nombre."),
        ("3.", "Aquí, seleccione las filas desde la fila 3, SOLO columnas A a BS (no copie desde la columna BU: son datos de revisión de SAML), y cópielas (Ctrl+C)."),
        ("4.", "En el maestro, haga clic en la celda A de su ÚLTIMA obra → clic derecho → «Insertar celdas copiadas» → «Desplazar las celdas hacia abajo»."),
        (None, "Así las obras nuevas quedan DENTRO de los rangos del Resumen, Mercado 2027 y Top 15, que se amplían solos. "
               "No las pegue debajo de la última obra: quedarían fuera de esos totales. Las fórmulas se ajustan solas."),
        ("5.", "Si Meses por Terminar está vacío, la obra no suma en la proyección hasta que lo complete."),
        (None, None),
        ("Obras por hoja", None),
        *((nombre, cantidad) for nombre, cantidad in por_hoja.items()),
        (None, None),
        ("Importante", "Las fechas de Inicio de obra y Meses por Terminar que vienen del cronograma son PROGRAMADAS por la entidad y pueden cambiar."),
        (None, "SAML no modifica su archivo maestro original."),
    ]
    for etiqueta, valor in filas:
        ws.append([etiqueta, _celda(valor)])
        if etiqueta and valor is None:
            ws.cell(ws.max_row, 1).font = FUENTE_SECCION
        ws.cell(ws.max_row, 2).alignment = Alignment(wrap_text=True, vertical="top")
        if isinstance(valor, datetime):
            ws.cell(ws.max_row, 2).number_format = "dd/mm/yyyy hh:mm"
            ws.cell(ws.max_row, 2).alignment = Alignment(horizontal="left")
    ws["A1"].font = FUENTE_TITULO


def construir_libro_asesor(
    maestro: mae.ArchivoMaestro,
    filas: list[oa.FilaObra],
    maestros: mae.Maestros,
    generado: datetime,
    modo: str | None,
    actualizado: str | None = None,
) -> tuple[Workbook, dict[str, int]]:
    libro = Workbook()
    por_hoja: dict[str, list[oa.FilaObra]] = {}
    for fila in filas:
        por_hoja.setdefault(fila.hoja.nombre, []).append(fila)
    conteo = {nombre: len(lista) for nombre, lista in por_hoja.items()}
    _hoja_instrucciones(libro, maestro, conteo, generado, modo, actualizado)
    for nombre in maestro.hojas:  # mismo orden que en el maestro
        if nombre in por_hoja:
            _hoja_obras(libro, maestro.hojas[nombre], por_hoja[nombre], maestros)
    _hoja_listas(libro, maestros)
    return libro, conteo


# ----------------------------------------------------------------------
# Archivo de revision
# ----------------------------------------------------------------------
def _hoja_resumen(hoja: Worksheet, g: Generacion, entrada: Path, modo: str | None) -> None:
    m = g.resultado.metricas
    for letra, ancho in zip("ABCDE", (44, 40, 18, 30, 80)):
        hoja.column_dimensions[letra].width = ancho

    def titulo(texto: str) -> None:
        hoja.append([])
        hoja.append([texto])
        hoja.cell(row=hoja.max_row, column=1).font = FUENTE_SECCION

    def fila(etiqueta: str, valor: Any, formato: str | None = None) -> None:
        hoja.append([etiqueta, _celda(valor)])
        hoja.cell(row=hoja.max_row, column=1).font = Font(bold=True)
        celda = hoja.cell(row=hoja.max_row, column=2)
        celda.alignment = Alignment(horizontal="left", vertical="top", wrap_text=isinstance(valor, str) and len(valor) > 40)
        if formato:
            celda.number_format = formato

    def encabezado(*columnas: str) -> None:
        hoja.append(list(columnas))
        for indice in range(1, len(columnas) + 1):
            celda = hoja.cell(row=hoja.max_row, column=indice)
            celda.font = FUENTE_ENCABEZADO
            celda.fill = RELLENO_ENCABEZADO
            celda.border = BORDE

    hoja.append(["SAML · Revisión de la ejecución"])
    hoja["A1"].font = FUENTE_TITULO

    titulo("Ejecución")
    fila("Generado el", g.generado, "dd/mm/yyyy hh:mm:ss")
    if modo:
        fila("Modo de ejecución", modo)
    entrada = entrada.resolve()
    fila("Carpeta de entrada", str(entrada.relative_to(RAIZ)) if entrada.is_relative_to(RAIZ) else str(entrada))
    fila("Carpeta de salida", str(g.carpeta.relative_to(RAIZ)) if g.carpeta.is_relative_to(RAIZ) else str(g.carpeta))
    encabezado("CSV consolidado", "Fecha de extracción", "Filas", "SHA-256 (sin modificar)")
    for archivo in g.resultado.archivos:
        hoja.append([archivo.ruta.name, archivo.fecha_extraccion, archivo.filas, archivo.sha256])
        hoja.cell(row=hoja.max_row, column=2).number_format = "dd/mm/yyyy hh:mm:ss"
    encabezado("Maestro (solo lectura)", "Asesor", "Hojas de obras", "SHA-256 (sin modificar)")
    for maestro in g.maestros.archivos:
        hoja.append([maestro.ruta.name, maestro.asesor, ", ".join(maestro.hojas), maestro.sha256])

    titulo("Resultado")
    fila("Filas leídas de los CSV", m["filas_leidas"])
    fila("Procesos únicos por idContrato", m["procesos_unicos"])
    fila("Adjudicados", m["incluidos"])
    fila("Excluidos por estado (no adjudicados)", m["excluidos_por_estado"])
    fila("Ya estaban en los maestros (no se repiten)", len(g.reparto.ya_en_maestros))
    sin_asignar = [f for f in g.reparto.filas if f.archivo is g.reparto.archivo_sin_asignar]
    fila("Filas nuevas para los asesores", len(g.reparto.filas) - len(sin_asignar))
    fila("Sin asignar (GRECO decide el asesor)", len(sin_asignar))
    fila("Sin consumo de material (solo mantenimiento)", len(g.reparto.sin_consumo))
    fila("Fuera de zona (descartadas)", len(g.reparto.fuera_de_zona))
    fila("Sin clasificar (sin plantilla o sin maestro)", len(g.reparto.sin_clasificar))
    encabezado("Archivo generado", "Asesor", "Filas", "Por hoja")
    for archivo in g.archivos_asesores:
        hoja.append([archivo.ruta.name, archivo.maestro.asesor, archivo.filas, ", ".join(f"{h}: {n}" for h, n in archivo.por_hoja.items())])
    usados = {id(x.maestro) for x in g.archivos_asesores}
    sin_filas = [a.asesor for a in g.maestros.archivos if id(a) not in usados]
    if sin_filas:
        fila("Asesores sin obras nuevas en esta ejecución", ", ".join(sin_filas))

    titulo("Maestros actualizados (copias con las obras ya insertadas)")
    act = g.actualizacion
    if act is None:
        fila("Generados", "No (se ejecutó sin actualizar maestros).")
    elif not act.disponible:
        fila("Generados", f"No. {act.motivo} Use los archivos Nuevas_ e inserte las filas a mano.")
    else:
        for a in act.archivos:
            fila(a.ruta.name, ", ".join(f"{h}: {n} filas" for h, n in a.filas_por_hoja.items()))
            for aviso in a.avisos:
                fila("Aviso", aviso)
        for error in act.errores:
            fila("Error", error)

    titulo("Clasificación con LLM (entidades sin regla)")
    llm = g.reparto.llm
    if llm is None:
        fila("Usado", "No. " + (g.reparto.llm_aviso or "Se usaron solo las reglas."))
    else:
        fila("Usado", f"Sí, OpenAI · modelo {llm.modelo}")
        fila("Procesos enviados", llm.enviados)
        fila("Clasificados por el LLM", len(llm.decisiones))
        fila("Clasificados por reglas (respaldo)", llm.enviados - len(llm.decisiones))
        fila("Tokens (entrada / salida)", f"{llm.tokens_entrada:,} / {llm.tokens_salida:,}")
        fila("Datos enviados", "Entidad, objeto, ubicación, valor y ganadores de cada proceso, y la lista de clientes de Sugeine.")
        for error in llm.errores[:20]:
            fila("Aviso", error)

    titulo("Reglas de reparto (entidad contratante → hoja)")
    encabezado("Hoja", "Entidades", "Archivo del asesor")
    reglas = [(h, d) for h, _, d in oa.REGLAS_HOJA] + [
        (oa.HOJA_MUNICIPIOS, "Alcaldías municipales, ICCU, Fondo de Desarrollo de Proyectos de Cundinamarca e IDUVI, con ubicación en Cundinamarca")
    ]
    for nombre, descripcion in reglas:
        maestro = g.maestros.archivo_de_hoja(nombre)
        hoja.append([nombre, descripcion, maestro.ruta.name if maestro else "(ningún maestro tiene esta hoja)"])
    fila("Sugeine", "Si un ganador coincide con un cliente de 'Cobertura clientes', la obra va también a su archivo y en el del otro asesor queda con Asesor = Sugeine. Si solo se parece: '(posible Sugeine)'.")
    fila(
        "Entidades sin regla",
        "En orden: 1) ubicación fuera de Bogotá y Cundinamarca → hoja 'Fuera de zona'; 2) INVIAS / ANI → Municipios Aledaños; "
        "3) Gobernación de Cundinamarca con obra en municipios → Municipios Aledaños; 4) solo mantenimiento, sin adecuación ni obra → "
        "hoja 'Sin consumo de material'; 5) el resto → archivo 'Sin asignar (GRECO)' (o a Sugeine, hoja 'Edificaciones y Mant.', "
        "si el ganador es su cliente).",
    )

    titulo("Qué llena SAML en las 71 columnas")
    encabezado("Columna", "Campo", "Regla")
    for letra, campo, regla in MAPEO_COLUMNAS:
        hoja.append([letra, campo, regla])
        hoja.cell(row=hoja.max_row, column=3).alignment = Alignment(wrap_text=True, vertical="top")

    titulo("Incidencias por tipo")
    encabezado("Tipo", "Cantidad")
    incidencias = g.resultado.incidencias + g.reparto.incidencias
    for tipo, cantidad in Counter(i["Tipo"] for i in incidencias).most_common():
        hoja.append([tipo, cantidad])
    fila("Total", len(incidencias))

    titulo("Campos del CSV que no se exportan")
    fila("Excluidos", ", ".join(lim.CAMPOS_EXCLUIDOS) + ". Son identificadores firmados de la plataforma o estado del usuario.")


def construir_libro_revision(g: Generacion, entrada: Path, modo: str | None) -> Workbook:
    libro = Workbook()
    resumen = libro.active
    resumen.title = "Resumen"
    columnas = [c.nombre for c in lim.COLUMNAS_PROCESOS]
    _escribir_tabla(libro.create_sheet("Sin clasificar"), ["Motivo", *columnas], g.reparto.sin_clasificar, congelar="C2")
    _escribir_tabla(libro.create_sheet(oa.SIN_CONSUMO), ["Motivo", *columnas], g.reparto.sin_consumo, congelar="C2")
    _escribir_tabla(libro.create_sheet(oa.FUERA_DE_ZONA), ["Motivo", *columnas], g.reparto.fuera_de_zona, congelar="C2")
    _escribir_tabla(libro.create_sheet("Ya en maestros"), ["Ya está en", *columnas], g.reparto.ya_en_maestros, congelar="C2")
    procesos = [
        {**d, "Destino": g.reparto.destinos.get(d["idContrato"], ""), "Clasificado por": g.reparto.metodo.get(d["idContrato"], "")}
        for d in g.resultado.datos
    ]
    _escribir_tabla(libro.create_sheet("Procesos"), ["Destino", "Clasificado por", *columnas], procesos, congelar="D2")
    _escribir_tabla(libro.create_sheet("Ganadores"), [c.nombre for c in lim.COLUMNAS_GANADORES], g.resultado.ganadores)
    incidencias = sorted(
        g.resultado.incidencias + g.reparto.incidencias, key=lambda i: (i["Tipo"], i["CodigoProceso"], i["Campo"])
    )
    _escribir_tabla(libro.create_sheet("Incidencias"), list(lim.COLUMNAS_INCIDENCIAS), incidencias)
    _hoja_resumen(resumen, g, entrada, modo)
    return libro


# ----------------------------------------------------------------------
# Punto de entrada
# ----------------------------------------------------------------------
def _guardar(libro: Workbook, destino: Path) -> None:
    with destino.open("xb") as archivo:  # "x": falla si ya existe
        libro.save(archivo)


def generar(
    entrada: Path,
    maestros_dir: Path,
    salida: Path,
    modo: str | None = None,
    clasificador: Any = None,
    aviso_llm: str = "",
    actualizar_maestros: bool = True,
) -> Generacion:
    """Limpia los CSV de `entrada`, los reparte segun los maestros y escribe una carpeta nueva en `salida`.

    Lanza FileNotFoundError si no hay CSV, maestros.ErrorMaestros si faltan los maestros o
    no tienen la plantilla esperada, y FileExistsError si la carpeta de destino ya existe.
    """
    resultado = lim.procesar(entrada)
    maestros = mae.cargar(maestros_dir)
    reparto = oa.repartir(resultado, maestros, clasificador)
    reparto.llm_aviso = aviso_llm
    generado = datetime.now().replace(microsecond=0)
    sello = f"{generado:%Y%m%d_%H%M%S}"
    salida.mkdir(parents=True, exist_ok=True)
    carpeta = salida / f"SAML_{sello}"
    carpeta.mkdir()  # falla si ya existe: nunca se sobrescribe una ejecucion

    # Primero las copias actualizadas de los maestros (con Excel, si esta disponible).
    actualizacion = am.actualizar(reparto, maestros, carpeta, sello) if actualizar_maestros else None
    actualizados = {id(a.maestro): a.ruta.name for a in (actualizacion.archivos if actualizacion else [])}

    archivos: list[ArchivoAsesor] = []
    destinos = list(maestros.archivos) + ([reparto.archivo_sin_asignar] if reparto.archivo_sin_asignar else [])
    for maestro in destinos:
        filas = reparto.filas_de(maestro)
        if not filas:
            continue
        libro, conteo = construir_libro_asesor(maestro, filas, maestros, generado, modo, actualizados.get(id(maestro)))
        destino = carpeta / f"Nuevas_{maestro.ruta.stem}_{sello}.xlsx"
        _guardar(libro, destino)
        archivos.append(ArchivoAsesor(destino, maestro, len(filas), conteo))

    revision = carpeta / f"SAML_revision_{sello}.xlsx"
    g = Generacion(carpeta, archivos, revision, resultado, reparto, maestros, generado, actualizacion)
    _guardar(construir_libro_revision(g, entrada, modo), revision)
    return g


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--entrada", type=Path, default=ENTRADA_POR_DEFECTO, help="carpeta con los CSV de la ingesta")
    parser.add_argument("--maestros", type=Path, default=MAESTROS_POR_DEFECTO, help="carpeta con los 4 Excel maestros")
    parser.add_argument("--salida", type=Path, default=SALIDA_POR_DEFECTO, help="carpeta donde se crea la salida")
    args = parser.parse_args()

    try:
        g = generar(args.entrada, args.maestros, args.salida)
    except (FileNotFoundError, mae.ErrorMaestros) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    except FileExistsError as exc:
        print(f"ERROR: {exc.filename} ya existe; no se sobrescribe. Vuelve a ejecutar en unos segundos.", file=sys.stderr)
        return 1

    m = g.resultado.metricas
    print(f"CSV consolidados: {len(g.resultado.archivos)} | adjudicados: {m['incluidos']}")
    print(
        f"Filas nuevas: {len(g.reparto.filas)} | ya en maestros: {len(g.reparto.ya_en_maestros)} | "
        f"sin consumo: {len(g.reparto.sin_consumo)} | fuera de zona: {len(g.reparto.fuera_de_zona)} | "
        f"sin clasificar: {len(g.reparto.sin_clasificar)}"
    )
    for archivo in g.archivos_asesores:
        print(f"Archivo para {archivo.maestro.asesor}: {archivo.ruta}")
    print(f"Revision: {g.revision}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
