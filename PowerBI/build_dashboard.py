"""
Generador del tablero SAML (Power BI, formato PBIP: TMDL + PBIR).

Genera, a partir de la fuente intermedia de licitaciones:
  * SAML_Dashboard.SemanticModel/definition/  -> modelo semantico en TMDL
  * SAML_Dashboard.Report/                     -> tema, fondos de pagina y 5 paginas con sus visuales

Uso (desde la raiz del repositorio o desde PowerBI/):
    python PowerBI/build_dashboard.py
    python PowerBI/build_dashboard.py --fuente "C:\\ruta\\a\\otro_archivo.xlsx"

Requiere Pillow (para dibujar los fondos). Despues de ejecutarlo, abrir SAML_Dashboard.pbip
en Power BI Desktop y pulsar "Actualizar".
"""
from __future__ import annotations

import argparse
import json
import shutil
import uuid
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter, ImageFont

AQUI = Path(__file__).resolve().parent
RAIZ = AQUI.parent
MODELO = AQUI / "SAML_Dashboard.SemanticModel"
REPORTE = AQUI / "SAML_Dashboard.Report"
FUENTES_TTF = RAIZ / "docs" / "SAML_presentaciones" / "fuentes"
FUENTE_POR_DEFECTO = RAIZ / "fuentes_powerbi" / "licitaciones_intermedio_20261012_172632.xlsx"

# --------------------------------------------------------------------------------------
# Identidad visual (tomada de las presentaciones SAML: carbon, papel calido y ambar)
# --------------------------------------------------------------------------------------
CARBON = "#1E2122"
CARBON_2 = "#2B3032"
CARBON_3 = "#262A2C"
PAPEL = "#F6F3EE"
PAPEL_2 = "#ECE8E1"
ARENA = "#E1DBD0"
AMBAR = "#E8A33D"
AMBAR_OSC = "#9A5E12"
GRIS = "#5F635E"
GRIS_CLARO = "#A9ACA6"
PETROLEO = "#2F5D62"
OXIDO = "#B5562B"
OLIVA = "#8C8F6B"
BLANCO = "#FFFFFF"
PALETA = [AMBAR, PETROLEO, OXIDO, OLIVA, GRIS, "#D9C29C", AMBAR_OSC, "#4F7A8A", "#C9C2B6", CARBON]

W, H = 1280, 720
ESCALA = 2  # los fondos se dibujan al doble para que se vean nitidos

NS = uuid.UUID("5a3d7c1e-0b6f-4c55-9f1e-2a7d4c9b8e01")


def guid(*partes: str) -> str:
    return str(uuid.uuid5(NS, "|".join(partes)))


def hexid(*partes: str) -> str:
    return uuid.uuid5(NS, "|".join(partes)).hex[:20]


# ======================================================================================
# 1. MODELO SEMANTICO (TMDL)
# ======================================================================================

def q(nombre: str) -> str:
    """Nombre TMDL: entre comillas simples si hace falta."""
    seguro = nombre.replace("_", "").isalnum()
    return nombre if seguro else "'" + nombre.replace("'", "''") + "'"


def bloque(texto: str, tabs: int) -> str:
    pre = "\t" * tabs
    return "\n".join(pre + linea if linea.strip() else "" for linea in texto.strip("\n").splitlines())


M_COMUN = """
    Origen = Excel.Workbook(File.Contents(RutaFuente), null, true),
    Texto = (x) => if x = null then "" else Text.From(x),
    Clasificar = (texto as text, reglas as list, porDefecto as text) as text =>
        let
            Coincide = List.Select(reglas, (r) => List.AnyTrue(List.Transform(r{1}, (k) => Text.Contains(texto, k))))
        in
            if List.IsEmpty(Coincide) then porDefecto else Coincide{0}{0},
"""

M_PROCESOS = """
let
""" + M_COMUN + """
    Hoja = Origen{[Item="Datos",Kind="Sheet"]}[Data],
    Encabezados = Table.PromoteHeaders(Hoja, [PromoteAllScalars=true]),
    Seleccion = Table.SelectColumns(Encabezados, {
        "idContrato", "CodigoProceso", "Entidad", "Empresa", "Valor Millones", "Valor (origen, pesos)",
        "Fuente", "Enlace", "Nombre (origen)", "Objeto (origen)", "Fecha publicación (origen)",
        "Fecha vencimiento (origen)", "Ubicación (origen)", "Num. ganadores", "Num. proponentes",
        "Adjudicación programada (cronograma)", "Estado (origen)", "Fecha extracción"}),
    Renombrado = Table.RenameColumns(Seleccion, {
        {"Empresa", "Adjudicatarios"}, {"Valor Millones", "Valor (MM)"}, {"Valor (origen, pesos)", "Valor (pesos)"},
        {"Nombre (origen)", "Nombre del proceso"}, {"Objeto (origen)", "Objeto"},
        {"Fecha publicación (origen)", "Fecha publicación"}, {"Fecha vencimiento (origen)", "Fecha cierre"},
        {"Ubicación (origen)", "Ubicación"}, {"Adjudicación programada (cronograma)", "Adjudicación programada"},
        {"Estado (origen)", "Estado"}}),
    Tipos = Table.TransformColumnTypes(Renombrado, {
        {"idContrato", type text}, {"CodigoProceso", type text}, {"Entidad", type text}, {"Adjudicatarios", type text},
        {"Valor (MM)", type number}, {"Valor (pesos)", type number}, {"Fuente", type text}, {"Enlace", type text},
        {"Nombre del proceso", type text}, {"Objeto", type text}, {"Fecha publicación", type datetime},
        {"Fecha cierre", type datetime}, {"Ubicación", type text}, {"Num. ganadores", Int64.Type},
        {"Num. proponentes", Int64.Type}, {"Adjudicación programada", type text}, {"Estado", type text},
        {"Fecha extracción", type datetime}}),
    EntidadLimpia = Table.TransformColumns(Tipos, {{"Entidad", each Text.Trim(Text.Trim(Text.Trim(Texto(_)), {"*", "+", "/", "-", "(", ")", " "})), type text}}),
    ReglasSector = {
        {"Salud", {"SALUD", "HOSPITAL", "SUBRED"}},
        {"Defensa y seguridad", {"ARMADA", "EJERCITO", "EJÉRCITO", "FUERZA", "MILITAR", "POLICIA", "POLICÍA", "CARABINEROS",
            "AEREO", "AÉREO", "GUERRA", "DEFENSA", "ANTINARCOTICOS", "DITRA", "DMBOG", "ECSAN", "CENAC", "INTELIGENCIA",
            "INTERPOL", "UNP", "CASAS FISCALES", "RECLUTAMIENTO"}},
        {"Justicia, control y registro", {"FISCAL", "PROCURADUR", "JUDICIAL", "DESAJ", "JEP", "USPEC", "PENITENCI",
            "REGISTRADUR", "NOTARIADO", "INVALIDEZ", "MIGRACION", "MIGRACIÓN", "RESTITUCION", "RESTITUCIÓN"}},
        {"Educación y cultura", {"EDUCACI", "UNIVERSIDAD", "ESAP", "SENA ", "CULTURA", "INRAVISI"}},
        {"Movilidad e infraestructura", {"INVIAS", "AEROCIVIL", "CAMINOS", "TRANSMILENIO", "DESARROLLO URBANO",
            "RENOVACION", "RENOVACIÓN", "IDUVI", "PROYECTOS DE CUNDINAMARCA", "FONTUR"}},
        {"Gobierno distrital y local", {"ALCALDIA LOCAL", "ALCALDÍA LOCAL", "DESARROLLO LOCAL", "DISTRITAL", "DISTRITO",
            "BOMBEROS"}},
        {"Gobierno departamental y municipal", {"DEPARTAMENTO DE CUNDINAMARCA", "MUNICIPIO", "MUNICIPAL"}}
    },
    Sector = Table.AddColumn(EntidadLimpia, "Sector", each Clasificar(Text.Upper([Entidad]), ReglasSector, "Otras entidades nacionales"), type text),
    Modalidad = Table.AddColumn(Sector, "Modalidad", each
        let t = Text.Lower(Texto([Nombre del proceso])) in
        if Text.Contains(t, "subasta") then "Subasta inversa"
        else if Text.Contains(t, "menor cuant") then "Selección abreviada"
        else if Text.Contains(t, "obra pública") or Text.Contains(t, "obra publica") then "Licitación de obra pública"
        else if Text.Contains(t, "régimen especial") or Text.Contains(t, "regimen especial") then "Régimen especial"
        else if Text.StartsWith(t, "licitaci") then "Licitación pública"
        else "Otra modalidad", type text),
    ReglasIntervencion = {
        {"Construcción", {"construcci"}},
        {"Mejoramiento y rehabilitación", {"mejoramiento", "rehabilit", "reparaci", "recuperaci", "reforzamiento"}},
        {"Adecuación", {"adecuaci", "remodelaci"}},
        {"Mantenimiento", {"mantenimiento"}},
        {"Suministro e instalación", {"suministro", "instalaci"}},
        {"Estudios y diseños", {"estudios", "diseño", "diseno"}}
    },
    Intervencion = Table.AddColumn(Modalidad, "Tipo de intervención", each Clasificar(Text.Lower(Texto([Objeto])), ReglasIntervencion, "Otras obras"), type text),
    Alcance = Table.AddColumn(Intervencion, "Alcance territorial", each
        let u = Texto([Ubicación]), deps = List.Count(Text.Split(u, ";")) in
        if Text.Contains(Text.Lower(u), "todo el pais") or Text.Contains(Text.Lower(u), "todo el país") or deps >= 5 then "Nacional"
        else if deps > 1 then "Varios departamentos"
        else if u = "Cundinamarca : Bogotá D.C." then "Bogotá D.C."
        else if Text.Contains(u, "Bogotá") then "Bogotá y municipios"
        else if Text.StartsWith(u, "Cundinamarca") then "Municipios de Cundinamarca"
        else "Otros departamentos", type text),
    RangoValor = Table.AddColumn(Alcance, "Rango de valor", each
        let v = [#"Valor (MM)"] in
        if v = null then "Sin valor"
        else if v < 500 then "< 500 MM"
        else if v < 1000 then "500 – 1.000 MM"
        else if v < 5000 then "1.000 – 5.000 MM"
        else if v < 20000 then "5.000 – 20.000 MM"
        else "> 20.000 MM", type text),
    OrdenRango = Table.AddColumn(RangoValor, "Orden rango valor", each
        let v = [#"Valor (MM)"] in
        if v = null then 6 else if v < 500 then 1 else if v < 1000 then 2 else if v < 5000 then 3 else if v < 20000 then 4 else 5, Int64.Type),
    Competencia = Table.AddColumn(OrdenRango, "Nivel de competencia", each
        let p = [#"Num. proponentes"] in
        if p = null or p = 0 then "Sin dato"
        else if p = 1 then "1 proponente"
        else if p <= 5 then "2 – 5"
        else if p <= 15 then "6 – 15"
        else if p <= 50 then "16 – 50"
        else "Más de 50", type text),
    OrdenCompetencia = Table.AddColumn(Competencia, "Orden competencia", each
        let p = [#"Num. proponentes"] in
        if p = null or p = 0 then 6 else if p = 1 then 1 else if p <= 5 then 2 else if p <= 15 then 3 else if p <= 50 then 4 else 5, Int64.Type),
    Resultado = Table.AddColumn(OrdenCompetencia, "Resultado de adjudicación", each
        let g = [#"Num. ganadores"] in
        if g = null or g = 0 then "Sin ganador registrado" else if g = 1 then "Ganador único" else "Varios ganadores", type text),
    FechaDia = Table.AddColumn(Resultado, "Fecha publicación (día)", each if [Fecha publicación] = null then null else DateTime.Date([Fecha publicación]), type date),
    Dias = Table.AddColumn(FechaDia, "Días publicación a cierre", each
        if [Fecha publicación] = null or [Fecha cierre] = null then null
        else Duration.Days(DateTime.Date([Fecha cierre]) - DateTime.Date([Fecha publicación])), Int64.Type),
    ObjetoCorto = Table.AddColumn(Dias, "Objeto (resumen)", each
        let o = Text.Trim(Texto([Objeto])) in if Text.Length(o) > 95 then Text.Start(o, 92) & "…" else o, type text)
in
    ObjetoCorto
"""

M_GANADORES = """
let
""" + M_COMUN + """
    Hoja = Origen{[Item="Ganadores",Kind="Sheet"]}[Data],
    Encabezados = Table.PromoteHeaders(Hoja, [PromoteAllScalars=true]),
    Seleccion = Table.SelectColumns(Encabezados, {"idContrato", "CodigoProceso", "Orden", "Empresa", "Es consorcio (origen)",
        "Está en proponentes", "Repeticiones en origen"}),
    Renombrado = Table.RenameColumns(Seleccion, {{"Es consorcio (origen)", "Es consorcio"}}),
    Tipos = Table.TransformColumnTypes(Renombrado, {{"idContrato", type text}, {"CodigoProceso", type text},
        {"Orden", Int64.Type}, {"Empresa", type text}, {"Es consorcio", type text}, {"Está en proponentes", type text},
        {"Repeticiones en origen", Int64.Type}}),
    Limpia = Table.TransformColumns(Tipos, {{"Empresa", each Text.Trim(Texto(_)), type text}}),
    TipoGanador = Table.AddColumn(Limpia, "Tipo de ganador", each
        if [Es consorcio] = "Sí" then "Consorcio o UT"
        else if [Es consorcio] = "No" then "Empresa individual" else "Sin dato", type text)
in
    TipoGanador
"""

M_INCIDENCIAS = """
let
""" + M_COMUN + """
    Hoja = Origen{[Item="Incidencias",Kind="Sheet"]}[Data],
    Encabezados = Table.PromoteHeaders(Hoja, [PromoteAllScalars=true]),
    Seleccion = Table.SelectColumns(Encabezados, {"Tipo", "Alcance", "idContrato", "CodigoProceso", "Hoja", "Campo", "Detalle"}),
    Tipos = Table.TransformColumnTypes(Seleccion, {{"Tipo", type text}, {"Alcance", type text}, {"idContrato", type text},
        {"CodigoProceso", type text}, {"Hoja", type text}, {"Campo", type text}, {"Detalle", type text}})
in
    Tipos
"""

M_MAPEO = """
let
""" + M_COMUN + """
    Hoja = Origen{[Item="Resumen",Kind="Sheet"]}[Data],
    Inicio = List.PositionOf(List.Transform(Hoja[Column1], Texto), "Mapeo de columnas"),
    Desde = Table.Skip(Hoja, Inicio + 1),
    Encabezados = Table.PromoteHeaders(Desde, [PromoteAllScalars=true]),
    Seleccion = Table.SelectColumns(Encabezados, {"Hoja", "Columna", "Estado", "Origen", "Regla"}),
    Filas = Table.SelectRows(Seleccion, each [Columna] <> null and Text.Trim(Texto([Columna])) <> ""),
    Tipos = Table.TransformColumnTypes(Filas, {{"Hoja", type text}, {"Columna", type text}, {"Estado", type text},
        {"Origen", type text}, {"Regla", type text}}),
    Orden = Table.AddColumn(Tipos, "Orden estado", each
        if [Estado] = "Mapeado" then 1 else if [Estado] = "Trazabilidad" then 2 else if [Estado] = "Provisional" then 3
        else if [Estado] = "Sin fuente" then 4 else if [Estado] = "Manual GRECO" then 5 else 6, Int64.Type),
    Disponible = Table.AddColumn(Orden, "Con dato desde la API", each
        if List.Contains({"Mapeado", "Trazabilidad", "Provisional"}, [Estado]) then "Sí" else "No", type text)
in
    Disponible
"""

DAX_CALENDARIO = """
VAR Inicio = DATE(YEAR(MIN(FactProcesos[Fecha publicación (día)])), 1, 1)
VAR Fin = DATE(YEAR(MAX(FactProcesos[Fecha publicación (día)])), 12, 31)
RETURN
    ADDCOLUMNS(
        CALENDAR(Inicio, Fin),
        "Año", YEAR([Date]),
        "Mes num", MONTH([Date]),
        "Mes", FORMAT([Date], "mmm"),
        "Año mes num", YEAR([Date]) * 100 + MONTH([Date]),
        "Mes año", FORMAT([Date], "mmm yyyy"),
        "Trimestre", "T" & QUARTER([Date]),
        "Semana", WEEKNUM([Date], 2)
    )
"""

# (nombre, tipo, opciones)
COLS_PROCESOS = [
    ("idContrato", "string", {"summarizeBy": "none"}),
    ("CodigoProceso", "string", {"summarizeBy": "none"}),
    ("Entidad", "string", {}),
    ("Adjudicatarios", "string", {}),
    ("Valor (MM)", "double", {"formatString": "#,0.0", "summarizeBy": "sum"}),
    ("Valor (pesos)", "double", {"formatString": "#,0", "summarizeBy": "sum", "isHidden": True}),
    ("Fuente", "string", {}),
    ("Enlace", "string", {"dataCategory": "WebUrl"}),
    ("Nombre del proceso", "string", {}),
    ("Objeto", "string", {}),
    ("Fecha publicación", "dateTime", {"formatString": "dd/mm/yyyy hh:nn"}),
    ("Fecha cierre", "dateTime", {"formatString": "dd/mm/yyyy hh:nn"}),
    ("Ubicación", "string", {}),
    ("Num. ganadores", "int64", {"formatString": "0", "summarizeBy": "sum"}),
    ("Num. proponentes", "int64", {"formatString": "0", "summarizeBy": "sum"}),
    ("Adjudicación programada", "string", {}),
    ("Estado", "string", {}),
    ("Fecha extracción", "dateTime", {"formatString": "dd/mm/yyyy hh:nn"}),
    ("Sector", "string", {"description": "Clasificación SAML de la entidad contratante por palabras clave del nombre."}),
    ("Modalidad", "string", {"description": "Modalidad de selección resumida desde 'Nombre (origen)'."}),
    ("Tipo de intervención", "string", {"description": "Clasificación SAML por palabras clave del objeto del contrato (orientativa)."}),
    ("Alcance territorial", "string", {"description": "Resumen de 'Ubicación (origen)'."}),
    ("Rango de valor", "string", {"sortByColumn": "Orden rango valor"}),
    ("Orden rango valor", "int64", {"isHidden": True, "summarizeBy": "none"}),
    ("Nivel de competencia", "string", {"sortByColumn": "Orden competencia"}),
    ("Orden competencia", "int64", {"isHidden": True, "summarizeBy": "none"}),
    ("Resultado de adjudicación", "string", {}),
    ("Fecha publicación (día)", "dateTime", {"formatString": "dd/mm/yyyy"}),
    ("Días publicación a cierre", "int64", {"formatString": "0", "summarizeBy": "none"}),
    ("Objeto (resumen)", "string", {}),
]
COLS_GANADORES = [
    ("idContrato", "string", {"summarizeBy": "none", "isHidden": True}),
    ("CodigoProceso", "string", {"summarizeBy": "none"}),
    ("Orden", "int64", {"formatString": "0", "summarizeBy": "none"}),
    ("Empresa", "string", {}),
    ("Es consorcio", "string", {}),
    ("Está en proponentes", "string", {}),
    ("Repeticiones en origen", "int64", {"formatString": "0", "summarizeBy": "none"}),
    ("Tipo de ganador", "string", {}),
]
COLS_INCIDENCIAS = [
    ("Tipo", "string", {}),
    ("Alcance", "string", {}),
    ("idContrato", "string", {"summarizeBy": "none", "isHidden": True}),
    ("CodigoProceso", "string", {"summarizeBy": "none"}),
    ("Hoja", "string", {}),
    ("Campo", "string", {}),
    ("Detalle", "string", {}),
]
COLS_MAPEO = [
    ("Hoja", "string", {}),
    ("Columna", "string", {}),
    ("Estado", "string", {"sortByColumn": "Orden estado"}),
    ("Origen", "string", {}),
    ("Regla", "string", {}),
    ("Orden estado", "int64", {"isHidden": True, "summarizeBy": "none"}),
    ("Con dato desde la API", "string", {}),
]
COLS_CALENDARIO = [
    ("Date", "dateTime", {"formatString": "dd/mm/yyyy", "isKey": True}),
    ("Año", "int64", {"formatString": "0", "summarizeBy": "none"}),
    ("Mes num", "int64", {"formatString": "0", "summarizeBy": "none", "isHidden": True}),
    ("Mes", "string", {"sortByColumn": "Mes num"}),
    ("Año mes num", "int64", {"formatString": "0", "summarizeBy": "none", "isHidden": True}),
    ("Mes año", "string", {"sortByColumn": "Año mes num"}),
    ("Trimestre", "string", {}),
    ("Semana", "int64", {"formatString": "0", "summarizeBy": "none"}),
]

SVG_BARRA = (
    "\"data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='100' height='12' viewBox='0 0 100 12'>"
    "<rect x='0' y='3' width='100' height='6' rx='3' fill='%23ECE8E1'/>"
    "<rect x='0' y='3' width='\" & W & \"' height='6' rx='3' fill='%23E8A33D'/></svg>\""
)

# (carpeta, nombre, formato, expresion, extras)
MEDIDAS = [
    # ---------------- Procesos
    ("1. Procesos", "Procesos adjudicados", "#,0", "DISTINCTCOUNT(FactProcesos[idContrato])", {}),
    ("1. Procesos", "Valor adjudicado (MM)", "#,0", "SUM(FactProcesos[Valor (MM)])", {}),
    ("1. Procesos", "Valor adjudicado (miles de millones)", "#,0.0", "DIVIDE([Valor adjudicado (MM)], 1000)", {}),
    ("1. Procesos", "Valor promedio (MM)", "#,0", "AVERAGE(FactProcesos[Valor (MM)])", {}),
    ("1. Procesos", "Valor mediano (MM)", "#,0", "MEDIAN(FactProcesos[Valor (MM)])", {}),
    ("1. Procesos", "Mayor proceso (MM)", "#,0", "MAX(FactProcesos[Valor (MM)])", {}),
    ("1. Procesos", "Entidades contratantes", "#,0", "DISTINCTCOUNT(FactProcesos[Entidad])", {}),
    ("1. Procesos", "% del valor", "0.0%",
     "DIVIDE([Valor adjudicado (MM)], CALCULATE([Valor adjudicado (MM)], ALLSELECTED()))", {}),
    ("1. Procesos", "% top 10 entidades", "0%", """
VAR Top10 = TOPN(10, VALUES(FactProcesos[Entidad]), [Valor adjudicado (MM)], DESC)
RETURN DIVIDE(SUMX(Top10, [Valor adjudicado (MM)]), [Valor adjudicado (MM)])""", {}),
    ("1. Procesos", "Valor (top 8 entidades)", "#,0", """
VAR Posicion = RANKX(ALLSELECTED(FactProcesos[Entidad]), [Valor adjudicado (MM)], , DESC, Dense)
RETURN IF(Posicion <= 8, [Valor adjudicado (MM)])""", {}),
    ("1. Procesos", "Días promedio publicación a cierre", "#,0", "AVERAGE(FactProcesos[Días publicación a cierre])", {}),
    # ---------------- Competencia
    ("2. Competencia", "Proponentes promedio", "#,0.0", "AVERAGE(FactProcesos[Num. proponentes])", {}),
    ("2. Competencia", "Proponentes mediana", "#,0", "MEDIAN(FactProcesos[Num. proponentes])", {}),
    ("2. Competencia", "Procesos con único proponente", "#,0",
     "CALCULATE([Procesos adjudicados], FactProcesos[Num. proponentes] = 1)", {}),
    ("2. Competencia", "% único proponente", "0%",
     "DIVIDE([Procesos con único proponente], [Procesos adjudicados])", {}),
    # ---------------- Ganadores
    ("3. Ganadores", "Ganadores registrados", "#,0", "COUNTROWS(FactGanadores)", {}),
    ("3. Ganadores", "Empresas ganadoras", "#,0", "DISTINCTCOUNT(FactGanadores[Empresa])", {}),
    ("3. Ganadores", "Ganadores consorcio o UT", "#,0",
     "CALCULATE(COUNTROWS(FactGanadores), FactGanadores[Es consorcio] = \"Sí\")", {}),
    ("3. Ganadores", "% consorcios o UT", "0%", "DIVIDE([Ganadores consorcio o UT], [Ganadores registrados])", {}),
    ("3. Ganadores", "Procesos con varios ganadores", "#,0",
     "CALCULATE([Procesos adjudicados], FactProcesos[Num. ganadores] > 1)", {}),
    ("3. Ganadores", "Procesos sin ganador", "#,0",
     "CALCULATE([Procesos adjudicados], FactProcesos[Num. ganadores] = 0)", {}),
    ("3. Ganadores", "Ganadores fuera de proponentes", "#,0",
     "CALCULATE(COUNTROWS(FactGanadores), FactGanadores[Está en proponentes] = \"No\")", {}),
    ("3. Ganadores", "Procesos ganados", "#,0", "DISTINCTCOUNT(FactGanadores[idContrato])", {}),
    ("3. Ganadores", "Empresas con más de un proceso", "#,0", """
COUNTROWS(
    FILTER(VALUES(FactGanadores[Empresa]), CALCULATE(DISTINCTCOUNT(FactGanadores[idContrato])) > 1)
)""", {}),
    ("3. Ganadores", "Valor procesos ganados (MM)", "#,0", "[Valor adjudicado (MM)]",
     {"description": "Valor total de los procesos en los que participa la empresa. No se reparte entre co-ganadores."}),
    ("3. Ganadores", "Barra valor empresa", None, """
VAR V = [Valor procesos ganados (MM)]
VAR Maximo = MAXX(ALLSELECTED(FactGanadores[Empresa]), [Valor procesos ganados (MM)])
VAR W = ROUND(DIVIDE(V, Maximo) * 98, 0) + 2
RETURN IF(ISBLANK(V) || NOT ISINSCOPE(FactGanadores[Empresa]), BLANK(), """ + SVG_BARRA + ")",
     {"dataCategory": "ImageUrl"}),
    ("3. Ganadores", "Barra valor entidad", None, """
VAR V = [Valor adjudicado (MM)]
VAR Maximo = MAXX(ALLSELECTED(FactProcesos[CodigoProceso]), [Valor adjudicado (MM)])
VAR W = ROUND(SQRT(DIVIDE(V, Maximo)) * 98, 0) + 2
RETURN IF(ISBLANK(V) || NOT ISINSCOPE(FactProcesos[CodigoProceso]), BLANK(), """ + SVG_BARRA + ")",
     {"dataCategory": "ImageUrl", "description": "Barra en escala raíz cuadrada para que el proceso atípico no aplaste al resto."}),
    # ---------------- Calidad
    ("4. Calidad", "Total incidencias", "#,0", "COUNTROWS(FactIncidencias)", {}),
    ("4. Calidad", "Incidencias de registro", "#,0",
     "CALCULATE(COUNTROWS(FactIncidencias), FactIncidencias[Alcance] = \"Registro\")", {}),
    ("4. Calidad", "Procesos con incidencias", "#,0",
     "CALCULATE(DISTINCTCOUNT(FactIncidencias[idContrato]), NOT ISBLANK(FactIncidencias[idContrato]))", {}),
    ("4. Calidad", "% procesos sin incidencias", "0%",
     "1 - DIVIDE([Procesos con incidencias], [Procesos adjudicados])", {}),
    ("4. Calidad", "Incidencias por proceso", "0.00", "DIVIDE([Incidencias de registro], [Procesos adjudicados])", {}),
    ("4. Calidad", "Incidencias (top 8 campos)", "#,0", """
VAR Posicion = RANKX(ALLSELECTED(FactIncidencias[Campo]), [Total incidencias], , DESC, Dense)
RETURN IF(Posicion <= 8, [Total incidencias])""", {}),
    ("4. Calidad", "Columnas documentadas", "#,0", "COUNTROWS(DimMapeoColumnas)", {}),
    ("4. Calidad", "Columnas con dato de la API", "#,0",
     "CALCULATE(COUNTROWS(DimMapeoColumnas), DimMapeoColumnas[Con dato desde la API] = \"Sí\")", {}),
    ("4. Calidad", "% cobertura de columnas", "0%", "DIVIDE([Columnas con dato de la API], [Columnas documentadas])", {}),
    # ---------------- Textos (subtitulos dinamicos)
    ("5. Textos", "Texto corte", None,
     "\"Corte \" & FORMAT(MAX(FactProcesos[Fecha extracción]), \"dd/mm/yyyy · hh:nn\")", {}),
    ("5. Textos", "Texto fuente", None,
     "FORMAT(CALCULATE([Procesos adjudicados], ALL(FactProcesos)), \"#,0\") & \" procesos · SECOP I y II\"", {}),
    ("5. Textos", "Sub procesos", None,
     "FORMAT([Entidades contratantes], \"#,0\") & \" entidades contratantes\"", {}),
    ("5. Textos", "Sub valor", None, "\"Top 10 entidades: \" & FORMAT([% top 10 entidades], \"0%\") & \" del valor\"", {}),
    ("5. Textos", "Sub mediana", None, "\"Promedio: \" & FORMAT([Valor promedio (MM)], \"#,0\") & \" MM\"", {}),
    ("5. Textos", "Sub empresas", None, "FORMAT([% consorcios o UT], \"0%\") & \" son consorcios o UT\"", {}),
    ("5. Textos", "Sub proponentes", None,
     "FORMAT([Procesos con único proponente], \"0\") & \" procesos con un solo oferente\"", {}),
    ("5. Textos", "Sub ganados", None,
     "FORMAT([Ganadores registrados], \"#,0\") & \" adjudicaciones registradas\"", {}),
    ("5. Textos", "Sub consorcios", None,
     "FORMAT([Ganadores consorcio o UT], \"#,0\") & \" de \" & FORMAT([Ganadores registrados], \"#,0\") & \" ganadores\"", {}),
    ("5. Textos", "Sub varios ganadores", None,
     "FORMAT(DIVIDE([Procesos con varios ganadores], [Procesos adjudicados]), \"0%\") & \" de los procesos\"", {}),
    ("5. Textos", "Sub recurrentes", None, "\"Ganaron dos o más procesos\"", {}),
    ("5. Textos", "Sub fuera proponentes", None, "\"No figuran en la lista de oferentes\"", {}),
    ("5. Textos", "Sub incidencias", None,
     "FORMAT([Incidencias de registro], \"#,0\") & \" de registro · \" & FORMAT([Total incidencias] - [Incidencias de registro], \"#,0\") & \" de columna\"", {}),
    ("5. Textos", "Sub procesos incidencias", None,
     "FORMAT(DIVIDE([Procesos con incidencias], [Procesos adjudicados]), \"0%\") & \" de los procesos\"", {}),
    ("5. Textos", "Sub sin incidencias", None, "\"Procesos listos sin revisión\"", {}),
    ("5. Textos", "Sub por proceso", None, "\"Incidencias de registro / proceso\"", {}),
    ("5. Textos", "Sub cobertura", None,
     "FORMAT([Columnas con dato de la API], \"0\") & \" de \" & FORMAT([Columnas documentadas], \"0\") & \" columnas con dato\"", {}),
    # ---------------- Ficha del proceso (drillthrough)
    ("6. Ficha", "Ficha · Código", None, "SELECTEDVALUE(FactProcesos[CodigoProceso], \"Selecciona un proceso\")", {}),
    ("6. Ficha", "Ficha · Entidad", None, "SELECTEDVALUE(FactProcesos[Entidad], \"—\")", {}),
    ("6. Ficha", "Ficha · Sector", None, "SELECTEDVALUE(FactProcesos[Sector], \"—\")", {}),
    ("6. Ficha", "Ficha · Modalidad", None, "SELECTEDVALUE(FactProcesos[Nombre del proceso], \"—\")", {}),
    ("6. Ficha", "Ficha · Ubicación", None, "SELECTEDVALUE(FactProcesos[Ubicación], \"—\")", {}),
    ("6. Ficha", "Ficha · Publicación", None,
     "IF(HASONEVALUE(FactProcesos[idContrato]), FORMAT(MAX(FactProcesos[Fecha publicación]), \"dd/mm/yyyy hh:nn\"), \"—\")", {}),
    ("6. Ficha", "Ficha · Cierre", None,
     "IF(HASONEVALUE(FactProcesos[idContrato]), FORMAT(MAX(FactProcesos[Fecha cierre]), \"dd/mm/yyyy hh:nn\"), \"—\")", {}),
    ("6. Ficha", "Ficha · Adjudicación programada", None,
     "SELECTEDVALUE(FactProcesos[Adjudicación programada], \"Sin cronograma\")", {}),
    ("6. Ficha", "Ficha · Objeto", None, "SELECTEDVALUE(FactProcesos[Objeto], \"Haz clic derecho sobre un proceso en Licitaciones → Obtener detalles → Ficha del proceso.\")", {}),
    ("6. Ficha", "Ficha · Enlace", None, "SELECTEDVALUE(FactProcesos[Enlace])", {"dataCategory": "WebUrl"}),
    ("6. Ficha", "Ficha · Días", "0", "MAX(FactProcesos[Días publicación a cierre])", {}),
    ("6. Ficha", "Ficha · Proponentes", "0", "SUM(FactProcesos[Num. proponentes])", {}),
    ("6. Ficha", "Ficha · Ganadores", "0", "SUM(FactProcesos[Num. ganadores])", {}),
]


def tmdl_columna(tabla: str, nombre: str, tipo: str, op: dict, calculada: bool = False) -> str:
    lineas = [f"\tcolumn {q(nombre)}"]
    lineas.append(f"\t\tdataType: {tipo}")
    if calculada:
        lineas.append("\t\tisNameInferred")
    if op.get("isKey"):
        lineas.append("\t\tisKey")
    if "formatString" in op:
        lineas.append(f"\t\tformatString: {op['formatString']}")
    if op.get("isHidden"):
        lineas.append("\t\tisHidden")
    if "dataCategory" in op:
        lineas.append(f"\t\tdataCategory: {op['dataCategory']}")
    lineas.append(f"\t\tlineageTag: {guid('col', tabla, nombre)}")
    if "description" in op:
        pass  # las descripciones se escriben como comentario /// antes de la columna
    lineas.append(f"\t\tsummarizeBy: {op.get('summarizeBy', 'none')}")
    lineas.append(f"\t\tsourceColumn: {'[' + nombre + ']' if calculada else nombre}")
    if "sortByColumn" in op:
        lineas.append(f"\t\tsortByColumn: {q(op['sortByColumn'])}")
    lineas.append("")
    lineas.append("\t\tannotation SummarizationSetBy = Automatic")
    texto = "\n".join(lineas)
    if "description" in op:
        texto = f"\t/// {op['description']}\n" + texto
    return texto


def tmdl_tabla_m(nombre: str, columnas, m: str, extra: str = "") -> str:
    partes = [f"table {q(nombre)}", f"\tlineageTag: {guid('tabla', nombre)}", ""]
    if extra:
        partes += [extra, ""]
    for c in columnas:
        partes += [tmdl_columna(nombre, *c), ""]
    partes += [
        f"\tpartition {q(nombre)} = m",
        "\t\tmode: import",
        "\t\tsource =",
        bloque(m, 4),
        "",
        "\tannotation PBI_ResultType = Table",
        "",
    ]
    return "\n".join(partes)


def tmdl_medidas() -> str:
    partes = ["table Medidas", f"\tlineageTag: {guid('tabla', 'Medidas')}", ""]
    for carpeta, nombre, fmt, expr, extra in MEDIDAS:
        if "description" in extra:
            partes.append(f"\t/// {extra['description']}")
        expr = expr.strip("\n")
        if "\n" in expr:
            partes.append(f"\tmeasure {q(nombre)} =")
            partes.append(bloque(expr, 3))
        else:
            partes.append(f"\tmeasure {q(nombre)} = {expr}")
        if fmt:
            partes.append(f"\t\tformatString: {fmt}")
        partes.append(f"\t\tdisplayFolder: {carpeta}")
        if "dataCategory" in extra:
            partes.append(f"\t\tdataCategory: {extra['dataCategory']}")
        partes.append(f"\t\tlineageTag: {guid('medida', nombre)}")
        partes.append("")
    partes += [
        "\tcolumn Nota",
        "\t\tdataType: string",
        "\t\tisHidden",
        f"\t\tlineageTag: {guid('col', 'Medidas', 'Nota')}",
        "\t\tsummarizeBy: none",
        "\t\tsourceColumn: Nota",
        "",
        "\tpartition Medidas = m",
        "\t\tmode: import",
        "\t\tsource =",
        bloque('let\n    Origen = #table(type table [Nota = text], {})\nin\n    Origen', 4),
        "",
        "\tannotation PBI_ResultType = Table",
        "",
    ]
    return "\n".join(partes)


def tmdl_calendario() -> str:
    partes = ["table DimCalendario", f"\tlineageTag: {guid('tabla', 'DimCalendario')}", "\tdataCategory: Time", ""]
    for c in COLS_CALENDARIO:
        partes += [tmdl_columna("DimCalendario", *c, calculada=True), ""]
    partes += [
        "\tpartition DimCalendario = calculated",
        "\t\tmode: import",
        "\t\tsource =",
        bloque(DAX_CALENDARIO, 4),
        "",
    ]
    return "\n".join(partes)


def escribir_modelo(ruta_fuente: Path) -> None:
    defn = MODELO / "definition"
    if defn.exists():
        shutil.rmtree(defn)
    (defn / "tables").mkdir(parents=True)
    for viejo in ("model.bim",):
        if (MODELO / viejo).exists():
            (MODELO / viejo).unlink()

    (MODELO / "definition.pbism").write_text(json.dumps({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/semanticModel/definitionProperties/1.0.0/schema.json",
        "version": "4.2",
        "settings": {}
    }, indent=2), encoding="utf-8")

    (defn / "database.tmdl").write_text("database\n\tcompatibilityLevel: 1600\n\n", encoding="utf-8")

    tablas = ["Medidas", "FactProcesos", "FactGanadores", "FactIncidencias", "DimMapeoColumnas", "DimCalendario"]
    modelo = [
        "model Model",
        "\tculture: es-CO",
        "\tdefaultPowerBIDataSourceVersion: powerBI_V3",
        "\tdiscourageImplicitMeasures",
        "\tsourceQueryCulture: es-CO",
        "\tdataAccessOptions",
        "\t\tlegacyRedirects",
        "\t\treturnErrorValuesAsNull",
        "",
        "annotation __PBI_TimeIntelligenceEnabled = 0",
        "",
        'annotation PBI_QueryOrder = ["RutaFuente","Medidas","FactProcesos","FactGanadores","FactIncidencias","DimMapeoColumnas"]',
        "",
        "annotation PBIDesktopVersion = 2.153.1206.0",
        "",
    ] + [f"ref table {q(t)}" for t in tablas] + [""]
    (defn / "model.tmdl").write_text("\n".join(modelo), encoding="utf-8")

    ruta = str(ruta_fuente).replace('"', '""')
    (defn / "expressions.tmdl").write_text(
        "/// Ruta del Excel fuente. Cambiarla en Power Query (Transformar datos > Parámetros) cuando llegue la salida final de SAML.\n"
        f'expression RutaFuente = "{ruta}" meta [IsParameterQuery=true, Type="Text", IsParameterQueryRequired=true]\n'
        f"\tlineageTag: {guid('expr', 'RutaFuente')}\n\n"
        "\tannotation PBI_ResultType = Text\n\n", encoding="utf-8")

    t = defn / "tables"
    t.joinpath("Medidas.tmdl").write_text(tmdl_medidas(), encoding="utf-8")
    t.joinpath("FactProcesos.tmdl").write_text(tmdl_tabla_m("FactProcesos", COLS_PROCESOS, M_PROCESOS), encoding="utf-8")
    t.joinpath("FactGanadores.tmdl").write_text(tmdl_tabla_m("FactGanadores", COLS_GANADORES, M_GANADORES), encoding="utf-8")
    t.joinpath("FactIncidencias.tmdl").write_text(tmdl_tabla_m("FactIncidencias", COLS_INCIDENCIAS, M_INCIDENCIAS), encoding="utf-8")
    t.joinpath("DimMapeoColumnas.tmdl").write_text(tmdl_tabla_m("DimMapeoColumnas", COLS_MAPEO, M_MAPEO), encoding="utf-8")
    t.joinpath("DimCalendario.tmdl").write_text(tmdl_calendario(), encoding="utf-8")

    rel = [
        ("FactGanadores", "idContrato", "FactProcesos", "idContrato", "bothDirections"),
        ("FactIncidencias", "idContrato", "FactProcesos", "idContrato", None),
        ("FactProcesos", "Fecha publicación (día)", "DimCalendario", "Date", None),
    ]
    lineas = []
    for ft, fc, tt, tc, cruce in rel:
        lineas.append(f"relationship {guid('rel', ft, fc, tt, tc)}")
        if cruce:
            lineas.append(f"\tcrossFilteringBehavior: {cruce}")
        lineas.append(f"\tfromColumn: {q(ft)}.{q(fc)}")
        lineas.append(f"\ttoColumn: {q(tt)}.{q(tc)}")
        lineas.append("")
    (defn / "relationships.tmdl").write_text("\n".join(lineas), encoding="utf-8")


# ======================================================================================
# 2. FONDOS DE PAGINA (Pillow)
# ======================================================================================

def fuente(nombre: str, tam: float) -> ImageFont.FreeTypeFont:
    return ImageFont.truetype(str(FUENTES_TTF / nombre), int(tam * ESCALA))


F_BOLD = "Archivo-Bold.ttf"
F_REG = "Archivo-Regular.ttf"
F_SERIF_IT = "SourceSerif4-Italic.ttf"


def s(v: float) -> int:
    return int(round(v * ESCALA))


class Lienzo:
    def __init__(self) -> None:
        self.img = Image.new("RGB", (s(W), s(H)), PAPEL)
        self.sombras = Image.new("RGBA", (s(W), s(H)), (0, 0, 0, 0))
        self.paneles: list = []
        self.textos: list = []

    # -- primitivas diferidas: primero sombras, luego paneles, luego textos
    def panel(self, x, y, w, h, relleno=BLANCO, borde=ARENA, radio=10, sombra=True):
        self.paneles.append((x, y, w, h, relleno, borde, radio, sombra))

    def texto(self, x, y, txt, f, tam, color, ancla="la", tracking=0.0):
        self.textos.append((x, y, txt, f, tam, color, ancla, tracking))

    def render(self) -> Image.Image:
        ds = ImageDraw.Draw(self.sombras)
        for x, y, w, h, *_r, radio, sombra in self.paneles:
            if sombra:
                ds.rounded_rectangle([s(x), s(y + 3), s(x + w), s(y + h + 3)], s(radio), fill=(30, 33, 34, 38))
        sombra = self.sombras.filter(ImageFilter.GaussianBlur(s(6)))
        base = self.img.convert("RGBA")
        base.alpha_composite(sombra)
        d = ImageDraw.Draw(base)
        for x, y, w, h, relleno, borde, radio, _ in self.paneles:
            d.rounded_rectangle([s(x), s(y), s(x + w), s(y + h)], s(radio), fill=relleno,
                                outline=borde, width=s(1) if borde else 0)
        for x, y, txt, f, tam, color, ancla, tracking in self.textos:
            fnt = fuente(f, tam)
            if tracking:
                ancho = sum(fnt.getlength(c) for c in txt) + s(tracking) * (len(txt) - 1)
                cx = s(x) - (ancho / 2 if ancla[0] == "m" else 0)
                for c in txt:
                    d.text((cx, s(y)), c, font=fnt, fill=color, anchor="l" + ancla[1])
                    cx += fnt.getlength(c) + s(tracking)
            else:
                d.text((s(x), s(y)), txt, font=fnt, fill=color, anchor=ancla)
        self.extra(d)
        return base.convert("RGB")

    def extra(self, d):  # se sobrescribe para dibujos sobre todo lo demas
        pass


# ---- iconos simples de navegacion
def icono(d: ImageDraw.ImageDraw, tipo: str, x: float, y: float, color: str) -> None:
    if tipo == "panorama":
        for i in range(2):
            for j in range(2):
                d.rounded_rectangle([s(x + i * 8), s(y + j * 8), s(x + i * 8 + 6), s(y + j * 8 + 6)], s(1.5), fill=color)
    elif tipo == "licitaciones":
        d.rounded_rectangle([s(x + 1), s(y), s(x + 13), s(y + 14)], s(2), outline=color, width=s(1.6))
        for k in range(3):
            d.line([s(x + 4), s(y + 4 + k * 3), s(x + 10), s(y + 4 + k * 3)], fill=color, width=s(1.4))
    elif tipo == "ganadores":
        d.rectangle([s(x), s(y + 7), s(x + 4), s(y + 14)], fill=color)
        d.rectangle([s(x + 5), s(y + 2), s(x + 9), s(y + 14)], fill=color)
        d.rectangle([s(x + 10), s(y + 5), s(x + 14), s(y + 14)], fill=color)
    elif tipo == "calidad":
        d.ellipse([s(x), s(y), s(x + 14), s(y + 14)], outline=color, width=s(1.6))
        d.line([s(x + 3.5), s(y + 7.5), s(x + 6), s(y + 10), s(x + 10.5), s(y + 4.5)], fill=color, width=s(1.8))


NAV = [("panorama", "Panorama"), ("licitaciones", "Licitaciones"), ("ganadores", "Ganadores"), ("calidad", "Calidad de datos")]
NAV_Y0, NAV_PASO, NAV_ALTO = 128, 46, 38
SIDEBAR = 208


def dibujar_sidebar(c: Lienzo, activo: str | None) -> None:
    c.panel(0, 0, SIDEBAR, H, relleno=CARBON, borde=None, radio=0, sombra=False)
    # logotipo
    c.panel(22, 24, 38, 38, relleno=AMBAR, borde=None, radio=9, sombra=False)
    c.texto(72, 25, "SAML", F_BOLD, 23, PAPEL)
    c.texto(73, 54, "Monitoreo de licitaciones", F_REG, 9.5, GRIS_CLARO)
    c.panel(20, 90, SIDEBAR - 40, 1, relleno="#33383A", borde=None, radio=0, sombra=False)
    c.texto(24, 106, "NAVEGACIÓN", F_BOLD, 8, "#7C807B", tracking=1.4)
    for i, (clave, etiqueta) in enumerate(NAV):
        y = NAV_Y0 + i * NAV_PASO
        if clave == activo:
            c.panel(12, y, SIDEBAR - 24, NAV_ALTO, relleno=CARBON_2, borde=None, radio=8, sombra=False)
            c.panel(12, y + 9, 4, NAV_ALTO - 18, relleno=AMBAR, borde=None, radio=2, sombra=False)
            c.texto(58, y + NAV_ALTO / 2, etiqueta, F_BOLD, 12.5, PAPEL, ancla="lm")
        else:
            c.texto(58, y + NAV_ALTO / 2, etiqueta, F_REG, 12.5, GRIS_CLARO, ancla="lm")
    # bloque de corte de datos
    c.panel(16, 556, SIDEBAR - 32, 92, relleno=CARBON_3, borde="#33383A", radio=10, sombra=False)
    c.texto(SIDEBAR / 2, 572, "CORTE DE DATOS", F_BOLD, 8, AMBAR, ancla="mm", tracking=1.4)
    c.texto(SIDEBAR / 2, 672, "Fuente: licitaciones.info", F_REG, 8.5, "#7C807B", ancla="mm")
    c.texto(SIDEBAR / 2, 688, "Versión intermedia de validación", F_REG, 8.5, "#7C807B", ancla="mm")

    viejo_extra = c.extra

    def extra(d):
        viejo_extra(d)
        # marca del logotipo: tres barras de distinta altura (licitaciones que crecen)
        for k, alto in enumerate((12, 20, 16)):
            x = 31 + k * 7.5
            d.rounded_rectangle([s(x), s(53 - alto), s(x + 5), s(53)], s(1.5), fill=CARBON)
        for i, (clave, _) in enumerate(NAV):
            y = NAV_Y0 + i * NAV_PASO + NAV_ALTO / 2 - 7
            icono(d, clave, 30, y, AMBAR if clave == activo else "#7C807B")

    c.extra = extra


def dibujar_encabezado(c: Lienzo, titulo: str, subtitulo: str) -> None:
    c.texto(224, 22, titulo, F_BOLD, 25, CARBON)
    c.texto(225, 58, subtitulo, F_SERIF_IT, 12.5, GRIS)


# ======================================================================================
# 3. REPORTE (PBIR)
# ======================================================================================

def L(v):
    if isinstance(v, bool):
        val = "true" if v else "false"
    elif isinstance(v, (int, float)):
        val = f"{v}D"
    else:
        val = "'" + str(v).replace("'", "''") + "'"
    return {"expr": {"Literal": {"Value": val}}}


def conv(k, v):
    if isinstance(v, dict):
        return v
    if isinstance(v, str) and v.startswith("#") and len(v) in (4, 7):
        return {"solid": {"color": L(v)}}
    return L(v)


def obj(*entradas):
    """obj({'prop': valor}, ...) o obj(({'prop': valor}, {'id': 'hover'}), ...)"""
    res = []
    for e in entradas:
        if isinstance(e, tuple):
            props, selector = e
            res.append({"properties": {k: conv(k, v) for k, v in props.items()}, "selector": selector})
        else:
            res.append({"properties": {k: conv(k, v) for k, v in e.items()}})
    return res


def C(tabla, col):
    return {"Column": {"Expression": {"SourceRef": {"Entity": tabla}}, "Property": col}}


def Me(nombre, tabla="Medidas"):
    return {"Measure": {"Expression": {"SourceRef": {"Entity": tabla}}, "Property": nombre}}


def proj(campo, nombre_visible=None):
    clave = "Column" if "Column" in campo else "Measure"
    t = campo[clave]["Expression"]["SourceRef"]["Entity"]
    p = campo[clave]["Property"]
    r = {"field": campo, "queryRef": f"{t}.{p}", "nativeQueryRef": p}
    if nombre_visible:
        r["displayName"] = nombre_visible
    return r


class Pagina:
    def __init__(self, clave, nombre, titulo, subtitulo, nav_activo, fondo):
        self.clave = clave
        self.id = hexid("pagina", clave)
        self.nombre = nombre
        self.titulo = titulo
        self.subtitulo = subtitulo
        self.visuales: list = []
        self.lienzo = Lienzo()
        self.fondo = fondo
        dibujar_sidebar(self.lienzo, nav_activo)
        dibujar_encabezado(self.lienzo, titulo, subtitulo)
        self.extra_json: dict = {}
        self.z = 1000

    # ---------------------------------------------------------------- utilidades
    def visual(self, x, y, w, h, tipo, query=None, objects=None, vco=None, extra=None):
        self.z += 1
        nombre = hexid("visual", self.clave, str(len(self.visuales)), tipo)
        v = {"visualType": tipo}
        if query:
            v["query"] = query
        if objects:
            v["objects"] = objects
        contenedor = {
            "title": obj({"show": False}),
            "background": obj({"show": False}),
            "border": obj({"show": False}),
            "dropShadow": obj({"show": False}),
            "padding": obj({"top": 0, "bottom": 0, "left": 0, "right": 0}),
        }
        if vco:
            contenedor.update(vco)
        v["visualContainerObjects"] = contenedor
        if extra:
            v.update(extra)
        self.visuales.append({
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/visualContainer/2.8.0/schema.json",
            "name": nombre,
            "position": {"x": x, "y": y, "z": self.z, "width": w, "height": h, "tabOrder": self.z},
            "visual": v,
        })
        return v

    def panel(self, x, y, w, h, titulo, subtitulo=None):
        """Dibuja la tarjeta en el fondo y devuelve el area util (x, y, w, h) para el visual."""
        c = self.lienzo
        c.panel(x, y, w, h)
        c.panel(x + 16, y + 17, 3, 12, relleno=AMBAR, borde=None, radio=1.5, sombra=False)
        c.texto(x + 26, y + 23, titulo, F_BOLD, 12.5, CARBON, ancla="lm")
        top = 40
        if subtitulo:
            c.texto(x + 26, y + 41, subtitulo, F_SERIF_IT, 10, GRIS, ancla="lm")
            top = 54
        return (x + 10, y + top, w - 20, h - top - 8)

    def kpi(self, x, y, w, h, etiqueta, medida, sub=None, heroe=False):
        c = self.lienzo
        if heroe:
            c.panel(x, y, w, h, relleno=CARBON, borde=None)
            c.texto(x + w / 2, y + 20, etiqueta.upper(), F_BOLD, 8.2, AMBAR, ancla="mm", tracking=1.1)
        else:
            c.panel(x, y, w, h)
            c.panel(x + w / 2 - 14, y + 1, 28, 3, relleno=AMBAR, borde=None, radio=1.5, sombra=False)
            c.texto(x + w / 2, y + 20, etiqueta.upper(), F_BOLD, 8.2, GRIS, ancla="mm", tracking=1.1)
        self.tarjeta(x + 6, y + 30, w - 12, 44, Me(medida), 25, PAPEL if heroe else CARBON, "Segoe UI Semibold")
        if sub:
            self.tarjeta(x + 6, y + 72, w - 12, 24, Me(sub), 9, GRIS_CLARO if heroe else GRIS, "Segoe UI")

    def tarjeta(self, x, y, w, h, campo, tam, color, familia, wrap=False):
        objects = {
            "labels": obj({"color": color, "fontSize": tam, "fontFamily": familia, "labelDisplayUnits": 1}),
            "categoryLabels": obj({"show": False}),
        }
        if wrap:
            objects["wordWrap"] = obj({"show": True})
        return self.visual(x, y, w, h, "card", {"queryState": {"Values": {"projections": [proj(campo)]}}}, objects)

    def barras(self, area, categoria, valor, color=AMBAR, orden_valor=True, horizontal=True, etiquetas=True,
               sort_cat=None, tooltips=None, formato_unidades=1):
        x, y, w, h = area
        qs = {"Category": {"projections": [proj(categoria)]}, "Y": {"projections": [proj(valor)]}}
        if tooltips:
            qs["Tooltips"] = {"projections": [proj(t) for t in tooltips]}
        query = {"queryState": qs}
        if orden_valor:
            query["sortDefinition"] = {"sort": [{"field": valor, "direction": "Descending"}], "isDefaultSort": True}
        elif sort_cat:
            query["sortDefinition"] = {"sort": [{"field": sort_cat, "direction": "Ascending"}], "isDefaultSort": True}
        objects = {
            "dataPoint": obj({"fill": color}),
            "labels": obj({"show": etiquetas, "color": CARBON, "fontSize": 8.5, "labelDisplayUnits": formato_unidades,
                           "labelPosition": "OutsideEnd" if horizontal else "OutsideEnd"}),
            "categoryAxis": obj({"show": True, "labelColor": GRIS, "fontSize": 8, "showAxisTitle": False,
                                 "innerPadding": 14 if horizontal else 28, "maxMarginFactor": 45 if horizontal else 25}),
            "valueAxis": obj({"show": not horizontal, "labelColor": GRIS_CLARO, "fontSize": 8, "showAxisTitle": False,
                              "gridlineShow": not horizontal, "gridlineColor": PAPEL_2, "gridlineStyle": "dotted",
                              "labelDisplayUnits": formato_unidades}),
            "legend": obj({"show": False}),
        }
        return self.visual(x, y, w, h, "clusteredBarChart" if horizontal else "clusteredColumnChart", query, objects)

    def dona(self, area, categoria, valor, leyenda=False):
        x, y, w, h = area
        query = {"queryState": {"Category": {"projections": [proj(categoria)]}, "Y": {"projections": [proj(valor)]}},
                 "sortDefinition": {"sort": [{"field": valor, "direction": "Descending"}], "isDefaultSort": True}}
        objects = {
            "legend": obj({"show": leyenda, "position": "BottomCenter", "labelColor": CARBON, "fontSize": 8.5, "showTitle": False}),
            "labels": obj({"show": True, "labelStyle": "Percent of total" if leyenda else "Category, percent of total",
                           "color": CARBON, "fontSize": 8,
                           "labelPosition": "preferOutside"}),
            "slices": obj({"innerRadiusRatio": 62}),
        }
        return self.visual(x, y, w, h, "donutChart", query, objects)

    def tabla(self, area, campos, orden=None, desc=True, img_alto=12, tam=9, anchos=None, ajustar=False):
        x, y, w, h = area
        query = {"queryState": {"Values": {"projections": [proj(c, n) for c, n in campos]}}}
        if orden is not None:
            query["sortDefinition"] = {"sort": [{"field": orden, "direction": "Descending" if desc else "Ascending"}],
                                       "isDefaultSort": True}
        objects = {
            "grid": obj({"gridVertical": False, "gridHorizontal": True, "gridHorizontalColor": PAPEL_2,
                         "rowPadding": 4, "outlineColor": BLANCO, "textSize": tam, "imageHeight": img_alto,
                         "imageWidth": 100}),
            "columnHeaders": obj({"fontColor": GRIS, "backColor": BLANCO, "fontFamily": "Segoe UI Semibold",
                                  "fontSize": tam, "outline": "Bottom only", "wordWrap": True}),
            "values": obj({"fontColorPrimary": CARBON, "backColorPrimary": BLANCO, "fontColorSecondary": CARBON,
                           "backColorSecondary": "#FBFAF7", "fontSize": tam, "urlIcon": True, "wordWrap": ajustar}),
            "total": obj({"totals": False}),
        }
        if anchos:
            objects["columnWidth"] = []
            for (campo, _), ancho in zip(campos, anchos):
                ref = proj(campo)["queryRef"]
                objects["columnWidth"].append({"properties": {"value": L(ancho)}, "selector": {"metadata": ref}})
        return self.visual(x, y, w, h, "tableEx", query, objects)

    def slicer(self, x, y, w, h, campo, titulo):
        query = {"queryState": {"Values": {"projections": [proj(campo, titulo)]}}}
        objects = {
            "data": obj({"mode": "Dropdown"}),
            "header": obj({"show": True, "fontColor": GRIS, "fontSize": 8.5, "fontFamily": "Segoe UI Semibold",
                           "outline": "None", "text": titulo}),
            "items": obj({"fontColor": CARBON, "fontSize": 9, "background": BLANCO, "outline": "None"}),
            "selection": obj({"selectAllCheckboxEnabled": True, "singleSelect": False}),
        }
        vco = {"background": obj({"show": True, "color": BLANCO, "transparency": 0}),
               "border": obj({"show": True, "color": ARENA, "radius": 8}),
               "padding": obj({"top": 4, "bottom": 2, "left": 8, "right": 6})}
        return self.visual(x, y, w, h, "slicer", query, objects, vco,
                           extra={"syncGroup": {"groupName": titulo, "fieldChanges": True, "filterChanges": True}})

    def boton_nav(self, destino_id, i):
        y = NAV_Y0 + i * NAV_PASO
        objects = {
            "icon": obj(({"shapeType": "blank"}, {"id": "default"})),
            "text": obj({"show": False}),
            "outline": obj({"show": False}),
            "fill": obj(({"show": True, "fillColor": BLANCO, "transparency": 100}, {"id": "default"}),
                        ({"show": True, "fillColor": BLANCO, "transparency": 92}, {"id": "hover"}),
                        ({"show": True, "fillColor": BLANCO, "transparency": 88}, {"id": "selected"})),
            "shape": obj({"tileShape": "rectangleRounded"}),
        }
        vco = {"visualLink": obj({"show": True, "type": "PageNavigation", "navigationSection": destino_id}),
               "visualTooltip": obj({"show": False})}
        self.visual(12, y, SIDEBAR - 24, NAV_ALTO, "actionButton", None, objects, vco)

    # ---------------------------------------------------------------- salida
    def json_pagina(self):
        p = {
            "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/page/2.1.0/schema.json",
            "name": self.id,
            "displayName": self.nombre,
            "displayOption": "FitToPage",
            "height": H,
            "width": W,
            "objects": {
                "background": obj({"image": {"image": {
                    "name": L(self.fondo),
                    "url": {"expr": {"ResourcePackageItem": {"PackageName": "RegisteredResources", "PackageType": 1,
                                                             "ItemName": self.fondo}}},
                    "scaling": L("Fit")}}, "transparency": 0}),
                "outspace": obj({"color": PAPEL_2, "transparency": 0}),
            },
        }
        p.update(self.extra_json)
        return p


def construir_paginas() -> list[Pagina]:
    paginas: list[Pagina] = []
    X0, X1 = 224, 1264
    ANCHO = X1 - X0

    def fila_kpis(pag: Pagina, y, specs, alto=104):
        n = len(specs)
        gap = 12
        w = (ANCHO - gap * (n - 1)) / n
        for i, sp in enumerate(specs):
            pag.kpi(round(X0 + i * (w + gap)), y, round(w), alto, *sp)

    # ------------------------------------------------------------ 1. PANORAMA
    p = Pagina("panorama", "Panorama", "Panorama de licitaciones adjudicadas",
               "Obras públicas adjudicadas en SECOP · Bogotá y Cundinamarca · cifras en millones de pesos (MM)",
               "panorama", "fondo_panorama.png")
    fila_kpis(p, 90, [
        ("Procesos adjudicados", "Procesos adjudicados", "Sub procesos", True),
        ("Valor adjudicado · MM", "Valor adjudicado (MM)", "Sub valor"),
        ("Valor mediano · MM", "Valor mediano (MM)", "Sub mediana"),
        ("Empresas ganadoras", "Empresas ganadoras", "Sub empresas"),
        ("Proponentes por proceso", "Proponentes promedio", "Sub proponentes"),
    ])
    # combo mensual
    a = p.panel(224, 206, 620, 240, "Ritmo de publicación", "Procesos (barras) y valor adjudicado en MM (línea) por mes de publicación")
    cat = C("DimCalendario", "Mes año")
    p.visual(*a, "lineClusteredColumnComboChart", {
        "queryState": {"Category": {"projections": [proj(cat)]},
                       "Y": {"projections": [proj(Me("Procesos adjudicados"))]},
                       "Y2": {"projections": [proj(Me("Valor adjudicado (MM)"))]}},
        "sortDefinition": {"sort": [{"field": C("DimCalendario", "Año mes num"), "direction": "Ascending"}], "isDefaultSort": True},
    }, {
        "dataPoint": obj(({"fill": AMBAR}, {"metadata": "Medidas.Procesos adjudicados"}),
                         ({"fill": PETROLEO}, {"metadata": "Medidas.Valor adjudicado (MM)"})),
        "lineStyles": obj({"strokeWidth": 3, "showMarker": True, "markerSize": 5, "lineStyle": "solid"}),
        "labels": obj({"show": True, "color": CARBON, "fontSize": 8.5, "labelDisplayUnits": 1},
                      ({"show": False}, {"metadata": "Medidas.Valor adjudicado (MM)"})),
        "categoryAxis": obj({"labelColor": GRIS, "fontSize": 9, "showAxisTitle": False}),
        "valueAxis": obj({"show": True, "labelColor": GRIS_CLARO, "fontSize": 8, "showAxisTitle": False,
                          "gridlineColor": PAPEL_2, "gridlineStyle": "dotted", "secShow": True,
                          "secLabelColor": GRIS_CLARO, "secFontSize": 8, "secShowAxisTitle": False, "secLabelDisplayUnits": 1}),
        "legend": obj({"show": True, "position": "TopRight", "labelColor": GRIS, "fontSize": 8.5}),
    })
    a = p.panel(856, 206, 408, 240, "¿Quién contrata?", "Valor adjudicado por sector de la entidad (clasificación SAML)")
    p.dona(a, C("FactProcesos", "Sector"), Me("Valor adjudicado (MM)"))
    a = p.panel(224, 458, 404, 246, "Top 8 entidades por valor", "Valor adjudicado en MM")
    p.barras(a, C("FactProcesos", "Entidad"), Me("Valor (top 8 entidades)"), tooltips=[Me("Procesos adjudicados")])
    a = p.panel(640, 458, 300, 246, "¿Qué se contrata?", "Procesos por tipo de intervención (según el objeto)")
    p.barras(a, C("FactProcesos", "Tipo de intervención"), Me("Procesos adjudicados"), color=PETROLEO,
             tooltips=[Me("Valor adjudicado (MM)")])
    a = p.panel(952, 458, 312, 246, "¿Cuánto vale un proceso?", "Procesos por rango de valor")
    p.barras(a, C("FactProcesos", "Rango de valor"), Me("Procesos adjudicados"), color=AMBAR, orden_valor=False,
             horizontal=False, sort_cat=C("FactProcesos", "Rango de valor"), tooltips=[Me("Valor adjudicado (MM)")])
    paginas.append(p)

    # ------------------------------------------------------------ 2. LICITACIONES
    p = Pagina("licitaciones", "Licitaciones", "Licitaciones",
               "Explora los procesos: filtra, compara y entra al detalle",
               "licitaciones", "fondo_licitaciones.png")
    sx = 548
    for i, (campo, titulo) in enumerate([(C("FactProcesos", "Sector"), "Sector"),
                                         (C("FactProcesos", "Modalidad"), "Modalidad"),
                                         (C("FactProcesos", "Tipo de intervención"), "Intervención"),
                                         (C("FactProcesos", "Alcance territorial"), "Alcance")]):
        p.slicer(sx + i * 180, 20, 168, 54, campo, titulo)
    a = p.panel(224, 90, 330, 270, "Modalidad de contratación", "Valor adjudicado (MM) por modalidad")
    p.visual(*a, "treemap", {
        "queryState": {"Group": {"projections": [proj(C("FactProcesos", "Modalidad"))]},
                       "Values": {"projections": [proj(Me("Valor adjudicado (MM)"))]}},
        "sortDefinition": {"sort": [{"field": Me("Valor adjudicado (MM)"), "direction": "Descending"}], "isDefaultSort": True},
    }, {
        "labels": obj({"show": True, "color": BLANCO, "fontSize": 9, "labelDisplayUnits": 1}),
        "categoryLabels": obj({"show": True, "color": BLANCO, "fontSize": 9.5}),
        "legend": obj({"show": False}),
    })
    a = p.panel(566, 90, 330, 270, "Nivel de competencia", "Procesos según el número de proponentes")
    p.barras(a, C("FactProcesos", "Nivel de competencia"), Me("Procesos adjudicados"), color=PETROLEO,
             orden_valor=False, horizontal=False, sort_cat=C("FactProcesos", "Nivel de competencia"),
             tooltips=[Me("Valor adjudicado (MM)")])
    a = p.panel(908, 90, 356, 270, "¿Más valor, más competencia?", "Proponentes promedio según el rango de valor")
    p.barras(a, C("FactProcesos", "Rango de valor"), Me("Proponentes promedio"), color=AMBAR, orden_valor=False,
             horizontal=False, sort_cat=C("FactProcesos", "Rango de valor"),
             tooltips=[Me("Procesos adjudicados"), Me("Proponentes mediana")])
    a = p.panel(224, 372, 1040, 332, "Detalle de procesos",
                "Clic derecho sobre una fila → Obtener detalles → Ficha del proceso · el ícono abre la ficha en licitaciones.info")
    p.tabla(a, [
        (C("FactProcesos", "CodigoProceso"), "Código"),
        (C("FactProcesos", "Entidad"), "Entidad"),
        (C("FactProcesos", "Objeto (resumen)"), "Objeto"),
        (C("FactProcesos", "Modalidad"), "Modalidad"),
        (C("FactProcesos", "Fecha publicación (día)"), "Publicado"),
        (C("FactProcesos", "Num. proponentes"), "Oferentes"),
        (Me("Valor adjudicado (MM)"), "Valor (MM)"),
        (Me("Barra valor entidad"), " "),
        (C("FactProcesos", "Enlace"), "Ficha"),
    ], orden=Me("Valor adjudicado (MM)"), anchos=[118, 172, 220, 128, 76, 70, 70, 104, 44])
    paginas.append(p)

    # ------------------------------------------------------------ 3. GANADORES
    p = Pagina("ganadores", "Ganadores", "Ganadores y mercado",
               "Quién gana, cómo se asocia y cuánta competencia enfrenta",
               "ganadores", "fondo_ganadores.png")
    p.slicer(1096, 20, 168, 54, C("FactProcesos", "Sector"), "Sector")
    p.slicer(916, 20, 168, 54, C("FactGanadores", "Tipo de ganador"), "Ganador")
    fila_kpis(p, 90, [
        ("Empresas ganadoras", "Empresas ganadoras", "Sub ganados", True),
        ("Consorcios y UT", "% consorcios o UT", "Sub consorcios"),
        ("Procesos con varios ganadores", "Procesos con varios ganadores", "Sub varios ganadores"),
        ("Empresas recurrentes", "Empresas con más de un proceso", "Sub recurrentes"),
        ("Ganadores no oferentes", "Ganadores fuera de proponentes", "Sub fuera proponentes"),
    ])
    a = p.panel(224, 206, 540, 498, "Ranking de empresas",
                "Valor total de los procesos ganados (MM) · el valor no se reparte entre co-ganadores")
    p.tabla(a, [
        (C("FactGanadores", "Empresa"), "Empresa"),
        (C("FactGanadores", "Tipo de ganador"), "Tipo"),
        (Me("Procesos ganados"), "Procesos"),
        (Me("Valor procesos ganados (MM)"), "Valor (MM)"),
        (Me("Barra valor empresa"), " "),
    ], orden=Me("Valor procesos ganados (MM)"), anchos=[180, 98, 62, 66, 104])
    a = p.panel(776, 206, 240, 240, "¿Solos o asociados?", "Ganadores por tipo")
    p.dona(a, C("FactGanadores", "Tipo de ganador"), Me("Ganadores registrados"))
    a = p.panel(1028, 206, 236, 240, "Ganadores por proceso", "Procesos según resultado")
    p.barras(a, C("FactProcesos", "Resultado de adjudicación"), Me("Procesos adjudicados"), color=OXIDO,
             horizontal=False)
    a = p.panel(776, 458, 488, 246, "Competencia por sector", "Proponentes promedio por proceso")
    p.barras(a, C("FactProcesos", "Sector"), Me("Proponentes promedio"), color=PETROLEO,
             tooltips=[Me("Procesos adjudicados"), Me("Proponentes mediana")])
    paginas.append(p)

    # ------------------------------------------------------------ 4. CALIDAD
    p = Pagina("calidad", "Calidad de datos", "Calidad de datos",
               "Qué tan completa y confiable es la fuente antes de llegar a los asesores de GRECO",
               "calidad", "fondo_calidad.png")
    fila_kpis(p, 90, [
        ("Incidencias registradas", "Total incidencias", "Sub incidencias", True),
        ("Procesos con incidencias", "Procesos con incidencias", "Sub procesos incidencias"),
        ("Procesos sin incidencias", "% procesos sin incidencias", "Sub sin incidencias"),
        ("Incidencias por proceso", "Incidencias por proceso", "Sub por proceso"),
        ("Cobertura de columnas", "% cobertura de columnas", "Sub cobertura"),
    ])
    a = p.panel(224, 206, 330, 240, "Incidencias por tipo", "Observaciones generadas por la transformación")
    p.barras(a, C("FactIncidencias", "Tipo"), Me("Total incidencias"), color=OXIDO)
    a = p.panel(566, 206, 330, 240, "Campos más observados", "Top 8 campos con incidencias")
    p.barras(a, C("FactIncidencias", "Campo"), Me("Incidencias (top 8 campos)"), color=AMBAR)
    a = p.panel(908, 206, 356, 240, "Mapeo de columnas", "Estado de las columnas del Excel de GRECO")
    p.dona(a, C("DimMapeoColumnas", "Estado"), Me("Columnas documentadas"))
    a = p.panel(224, 458, 1040, 246, "Bitácora de incidencias", "Detalle para revisión del equipo")
    p.tabla(a, [
        (C("FactIncidencias", "Tipo"), "Tipo"),
        (C("FactIncidencias", "Alcance"), "Alcance"),
        (C("FactIncidencias", "Hoja"), "Hoja"),
        (C("FactIncidencias", "Campo"), "Campo"),
        (C("FactIncidencias", "CodigoProceso"), "Proceso"),
        (C("FactIncidencias", "Detalle"), "Detalle"),
    ], orden=C("FactIncidencias", "Tipo"), desc=False, anchos=[120, 66, 70, 110, 150, 480], ajustar=True)
    paginas.append(p)

    # ------------------------------------------------------------ 5. FICHA (drillthrough)
    p = Pagina("ficha", "Ficha del proceso", "Ficha del proceso",
               "Detalle de un proceso adjudicado · se abre desde Licitaciones con clic derecho → Obtener detalles",
               None, "fondo_ficha.png")
    filtro = hexid("filtro", "ficha")
    campo_dt = C("FactProcesos", "CodigoProceso")
    p.extra_json = {
        "pageBinding": {"name": hexid("binding", "ficha"), "type": "Drillthrough", "referenceScope": "Default",
                        "parameters": [{"name": hexid("param", "ficha"), "boundFilter": filtro, "fieldExpr": campo_dt}]},
        "filterConfig": {"filters": [{"name": filtro, "field": campo_dt, "type": "Categorical",
                                      "howCreated": "Drillthrough"}]},
        "visibility": "HiddenInViewMode",
    }
    # boton volver
    p.visual(1124, 26, 140, 40, "actionButton", None, {
        "icon": obj(({"shapeType": "back", "lineColor": CARBON, "lineWeight": 2, "padding": 6,
                      "horizontalAlignment": "left"}, {"id": "default"})),
        "text": obj(({"show": True, "text": "Volver", "fontColor": CARBON, "fontSize": 11,
                      "fontFamily": "Segoe UI Semibold"}, {"id": "default"})),
        "outline": obj(({"show": True, "lineColor": ARENA, "weight": 1, "roundEdge": 8}, {"id": "default"})),
        "fill": obj(({"show": True, "fillColor": BLANCO, "transparency": 0}, {"id": "default"}),
                    ({"show": True, "fillColor": PAPEL_2, "transparency": 0}, {"id": "hover"})),
    }, {"visualLink": obj({"show": True, "type": "Back"})})
    # codigo del proceso (destacado)
    p.lienzo.panel(224, 90, 1040, 64, relleno=CARBON, borde=None)
    p.lienzo.texto(244, 122, "PROCESO", F_BOLD, 8.2, AMBAR, ancla="lm", tracking=1.2)
    p.tarjeta(320, 98, 820, 48, Me("Ficha · Código"), 20, PAPEL, "Segoe UI Semibold")
    a = p.panel(224, 166, 620, 236, "Identificación")
    p.visual(*a, "multiRowCard", {"queryState": {"Values": {"projections": [
        proj(Me("Ficha · Entidad"), "Entidad contratante"),
        proj(Me("Ficha · Sector"), "Sector"),
        proj(Me("Ficha · Modalidad"), "Modalidad y código"),
        proj(Me("Ficha · Ubicación"), "Ubicación"),
        proj(Me("Ficha · Publicación"), "Publicación"),
        proj(Me("Ficha · Cierre"), "Cierre de ofertas"),
        proj(Me("Ficha · Adjudicación programada"), "Adjudicación programada"),
        proj(Me("Ficha · Enlace"), "Ficha en licitaciones.info"),
    ]}}}, {
        "dataLabels": obj({"color": CARBON, "fontSize": 10, "fontFamily": "Segoe UI Semibold"}),
        "categoryLabels": obj({"show": True, "color": GRIS, "fontSize": 8}),
        "card": obj({"barShow": False, "outlineStyle": "None", "cardPadding": 6, "cardBackground": BLANCO}),
    })
    tiles = [("Valor adjudicado · MM", "Valor adjudicado (MM)"), ("Proponentes", "Ficha · Proponentes"),
             ("Ganadores", "Ficha · Ganadores"), ("Días de publicación", "Ficha · Días")]
    for i, (et, me) in enumerate(tiles):
        x = 856 + (i % 2) * 210
        y = 166 + (i // 2) * 124
        p.kpi(x, y, 198, 112, et, me, None, heroe=(i == 0))
    a = p.panel(224, 414, 1040, 92, "Objeto del contrato")
    p.tarjeta(a[0] + 6, a[1] - 4, a[2] - 12, a[3] + 6, Me("Ficha · Objeto"), 10.5, CARBON, "Segoe UI", wrap=True)
    a = p.panel(224, 518, 620, 186, "Ganadores del proceso")
    p.tabla(a, [
        (C("FactGanadores", "Orden"), "#"),
        (C("FactGanadores", "Empresa"), "Empresa"),
        (C("FactGanadores", "Tipo de ganador"), "Tipo"),
        (C("FactGanadores", "Está en proponentes"), "¿Ofertó?"),
    ], orden=C("FactGanadores", "Orden"), desc=False, anchos=[30, 330, 130, 70])
    a = p.panel(856, 518, 408, 186, "Incidencias del proceso")
    p.tabla(a, [
        (C("FactIncidencias", "Tipo"), "Tipo"),
        (C("FactIncidencias", "Campo"), "Campo"),
        (C("FactIncidencias", "Detalle"), "Detalle"),
    ], anchos=[100, 80, 196], ajustar=True)
    paginas.append(p)

    # ------------------------------------------------------------ comunes: navegacion y corte de datos
    ids = {pg.clave: pg.id for pg in paginas}
    for pg in paginas:
        for i, (clave, _) in enumerate(NAV):
            pg.boton_nav(ids[clave], i)
        pg.tarjeta(20, 582, SIDEBAR - 40, 30, Me("Texto corte"), 10.5, PAPEL, "Segoe UI Semibold")
        pg.tarjeta(20, 610, SIDEBAR - 40, 26, Me("Texto fuente"), 8.5, GRIS_CLARO, "Segoe UI")
    return paginas


def tema() -> dict:
    return {
        "name": "SAML Editorial",
        "dataColors": PALETA,
        "foreground": CARBON,
        "foregroundNeutralSecondary": GRIS,
        "foregroundNeutralTertiary": GRIS_CLARO,
        "background": BLANCO,
        "backgroundLight": PAPEL,
        "backgroundNeutral": ARENA,
        "tableAccent": AMBAR,
        "good": PETROLEO,
        "neutral": AMBAR,
        "bad": OXIDO,
        "maximum": AMBAR_OSC,
        "center": AMBAR,
        "minimum": "#F6E3C2",
        "null": "#C9C2B6",
        "hyperlink": PETROLEO,
        "visitedHyperlink": PETROLEO,
        "textClasses": {
            "callout": {"fontSize": 26, "fontFace": "Segoe UI Semibold", "color": CARBON},
            "title": {"fontSize": 12, "fontFace": "Segoe UI Semibold", "color": CARBON},
            "header": {"fontSize": 11, "fontFace": "Segoe UI Semibold", "color": CARBON},
            "label": {"fontSize": 9, "fontFace": "Segoe UI", "color": GRIS},
        },
        "visualStyles": {
            "*": {"*": {
                "background": [{"show": False}],
                "border": [{"show": False}],
                "dropShadow": [{"show": False}],
                "title": [{"show": False, "fontColor": {"solid": {"color": CARBON}}, "fontFamily": "Segoe UI Semibold"}],
                "visualHeader": [{"background": {"solid": {"color": BLANCO}}, "foreground": {"solid": {"color": GRIS}},
                                  "border": {"solid": {"color": ARENA}}, "transparency": 0}],
                "visualTooltip": [{"titleFontColor": {"solid": {"color": GRIS}}, "valueFontColor": {"solid": {"color": CARBON}},
                                   "background": {"solid": {"color": BLANCO}}, "fontSize": 9}],
            }},
            "page": {"*": {
                "background": [{"transparency": 0}],
                "outspace": [{"color": {"solid": {"color": PAPEL_2}}, "transparency": 0}],
            }},
        },
    }


def escribir_reporte() -> None:
    defn = REPORTE / "definition"
    pages_dir = defn / "pages"
    if pages_dir.exists():
        shutil.rmtree(pages_dir)
    pages_dir.mkdir(parents=True)
    rr = REPORTE / "StaticResources" / "RegisteredResources"
    if rr.exists():
        shutil.rmtree(rr)
    rr.mkdir(parents=True)

    paginas = construir_paginas()
    items = []
    for pg in paginas:
        pg.lienzo.render().save(rr / pg.fondo, optimize=True)
        items.append({"name": pg.fondo, "path": pg.fondo, "type": "Image"})
        carpeta = pages_dir / pg.id
        (carpeta / "visuals").mkdir(parents=True)
        (carpeta / "page.json").write_text(json.dumps(pg.json_pagina(), indent=2, ensure_ascii=False), encoding="utf-8")
        for v in pg.visuales:
            vd = carpeta / "visuals" / v["name"]
            vd.mkdir()
            (vd / "visual.json").write_text(json.dumps(v, indent=2, ensure_ascii=False), encoding="utf-8")

    (rr / "SAML_Editorial.json").write_text(json.dumps(tema(), indent=2, ensure_ascii=False), encoding="utf-8")
    items.append({"name": "SAML_Editorial.json", "path": "SAML_Editorial.json", "type": "CustomTheme"})

    (pages_dir / "pages.json").write_text(json.dumps({
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/pagesMetadata/1.0.0/schema.json",
        "pageOrder": [pg.id for pg in paginas],
        "activePageName": paginas[0].id,
    }, indent=2), encoding="utf-8")

    version = {"visual": "2.8.0", "report": "3.2.0", "page": "2.3.1"}
    reporte = {
        "$schema": "https://developer.microsoft.com/json-schemas/fabric/item/report/definition/report/3.2.0/schema.json",
        "themeCollection": {
            "baseTheme": {"name": "CY26SU04", "reportVersionAtImport": version, "type": "SharedResources"},
            "customTheme": {"name": "SAML_Editorial.json", "reportVersionAtImport": version, "type": "RegisteredResources"},
        },
        "objects": {"section": [{"properties": {"verticalAlignment": L("Middle")}}]},
        "resourcePackages": [
            {"name": "SharedResources", "type": "SharedResources",
             "items": [{"name": "CY26SU04", "path": "BaseThemes/CY26SU04.json", "type": "BaseTheme"}]},
            {"name": "RegisteredResources", "type": "RegisteredResources", "items": items},
        ],
        "settings": {
            "useStylableVisualContainerHeader": True,
            "exportDataMode": "AllowSummarized",
            "defaultDrillFilterOtherVisuals": True,
            "allowChangeFilterTypes": True,
            "useEnhancedTooltips": True,
            "useDefaultAggregateDisplayName": True,
        },
    }
    (defn / "report.json").write_text(json.dumps(reporte, indent=2, ensure_ascii=False), encoding="utf-8")


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--fuente", type=Path, default=FUENTE_POR_DEFECTO, help="Excel fuente (hojas Datos, Ganadores, Incidencias, Resumen)")
    args = ap.parse_args()
    fuente_xlsx = args.fuente.resolve()
    if not fuente_xlsx.exists():
        raise SystemExit(f"No existe la fuente: {fuente_xlsx}")
    escribir_modelo(fuente_xlsx)
    escribir_reporte()
    print(f"Modelo y reporte generados. Fuente: {fuente_xlsx}")
    print(f"Abrir: {AQUI / 'SAML_Dashboard.pbip'}")


if __name__ == "__main__":
    main()
