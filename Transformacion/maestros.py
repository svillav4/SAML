"""Lectura (solo lectura) de los 4 Excel maestros de GRECO, uno por asesor.

Los maestros viven en datos/maestros/ (carpeta excluida de git: traen NIT, ventas
y clientes). SAML nunca los modifica: solo los lee para
    - saber que hojas de obras tiene cada archivo y a que asesor pertenece;
    - tomar la plantilla de cada hoja (encabezados de las filas 1 y 2, formulas
      y formatos de la primera fila de datos y porcentajes de la fila 1), para
      que las filas nuevas sean identicas a las del maestro;
    - conocer los codigos de proceso/contrato ya cargados, y no repetir obras;
    - leer la lista de clientes de Sugeine (hoja "Cobertura clientes");
    - leer las listas desplegables (hoja "Listas").

Estructura esperada de cada hoja de obras: encabezados en la fila 2, datos desde
la fila 3 y 71 columnas (A:BS), con formulas en la fila 3 (Valor por Ejecutar,
potencial, proyeccion mensual y anual, columnas calculadas).
"""
from __future__ import annotations

import re
import unicodedata
import warnings
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

from openpyxl import load_workbook
from openpyxl.cell.cell import Cell
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

import limpieza_licitaciones as lim

FILA_ENCABEZADO = 2
FILA_PLANTILLA = 3
ULTIMA_COLUMNA = "BS"
HOJA_CLIENTES = "Cobertura clientes"
HOJA_LISTAS = "Listas"
ASESOR_SUGEINE = "Sugeine"

# Columnas que SAML llena (letra -> inicio del encabezado esperado en la fila 2).
# Si un maestro cambia el orden de estas columnas, la carga falla con un mensaje claro
# en vez de escribir datos en la columna equivocada.
COLUMNAS_ESPERADAS = {
    "A": "Entidad",
    "B": "Localidad",
    "C": "Latitud",
    "D": "Longitud",
    "E": "Asesor",
    "F": "Despacho",
    "G": "Contactad",
    "H": "Nit",
    "I": "Empresa",
    "J": "Notas",
    "K": "Nombre de la Obra",
    "L": "Tipo de contrato",
    "M": "Contrato No",
    "N": "Valor Millones",
    "O": "Fecha Adjudicación",
    "P": "Avance",
    "Q": "Meses por Terminar",
    "R": "Inicio de obra",
    "S": "Valor por Ejecutar",
    "AY": "Origen del registro",
    "AZ": "Proceso de contratación",
    "BA": "Fecha de cargue",
    "BB": "Meses de espera",
    "BH": "Fuente del valor",
    "BI": "Cliente Sugeine",
    "BJ": "Base de la asignación",
    "BK": "Evidencia e integrantes",
    "BM": "Asesor por entidad",
}


class ErrorMaestros(Exception):
    """Los maestros faltan o no tienen la estructura esperada."""


def normalizar(texto: Any) -> str:
    """Mayusculas sin tildes y con un solo espacio entre palabras alfanumericas."""
    t = unicodedata.normalize("NFKD", str(texto or "")).encode("ascii", "ignore").decode().upper()
    return re.sub(r"[^A-Z0-9]+", " ", t).strip()


def normalizar_codigo(codigo: Any) -> str:
    """Codigo de proceso comparable: sin espacios, guiones ni signos ('FDLBOSA LIC 001-2026-' == 'FDLBOSA-LIC-001-2026')."""
    return normalizar(codigo).replace(" ", "")


def _encabezado(valor: Any) -> str:
    return re.sub(r"\s+", " ", str(valor or "")).strip()


@dataclass(frozen=True)
class ObraExistente:
    """Una obra ya cargada en un maestro (para no repetirla)."""

    ubicacion: str  # "archivo / hoja (fila n)"
    empresa: str
    valor_millones: Optional[float]


@dataclass
class HojaMaestro:
    archivo: str
    nombre: str
    entidad: str  # valor de la columna A (categoria), p. ej. "IDU" o "Alcaldía Local"
    filas_datos: int
    ultima_fila: int  # fila de la ultima obra (las nuevas se insertan encima, dentro de los rangos del Resumen)
    inicio_proyeccion: Optional[datetime]  # celda X2: primer mes de la proyeccion
    porcentajes: dict[str, Any]  # T1, U1, V1
    encabezados: dict[str, str]  # letra -> texto de la fila 2
    filas_encabezado: list[list[Cell]] = field(default_factory=list, repr=False)  # filas 1 y 2
    plantilla: list[Cell] = field(default_factory=list, repr=False)  # fila 3 (formulas y formatos)
    anchos: dict[str, float] = field(default_factory=dict, repr=False)

    @property
    def columnas(self) -> list[str]:
        return [get_column_letter(i) for i in range(1, column_index_from_string(ULTIMA_COLUMNA) + 1)]


@dataclass
class ArchivoMaestro:
    ruta: Path
    asesor: str
    es_sugeine: bool
    sha256: str
    hojas: dict[str, HojaMaestro]


@dataclass
class Maestros:
    directorio: Path
    archivos: list[ArchivoMaestro]
    codigos: dict[str, list[ObraExistente]]  # codigo normalizado -> obras con ese codigo
    clientes_sugeine: dict[str, str]  # nombre normalizado -> nombre original
    localidades: dict[str, str]  # localidad normalizada -> escritura usada en los maestros
    tipos_contrato: list[str]
    asesores: list[str]
    meses_inicio: list[datetime]
    listas: list[list[Any]]  # contenido de la hoja Listas (valores) para copiarlo a las salidas

    @property
    def sugeine(self) -> Optional[ArchivoMaestro]:
        return next((a for a in self.archivos if a.es_sugeine), None)

    def archivo_de_hoja(self, hoja: str) -> Optional[ArchivoMaestro]:
        """Archivo del asesor dueno de la hoja (sin contar el de Sugeine, que se arma por clientes)."""
        return next((a for a in self.archivos if not a.es_sugeine and hoja in a.hojas), None)


def _es_hoja_de_obras(ws: Worksheet) -> bool:
    # "Archivo (no vigentes)" se parece, pero no tiene las columnas calculadas BN:BS: no es plantilla.
    return (
        _encabezado(ws["A2"].value) == "Entidad"
        and _encabezado(ws["S2"].value).startswith("Valor por Ejecutar")
        and _encabezado(ws["BP2"].value).startswith("Estado de la obra")
        and isinstance(ws["S3"].value, str)
        and ws["S3"].value.startswith("=")
    )


def _filas_con_datos(ws: Worksheet) -> list[int]:
    # Una fila es una obra si tiene Entidad, Empresa o Nombre de la Obra (las formulas llegan mas abajo).
    return [r for r in range(FILA_PLANTILLA, ws.max_row + 1) if any(ws.cell(r, c).value not in (None, "") for c in (1, 9, 11))]


def _validar_encabezados(ws: Worksheet, archivo: str) -> dict[str, str]:
    encabezados = {
        get_column_letter(c): _encabezado(ws.cell(FILA_ENCABEZADO, c).value)
        for c in range(1, column_index_from_string(ULTIMA_COLUMNA) + 1)
    }
    errores = [
        f"{letra}2 dice '{encabezados[letra]}' y se esperaba '{esperado}…'"
        for letra, esperado in COLUMNAS_ESPERADAS.items()
        if not normalizar(encabezados[letra]).startswith(normalizar(esperado))
    ]
    if errores:
        raise ErrorMaestros(f"{archivo} / {ws.title}: la plantilla no coincide con la esperada: " + "; ".join(errores))
    return encabezados


def _lista(ws: Worksheet, columna: str) -> list[Any]:
    valores = []
    for r in range(2, ws.max_row + 1):
        v = ws[f"{columna}{r}"].value
        if v in (None, ""):
            break
        valores.append(v)
    return valores


def cargar(directorio: Path) -> Maestros:
    """Lee todos los .xlsx de la carpeta de maestros. No modifica ningun archivo."""
    if not directorio.is_dir():
        raise ErrorMaestros(f"No existe la carpeta de maestros {directorio}. Copia ahí los 4 Excel de GRECO.")
    rutas = sorted(p for p in directorio.glob("*.xlsx") if not p.name.startswith("~$"))
    if not rutas:
        raise ErrorMaestros(f"No hay archivos .xlsx en {directorio}. Copia ahí los 4 Excel de GRECO.")

    archivos: list[ArchivoMaestro] = []
    codigos: dict[str, list[ObraExistente]] = {}
    clientes: dict[str, str] = {}
    localidades: Counter = Counter()
    listas_ws: Optional[Worksheet] = None

    for ruta in rutas:
        sha = lim.sha256_de(ruta)
        with warnings.catch_warnings():
            # openpyxl avisa que no lee las validaciones extendidas; da igual porque nunca se guarda el maestro.
            warnings.simplefilter("ignore", UserWarning)
            libro = load_workbook(ruta)  # con formulas (no data_only) para copiar la plantilla
        hojas: dict[str, HojaMaestro] = {}
        asesores_vistos: Counter = Counter()

        for ws in libro.worksheets:
            # Codigos ya cargados, en cualquier hoja con "Contrato No" (tambien "Archivo (no vigentes)").
            if _encabezado(ws["M2"].value) == "Contrato No":
                for r in range(FILA_PLANTILLA, ws.max_row + 1):
                    for letra in ("M", "AZ"):
                        codigo = normalizar_codigo(ws[f"{letra}{r}"].value)
                        if not codigo:
                            continue
                        ubicacion = f"{ruta.name} / {ws.title} (fila {r})"
                        registros = codigos.setdefault(codigo, [])
                        if all(reg.ubicacion != ubicacion for reg in registros):
                            valor = ws[f"N{r}"].value
                            registros.append(
                                ObraExistente(
                                    ubicacion=ubicacion,
                                    empresa=str(ws[f"I{r}"].value or ""),
                                    valor_millones=float(valor) if isinstance(valor, (int, float)) else None,
                                )
                            )
            if ws.title == HOJA_CLIENTES:
                for r in range(2, ws.max_row + 1):
                    nombre = ws.cell(r, 1).value
                    if nombre not in (None, ""):
                        clientes[normalizar(nombre)] = str(nombre).strip()
                continue
            if ws.title == HOJA_LISTAS and listas_ws is None:
                listas_ws = ws
                continue
            if not _es_hoja_de_obras(ws):
                continue

            encabezados = _validar_encabezados(ws, ruta.name)
            filas = _filas_con_datos(ws)
            entidades = Counter(str(ws.cell(r, 1).value).strip() for r in filas if ws.cell(r, 1).value)
            for r in filas:
                asesor = ws.cell(r, 5).value
                if asesor:
                    asesores_vistos[str(asesor).strip()] += 1
                if ws.cell(r, 2).value:
                    localidades[str(ws.cell(r, 2).value).strip()] += 1
            maximo = column_index_from_string(ULTIMA_COLUMNA)
            hojas[ws.title] = HojaMaestro(
                archivo=ruta.name,
                nombre=ws.title,
                entidad=entidades.most_common(1)[0][0] if entidades else ws.title,
                filas_datos=len(filas),
                ultima_fila=max(filas) if filas else FILA_PLANTILLA,
                inicio_proyeccion=ws["X2"].value if isinstance(ws["X2"].value, datetime) else None,
                porcentajes={c: ws[f"{c}1"].value for c in ("T", "U", "V")},
                encabezados=encabezados,
                filas_encabezado=[[ws.cell(f, c) for c in range(1, maximo + 1)] for f in (1, FILA_ENCABEZADO)],
                plantilla=[ws.cell(FILA_PLANTILLA, c) for c in range(1, maximo + 1)],
                anchos={
                    get_column_letter(c): ws.column_dimensions[get_column_letter(c)].width
                    for c in range(1, maximo + 1)
                    if ws.column_dimensions[get_column_letter(c)].width
                },
            )

        if not hojas:
            raise ErrorMaestros(f"{ruta.name} no tiene hojas de obras con la plantilla de GRECO.")
        es_sugeine = HOJA_CLIENTES in libro.sheetnames
        if es_sugeine:
            asesor = ASESOR_SUGEINE
        else:
            propios = Counter({a: n for a, n in asesores_vistos.items() if ASESOR_SUGEINE.lower() not in a.lower()})
            if not propios:
                raise ErrorMaestros(f"{ruta.name}: no se pudo identificar el asesor (columna E).")
            asesor = propios.most_common(1)[0][0]
        archivos.append(ArchivoMaestro(ruta=ruta, asesor=asesor, es_sugeine=es_sugeine, sha256=sha, hojas=hojas))

    if listas_ws is None:
        raise ErrorMaestros("Ningún maestro tiene la hoja 'Listas' (listas desplegables).")
    sugeines = [a for a in archivos if a.es_sugeine]
    if len(sugeines) > 1:
        raise ErrorMaestros("Hay más de un maestro con la hoja 'Cobertura clientes'; debe haber uno solo (Sugeine).")
    duplicadas = Counter(h for a in archivos if not a.es_sugeine for h in a.hojas)
    repetidas = [h for h, n in duplicadas.items() if n > 1]
    if repetidas:
        raise ErrorMaestros(f"La hoja {repetidas} está en más de un maestro de asesor; no se sabe a quién asignarla.")

    meses = [v for v in _lista(listas_ws, "D") if isinstance(v, datetime)]
    contenido_listas = [[c.value for c in fila] for fila in listas_ws.iter_rows()]
    por_norma: dict[str, Counter] = {}
    for escritura, n in localidades.items():
        por_norma.setdefault(normalizar(escritura), Counter())[escritura] += n
    return Maestros(
        directorio=directorio,
        archivos=archivos,
        codigos=codigos,
        clientes_sugeine=clientes,
        localidades={k: c.most_common(1)[0][0] for k, c in por_norma.items()},
        tipos_contrato=[str(v) for v in _lista(listas_ws, "A")],
        asesores=[str(v) for v in _lista(listas_ws, "F")],
        meses_inicio=meses,
        listas=contenido_listas,
    )
