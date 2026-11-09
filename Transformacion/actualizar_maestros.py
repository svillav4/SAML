"""Genera una COPIA actualizada de cada maestro con las obras nuevas ya insertadas, usando Excel.

Por que Excel y no openpyxl: los maestros tienen graficos y listas desplegables que
openpyxl borra al guardar. Excel (por automatizacion COM, solo en Windows con Excel
instalado) abre la copia, inserta las filas y la guarda sin perder nada.

Para cada hoja con obras nuevas:
    1. Inserta N filas completas ENCIMA de la ultima obra. Asi quedan dentro de los
       rangos del Resumen, Mercado 2027 y Top 15 (p. ej. I3:I93), que Excel amplia solo.
    2. Copia en ellas la ultima obra del propio maestro (formulas, formatos,
       desplegables y formato condicional de ese mismo libro).
    3. Escribe los datos de SAML en las columnas que no son formula (las demas quedan vacias).
    4. Revisa que los rangos de las hojas de resumen cubran las filas nuevas.

El maestro original nunca se toca: se trabaja sobre Obras_..._actualizado_<fecha>.xlsx en
la carpeta de salida. GRECO revisa esa copia y, si esta bien, reemplaza su archivo.
Si no hay Excel (p. ej. en el contenedor de n8n), se omite y quedan los archivos Nuevas_.
"""
from __future__ import annotations

import re
import shutil
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

import maestros as mae
import obras_asesores as oa

XL_SHIFT_DOWN = -4121
HOJAS_RESUMEN = ("Resumen", "Mercado 2027", "Top 15 2027")


@dataclass
class Actualizacion:
    maestro: mae.ArchivoMaestro
    ruta: Path
    filas_por_hoja: dict[str, int]
    avisos: list[str] = field(default_factory=list)


@dataclass
class ResultadoActualizacion:
    disponible: bool
    motivo: str = ""
    archivos: list[Actualizacion] = field(default_factory=list)
    errores: list[str] = field(default_factory=list)


def excel_disponible() -> tuple[bool, str]:
    if sys.platform != "win32":
        return False, "No es Windows: no hay Excel para actualizar copias de los maestros."
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return False, "Falta el paquete 'pywin32' (pip install -r Transformacion/requirements.txt)."
    return True, ""


def _valor(valor: Any) -> Any:
    if isinstance(valor, str):
        return ILLEGAL_CHARACTERS_RE.sub("", valor)
    if isinstance(valor, datetime):
        return valor.replace(tzinfo=None)
    return valor


def _rangos_cortos(libro: Any, hoja: str, ultima: int) -> list[str]:
    """Formulas de las hojas de resumen cuyos rangos sobre `hoja` terminan antes de `ultima`."""
    patron = re.compile(rf"'?{re.escape(hoja)}'?!\$?[A-Z]{{1,3}}\$?3:\$?[A-Z]{{1,3}}\$?(\d+)")
    cortos = []
    for nombre in HOJAS_RESUMEN:
        try:
            ws = libro.Worksheets(nombre)
        except Exception:  # noqa: BLE001 - el maestro puede no tener esa hoja
            continue
        usado = ws.UsedRange
        formulas = usado.Formula
        filas = formulas if isinstance(formulas, tuple) else ((formulas,),)
        for i, fila in enumerate(filas):
            for j, formula in enumerate(fila):
                if isinstance(formula, str) and formula.startswith("="):
                    for fin in patron.findall(formula):
                        if int(fin) < ultima:
                            celda = ws.Cells(usado.Row + i, usado.Column + j).Address.replace("$", "")
                            cortos.append(f"{nombre}!{celda}")
    return sorted(set(cortos))


def _insertar(libro: Any, hoja: mae.HojaMaestro, filas: list[oa.FilaObra]) -> list[str]:
    ws = libro.Worksheets(hoja.nombre)
    # Si el maestro se guardo con varias hojas agrupadas, Excel no deja insertar filas ("selecciones multiples").
    ws.Activate()
    ws.Select(True)
    avisos = []
    # Con un filtro activo (filas ocultas) Excel no deja copiar sobre el rango: se muestran todas las filas.
    if ws.FilterMode:
        ws.ShowAllData()
        avisos.append(f"{hoja.nombre}: tenía un filtro activo; en la copia se muestran todas las filas (vuelva a filtrar si lo necesita).")
    n = len(filas)
    ultima = hoja.ultima_fila
    columna_final = mae.ULTIMA_COLUMNA
    if ultima > mae.FILA_PLANTILLA:
        destino = ultima  # encima de la ultima obra: dentro de los rangos
        modelo = ultima + n  # la ultima obra baja n filas
    else:
        destino = ultima + 1  # una sola obra en la hoja: se insertan debajo de ella
        modelo = ultima
    ws.Rows(f"{destino}:{destino + n - 1}").Insert(XL_SHIFT_DOWN)
    ws.Range(f"A{modelo}:{columna_final}{modelo}").Copy(ws.Range(f"A{destino}:{columna_final}{destino + n - 1}"))

    letras_datos = [c.column_letter for c in hoja.plantilla if not (isinstance(c.value, str) and c.value.startswith("="))]
    for desplazamiento, fila in enumerate(filas):
        r = destino + desplazamiento
        for letra in letras_datos:
            ws.Range(f"{letra}{r}").Value = _valor(fila.valores.get(letra))

    errores = 0
    for celda in ws.Range(f"A{destino}:{columna_final}{destino + n - 1}"):
        texto = str(celda.Text)
        if texto.startswith("#") and texto[1:2].isalpha():
            errores += 1
    if errores:
        avisos.append(f"{hoja.nombre}: {errores} celdas con error en las filas insertadas.")
    nueva_ultima = max(ultima, destino + n - 1) if ultima == mae.FILA_PLANTILLA else ultima + n
    cortos = _rangos_cortos(libro, hoja.nombre, nueva_ultima)
    if cortos:
        avisos.append(
            f"{hoja.nombre}: estas fórmulas del resumen no llegan hasta la fila {nueva_ultima}; amplíe su rango: "
            + ", ".join(cortos[:10])
        )
    return avisos


def actualizar(reparto: oa.Reparto, maestros: mae.Maestros, carpeta: Path, sello: str) -> ResultadoActualizacion:
    filas = {id(a): reparto.filas_de(a) for a in maestros.archivos}
    return _actualizar(maestros, lambda maestro: filas.get(id(maestro), []), carpeta, sello)


def _actualizar(maestros: mae.Maestros, filas_de, carpeta: Path, sello: str) -> ResultadoActualizacion:
    disponible, motivo = excel_disponible()
    resultado = ResultadoActualizacion(disponible=disponible, motivo=motivo)
    if not disponible:
        return resultado

    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        for maestro in maestros.archivos:
            filas = filas_de(maestro)
            if not filas:
                continue
            destino = carpeta / f"{maestro.ruta.stem}_actualizado_{sello}{maestro.ruta.suffix}"
            shutil.copy2(maestro.ruta, destino)
            libro = None
            try:
                libro = excel.Workbooks.Open(str(destino.resolve()))
                por_hoja: dict[str, list[oa.FilaObra]] = {}
                for fila in filas:
                    por_hoja.setdefault(fila.hoja.nombre, []).append(fila)
                avisos: list[str] = []
                for nombre, lista in por_hoja.items():
                    avisos += _insertar(libro, maestro.hojas[nombre], lista)
                excel.CalculateFull()
                libro.Save()
                resultado.archivos.append(Actualizacion(maestro, destino, {n: len(l) for n, l in por_hoja.items()}, avisos))
            except Exception as exc:  # noqa: BLE001 - cualquier fallo de Excel deja la copia marcada y sigue
                resultado.errores.append(f"{maestro.ruta.name}: no se pudo actualizar la copia ({exc}).")
                if libro is not None:
                    libro.Close(False)
                    libro = None
                destino.unlink(missing_ok=True)
            finally:
                if libro is not None:
                    libro.Close(False)
    except Exception as exc:  # noqa: BLE001 - Excel no arranco
        resultado.disponible = False
        resultado.motivo = f"No se pudo abrir Excel ({exc})."
    finally:
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()
    return resultado


# ----------------------------------------------------------------------
# Uso independiente: generar las copias actualizadas a partir de una salida ya hecha
# (por ejemplo, una ejecucion de n8n en Docker, donde no hay Excel).
# ----------------------------------------------------------------------
HOJAS_AUXILIARES = {"Instrucciones", mae.HOJA_LISTAS}


def _huellas_de_la_revision(carpeta: Path) -> dict[str, str]:
    """SHA-256 con que se generaron los archivos (hoja Resumen del Excel de revision): maestro -> huella."""
    from openpyxl import load_workbook

    revisiones = sorted(carpeta.glob("SAML_revision_*.xlsx"))
    if not revisiones:
        return {}
    libro = load_workbook(revisiones[-1], read_only=True)
    huellas = {}
    for fila in libro["Resumen"].iter_rows(values_only=True):
        if fila and isinstance(fila[0], str) and fila[0].endswith(".xlsx") and len(fila) > 3:
            if isinstance(fila[3], str) and len(fila[3]) == 64:
                huellas[fila[0]] = fila[3]
    libro.close()
    return huellas


def _filas_de_nuevas(ruta: Path, maestro: mae.ArchivoMaestro) -> list[Any]:
    """Lee un Nuevas_<maestro>_<fecha>.xlsx y devuelve sus filas como las arma SAML (solo columnas de datos)."""
    import warnings
    from types import SimpleNamespace

    from openpyxl import load_workbook

    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        libro = load_workbook(ruta)
    filas = []
    for ws in libro.worksheets:
        if ws.title in HOJAS_AUXILIARES or ws.title not in maestro.hojas:
            continue
        hoja = maestro.hojas[ws.title]
        datos = [c.column_letter for c in hoja.plantilla if not (isinstance(c.value, str) and c.value.startswith("="))]
        for r in range(mae.FILA_PLANTILLA, ws.max_row + 1):
            if all(ws[f"{letra}{r}"].value in (None, "") for letra in ("A", "I", "K")):
                continue
            filas.append(SimpleNamespace(hoja=hoja, valores={letra: ws[f"{letra}{r}"].value for letra in datos}))
    return filas


def desde_carpeta(carpeta: Path, maestros_dir: Path) -> ResultadoActualizacion:
    maestros = mae.cargar(maestros_dir)
    sello = carpeta.name.removeprefix("SAML_").removesuffix("_REVISAR")
    huellas = _huellas_de_la_revision(carpeta)
    errores: list[str] = []
    ya_hechos: list[str] = []
    filas_por_maestro: dict[int, list[Any]] = {}
    for maestro in maestros.archivos:
        nuevas = sorted(carpeta.glob(f"Nuevas_{maestro.ruta.stem}_*.xlsx"))
        if not nuevas:
            continue
        if list(carpeta.glob(f"{maestro.ruta.stem}_actualizado_*.xlsx")):
            ya_hechos.append(f"{maestro.ruta.name}: ya tiene su copia actualizada en esta carpeta; no se repite.")
            continue
        esperada = huellas.get(maestro.ruta.name)
        if esperada and esperada != maestro.sha256:
            errores.append(
                f"{maestro.ruta.name} cambió desde que se generó esta salida (quizá ya se le agregaron obras): "
                "vuelva a ejecutar SAML para no duplicar obras."
            )
            continue
        filas_por_maestro[id(maestro)] = _filas_de_nuevas(nuevas[-1], maestro)
    resultado = _actualizar(maestros, lambda m: filas_por_maestro.get(id(m), []), carpeta, sello)
    resultado.errores = errores + resultado.errores
    resultado.motivo = " ".join(ya_hechos)
    return resultado


def main() -> int:
    import argparse

    raiz = Path(__file__).resolve().parent.parent
    parser = argparse.ArgumentParser(
        description="Genera las copias <maestro>_actualizado_<fecha>.xlsx de una salida de SAML ya generada (p. ej. desde n8n)."
    )
    parser.add_argument("carpeta", nargs="?", type=Path, help="carpeta datos/salida/SAML_<fecha> (por defecto, la más reciente)")
    parser.add_argument("--maestros", type=Path, default=raiz / "datos" / "maestros", help="carpeta de los maestros")
    args = parser.parse_args()

    carpeta = args.carpeta
    if carpeta is None:
        candidatas = sorted(p for p in (raiz / "datos" / "salida").glob("SAML_*") if p.is_dir())
        if not candidatas:
            print("ERROR: no hay carpetas SAML_* en datos/salida.")
            return 1
        carpeta = candidatas[-1]
    print(f"Carpeta: {carpeta}")
    try:
        resultado = desde_carpeta(carpeta, args.maestros)
    except mae.ErrorMaestros as exc:
        print(f"ERROR: {exc}")
        return 1
    if not resultado.disponible:
        print(f"ERROR: {resultado.motivo}")
        return 1
    for a in resultado.archivos:
        print(f"Maestro actualizado para {a.maestro.asesor}: {a.ruta} ({', '.join(f'{h}: {n}' for h, n in a.filas_por_hoja.items())})")
        for aviso in a.avisos:
            print(f"  aviso: {aviso}")
    for error in resultado.errores:
        print(f"  ERROR: {error}")
    if resultado.motivo:
        print(resultado.motivo)
    if not resultado.archivos and not resultado.errores and not resultado.motivo:
        print("No hay archivos Nuevas_ de maestros en esa carpeta.")
    return 1 if resultado.errores and not resultado.archivos else 0


if __name__ == "__main__":
    sys.exit(main())
