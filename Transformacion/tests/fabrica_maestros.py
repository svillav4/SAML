"""Crea maestros SINTETICOS con la misma estructura que los Excel de GRECO (no son datos reales).

Cada hoja de obras tiene encabezados en la fila 2, 71 columnas (A:BS), porcentajes en
T1:V1, el primer mes de la proyeccion en X2 y una obra de ejemplo en la fila 3 con
formulas equivalentes a las del maestro.
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.utils import column_index_from_string, get_column_letter

ENCABEZADOS = {
    "A": "Entidad", "B": "Localidad", "C": "Latitud", "D": "Longitud", "E": "Asesor",
    "F": "Despacho/Disposicion", "G": "Contactado", "H": "Nit", "I": "Empresa", "J": "Notas",
    "K": "Nombre de la Obra", "L": "Tipo de contrato", "M": "Contrato No", "N": "Valor\nMillones",
    "O": "Fecha Adjudicación", "P": "Avance\n%", "Q": "Meses por Terminar", "R": "Inicio de obra",
    "S": "Valor por Ejecutar", "T": "Potencial Infraestructura", "U": "Meta 25% Material Reciclado",
    "V": "MS Greco", "W": "Meta\nmes", "AY": "Origen del registro", "AZ": "Proceso de contratación",
    "BA": "Fecha de cargue", "BB": "Meses de espera antes de iniciar compras", "BH": "Fuente del valor (millones)",
    "BI": "Cliente Sugeine (nivel)", "BJ": "Base de la asignación", "BK": "Evidencia e integrantes del consorcio",
    "BL": "Ventas 2026 de los clientes enlazados", "BM": "Asesor por entidad (distribución 10-oct-26)",
    "BN": "Inicio de proyección (calculado)", "BO": "Fin de proyección (calculado)",
    "BP": "Estado de la obra (calculado)", "BQ": "Año de adjudicación (calculado)",
    "BR": "¿Consume material? (calculado)", "BS": "¿Suma a 2027? / motivo (calculado)",
}
FORMULAS = {
    "S": '=IF(OR(L{r}="Conservación, mantenimiento y reparcheo",L{r}="Repetido",E{r}="Sugeine"),0,+(1-P{r})*N{r})',
    "T": "=IF(ISNUMBER(S{r}),S{r}*$T$1,0)",
    "U": "=IF(ISNUMBER(T{r}),T{r}*$U$1,0)",
    "V": "=IF(ISNUMBER(U{r}),U{r}*$V$1,0)",
    "W": "=IFERROR(V{r}/Q{r},0)",
    "X": '=IF(AND(X$2>=$BN{r},X$2<=$BO{r}),$W{r},"")',
    "Y": '=IF(AND(Y$2>=$BN{r},Y$2<=$BO{r}),$W{r},"")',
    "BN": "=DATE(YEAR(MAX(N($R{r}),$X$2)),MONTH(MAX(N($R{r}),$X$2))+N($BB{r}),1)",
    "BO": "=DATE(YEAR($BN{r}),MONTH($BN{r})+INT(MAX(0,N($Q{r})))-1,1)",
    "BP": '=IF(AND($I{r}="",$K{r}=""),"",IF($E{r}="Sugeine","En archivo de Sugeine","En construcción"))',
}
TIPOS = [
    "Conservación, mantenimiento y reparcheo", "Señalización", "Semaforización", "Cámaras de seguridad", "N/A Obra",
    "Interventoría", "Suspendido", "Repetido", "Construcción de obra", "Redes de acueducto y alcantarillado",
    "Parques y escenarios deportivos", "Edificaciones y adecuaciones", "Obra privada", "Por clasificar",
]
ASESORES = ["Gustavo", "Mariana", "Asesor 3 (por contratar)", "Sugeine", "Sin asignar (GRECO)", "Mariana (posible Sugeine)", "Gustavo (posible Sugeine)"]
INICIO_PROYECCION = datetime(2026, 10, 1)


def _hoja_obras(libro: Workbook, nombre: str, entidad: str, asesor: str, obras: list[dict]) -> None:
    ws = libro.create_sheet(nombre)
    ws["T1"], ws["U1"], ws["V1"] = 0.4, 0.25, 0.3
    ws["X1"], ws["Y1"] = "=YEAR(X2)", "=YEAR(Y2)"
    ws["BC1"] = "=SUM(BC3:BC116)"
    for c in range(1, column_index_from_string("BS") + 1):
        letra = get_column_letter(c)
        ws[f"{letra}2"] = ENCABEZADOS.get(letra, f"Columna {letra}")
    ws["X2"], ws["Y2"] = INICIO_PROYECCION, "=EDATE(X2,1)"
    ws["E2"].number_format = "General"
    for desplazamiento, obra in enumerate(obras or [{}]):
        r = 3 + desplazamiento
        ws[f"A{r}"], ws[f"E{r}"] = entidad, obra.get("asesor", asesor)
        ws[f"I{r}"], ws[f"K{r}"] = obra.get("empresa", "EMPRESA EXISTENTE SAS"), "Obra existente"
        ws[f"M{r}"], ws[f"N{r}"] = obra.get("codigo", f"EXISTENTE-{nombre[:3]}-{r}"), obra.get("valor", 1000)
        ws[f"B{r}"] = obra.get("localidad", "Bogotá D.C.")
        ws[f"P{r}"], ws[f"Q{r}"], ws[f"R{r}"] = 0.15, 8, INICIO_PROYECCION
        ws[f"N{r}"].number_format = "#,##0"
        ws[f"R{r}"].number_format = "mmm-yy"
        for letra, formula in FORMULAS.items():
            ws[f"{letra}{r}"] = formula.format(r=r)


def _hoja_listas(libro: Workbook) -> None:
    ws = libro.create_sheet("Listas")
    ws.append(["Tipo de contrato", "Efecto", None, "Inicio de obra (lista de meses)", None, "Asesor (lista)"])
    meses = [datetime(2026 + (9 + i) // 12, (9 + i) % 12 + 1, 1) for i in range(63)]
    for i in range(63):
        ws.append([TIPOS[i] if i < len(TIPOS) else None, None, None, meses[i], None, ASESORES[i] if i < len(ASESORES) else None])


def crear_maestros(
    carpeta: Path,
    *,
    existentes_alcaldias: list[dict] | None = None,
    clientes: list[str] | None = None,
    con_idrd: bool = False,
) -> list[Path]:
    """Tres maestros: Gustavo (IDU), Mariana (Alcaldías Locales y Municipios Aledaños) y Sugeine.
    Con con_idrd=True, un cuarto maestro de Asesor 3 con la hoja IDRD (plantilla de "Edificaciones")."""
    carpeta.mkdir(parents=True, exist_ok=True)
    rutas = []

    libro = Workbook()
    libro.remove(libro.active)
    _hoja_obras(libro, "IDU", "IDU", "Gustavo", [{}])
    _hoja_listas(libro)
    rutas.append(carpeta / "Obras_IDU_Prueba.xlsx")
    libro.save(rutas[-1])

    libro = Workbook()
    libro.remove(libro.active)
    _hoja_obras(libro, "Alcaldías Locales", "Alcaldía Local", "Mariana", existentes_alcaldias or [{}])
    _hoja_obras(libro, "Municipios Aledaños", "Municipio", "Mariana", [{}])
    _hoja_listas(libro)
    rutas.append(carpeta / "Obras_Alcaldias_Prueba.xlsx")
    libro.save(rutas[-1])

    libro = Workbook()
    libro.remove(libro.active)
    _hoja_obras(libro, "IDU", "IDU", "Sugeine", [{}])
    _hoja_obras(libro, "Alcaldías Locales", "Alcaldía Local", "Sugeine", [{}])
    ws = libro.create_sheet("Cobertura clientes")
    ws.append(["Cliente (razón social)", "Origen del cliente"])
    for cliente in clientes or ["CONSORCIO VIAS CB 2024", "INCITECO"]:
        ws.append([cliente, "Sintético"])
    _hoja_listas(libro)
    rutas.append(carpeta / "Obras_Sugeine_Prueba.xlsx")
    libro.save(rutas[-1])

    if con_idrd:
        libro = Workbook()
        libro.remove(libro.active)
        _hoja_obras(libro, "IDRD", "IDRD", "Asesor 3 (por contratar)", [{}])
        libro["IDRD"]["U1"] = 0.23
        _hoja_listas(libro)
        rutas.append(carpeta / "Obras_IDRD_Prueba.xlsx")
        libro.save(rutas[-1])
    return rutas
