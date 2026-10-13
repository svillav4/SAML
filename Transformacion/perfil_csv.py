"""Perfil de solo lectura de un CSV generado por Ingesta/extract.py.

Sirve para conocer la estructura de los datos antes de definir la limpieza,
sin imprimir el contenido de los registros:
    - numero de filas y columnas;
    - por columna: cuantas filas tienen valor, cuantos valores distintos hay
      y que tipo de dato parece contener (texto, numero o estructura anidada);
    - para las columnas anidadas (p. ej. "fechas" o "documentos", que
      extract.py guarda como texto de Python), las claves internas que usan.

Solo muestra valores concretos de las columnas que se pidan con --valores
(por ejemplo, el estado del proceso), y nunca modifica el archivo.

Uso (desde la raiz del proyecto):
    .venv/bin/python Transformacion/perfil_csv.py                      # CSV mas reciente
    .venv/bin/python Transformacion/perfil_csv.py ruta/al/archivo.csv
    .venv/bin/python Transformacion/perfil_csv.py --valores estado_agrupado
"""
from __future__ import annotations

import argparse
import ast
import csv
import sys
from collections import Counter
from pathlib import Path

INGESTA_OUTPUT = Path(__file__).resolve().parent.parent / "Ingesta" / "output"


def csv_mas_reciente() -> Path:
    candidatos = sorted(INGESTA_OUTPUT.glob("licitaciones_*.csv"))
    if not candidatos:
        sys.exit(f"No hay archivos licitaciones_*.csv en {INGESTA_OUTPUT}. Ejecuta primero la ingesta.")
    return candidatos[-1]


def interpretar_anidado(valor: str):
    """Devuelve el dict/list si el texto es una estructura de Python o JSON; si no, None."""
    texto = valor.strip()
    if not texto or texto[0] not in "[{":
        return None
    try:
        return ast.literal_eval(texto)
    except (ValueError, SyntaxError):
        pass
    try:
        import json

        return json.loads(texto)
    except ValueError:
        return None


def claves_de(estructura) -> set[str]:
    if isinstance(estructura, dict):
        return set(map(str, estructura.keys()))
    if isinstance(estructura, list):
        claves: set[str] = set()
        for elemento in estructura:
            if isinstance(elemento, dict):
                claves.update(map(str, elemento.keys()))
        return claves
    return set()


def es_numero(valor: str) -> bool:
    try:
        float(valor.replace(",", ""))
        return True
    except ValueError:
        return False


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("csv", nargs="?", type=Path, help="CSV a revisar (por defecto, el mas reciente de Ingesta/output)")
    parser.add_argument(
        "--valores",
        action="append",
        default=[],
        metavar="COLUMNA",
        help="muestra los valores distintos (con conteo) de esta columna; se puede repetir",
    )
    args = parser.parse_args()

    ruta = args.csv or csv_mas_reciente()
    csv.field_size_limit(sys.maxsize)

    with ruta.open(newline="", encoding="utf-8") as f:
        lector = csv.DictReader(f)
        columnas = lector.fieldnames or []
        filas = list(lector)

    print(f"Archivo: {ruta}")
    if columnas == ["Sin resultados"] or not filas:
        print("El CSV no contiene registros (la API no devolvio procesos con los filtros del .env).")
        return

    print(f"Filas: {len(filas)}  |  Columnas: {len(columnas)}\n")
    print(f"{'columna':<40} {'con valor':>10} {'distintos':>10}  tipo")
    print("-" * 80)

    for columna in columnas:
        valores = [(fila.get(columna) or "").strip() for fila in filas]
        con_valor = [v for v in valores if v]
        anidados = [e for e in (interpretar_anidado(v) for v in con_valor) if e is not None]

        if con_valor and len(anidados) == len(con_valor):
            claves: set[str] = set()
            for estructura in anidados:
                claves |= claves_de(estructura)
            tipo = "anidado: " + (", ".join(sorted(claves)) if claves else "lista sin claves")
        elif con_valor and all(es_numero(v) for v in con_valor):
            tipo = "numero"
        elif con_valor:
            tipo = "texto"
        else:
            tipo = "(siempre vacio)"

        print(f"{columna:<40} {len(con_valor):>10} {len(set(con_valor)):>10}  {tipo}")

    for columna in args.valores:
        if columna not in columnas:
            print(f"\n--valores {columna}: la columna no existe en el CSV.")
            continue
        conteo = Counter((fila.get(columna) or "").strip() or "(vacio)" for fila in filas)
        print(f"\nValores de '{columna}':")
        for valor, n in conteo.most_common(30):
            print(f"  {n:>6}  {valor[:80]}")
        if len(conteo) > 30:
            print(f"  ... y {len(conteo) - 30} valores mas")


if __name__ == "__main__":
    main()
