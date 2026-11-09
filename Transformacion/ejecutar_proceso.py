"""Punto de entrada del proceso local de licitaciones adjudicadas de GRECO.

Coordina las etapas existentes, sin duplicar su logica:
    1. (solo con --actualizar) Ingesta/extract.py con las credenciales de Ingesta/.env
    2. Limpieza, reparto entre los maestros y Excel por asesor (generar_archivos_asesores.generar)
    3. Verificacion de los archivos generados y de que los CSV y los maestros no cambiaron

Modos:
    (por defecto)   DEMOSTRACION: usa solo los CSV ya guardados en Ingesta/output.
                    No consulta Licitaciones.info ni necesita credenciales.
    --actualizar    ACTUALIZACION: ejecuta primero la ingesta. Si falla, o no deja
                    exactamente un CSV nuevo, el proceso se detiene y NO genera archivos.

Opciones:
    --maestros DIR  carpeta con los 4 Excel maestros de GRECO (por defecto datos/maestros)
    --salida DIR    carpeta donde se crea SAML_<fecha_hora>/ (por defecto datos/salida)

Codigos de salida: 0 correcto | 1 faltan CSV, maestros o configuracion | 2 fallo la ingesta |
3 la verificacion fallo (la carpeta de salida se renombra con sufijo _REVISAR).

Uso (desde la raiz del proyecto):
    .venv/bin/python Transformacion/ejecutar_proceso.py
    .venv/bin/python Transformacion/ejecutar_proceso.py --actualizar
"""
from __future__ import annotations

import argparse
import os
import re
import subprocess
import sys
import warnings
from pathlib import Path

from dotenv import dotenv_values
from openpyxl import load_workbook

import clasificador_llm as llm
import generar_archivos_asesores as gen
import limpieza_licitaciones as lim
import maestros as mae

RAIZ = gen.RAIZ
INGESTA = RAIZ / "Ingesta"
ENTRADA = gen.ENTRADA_POR_DEFECTO
ARCHIVO_ENV = INGESTA / ".env"

PATRON_JWT = re.compile(r"eyJ[A-Za-z0-9_-]{5,}\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+")
PATRON_CORREO = re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+")
PATRON_BEARER = re.compile(r"Bearer\s+\S+", re.IGNORECASE)
PATRON_CLAVE_API = re.compile(r"sk-[A-Za-z0-9_-]{10,}")
HOJAS_AUXILIARES = {"Instrucciones", mae.HOJA_LISTAS}


def aviso(texto: str = "") -> None:
    print(texto, flush=True)


def relativa(ruta: Path) -> str:
    ruta = ruta.resolve()
    return str(ruta.relative_to(RAIZ.resolve())).replace("\\", "/") if ruta.is_relative_to(RAIZ.resolve()) else str(ruta)


# ----------------------------------------------------------------------
# Configuracion de la ingesta (sin mostrar valores secretos)
# ----------------------------------------------------------------------
def valor_efectivo(nombre: str, archivo_env: dict) -> str:
    """Igual que config.py: una variable ya definida en el entorno prevalece sobre el .env."""
    if nombre in os.environ:
        return os.environ[nombre] or ""
    return archivo_env.get(nombre) or ""


def secretos_locales(archivo_env: dict) -> list[str]:
    return [v for v in (valor_efectivo(k, archivo_env) for k in ("LICITACIONES_EMAIL", "LICITACIONES_PASSWORD")) if v]


def ocultar(linea: str, secretos: list[str]) -> str:
    for secreto in secretos:
        linea = linea.replace(secreto, "<oculto>")
    linea = PATRON_JWT.sub("<JWT oculto>", linea)
    linea = PATRON_BEARER.sub("Bearer <oculto>", linea)
    linea = PATRON_CLAVE_API.sub("<clave oculta>", linea)
    return PATRON_CORREO.sub("<correo oculto>", linea)


def comprobar_configuracion() -> dict | None:
    if not ARCHIVO_ENV.exists():
        aviso(f"ERROR: no existe {relativa(ARCHIVO_ENV)}. Copia Ingesta/.env.example y completa las credenciales.")
        return None
    archivo_env = dotenv_values(ARCHIVO_ENV)
    faltan = [k for k in ("LICITACIONES_EMAIL", "LICITACIONES_PASSWORD") if not valor_efectivo(k, archivo_env)]
    if faltan:
        aviso(f"ERROR: faltan valores en Ingesta/.env: {', '.join(faltan)}.")
        return None
    estado = valor_efectivo("LICITACIONES_ESTADO", archivo_env)
    if estado != lim.ESTADO_ADJUDICADO:
        aviso(
            f"ERROR: LICITACIONES_ESTADO es '{estado or '(vacio)'}' y este proceso esta limitado a "
            f"'{lim.ESTADO_ADJUDICADO}'. Ajusta Ingesta/.env antes de actualizar."
        )
        return None
    aviso(f"Configuracion: credenciales definidas (valores ocultos); filtro de la ingesta LICITACIONES_ESTADO={estado}.")
    return archivo_env


# ----------------------------------------------------------------------
# Etapas
# ----------------------------------------------------------------------
def ejecutar_ingesta(secretos: list[str]) -> Path | None:
    """Ejecuta Ingesta/extract.py tal como esta. Devuelve el CSV nuevo o None si fallo."""
    antes = {p.name for p in ENTRADA.glob("licitaciones_*.csv")}
    aviso("Ejecutando Ingesta/extract.py (salida de la ingesta, con datos sensibles ocultos):")
    with subprocess.Popen(
        [sys.executable, "extract.py"],
        cwd=INGESTA,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
    ) as proceso:
        assert proceso.stdout is not None
        for linea in proceso.stdout:
            aviso("  | " + ocultar(linea.rstrip("\n"), secretos))
        codigo = proceso.wait()

    if codigo != 0:
        aviso(
            f"\nERROR: la ingesta termino con codigo {codigo}. El proceso se detiene y NO se generan archivos.\n"
            "Revisa el mensaje anterior: credenciales rechazadas (401), bloqueo de Cloudflare/WAF, "
            "falta de conexion o un error de la API."
        )
        return None

    nuevos = sorted({p.name for p in ENTRADA.glob("licitaciones_*.csv")} - antes)
    if len(nuevos) != 1:
        aviso(
            f"\nERROR: la ingesta termino sin error pero dejo {len(nuevos)} CSV nuevos (se esperaba 1). "
            "El proceso se detiene y NO se generan archivos."
        )
        return None
    return ENTRADA / nuevos[0]


def _filas(hoja) -> list[dict]:
    filas = hoja.iter_rows(min_row=1, values_only=True)
    encabezados = list(next(filas, []))
    return [dict(zip(encabezados, fila)) for fila in filas]


def _tiene_jwt(libro) -> bool:
    return any(
        isinstance(v, str) and PATRON_JWT.search(v)
        for hoja in libro.worksheets
        for fila in hoja.iter_rows(values_only=True)
        for v in fila
    )


def verificar(g: gen.Generacion, huellas_csv: dict[str, str]) -> list[str]:
    """Comprueba los archivos generados y que los CSV y los maestros siguen intactos."""
    problemas: list[str] = []

    for archivo in lim.listar_csvs(ENTRADA):
        anterior = huellas_csv.get(archivo.ruta.name)
        if anterior and anterior != archivo.sha256:
            problemas.append(f"El CSV {archivo.ruta.name} cambio durante el proceso.")
    for maestro in g.maestros.archivos:
        if lim.sha256_de(maestro.ruta) != maestro.sha256:
            problemas.append(f"El maestro {maestro.ruta.name} cambio durante el proceso.")

    ya_cargados = {mae.normalizar_codigo(p["CodigoProceso"]) for p in g.reparto.ya_en_maestros}
    total_filas = 0
    for archivo in g.archivos_asesores:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UserWarning)
            libro = load_workbook(archivo.ruta)
        if _tiene_jwt(libro):
            problemas.append(f"{archivo.ruta.name}: hay un valor con forma de JWT.")
        filas_archivo = 0
        for hoja in libro.worksheets:
            if hoja.title in HOJAS_AUXILIARES:
                continue
            codigos: list[str] = []
            for r in range(mae.FILA_PLANTILLA, hoja.max_row + 1):
                if hoja[f"A{r}"].value in (None, ""):
                    continue
                filas_archivo += 1
                asesor = hoja[f"E{r}"].value
                if asesor not in g.maestros.asesores:
                    problemas.append(f"{archivo.ruta.name} / {hoja.title} fila {r}: el Asesor '{asesor}' no está en la lista del maestro.")
                if not str(hoja[f"S{r}"].value or "").startswith("="):
                    problemas.append(f"{archivo.ruta.name} / {hoja.title} fila {r}: falta la fórmula de Valor por Ejecutar.")
                codigos.append(mae.normalizar_codigo(hoja[f"M{r}"].value))
            repetidos = {c for c in codigos if c and codigos.count(c) > 1}
            if repetidos:
                problemas.append(f"{archivo.ruta.name} / {hoja.title}: códigos repetidos {sorted(repetidos)}.")
            if ya_cargados & set(codigos):
                problemas.append(f"{archivo.ruta.name} / {hoja.title}: incluye obras que ya estaban en los maestros.")
        if filas_archivo != archivo.filas:
            problemas.append(f"{archivo.ruta.name}: tiene {filas_archivo} filas y se esperaban {archivo.filas}.")
        total_filas += filas_archivo
        libro.close()

    # Copias actualizadas de los maestros: cada hoja debe tener sus obras de antes mas las nuevas, con sus codigos.
    for a in (g.actualizacion.archivos if g.actualizacion else []):
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UserWarning)
            libro = load_workbook(a.ruta, read_only=True)
        nuevas = {}
        for fila in g.reparto.filas_de(a.maestro):
            nuevas.setdefault(fila.hoja.nombre, set()).add(mae.normalizar_codigo(fila.valores.get("M")))
        for nombre, cantidad in a.filas_por_hoja.items():
            ws = libro[nombre]
            filas_obra = [f for f in ws.iter_rows(min_row=mae.FILA_PLANTILLA, max_col=13, values_only=True) if any(f[c] not in (None, "") for c in (0, 8, 10))]
            esperadas = a.maestro.hojas[nombre].filas_datos + cantidad
            if len(filas_obra) != esperadas:
                problemas.append(f"{a.ruta.name} / {nombre}: tiene {len(filas_obra)} obras y se esperaban {esperadas}.")
            codigos = {mae.normalizar_codigo(f[12]) for f in filas_obra}
            if nuevas.get(nombre, set()) - codigos:
                problemas.append(f"{a.ruta.name} / {nombre}: faltan obras nuevas en la copia actualizada.")
        libro.close()
        if a.ruta.stat().st_size == 0:
            problemas.append(f"{a.ruta.name}: el archivo quedó vacío.")

    libro = load_workbook(g.revision, read_only=True)
    procesos = _filas(libro["Procesos"])
    estados = {p.get("Estado (origen)") for p in procesos}
    if estados - {lim.ESTADO_ADJUDICADO}:
        problemas.append(f"Revisión: hay estados distintos de '{lim.ESTADO_ADJUDICADO}': {sorted(map(str, estados))}.")
    ids = [p.get("idContrato") for p in procesos]
    if len(ids) != len(set(ids)):
        problemas.append("Revisión: hay idContrato duplicados en Procesos.")
    if _tiene_jwt(libro):
        problemas.append("Revisión: hay un valor con forma de JWT.")
    libro.close()

    if not problemas:
        aviso(
            f"Verificacion: OK. {len(procesos)} procesos adjudicados; {total_filas} filas nuevas en "
            f"{len(g.archivos_asesores)} archivos de asesores, con asesores de la lista, fórmulas y sin códigos repetidos "
            "ni obras ya cargadas; sin JWT; CSV y maestros sin cambios."
        )
    return problemas


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument(
        "--actualizar",
        action="store_true",
        help="ejecuta primero la ingesta desde Licitaciones.info; si falla, no genera archivos",
    )
    parser.add_argument("--maestros", type=Path, default=gen.MAESTROS_POR_DEFECTO, help="carpeta con los Excel maestros (por defecto datos/maestros)")
    parser.add_argument("--salida", type=Path, default=gen.SALIDA_POR_DEFECTO, help="carpeta de salida (por defecto datos/salida)")
    parser.add_argument("--sin-llm", action="store_true", help="no usa el LLM: las entidades sin regla se clasifican con reglas")
    parser.add_argument(
        "--sin-actualizar-maestros", action="store_true", help="no genera las copias actualizadas de los maestros (solo los Nuevas_)"
    )
    args = parser.parse_args()

    # Los maestros se revisan antes de consultar el portal: sin ellos no se puede repartir.
    try:
        mae.cargar(args.maestros)
    except mae.ErrorMaestros as exc:
        aviso(f"ERROR: {exc}")
        return 1

    if args.actualizar:
        aviso("=== MODO ACTUALIZACION: se consulta Licitaciones.info con las credenciales de Ingesta/.env ===")
        archivo_env = comprobar_configuracion()
        if archivo_env is None:
            return 1
        nuevo = ejecutar_ingesta(secretos_locales(archivo_env))
        if nuevo is None:
            return 2
        fecha = lim.fecha_desde_nombre(nuevo)
        filas = len(lim.leer_filas(lim.ArchivoFuente(nuevo, fecha, "")))
        aviso(f"\nIngesta correcta: {relativa(nuevo)} ({filas} filas, extraido el {fecha:%d/%m/%Y %H:%M:%S}).")
        if filas == 0:
            aviso("La consulta no devolvio procesos; se consolidaran los CSV que ya estaban guardados.")
        modo = (
            f"Actualización: ingesta ejecutada y guardada en {nuevo.name} ({filas} filas, extraída el "
            f"{fecha:%d/%m/%Y %H:%M:%S}); consolidada con los demás CSV locales."
        )
    else:
        aviso("=== MODO DEMOSTRACION: se usan CSV ya guardados localmente; NO se consulta Licitaciones.info ===")
        modo = "Demostración: se usaron CSV ya guardados localmente; no se consultó Licitaciones.info."

    archivos = lim.listar_csvs(ENTRADA)
    if not archivos:
        aviso(f"ERROR: no hay CSV en {relativa(ENTRADA)}. Ejecuta con --actualizar para obtenerlos.")
        return 1
    aviso(f"\nCSV que se consolidan ({len(archivos)}), del mas antiguo al mas reciente:")
    for archivo in archivos:
        fecha = f"{archivo.fecha_extraccion:%d/%m/%Y %H:%M:%S}" if archivo.fecha_extraccion else "fecha no legible"
        aviso(f"  - {relativa(archivo.ruta)} (extraido el {fecha})")
    huellas_csv = {a.ruta.name: a.sha256 for a in archivos}

    if args.sin_llm:
        clasificador, aviso_llm = None, "Se ejecutó con --sin-llm: se usan solo las reglas."
    else:
        clasificador, aviso_llm = llm.crear_clasificador()
    if clasificador is not None:
        aviso(f"\nClasificacion con LLM: OpenAI, modelo {clasificador.modelo} (solo entidades sin regla de reparto).")
    else:
        aviso(f"\nClasificacion con LLM: no. {aviso_llm}")

    try:
        g = gen.generar(
            ENTRADA, args.maestros, args.salida, modo, clasificador, aviso_llm, actualizar_maestros=not args.sin_actualizar_maestros
        )
    except mae.ErrorMaestros as exc:
        aviso(f"ERROR: {exc}")
        return 1
    except FileExistsError as exc:
        aviso(f"ERROR: {exc.filename} ya existe; no se sobrescribe. Vuelve a ejecutar en unos segundos.")
        return 1

    m = g.resultado.metricas
    aviso(
        f"\nLimpieza: {m['filas_leidas']} filas leidas, {m['procesos_unicos']} procesos unicos, "
        f"{m['incluidos']} adjudicados, {m['excluidos_por_estado']} excluidos por estado."
    )
    if g.reparto.llm is not None:
        r = g.reparto.llm
        aviso(
            f"LLM: {len(r.decisiones)} de {r.enviados} procesos clasificados por el LLM "
            f"({r.tokens_entrada:,} tokens de entrada, {r.tokens_salida:,} de salida)."
        )
        for error in r.errores[:10]:
            aviso(f"  aviso LLM: {ocultar(error, [])}")
    aviso(
        f"Reparto: {len(g.reparto.filas)} filas nuevas, {len(g.reparto.ya_en_maestros)} ya estaban en los maestros, "
        f"{len(g.reparto.sin_consumo)} sin consumo de material, {len(g.reparto.fuera_de_zona)} fuera de zona, "
        f"{len(g.reparto.sin_clasificar)} sin clasificar."
    )

    problemas = verificar(g, huellas_csv)
    if problemas:
        revisar = g.carpeta.with_name(g.carpeta.name + "_REVISAR")
        g.carpeta.rename(revisar)
        aviso("Verificacion: FALLO. No uses estos archivos sin revisarlos:")
        for problema in problemas:
            aviso(f"  - {problema}")
        aviso(f"Carpeta marcada para revision: {relativa(revisar)}")
        return 3

    aviso("")
    for archivo in g.archivos_asesores:
        aviso(f"Archivo para {archivo.maestro.asesor}: {relativa(archivo.ruta)} ({archivo.filas} filas)")
    if not g.archivos_asesores:
        aviso("No hay obras nuevas para ningún asesor en esta ejecución.")
    act = g.actualizacion
    if act is not None and act.disponible:
        for a in act.archivos:
            aviso(f"Maestro actualizado para {a.maestro.asesor}: {relativa(a.ruta)}")
            for detalle in a.avisos:
                aviso(f"  aviso: {detalle}")
        for error in act.errores:
            aviso(f"  ERROR al actualizar un maestro: {error}")
    elif act is not None:
        aviso(f"Maestros actualizados: no. {act.motivo} Inserta las filas a mano desde los archivos Nuevas_.")
    aviso(f"Revision generada: {relativa(g.revision)}")
    aviso(f"Carpeta de salida: {relativa(g.carpeta)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
