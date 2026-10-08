"""Punto de entrada para extraer licitaciones desde la API de Licitaciones.Info.

Flujo:
    1. LOGIN                -> obtiene el token JWT (api_client.login).
    2. PERFILES DE NEGOCIO  -> si no se definieron perfiles en el .env,
       se obtienen automaticamente (api_client.get_perfiles).
    3. CONSULTAR PROCESOS   -> trae las licitaciones de esos perfiles,
       recorriendo todas las paginas (api_client.consultar_todos_los_procesos).
    4. Guarda el resultado en output/licitaciones_<timestamp>.csv

Toda la configuracion (credenciales, pais, perfiles, palabra clave a buscar,
filtro de nuevos, campos adicionales, etc.) se define en el archivo .env de
la raiz del proyecto -- ver .env.example para la lista completa de opciones.

Documentacion de la API:
https://gusty-golf-290.notion.site/Documentacion_API-2179c7a942c78012a3fee5b20050814a

Uso:
    python extract.py
"""
from __future__ import annotations

import csv
import logging
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional

from api_client import LicitacionesAPIClient, LicitacionesAPIError
from config import OUTPUT_DIR, load_settings

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger("extract")


def resolver_perfiles(client: LicitacionesAPIClient, perfiles_env: Optional[str]) -> str:
    """Usa los perfiles definidos en el .env, o los obtiene automaticamente."""
    if perfiles_env:
        return perfiles_env

    perfiles = client.get_perfiles()
    if not perfiles:
        raise LicitacionesAPIError(
            "El usuario no tiene perfiles de negocio configurados en Licitaciones.Info."
        )

    ids = []
    for perfil in perfiles:
        if isinstance(perfil, dict):
            ids.append(str(perfil.get("id") or perfil.get("id_perfil") or perfil))
        else:
            ids.append(str(perfil))

    logger.info("Perfiles obtenidos automaticamente: %s", ", ".join(ids))
    return ",".join(ids)


def guardar_csv(procesos: list) -> Path:
    """Guarda las licitaciones encontradas en un CSV con marca de tiempo."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    destino = OUTPUT_DIR / f"licitaciones_{timestamp}.csv"

    if not procesos:
        destino.write_text("Sin resultados\n", encoding="utf-8")
        return destino

    columnas: list[str] = []
    for proceso in procesos:
        if not isinstance(proceso, dict):
            continue
        for campo in proceso.keys():
            if campo not in columnas:
                columnas.append(campo)

    with destino.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=columnas)
        writer.writeheader()
        for proceso in procesos:
            if isinstance(proceso, dict):
                writer.writerow({k: proceso.get(k, "") for k in columnas})

    return destino


def extract_data() -> Path:
    """Orquesta LOGIN -> PERFILES -> CONSULTAR PROCESOS y guarda el CSV resultante."""
    settings = load_settings()

    client = LicitacionesAPIClient(
        base_url=settings.base_url,
        email=settings.email,
        password=settings.password,
    )

    client.login()

    perfiles = resolver_perfiles(client, settings.perfiles)

    query = settings.build_query()

    procesos = client.consultar_todos_los_procesos(
        perfiles=perfiles,
        limit=settings.page_limit,
        filtrar_nuevos=settings.filtrar_nuevos or None,
        query=query,
        campos_adicionales=settings.campos_adicionales,
    )

    logger.info("Se encontraron %d licitaciones.", len(procesos))

    destino = guardar_csv(procesos)
    logger.info("Resultado guardado en %s", destino)
    return destino


if __name__ == "__main__":
    try:
        extract_data()
    except LicitacionesAPIError as exc:
        logger.error("Error consultando la API: %s", exc)
        sys.exit(1)
