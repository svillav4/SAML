"""Configuracion de la integracion con la API de Licitaciones.Info.

Lee las credenciales y parametros desde variables de entorno (archivo .env
en la raiz del proyecto). Ver .env.example para la lista completa.
"""
from __future__ import annotations

import json
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional

from dotenv import load_dotenv

load_dotenv()

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

# Subdominios disponibles por pais, segun la documentacion de la API:
# https://gusty-golf-290.notion.site/Documentacion_API-2179c7a942c78012a3fee5b20050814a
COUNTRY_SUBDOMAINS = {
    "colombia": "col",
    "ecuador": "ecu",
    "panama": "pan",
    "costa_rica": "cri",
    "chile": "chl",
    "peru": "per",
}


def _env_bool(nombre: str, default: bool = False) -> bool:
    valor = os.getenv(nombre)
    if valor is None:
        return default
    return valor.strip().lower() in ("1", "true", "si", "sí", "yes")


@dataclass
class Settings:
    email: str
    password: str
    country: str
    perfiles: Optional[str]
    page_limit: int
    keyword: Optional[str]
    estado: Optional[str]
    query_extra: Optional[str]
    filtrar_nuevos: bool
    campos_adicionales: Optional[str]

    @property
    def subdomain(self) -> str:
        pais = self.country.lower().strip()
        return COUNTRY_SUBDOMAINS.get(pais, pais)

    @property
    def base_url(self) -> str:
        return f"https://{self.subdomain}.licitaciones.info/api/client"

    def build_query(self) -> Optional[dict[str, Any]]:
        """Arma el dict que se envia como parametro "query" a la API,
        combinando los filtros individuales (keyword, estado) con
        cualquier filtro extra definido en LICITACIONES_QUERY_EXTRA.
        """
        query: dict[str, Any] = {}

        if self.keyword:
            query["Nombre"] = self.keyword
        if self.estado:
            query["estado_agrupado"] = self.estado

        if self.query_extra:
            try:
                extra = json.loads(self.query_extra)
            except json.JSONDecodeError as exc:
                raise RuntimeError(
                    f"LICITACIONES_QUERY_EXTRA no es un JSON valido: {exc}"
                ) from exc
            if not isinstance(extra, dict):
                raise RuntimeError("LICITACIONES_QUERY_EXTRA debe ser un objeto JSON (ej: {\"campo\": \"valor\"}).")
            query.update(extra)

        return query or None


def load_settings() -> Settings:
    email = os.getenv("LICITACIONES_EMAIL")
    password = os.getenv("LICITACIONES_PASSWORD")
    country = os.getenv("LICITACIONES_COUNTRY", "colombia")
    perfiles = os.getenv("LICITACIONES_PERFILES") or None
    page_limit = int(os.getenv("LICITACIONES_PAGE_LIMIT", "30"))
    keyword = os.getenv("LICITACIONES_KEYWORD") or None
    estado = os.getenv("LICITACIONES_ESTADO") or None
    query_extra = os.getenv("LICITACIONES_QUERY_EXTRA") or None
    filtrar_nuevos = _env_bool("LICITACIONES_FILTRAR_NUEVOS", default=False)
    campos_adicionales = os.getenv("LICITACIONES_CAMPOS_ADICIONALES", "fechas,documentos") or None

    if not email or not password:
        raise RuntimeError(
            "Debes definir LICITACIONES_EMAIL y LICITACIONES_PASSWORD "
            "en un archivo .env (copia .env.example como punto de partida)."
        )

    return Settings(
        email=email,
        password=password,
        country=country,
        perfiles=perfiles,
        page_limit=page_limit,
        keyword=keyword,
        estado=estado,
        query_extra=query_extra,
        filtrar_nuevos=filtrar_nuevos,
        campos_adicionales=campos_adicionales,
    )
