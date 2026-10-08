"""Cliente para consumir la API de Licitaciones.Info.

La API es RESTful y expone tres servicios principales:
    - LOGIN                          POST /auth-token
    - PERFILES DE NEGOCIO            GET  /perfiles/consultar
    - PROCESOS DE CONTRATACION:
        - CONSULTAR PROCESOS                     POST /contratos/consultar
        - CONSULTAR PROCESOS POR SERVICIO ESPECIAL
                                                  POST /contratos/servicios-especiales/consultar

Autenticacion: BearerAuth. El LOGIN devuelve un token JWT que debe enviarse
como header "Authorization: Bearer <token>" en el resto de peticiones.
"""
from __future__ import annotations

import json
import logging
from typing import Any, Optional

import requests

logger = logging.getLogger(__name__)


class LicitacionesAPIError(Exception):
    """Error generico al comunicarse con la API de Licitaciones.Info."""


class LicitacionesAuthError(LicitacionesAPIError):
    """Credenciales invalidas o token no autorizado (401)."""


class LicitacionesAPIClient:

    def __init__(self, base_url: str, email: str, password: str, timeout: int = 30):
        self.base_url = base_url.rstrip("/")
        self.email = email
        self.password = password
        self.timeout = timeout
        self._token: Optional[str] = None
        self._token_expires_at: Optional[str] = None
        self._session = requests.Session()
        # Algunos WAF/Cloudflare bloquean el User-Agent por defecto de requests
        # (python-requests/x.y) en ciertos endpoints. Usamos headers de un
        # navegador normal para reducir la probabilidad de que nos bloqueen.
        self._session.headers.update(
            {
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                    "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
                ),
                "Accept": "application/json, text/plain, */*",
                "Accept-Language": "es-CO,es;q=0.9,en;q=0.8",
            }
        )

    # ------------------------------------------------------------------
    # LOGIN
    # def login: obtiene un token JWT de la API usando las credenciales de email y password
    # ------------------------------------------------------------------
    def login(self) -> str:
        url = f"{self.base_url}/auth-token"                         # rellenar el URL para el endpoint de login
        payload = {"email": self.email, "password": self.password}  # crear el payload con las credenciales de email y password

        response = self._session.post(url, json=payload, timeout=self.timeout)  # Recibir la respuesta de la API al hacer la solicitud POST al endpoint de login
        if response.status_code == 401:
            raise LicitacionesAuthError(
                f"Acceso no autorizado (401) al hacer LOGIN. Respuesta: {response.text[:300]}"
            )

        data = self._safe_json(response)    # Usar el método _safe_json (que es nuestro metodo) para obtener la respuesta en formato JSON

        # La API de Licitaciones.Info NO devuelve el JWT directo en "token",
        # sino anidado junto con su fecha de expiracion:
        #   {"success": true, "token": {"accessToken": "<JWT>", "expires_at": "..."}, "user": {...}}
        token_field = data.get("token") if isinstance(data, dict) else None

        if isinstance(token_field, dict):
            token = token_field.get("accessToken")
            expires_at = token_field.get("expires_at")
        elif isinstance(token_field, str):
            # Por si en el futuro la API cambia y devuelve el JWT directo.
            token = token_field
            expires_at = None
        else:
            token = None
            expires_at = None

        if not token:
            raise LicitacionesAuthError(
                f"LOGIN sin token en la respuesta (status {response.status_code}): {data}"
            )

        self._token = token                     # Guardar el token en la instancia del cliente para usarlo en futuras solicitudes
        self._token_expires_at = expires_at
        if expires_at:
            logger.info(f"Login exitoso en {url} (token valido hasta {expires_at})")
        else:
            logger.info(f"Login exitoso en {url} (token de {len(token)} caracteres)")
        return token

    @property
    def token(self) -> str:
        if not self._token:
            raise LicitacionesAuthError("No hay token activo: llama a login() primero.")
        return self._token

    def _auth_headers(self) -> dict:
        return {"Authorization": f"Bearer {self.token}"}

    # ------------------------------------------------------------------
    # PERFILES DE NEGOCIO
    # def get_perfiles: obtiene los perfiles de negocio del usuario autenticado
    # ------------------------------------------------------------------
    def get_perfiles(self) -> list:
        url = f"{self.base_url}/perfiles/consultar"
        response = self._session.get(url, headers=self._auth_headers(), timeout=self.timeout)

        if response.status_code == 204:
            return []
        if response.status_code == 401:
            raise LicitacionesAuthError(
                f"Token invalido o vencido al consultar perfiles. Respuesta: {response.text[:300]}"
            )

        data = self._safe_json(response)
        if isinstance(data, list):
            return data
        if isinstance(data, dict):
            return data.get("data") or data.get("perfiles") or []
        return []

    # ------------------------------------------------------------------
    # CONSULTAR PROCESOS (licitaciones de los ultimos 8 meses)
    # ------------------------------------------------------------------
    def consultar_procesos(
        self,
        perfiles: str,
        page: int = 1,
        limit: int = 30,
        filtrar_nuevos: Optional[bool] = None,
        query: Optional[dict] = None,
        ascending: Optional[int] = None,
        campos_adicionales: Optional[str] = None,
    ) -> dict:
        """POST /contratos/consultar -> procesos de contratacion (licitaciones)."""
        url = f"{self.base_url}/contratos/consultar"

        body: dict[str, Any] = {
            "perfiles": perfiles,
            "page": page,
            "limit": min(limit, 30),  # la API limita a 30 registros por pagina
        }
        if filtrar_nuevos is not None:
            body["filtrar_nuevos"] = int(filtrar_nuevos)
        if query is not None:
            body["query"] = json.dumps(query, ensure_ascii=False)
        if ascending is not None:
            body["ascending"] = ascending
        if campos_adicionales is not None:
            body["campos_adicionales"] = campos_adicionales

        response = self._session.post(
            url, json=body, headers=self._auth_headers(), timeout=self.timeout
        )

        if response.status_code == 204:
            return {"procesos": []}
        if response.status_code == 401:
            raise LicitacionesAuthError(
                f"Token invalido o vencido al consultar procesos. Respuesta: {response.text[:300]}"
            )

        return self._safe_json(response)

    def consultar_todos_los_procesos(self, perfiles: str, **kwargs) -> list:
        """Recorre todas las paginas de CONSULTAR PROCESOS y arma una sola lista."""
        limit = kwargs.pop("limit", 30)
        page = kwargs.pop("page", 1)
        procesos: list = []

        while True:
            data = self.consultar_procesos(perfiles, page=page, limit=limit, **kwargs)
            lote = self._extract_procesos(data)
            if not lote:
                break
            procesos.extend(lote)
            if len(lote) < limit:
                break
            page += 1

        return procesos

    # ------------------------------------------------------------------
    # CONSULTAR PROCESOS POR SERVICIO ESPECIAL
    # ------------------------------------------------------------------
    def consultar_procesos_servicio_especial(
        self,
        servicio_especial: Optional[int] = None,
        filtrar_nuevos: Optional[bool] = None,
        page: int = 1,
        limit: int = 30,
        query: Optional[dict] = None,
        ascending: Optional[int] = None,
    ) -> dict:
        """POST /contratos/servicios-especiales/consultar."""
        url = f"{self.base_url}/contratos/servicios-especiales/consultar"

        body: dict[str, Any] = {"page": page, "limit": min(limit, 30)}
        if servicio_especial is not None:
            body["servicio_especial"] = servicio_especial
        if filtrar_nuevos is not None:
            body["filtrar_nuevos"] = int(filtrar_nuevos)
        if query is not None:
            body["query"] = json.dumps(query, ensure_ascii=False)
        if ascending is not None:
            body["ascending"] = ascending

        response = self._session.post(
            url, json=body, headers=self._auth_headers(), timeout=self.timeout
        )

        if response.status_code == 204:
            return {"procesos": []}
        if response.status_code == 401:
            raise LicitacionesAuthError(
                "Token invalido o vencido al consultar procesos por servicio especial. "
                f"Respuesta: {response.text[:300]}"
            )

        return self._safe_json(response)

    # ------------------------------------------------------------------
    # Utilidades internas
    # ------------------------------------------------------------------
    @staticmethod
    def _safe_json(response: requests.Response) -> Any:
        try:
            return response.json()
        except ValueError as exc:
            cuerpo = response.text or ""
            if "cloudflare" in cuerpo.lower() or "attention required" in cuerpo.lower():
                raise LicitacionesAPIError(
                    f"La peticion fue bloqueada por Cloudflare/WAF (status {response.status_code}) "
                    "antes de llegar a la API real, no es un error de tus credenciales ni de tu "
                    "codigo. Intenta de nuevo en unos minutos; si persiste, contacta a "
                    "Licitaciones.Info para pedir que tu IP/servidor quede en lista blanca."
                ) from exc
            raise LicitacionesAPIError(
                f"Respuesta no-JSON de la API (status {response.status_code}): {cuerpo[:500]}"
            ) from exc

    @staticmethod
    def _extract_procesos(data: Any) -> list:
        if isinstance(data, list):
            return data
        if isinstance(data, dict):
            for key in ("procesos", "data", "results", "items"):
                valor = data.get(key)
                if isinstance(valor, list):
                    return valor
        return []
