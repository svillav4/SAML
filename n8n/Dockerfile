# n8n local con Python para ejecutar el proceso SAML (Ingesta + Transformacion).
#
# La imagen oficial n8nio/n8n 2.x es una "Docker Hardened Image" (Alpine 3.24)
# sin gestor de paquetes ni Python. Por eso se copia el interprete oficial de
# python:<version>-alpine3.24 (misma version que el .venv local del proyecto y
# misma base Alpine/musl) junto con las bibliotecas que la imagen de n8n no trae,
# y se crea un entorno virtual con las dependencias del proyecto.
#
# Construir y levantar: ver n8n/README.md (docker compose -f n8n/docker-compose.yml ...)

ARG N8N_VERSION=2.42.7
ARG PYTHON_VERSION=3.14.7
ARG ALPINE_VERSION=3.24

FROM python:${PYTHON_VERSION}-alpine${ALPINE_VERSION} AS python

FROM n8nio/n8n:${N8N_VERSION}

USER root

# Interprete y biblioteca estandar de Python (se instalan en /usr/local en la imagen oficial).
COPY --from=python /usr/local/bin/python3.14 /usr/local/bin/python3.14
COPY --from=python /usr/local/lib/libpython3.14.so.1.0 /usr/local/lib/libpython3.14.so.1.0
COPY --from=python /usr/local/lib/python3.14 /usr/local/lib/python3.14

# Bibliotecas del sistema que usan modulos de Python y no existen en la imagen de n8n.
COPY --from=python /usr/lib/libffi.so.8* /usr/lib/libreadline.so.8* /usr/lib/libpanelw.so.6* \
     /usr/lib/libsqlite3.so.0* /usr/lib/libuuid.so.1* /usr/local/lib/

# Solo las dependencias declaradas del proyecto (el contexto de build excluye todo lo demas,
# ver Dockerfile.dockerignore): nunca se copian .env, CSV ni Excel a la imagen.
COPY Ingesta/requirements.txt /opt/saml/requirements-ingesta.txt
COPY Transformacion/requirements.txt /opt/saml/requirements-transformacion.txt

RUN ln -s python3.14 /usr/local/bin/python3 \
 && python3 -m venv /opt/saml-venv \
 && /opt/saml-venv/bin/pip install --no-cache-dir --disable-pip-version-check \
      -r /opt/saml/requirements-ingesta.txt -r /opt/saml/requirements-transformacion.txt \
 && /opt/saml-venv/bin/python -c "import ssl, requests, dotenv, openpyxl"

# Los codigos del proyecto se montan en solo lectura: no escribir __pycache__.
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

USER node
