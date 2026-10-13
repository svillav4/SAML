# n8n local en Docker: proceso SAML de licitaciones adjudicadas

Instancia **local** de n8n (2.42.7) con Python 3.14.7 para ejecutar el proceso del repositorio (`Transformacion/ejecutar_proceso.py`) desde n8n con un disparador manual.

- No usa la instancia remota del equipo ni Supabase.
- No necesita API key de n8n.
- No guarda credenciales en n8n.

| Archivo | Propósito |
|---|---|
| `Dockerfile` | Imagen `saml-n8n:2.42.7-py3.14.7`: n8n oficial más Python y las dependencias del proyecto |
| `Dockerfile.dockerignore` | El contexto de build solo incluye los dos `requirements.txt`; nunca `.env`, CSV ni Excel |
| `docker-compose.yml` | Proyecto Compose `saml-n8n`: contenedor, puerto local, volumen y montajes |
| `flujo_local_licitaciones.json` | Workflow **principal** (id `samlLicitaciones`): proceso completo, de la consulta a Licitaciones.info al Excel. Inactivo y sin credenciales |
| `flujo_demo_csv_locales.json` | Workflow de **demo** (id `samlDemoLocal`): Excel con los CSV ya guardados, sin consultar el portal |

## Iniciar y detener

Desde la raíz del repositorio, siempre con el contexto `desktop-linux`. El contexto `default` apunta a `/var/run/docker.sock`, que no existe en este equipo.

| Acción | Comando |
|---|---|
| Construir y levantar (la primera vez, o tras cambiar el Dockerfile) | `docker --context desktop-linux compose -f n8n/docker-compose.yml up -d --build` |
| Detener | `docker --context desktop-linux compose -f n8n/docker-compose.yml stop` |
| Volver a iniciar | `docker --context desktop-linux compose -f n8n/docker-compose.yml start` |
| Ver estado y registros | `docker --context desktop-linux compose -f n8n/docker-compose.yml ps` y `... logs -f` |

URL: **http://localhost:5678**. El puerto solo se publica en `127.0.0.1`.

La configuración, la cuenta de propietario, la clave de cifrado y los workflows viven en el volumen Docker `saml_n8n_data`. **No uses `down -v`**: borraría ese volumen. `stop` y `start` lo conservan.

## Primer acceso

n8n pide crear una **cuenta de propietario local** ("Set up owner account"). Esa cuenta solo existe en esta instancia, guardada en el volumen. Créala tú con tus propios datos; no se guardan en el repositorio.

## Importar los workflows

Los dos ya quedaron importados en tu proyecto personal. Si hubiera que volver a importarlos, hay dos formas:

- **Interfaz:** *Workflows* → *Create Workflow* → menú `⋯` → *Import from File…* → elige el JSON → *Save*.
- **CLI** (reemplaza el workflow con el mismo id). Recarga la página de n8n después:

  ```bash
  docker --context desktop-linux exec saml-n8n n8n import:workflow --input=/proyecto/n8n/flujo_local_licitaciones.json
  docker --context desktop-linux exec saml-n8n n8n import:workflow --input=/proyecto/n8n/flujo_demo_csv_locales.json
  ```

## Ejecutar el proceso completo

1. Abre **"SAML - Proceso completo: Licitaciones.info a Excel intermedio"**.
2. Pulsa **Execute workflow**, el botón de la parte inferior del lienzo, o el ▶ del disparador **"Ejecutar proceso completo"**. Hay un solo disparador, así que no hay que elegir.
3. Espera a que los tres nodos queden en verde y abre **"Resultado: Excel generado"**.

| Nodo | Qué hace |
|---|---|
| Ejecutar proceso completo | Disparador manual |
| Ingesta (Licitaciones.info) y luego Excel | Ejecuta `cd /proyecto && /opt/saml-venv/bin/python Transformacion/ejecutar_proceso.py --actualizar` (detalle abajo) |
| Resultado: Excel generado | Devuelve los campos de abajo, y **falla** si no hubo ingesta correcta antes del Excel, si el código de salida no es 0 o si la verificación no es OK |

El nodo de ingesta hace, en este orden:
1. Comprueba `Ingesta/.env`.
2. Ejecuta `Ingesta/extract.py` y guarda el CSV nuevo en `Ingesta/output/`.
3. Solo si la ingesta termina bien: limpia, se queda con los procesos `Adjudicado`, genera el Excel nuevo en `datos/intermedio/` y lo verifica.

Campos que devuelve el nodo de resultado:
- `ingesta_csv_nuevo`
- `excel_nombre` y `excel_ruta_en_el_repositorio`
- `verificacion`
- `csv_consolidados`

Si faltan credenciales, el error dice qué variable completar en `Ingesta/.env` (por ejemplo `LICITACIONES_PASSWORD`), sin mostrar valores. Si la ingesta falla (credenciales rechazadas, red, API), la ejecución queda en rojo y **no** se genera Excel.

Con `LICITACIONES_FILTRAR_NUEVOS=true`, una consulta puede no devolver procesos nuevos. En ese caso el Excel consolida los CSV ya guardados y el nodo lo indica.

## Demo sin consultar el portal

Abre **"SAML - Demo con CSV existentes (sin consultar el portal)"** y pulsa **Execute workflow**. Usa solo los CSV de `Ingesta/output/`.

> **No ejecutes estos workflows con `n8n execute` desde la CLI** salvo para pruebas controladas: el principal consulta el portal.

## Cómo accede el contenedor al proyecto

| Ruta en el contenedor | Origen en el equipo | Acceso |
|---|---|---|
| `/proyecto/Ingesta` | `Ingesta/` (incluye `.env`) | Solo lectura |
| `/proyecto/Transformacion` | `Transformacion/` | Solo lectura |
| `/proyecto/n8n` | `n8n/` | Solo lectura |
| `/proyecto/Ingesta/output` | `Ingesta/output/` | Escritura (CSV de la ingesta) |
| `/proyecto/datos/intermedio` | `datos/intermedio/` | Escritura (Excel generados) |
| `/opt/saml-venv` | Dentro de la imagen | Python 3.14.7 con `requests`, `python-dotenv` y `openpyxl` |

**Credenciales.** `Ingesta/.env` solo lo lee Python, y solo en el modo `--actualizar`. No se carga como variables de entorno de n8n, y `N8N_BLOCK_ENV_ACCESS_IN_NODE=true` impide leer el entorno desde expresiones. La salida del proceso oculta correos, contraseñas y JWT.

## Seguridad de habilitar Execute Command

n8n 2.x trae **desactivados** por defecto los nodos Execute Command y Local File Trigger. Aquí se habilitan con `NODES_EXCLUDE=[]` porque el workflow necesita lanzar Python. Efectos:

- **Cualquier persona con acceso al editor puede ejecutar comandos de shell** dentro del contenedor, como el usuario `node`. Eso incluye:
  - leer todo lo montado, incluido `Ingesta/.env`, y los datos de n8n (base de datos y clave de cifrado);
  - escribir en `Ingesta/output/` y `datos/intermedio/`;
  - hacer conexiones de red salientes.
- **Mitigaciones de esta instancia:**
  - el puerto solo escucha en `127.0.0.1`;
  - el editor exige la cuenta de propietario;
  - el código y el `.env` están montados en solo lectura;
  - no se monta el socket de Docker ni otras carpetas del equipo.
- **Reglas de uso:** no publiques el puerto en la red, no importes workflows de terceros en esta instancia y detén el contenedor cuando no lo uses. Esta configuración es solo para demostración local, no para producción.

## Notas

- `N8N_SECURE_COOKIE=false` está aplicado solo a esta instancia. Safari no trata `http://localhost` como origen seguro y, sin este ajuste, muestra el aviso de cookie segura y no permite iniciar sesión. Es aceptable porque n8n solo escucha en `127.0.0.1`. **No** lo uses si la instancia se publica en una red.

- En los registros aparece "Failed to start Python task runner…". Se refiere al nodo Code en **Python** de n8n, que este workflow no usa (su nodo Code es JavaScript). El Python del proyecto se ejecuta por Execute Command con `/opt/saml-venv`.
- La imagen oficial de n8n 2.x es una "Docker Hardened Image" sin `apk` ni Python. Por eso el `Dockerfile` copia el intérprete oficial de `python:3.14.7-alpine3.24`, que tiene la misma base Alpine 3.24 y musl, junto con las 5 bibliotecas del sistema que faltaban.
