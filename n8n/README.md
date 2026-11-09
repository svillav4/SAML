# n8n local en Docker: proceso SAML de licitaciones adjudicadas

Instancia **local** de n8n (2.42.7) con Python 3.14.7 para ejecutar el proceso del repositorio (`Transformacion/ejecutar_proceso.py`) desde n8n con un disparador manual.

- No usa la instancia remota del equipo ni Supabase.
- No necesita API key de n8n.
- No guarda credenciales en n8n.

| Archivo | Propósito |
|---|---|
| `Dockerfile` | Imagen `saml-n8n:2.42.7-py3.14.7`: n8n oficial más Python y las dependencias del proyecto |
| `Dockerfile.dockerignore` | El contexto de build solo incluye los dos `requirements.txt`; nunca `.env`, CSV ni Excel (ni los maestros) |
| `docker-compose.yml` | Proyecto Compose `saml-n8n`: contenedor, puerto local, volumen y montajes |
| `flujo_local_licitaciones.json` | Workflow **principal** (id `samlLicitaciones`): proceso completo, de la consulta a Licitaciones.info a los archivos por asesor. Inactivo y sin credenciales |
| `flujo_demo_csv_locales.json` | Workflow de **demo** (id `samlDemoLocal`): archivos por asesor con los CSV ya guardados, sin consultar el portal |

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

0. Copia los 4 Excel maestros de GRECO en `datos/maestros/` (se montan en solo lectura).
1. Abre **"SAML - Proceso completo: Licitaciones.info a archivos por asesor"**.
2. Pulsa **Execute workflow**, el botón de la parte inferior del lienzo, o el ▶ del disparador **"Ejecutar proceso completo"**. Hay un solo disparador, así que no hay que elegir.
3. Espera a que los tres nodos queden en verde y abre **"Resultado: archivos por asesor"**.

| Nodo | Qué hace |
|---|---|
| Ejecutar proceso completo | Disparador manual |
| Ingesta (Licitaciones.info) y luego archivos por asesor | Ejecuta `cd /proyecto && /opt/saml-venv/bin/python Transformacion/ejecutar_proceso.py --actualizar` (detalle abajo) |
| Resultado: archivos por asesor | Devuelve los campos de abajo, y **falla** si no hubo ingesta correcta antes de los archivos, si no se generó la revisión, si el código de salida no es 0 o si la verificación no es OK |

El nodo de ingesta hace, en este orden:
1. Comprueba `Ingesta/.env`.
2. Ejecuta `Ingesta/extract.py` y guarda el CSV nuevo en `Ingesta/output/`.
3. Solo si la ingesta termina bien: limpia, se queda con los procesos `Adjudicado`, los reparte entre los maestros de `datos/maestros/`, crea `datos/salida/SAML_<fecha>/` con un Excel por asesor y el de revisión, y los verifica.

En el contenedor no hay Excel, así que la ejecución de n8n **no** genera las copias `<maestro>_actualizado_<fecha>.xlsx`. Para tenerlas, después de ejecutar el flujo, en el equipo Windows con Excel haz **doble clic en `Actualizar_maestros.bat`** (raíz del repositorio): toma la carpeta más reciente de `datos/salida/` y genera ahí las copias de los maestros con las obras insertadas (comando equivalente: `.venv\Scripts\python Transformacion\actualizar_maestros.py [carpeta]`). No repite una copia que ya exista y no actualiza un maestro que haya cambiado desde que se generó esa salida (para no duplicar obras).

Campos que devuelve el nodo de resultado:
- `ingesta_csv_nuevo`
- `reparto` (filas nuevas, ya en maestros y sin clasificar)
- `archivos_por_asesor`: asesor, archivo, ruta y filas de cada Excel de obras nuevas
- `revision` y `carpeta_salida`
- `verificacion`
- `csv_consolidados`

Si faltan credenciales, el error dice qué variable completar en `Ingesta/.env` (por ejemplo `LICITACIONES_PASSWORD`), sin mostrar valores. Si la ingesta falla (credenciales rechazadas, red, API), la ejecución queda en rojo y **no** se generan archivos. Si faltan los maestros, falla antes de consultar el portal.

Con `LICITACIONES_FILTRAR_NUEVOS=true`, una consulta puede no devolver procesos nuevos. En ese caso se consolidan los CSV ya guardados y el nodo lo indica.

## Demo sin consultar el portal

Abre **"SAML - Demo con CSV existentes (sin consultar el portal)"** y pulsa **Execute workflow**. Usa solo los CSV de `Ingesta/output/`.

> **No ejecutes estos workflows con `n8n execute` desde la CLI** salvo para pruebas controladas: el principal consulta el portal. Con la instancia corriendo, la CLI necesita otro puerto interno: `docker --context desktop-linux exec -e N8N_RUNNERS_BROKER_PORT=5690 saml-n8n n8n execute --id samlDemoLocal`. En Windows con Git Bash, antepón `MSYS_NO_PATHCONV=1` para que no cambie las rutas `/proyecto/...`.

**Probado el 9-nov-2026** en Docker Desktop (Windows): ambos flujos terminan en verde; el completo inicia sesión en Licitaciones.info, trae las adjudicadas, clasifica con el LLM y genera los archivos por asesor y la revisión (unos 6 minutos).

## Cómo accede el contenedor al proyecto

| Ruta en el contenedor | Origen en el equipo | Acceso |
|---|---|---|
| `/proyecto/Ingesta` | `Ingesta/` (incluye `.env`) | Solo lectura |
| `/proyecto/Transformacion` | `Transformacion/` (incluye `Transformacion/.env` con la clave de OpenAI, si existe) | Solo lectura |
| `/proyecto/n8n` | `n8n/` | Solo lectura |
| `/proyecto/Ingesta/output` | `Ingesta/output/` | Escritura (CSV de la ingesta) |
| `/proyecto/datos/maestros` | `datos/maestros/` | Solo lectura (los 4 Excel maestros de GRECO) |
| `/proyecto/datos/salida` | `datos/salida/` | Escritura (archivos por asesor y revisión) |
| `/opt/saml-venv` | Dentro de la imagen | Python 3.14.7 con `requests`, `python-dotenv`, `openpyxl` y `openai` |

**Credenciales.** `Ingesta/.env` solo lo lee Python, y solo en el modo `--actualizar`. No se carga como variables de entorno de n8n, y `N8N_BLOCK_ENV_ACCESS_IN_NODE=true` impide leer el entorno desde expresiones. La salida del proceso oculta correos, contraseñas y JWT.

## Seguridad de habilitar Execute Command

n8n 2.x trae **desactivados** por defecto los nodos Execute Command y Local File Trigger. Aquí se habilitan con `NODES_EXCLUDE=[]` porque el workflow necesita lanzar Python. Efectos:

- **Cualquier persona con acceso al editor puede ejecutar comandos de shell** dentro del contenedor, como el usuario `node`. Eso incluye:
  - leer todo lo montado, incluido `Ingesta/.env`, y los datos de n8n (base de datos y clave de cifrado);
  - escribir en `Ingesta/output/` y `datos/salida/`;
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
