# Transformación: Excel intermedio de licitaciones adjudicadas

Etapa local posterior a la ingesta. Lee los CSV que genera `Ingesta/extract.py`, los limpia y produce un **Excel intermedio** para revisar y presentar las licitaciones adjudicadas.

> **Versión de demostración y validación.** No es el producto final de GRECO ni reemplaza la integración con "Obras V6", Supabase o n8n, que se definirá con el Product Owner. Las columnas GRECO son provisionales: se tomaron de una captura de "Obras V6" y deben confirmarse con el `.xlsx` original.

## Requisitos

Usa el entorno virtual `.venv/` de la raíz del proyecto. Si todavía no existe:

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r Ingesta/requirements.txt -r Transformacion/requirements.txt
```

## Proceso completo con un solo comando (recomendado)

`ejecutar_proceso.py` coordina las etapas existentes: la ingesta (opcional), la limpieza, el Excel intermedio y una verificación final. Se ejecuta desde la raíz del repositorio.

| Comando | Qué hace |
|---|---|
| `.venv/bin/python Transformacion/ejecutar_proceso.py` | **Demostración** (modo por defecto). Usa solo los CSV ya guardados en `Ingesta/output/`, **sin consultar Licitaciones.info** y sin necesitar credenciales. Muestra qué CSV usa y cuándo se extrajeron, y genera un Excel nuevo |
| `.venv/bin/python Transformacion/ejecutar_proceso.py --actualizar` | **Actualización.** Primero comprueba que `Ingesta/.env` tenga credenciales y `LICITACIONES_ESTADO=Adjudicado`. Luego ejecuta `Ingesta/extract.py` sin modificarlo y, **solo si termina bien y deja exactamente un CSV nuevo**, genera el Excel con todos los CSV locales. Si la ingesta falla, se detiene y **no** genera Excel |
| `--salida DIR` | Cambia la carpeta del Excel (por defecto `datos/intermedio`) |

Después de generar el Excel, el proceso verifica:
- que todos los procesos de Datos tienen estado `Adjudicado` y que no hay `idContrato` repetidos;
- que las columnas sin fuente siguen vacías: NIT, Localidad, Latitud, Longitud, Nombre de la Obra, Contrato No, Fecha Adjudicación e Integrantes del consorcio;
- que no hay valores con forma de JWT;
- que los CSV de origen no cambiaron.

Si algo falla, el Excel se renombra con el sufijo `_REVISAR`.

La salida de la ingesta se muestra con correos, contraseñas, tokens Bearer y JWT ocultos.

Códigos de salida: `0` correcto · `1` faltan CSV o configuración · `2` falló la ingesta · `3` falló la verificación.

## Etapas por separado

1. *(Opcional)* Ejecuta la ingesta para traer datos nuevos. Consulta la configuración del `.env` en `Ingesta/.env.example`:

   ```bash
   cd Ingesta && ../.venv/bin/python extract.py && cd ..
   ```

2. Desde la raíz del proyecto, genera el Excel intermedio:

   ```bash
   .venv/bin/python Transformacion/generar_excel_intermedio.py
   ```

   El resultado queda en `datos/intermedio/licitaciones_intermedio_<AAAAMMDD_HHMMSS>.xlsx`. La carpeta `datos/` está ignorada por Git. Cada ejecución crea un archivo nuevo: nunca se sobrescribe uno existente ni se modifican los CSV de `Ingesta/output/`.

Opciones: `--entrada <carpeta con los CSV>` y `--salida <carpeta destino>`.

## Qué hace

| Paso | Regla |
|---|---|
| Entrada | Todos los `licitaciones_*.csv` presentes en `Ingesta/output/`, que solo se leen |
| Consolidación | Una fila por `idContrato`. Si un proceso aparece en varios CSV, se usa la versión del más reciente (según la fecha del nombre del archivo). Solo se consolidan los archivos que estén en la carpeta |
| Filtro | Solo procesos con `estado_agrupado = Adjudicado`. Los demás se excluyen y quedan registrados en la hoja Incidencias |
| Textos | Se quitan espacios sobrantes y saltos de línea; no se cambian mayúsculas ni contenido |
| Valor | La API lo entrega en pesos. `Valor Millones = Valor / 1.000.000`. Un valor de 0 se deja vacío y se registra |
| Fechas | `FechaPublicacion` y `FechaVencimiento` pasan a fecha de Excel. **No** se usan como fecha de adjudicación |
| Ganadores | Una fila por ganador en la hoja Ganadores, relacionada por `idContrato`. Las entradas idénticas repetidas se muestran una vez, indicando cuántas veces aparecían. Los proponentes no se tratan como ganadores |
| Exclusiones | `jwt_contratista_id`, `pivot` y los campos de estado del usuario (`leido`, `en_favorito`, etc.) no se exportan al Excel ni a los mensajes |

## Hojas del Excel

| Hoja | Contenido |
|---|---|
| **Datos** | Una fila por licitación: primero las 20 columnas provisionales de Obras V6 y después las de trazabilidad (`idContrato`, `CodigoProceso`, fuente, enlace y datos de origen). El color del encabezado indica el estado del mapeo, y cada encabezado tiene una nota con su origen y su regla |
| **Ganadores** | Una fila por ganador distinto de cada proceso |
| **Incidencias** | Columnas sin fuente, mapeos provisionales, campos vacíos, problemas de calidad, duplicados, exclusiones y casos para revisión manual |
| **Resumen** | Archivos consolidados (con su huella SHA-256), fecha de extracción, conteos, confirmación del filtro `Adjudicado`, unidad de Valor, incidencias por tipo, leyenda de colores y mapeo completo de columnas |

### Estado del mapeo de las columnas GRECO

| Estado | Columnas |
|---|---|
| Mapeado | `Entidad` (← `EntidadContratante`), `Valor Millones` (← `Valor` / 1.000.000) |
| Provisional | `Empresa` (← nombre de los ganadores) |
| Sin fuente en la API (vacías) | `Localidad`, `Latitud`, `Longitud`, `NIT`, `Nombre de la Obra`, `Contrato No`, `Fecha Adjudicación`. También `Direccion` (auxiliar) y, en Ganadores, `Integrantes del consorcio` |
| Manuales de GRECO (vacías) | `Asesor`, `Despacho/Disposición`, `Contactada`, `Notas` |
| Calculadas en Obras V6 (vacías) | `Avance %`, `Meses por Terminar`, `Valor por Ejecutar`, `Potencial Infraestructura`, `Meta 25% Material`, `MS Greco` |

Por qué algunas quedan vacías:

- **NIT:** no viene en la API. `id_contratista_verificado` es un identificador interno de licitaciones.info (de 2 a 6 dígitos), no un NIT.
- **Fecha Adjudicación:** el cronograma solo trae fechas *programadas*, a veces varias por proceso. Se muestran como referencia en `Adjudicación programada (cronograma)`, pero no se toman como fecha de adjudicación.
- **Nombre de la Obra:** el candidato es `Objeto (origen)`. `Nombre (origen)` es la modalidad más el código del proceso. Falta confirmarlo con Obras V6.
- **Contrato No:** la API no trae número de contrato. `CodigoProceso` es el código del proceso de selección, no el del contrato.

## Pruebas

Usan datos **sintéticos** en carpetas temporales:

```bash
.venv/bin/python -m unittest discover -s Transformacion/tests -v
```

## Archivos

| Archivo | Propósito |
|---|---|
| `ejecutar_proceso.py` | Punto de entrada único: modo demostración o `--actualizar`, más la verificación final |
| `limpieza_licitaciones.py` | Reglas de consolidación, filtro, normalización e incidencias. No escribe archivos |
| `generar_excel_intermedio.py` | Punto de entrada: ejecuta la limpieza y escribe el Excel con `openpyxl` |
| `perfil_csv.py` | Perfil de solo lectura de un CSV de la ingesta (columnas, tipos, completitud), sin imprimir los registros |
| `tests/test_limpieza.py` | Pruebas de las reglas con datos sintéticos |
| `tests/test_ejecutar_proceso.py` | Pruebas del punto de entrada, con una ingesta simulada (no consultan el portal) |
| `requirements.txt` | Dependencias adicionales de esta etapa |
