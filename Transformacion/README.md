# Transformación: obras nuevas para cada asesor de GRECO

Etapa local posterior a la ingesta. Lee los CSV que genera `Ingesta/extract.py`, limpia las licitaciones **adjudicadas**, las reparte entre los **4 Excel maestros** de GRECO (uno por asesor) y genera, para cada asesor, un Excel con sus **obras nuevas** en el mismo formato de 71 columnas de su maestro, listo para copiar.

SAML **nunca modifica los maestros**: solo los lee. El asesor decide qué filas agrega a su archivo.

## Requisitos

1. Entorno virtual `.venv/` en la raíz del proyecto:

   ```bash
   python -m venv .venv
   .venv/Scripts/python -m pip install -r Ingesta/requirements.txt -r Transformacion/requirements.txt   # Windows
   .venv/bin/python -m pip install -r Ingesta/requirements.txt -r Transformacion/requirements.txt       # macOS / Linux
   ```

2. Los 4 maestros en **`datos/maestros/`** (por ejemplo `Obras_IDU_Gustavo_VF1.xlsx`, `Obras_Acueducto_Alcaldias_Mariana_VF1.xlsx`, `Obras_IDRD_Privadas_Municipios_Asesor3_VF1.xlsx` y `Obras_Sugeine_VF1.xlsx`). La carpeta `datos/` **no se versiona**: los maestros traen NIT, ventas y clientes, y el repositorio es público. Cuando haya una versión nueva (VF2…), se reemplazan ahí.

En los ejemplos se usa `.venv/Scripts/python` (Windows). En macOS o Linux es `.venv/bin/python`.

## Proceso completo con un solo comando (recomendado)

Desde la raíz del repositorio:

| Comando | Qué hace |
|---|---|
| `.venv/Scripts/python Transformacion/ejecutar_proceso.py` | **Demostración** (por defecto). Usa solo los CSV ya guardados en `Ingesta/output/`, **sin consultar Licitaciones.info** |
| `.venv/Scripts/python Transformacion/ejecutar_proceso.py --actualizar` | **Actualización.** Comprueba `Ingesta/.env` (credenciales y `LICITACIONES_ESTADO=Adjudicado`), ejecuta `Ingesta/extract.py` y, **solo si termina bien y deja exactamente un CSV nuevo**, genera los archivos. Si la ingesta falla, se detiene y **no** genera nada |
| `--maestros DIR` | Carpeta de los maestros (por defecto `datos/maestros`) |
| `--salida DIR` | Carpeta de salida (por defecto `datos/salida`) |

Antes de consultar el portal se comprueba que los maestros existan y tengan la plantilla esperada.

Cada ejecución crea una carpeta nueva `datos/salida/SAML_<AAAAMMDD_HHMMSS>/` (nunca sobrescribe otra) con:

| Archivo | Contenido |
|---|---|
| `<maestro>_actualizado_<fecha>.xlsx` | **Copia de cada maestro con las obras nuevas ya insertadas** (con Excel, en Windows). Es la que se usa |
| `Nuevas_Sin_asignar_GRECO_<fecha>.xlsx` | Obras de entidades sin regla, en la misma plantilla, para que GRECO elija el asesor |
| `Nuevas_<maestro>_<fecha>.xlsx` | Uno por asesor con obras nuevas. Hoja **Instrucciones**, una hoja por cada hoja de destino (mismo nombre que en el maestro: `IDU`, `Alcaldías Locales`, `Municipios Aledaños`…) y la hoja **Listas** con los desplegables |
| `SAML_revision_<fecha>.xlsx` | **Resumen** (CSV y maestros con su SHA-256, conteos, reglas y mapeo), **Sin clasificar** (solo si falta una plantilla), **Sin consumo de material**, **Fuera de zona**, **Ya en maestros**, **Procesos** (todos los adjudicados con su destino), **Ganadores** e **Incidencias** |

Después de generar los archivos, el proceso verifica que:
- todos los procesos son `Adjudicado` y no hay `idContrato` repetidos;
- cada fila nueva tiene un Asesor de la lista del maestro y la fórmula de Valor por Ejecutar;
- no hay códigos repetidos en una hoja ni obras que ya estaban en los maestros;
- no hay valores con forma de JWT;
- ni los CSV ni los maestros cambiaron (huella SHA-256 antes y después);
- cada copia actualizada tiene, en cada hoja, sus obras de antes más las nuevas, con sus códigos.

Si algo falla, la carpeta se renombra con el sufijo `_REVISAR`.

Códigos de salida: `0` correcto · `1` faltan CSV, maestros o configuración · `2` falló la ingesta · `3` falló la verificación.

## Cómo llegan las obras al maestro

**Automático (Windows con Excel instalado).** Para cada asesor con obras nuevas, SAML copia su maestro a la carpeta de salida (`<maestro>_actualizado_<fecha>.xlsx`), lo abre con **Excel** (automatización COM con `pywin32`; no con openpyxl, que borraría los gráficos y los desplegables) y, en cada hoja:

1. inserta las filas nuevas **encima de la última obra**, para que queden dentro de los rangos fijos del Resumen, Mercado 2027 y Top 15 (por ejemplo `I3:I93`), que Excel amplía solos;
2. copia en ellas la última obra del propio maestro (fórmulas, formatos, desplegables y formato condicional de ese libro);
3. escribe los datos de SAML en las columnas que no son fórmula;
4. revisa que las fórmulas del Resumen, Mercado 2027 y Top 15 cubran las filas nuevas y avisa si alguna no (pasa en hojas con una sola obra, donde las filas se insertan debajo).

El maestro original nunca se toca. GRECO revisa la copia y, si está bien, la usa en lugar de su archivo. Probado con los 4 maestros reales: gráficos intactos, Resumen actualizado y sin errores. Si el maestro se guardó con hojas agrupadas, SAML selecciona una sola hoja antes de insertar; si una hoja tiene un filtro activo (filas ocultas), en la copia se muestran todas las filas y se avisa. Para no generar las copias: `--sin-actualizar-maestros`.

**Después de n8n (sin Excel en Docker).** En el equipo Windows con Excel, doble clic en `Actualizar_maestros.bat` (o `.venv\Scripts\python Transformacion\actualizar_maestros.py [carpeta]`): genera las copias actualizadas a partir de los `Nuevas_…` de la carpeta más reciente de `datos/salida/`. No repite una copia ya generada y no actualiza un maestro que cambió desde esa salida.

**Manual (sin Excel en ningún equipo).** Se usan los archivos `Nuevas_…xlsx`:

1. Revisar cada fila (Tipo de contrato, Localidad, Inicio de obra, Meses por Terminar) y completar lo que falte.
2. Seleccionar las filas desde la fila 3, **solo columnas A a BS** (desde BU son datos de revisión de SAML), y copiarlas.
3. En el maestro, hoja del mismo nombre: clic en la celda A de la **última obra** → clic derecho → **Insertar celdas copiadas** → **Desplazar las celdas hacia abajo**. No pegarlas debajo de la última obra: quedarían fuera de los totales.

## Reglas de reparto

| Entidad contratante (nombre normalizado) | Hoja | Archivo |
|---|---|---|
| Instituto de Desarrollo Urbano, TransMilenio | IDU | Gustavo |
| Empresa de Acueducto (EAAB) | Acueducto | Mariana |
| Instituto Distrital de Recreación y Deporte (IDRD) | IDRD | Asesor 3 |
| Alcaldía Local de… / Fondo de Desarrollo Local de… | Alcaldías Locales | Mariana |
| Alcaldías municipales, ICCU, Fondo de Desarrollo de Proyectos de Cundinamarca, IDUVI — con ubicación en Cundinamarca | Municipios Aledaños | Asesor 3 |
| Cualquier otra (entidades nacionales, distritales o departamentales sin regla) | Ver «Entidades sin regla» | — |

**Entidades sin regla.** Se evalúan en este orden:

1. Ubicación fuera de Bogotá y Cundinamarca (incluye «Todo el país») → hoja **Fuera de zona** de la revisión (descartadas).
2. INVIAS o ANI (obra vial nacional) → **Municipios Aledaños** (Asesor 3).
3. Gobernación / Departamento de Cundinamarca con obra fuera de Bogotá → **Municipios Aledaños** (Asesor 3).
4. Solo mantenimiento (el objeto dice mantenimiento o conservación y no dice adecuación, construcción, remodelación, mejoramiento, rehabilitación…) → hoja **Sin consumo de material** de la revisión. La fórmula de GRECO ya les da 0 potencial.
5. El resto (obras en Bogotá) se reparte **por tipo de obra**: vial o espacio público → **IDU** (Gustavo); parques y escenarios deportivos → **IDRD** (Asesor 3); redes de acueducto y alcantarillado → **Acueducto** (Mariana); edificaciones y adecuaciones de sedes → **IDRD** (Asesor 3), con Tipo de contrato «Edificaciones y adecuaciones» y el nombre de la entidad en la columna A. Si el ganador es cliente de Sugeine, va también a su archivo (las edificaciones, a su hoja «Edificaciones y Mant.»).
6. Lo que no encaja en ningún tipo → archivo **`Nuevas_Sin_asignar_GRECO_<fecha>.xlsx`** con las 71 columnas y `Asesor = Sin asignar (GRECO)`, para que GRECO elija el asesor. Con el LLM esto no ocurre: siempre elige una hoja.

Con el CSV del 8-nov-2026 (138 adjudicadas) y el LLM (gpt-5-mini): Mariana 20, Asesor 3 65 (la mayoría edificaciones en IDRD), Gustavo 7, Sugeine 2, ninguna sin asignar; 22 sin consumo de material y 7 fuera de zona.

Con el CSV del 8-nov-2026 (138 adjudicadas): 13 para Mariana, 10 para Asesor 3, 2 para Gustavo, 1 para Sugeine, 53 sin asignar, 38 sin consumo de material, 4 fuera de zona y 17 que ya estaban en los maestros.

- El archivo y el asesor salen del maestro que tiene esa hoja; el asesor se identifica por la columna E de sus obras.
- **Sugeine:** si un ganador coincide con un cliente de su hoja «Cobertura clientes» (sin tildes, signos ni sufijos como S.A.S. o LTDA), la obra va **también** a su archivo, y en el del otro asesor queda con `Asesor = Sugeine` (su fórmula la pone en 0 para no contarla dos veces). Si el nombre solo se parece (similitud ≥ 0,88), queda como `<asesor> (posible Sugeine)` y se anota en Incidencias.
- **Obras repetidas:** un proceso no se agrega si su código (comparado sin espacios ni guiones) ya está en algún maestro **y** coincide la empresa o el valor (±2 %). Si solo coincide el código (códigos genéricos como `LP-001-2026` se repiten entre entidades), se agrega y se anota para revisar.

Las reglas están en `obras_asesores.py` (`REGLAS_HOJA` y `PATRON_MUNICIPAL`) para ajustarlas si GRECO cambia la distribución.

## Clasificación con LLM (OpenAI)

Las entidades **sin regla** (nacionales, distritales, departamentales…) se pueden clasificar con un LLM de OpenAI en lugar de las reglas de palabras clave. Las que tienen regla (IDU, alcaldías locales, acueducto, IDRD, municipios) siguen por reglas.

1. Copia `Transformacion/.env.example` como `Transformacion/.env` (no se versiona) y escribe `OPENAI_API_KEY`. Opcional: `OPENAI_MODEL` (por defecto `gpt-5-mini`) y `SAML_USAR_LLM=false` para apagarlo.
2. Ejecuta el proceso como siempre. Para forzar solo reglas: `--sin-llm`.

Qué hace (`clasificador_llm.py`):
- Envía, en lotes de 20, la entidad, el objeto, la ubicación, el valor y los ganadores de cada proceso, junto con las hojas de cada asesor, los tipos de contrato del maestro y la lista de **clientes de Sugeine** (decisión de GRECO: la lista sale del equipo hacia OpenAI).
- El LLM responde en JSON con un esquema fijo (salida estructurada): destino (una hoja de asesor, `Sin consumo de material` o `Fuera de zona`), tipo de contrato, localidad, si consume material, si el ganador es cliente de Sugeine, confianza (alta, media o baja) y una justificación.
- **Toda obra que consume material en la zona queda con un asesor**: el LLM no tiene la opción «Sin asignar». Si duda entre hojas, elige la más probable; con confianza baja se anota en Notas («confirmar el asesor») y en Incidencias.
- Sin clave (o si la API falla) se usan las reglas, y ahí sí puede haber obras en el archivo «Sin asignar (GRECO)».
- SAML vuelve a validar cada valor contra las listas del maestro. Una coincidencia exacta local con un cliente de Sugeine siempre vale; un "cliente" que no esté en la lista se descarta.
- La justificación queda en la columna de revisión **SAML · Clasificación** (y en el Motivo de las hojas de revisión); la hoja Procesos indica si cada uno se clasificó por **Reglas** o por **LLM**, y el Resumen muestra el modelo, los tokens y los avisos.
- Si no hay clave, la clave es rechazada, el modelo no existe o la API falla, esos procesos se clasifican con las reglas y se avisa (la ejecución no falla).

El texto de las instrucciones y la descripción de cada hoja están en `clasificador_llm.py` (`INSTRUCCIONES` y `DESCRIPCION_HOJAS`) para ajustarlos con GRECO.

## Qué llena SAML en las 71 columnas

| Columna | Regla |
|---|---|
| A Entidad | Categoría de la hoja como la usa el maestro (`IDU`, `Alcaldía Local`, `Municipio`…) |
| B Localidad | Alcaldías Locales: localidad del nombre de la entidad. Municipios: municipio de la ubicación. Resto: `Bogotá D.C.` |
| E Asesor | Dueño de la hoja, `Sugeine` o `<asesor> (posible Sugeine)` |
| I Empresa | Ganador(es), separados por ` \| ` |
| J Notas | `[SAML · fecha] ADJUDICADA · VERIFICAR INICIO —` origen de Inicio y Meses y pendientes |
| K Nombre de la Obra | Objeto del contrato |
| L Tipo de contrato | Propuesto por palabras clave del objeto, de la lista del maestro (confirmar) |
| M Contrato No y AZ Proceso | Código del proceso |
| N Valor Millones | Valor de la API / 1.000.000 |
| P Avance | 0 |
| Q Meses por Terminar | Del cronograma **solo** si inicio y fin son claros; si no, vacío |
| R Inicio de obra | Mes de inicio programado si es claro; si no, mes de la extracción (mínimo el primer mes de la proyección, X2) |
| AY, BA, BB, BH | Origen `NUEVA · SAML…`, fecha de cargue, 0 meses de espera y fuente del valor |
| BI:BM | Datos de Sugeine (solo si aplica) |
| Fórmulas (S:AX, BC:BG, BN:BS) | Copiadas de la primera fila del maestro y ajustadas a cada fila, con los porcentajes de la hoja (T1:V1) |
| C, D, F, G, H, O, BL | Vacías (las diligencia el asesor) |

**Inicio y Meses desde el cronograma.** El cronograma de SECOP trae fechas **programadas** y conserva las versiones anteriores (adendas), así que se compara por **mes**: si todas las fechas de «inicio de ejecución» caen en el mismo mes, ese es el Inicio de obra; si además hay un único mes de fin («plazo de ejecución» o «ejecución del contrato»), Meses por Terminar = meses desde el inicio (o desde X2, si empezó antes) hasta el fin. En cualquier otro caso Meses queda vacío y se anota en Incidencias. Con Meses vacío la obra no proyecta hasta que el asesor lo complete.

**Lo que la API no trae:** NIT, número de contrato, fecha real de adjudicación, plazo como campo, coordenadas ni integrantes de los consorcios.

## Etapas por separado

```bash
cd Ingesta && ../.venv/Scripts/python extract.py && cd ..          # opcional: datos nuevos
.venv/Scripts/python Transformacion/generar_archivos_asesores.py   # limpia, reparte y genera
```

Opciones de `generar_archivos_asesores.py`: `--entrada`, `--maestros` y `--salida`.

## Pruebas

Usan datos y maestros **sintéticos** en carpetas temporales (nunca los archivos reales):

```bash
.venv/Scripts/python -m unittest discover -s Transformacion/tests -v
```

## Archivos

| Archivo | Propósito |
|---|---|
| `ejecutar_proceso.py` | Punto de entrada único: modo demostración o `--actualizar`, más la verificación final |
| `limpieza_licitaciones.py` | Consolidación por `idContrato`, filtro `Adjudicado`, normalización, ganadores, fechas del cronograma e incidencias. No escribe archivos |
| `maestros.py` | Lectura (solo lectura) de los maestros: hojas, plantilla, asesores, códigos existentes, clientes de Sugeine y listas |
| `obras_asesores.py` | Reglas de reparto, Sugeine, duplicados y armado de las filas de 71 columnas |
| `actualizar_maestros.py` | Copias de los maestros con las obras insertadas, con Excel (Windows) |
| `clasificador_llm.py` | Clasificación con OpenAI de las entidades sin regla (salida estructurada, con respaldo por reglas) |
| `.env.example` | Plantilla de `Transformacion/.env` (clave y modelo de OpenAI) |
| `generar_archivos_asesores.py` | Escribe los Excel por asesor y el de revisión con `openpyxl` |
| `perfil_csv.py` | Perfil de solo lectura de un CSV de la ingesta (columnas, tipos, completitud) |
| `tests/` | Pruebas: limpieza, reparto y archivos, y el punto de entrada con una ingesta simulada |
| `requirements.txt` | Dependencias adicionales de esta etapa |
