# Tablero Power BI de SAML

Proyecto de Power BI en formato **PBIP** (modelo en TMDL y reporte en PBIR, ambos en texto y versionables).
Usa la misma identidad visual que las presentaciones de SAML: fondo carbón, papel cálido, acento ámbar y tipografía Archivo.

## Cómo abrirlo

1. Abrir `SAML_Dashboard.pbip` con Power BI Desktop (versión de 2025 o posterior).
2. Pulsar **Actualizar** (o "Actualizar ahora" en la barra amarilla) para cargar los datos.
3. Guardar con Ctrl+S si quieres conservar la caché local (la carpeta `.pbi/` no se versiona).

La fuente se define en el parámetro **RutaFuente** (Transformar datos > Administrar parámetros).
Por defecto apunta a `fuentes_powerbi/licitaciones_intermedio_20261012_172632.xlsx`.

## Páginas

| Página | Qué responde |
|---|---|
| **Panorama** | KPIs (procesos, valor, mediana, empresas, competencia), ritmo mensual de publicación, valor por sector, top de entidades, tipo de intervención y rangos de valor. |
| **Licitaciones** | Segmentadores (sector, modalidad, intervención, alcance), modalidades, nivel de competencia, competencia según valor y detalle de procesos con barra de valor y enlace a licitaciones.info. |
| **Ganadores** | Empresas ganadoras, consorcios y UT, procesos con varios ganadores, ranking de empresas y competencia por sector. |
| **Calidad de datos** | Incidencias por tipo y campo, estado del mapeo de las 48 columnas del Excel de GRECO y la bitácora de incidencias. |
| **Ficha del proceso** | Página de *drillthrough* oculta: clic derecho sobre un proceso en Licitaciones > **Obtener detalles** > Ficha del proceso. |

La barra lateral es la navegación: cada sección tiene un botón transparente encima.

## Modelo

| Tabla | Origen | Notas |
|---|---|---|
| `FactProcesos` | hoja *Datos* | Un proceso por `idContrato`. Columnas derivadas: Sector, Modalidad, Tipo de intervención, Alcance territorial, Rango de valor, Nivel de competencia, Resultado de adjudicación y Días de publicación a cierre. |
| `FactGanadores` | hoja *Ganadores* | Relación **bidireccional** con FactProcesos, así un filtro por empresa también filtra los procesos. |
| `FactIncidencias` | hoja *Incidencias* | Relación con FactProcesos. Las incidencias de columna no tienen `idContrato`. |
| `DimMapeoColumnas` | hoja *Resumen*, sección "Mapeo de columnas" | Estado de cada columna (Mapeado, Trazabilidad, Sin fuente, etc.). |
| `DimCalendario` | DAX | Calendario de años completos, relacionado con la fecha de publicación. |
| `Medidas` | — | Las medidas, organizadas en carpetas: Procesos, Competencia, Ganadores, Calidad, Textos y Ficha. |

Las columnas de la hoja *Datos* que vienen vacías desde la API (Localidad, Latitud, NIT, Fecha de adjudicación, campos de GRECO, etc.) no se cargan.
Aparecen como "Sin fuente", "Manual GRECO" o "Calculado GRECO" en la página de calidad.

### Clasificaciones de SAML (orientativas)

- **Sector**: se asigna con palabras clave del nombre de la entidad (por ejemplo POLICIA o ARMADA → Defensa y seguridad; ALCALDIA LOCAL → Gobierno distrital y local). El orden de las reglas importa: un hospital militar cuenta como Salud.
- **Tipo de intervención**: se asigna con palabras clave del objeto del contrato, con esta prioridad: construcción > mejoramiento > adecuación > mantenimiento > suministro > estudios.
- **Valor de procesos ganados**: es el valor total del proceso; no se reparte entre co-ganadores.

Las reglas están en `build_dashboard.py` (constantes `ReglasSector` y `ReglasIntervencion` dentro del código M).

## Regenerar el proyecto

Todo el contenido (modelo, tema, fondos y visuales) se genera con un solo script:

```powershell
python PowerBI\build_dashboard.py                      # con la fuente por defecto
python PowerBI\build_dashboard.py --fuente "C:\ruta\SAML_revision_<fecha>.xlsx"
```

Requiere Python 3 con Pillow, y Power BI Desktop debe estar cerrado al ejecutarlo.
Los cambios hechos a mano en Desktop se sobrescriben al regenerar: lleva los cambios permanentes al script.

## Cuando llegue la salida final de SAML

La salida final (`datos/salida/SAML_<fecha>/SAML_revision_<fecha>.xlsx`) debe conservar las hojas *Datos*, *Ganadores*, *Incidencias* y *Resumen* con los mismos encabezados.
Si es así, basta con cambiar `RutaFuente` o volver a ejecutar el script con `--fuente`.
Si cambian los nombres de las columnas, hay que ajustar los pasos `Table.SelectColumns` del código M.
No se deben sustituir los datos por cifras inventadas.
