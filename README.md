# SAML
Sistema Automatizado de Monitoreo de Licitaciones

---

Este repositorio centraliza la documentación, entregables, evidencias y elementos de planificación del proyecto.

## Integrantes:

- Juan José Botero García
- Emanuel Castañeda Cano
- Samuel Deossa Gómez
- Cristóbal Rodríguez Rojas
- Samuel Villa Valencia

---

## Estructura del repositorio: 
```text
SAML/
│
├── docs/
│   ├── Entregable1/ … Entregable3/
│   ├── Entregable4/  # documentación técnica y manual de usuario (versión final)
│   └── SAML_presentaciones/
├── Ejecutar_SAML.bat        # doble clic: proceso completo
├── Actualizar_maestros.bat  # doble clic: copias de los maestros tras una ejecución de n8n
├── Ingesta/          # extraccion desde la API de licitaciones.info -> Ingesta/output/*.csv
├── Transformacion/   # limpieza y reparto por asesor -> datos/salida/SAML_<fecha>/*.xlsx
├── datos/            # NO se versiona: maestros/ (4 Excel de GRECO) y salida/
├── n8n/              # n8n local en Docker + workflow (disparadores manuales del proceso local)
└── README.md
```

### Ingesta y Transformacion:

- Ingesta: configuración en `Ingesta/.env` (plantilla en `Ingesta/.env.example`); se ejecuta con `python extract.py` desde `Ingesta/`.
- Transformacion: limpia las licitaciones adjudicadas, las reparte entre los 4 Excel maestros de GRECO (uno por asesor, en `datos/maestros/`) y genera un Excel de obras nuevas por asesor, con las 71 columnas de su maestro, más un Excel de revisión. Instrucciones en [Transformacion/README.md](Transformacion/README.md).
  - Demo con los CSV ya guardados (no consulta el portal): `.venv/bin/python Transformacion/ejecutar_proceso.py`
  - Actualizar desde Licitaciones.info y generar los archivos: `.venv/bin/python Transformacion/ejecutar_proceso.py --actualizar`
  - En Windows el intérprete es `.venv/Scripts/python`.
- En Windows, el proceso completo también se ejecuta con doble clic en `Ejecutar_SAML.bat`.
- Copias actualizadas de los maestros (con Excel, en Windows): se generan solas al ejecutar desde la terminal; después de una ejecución de n8n, doble clic en `Actualizar_maestros.bat`.
- n8n: instancia local en Docker (http://localhost:5678) con el workflow "Proceso completo" (ingesta desde Licitaciones.info y luego archivos por asesor) y otro de demo con CSV locales. Instrucciones en [n8n/README.md](n8n/README.md).
- `datos/` y `.env` no se versionan: contienen credenciales, los maestros del cliente (NIT, ventas, clientes) y resultados generados. El repositorio es público: nunca subas esos archivos.

### docs:

Contiene la documentación y los entregables desarrollados durante las diferentes etapas del proyecto.

- Entregable 1: Acta de constitución, evidencias de reuniones con el Product Owner (PO) y retrospectiva.
- Entregable 2: Documentación correspondiente a la etapa de planificación, incluyendo el documento de formulación y sus anexos, evidencias de reuniones y enlace al backlog.
- Entregable 3: Registro de gestión de riesgos, anexo de uso de IA, presentaciones y retrospectiva.
- Entregable 4: **[Documentación técnica](docs/Entregable4/Documentacion_Tecnica_SAML.md)** (también en [PDF](docs/Entregable4/Documentacion_Tecnica_SAML.pdf)) y **manual de usuario** para GRECO ([PDF](docs/Entregable4/Manual_Usuario_SAML.pdf) · [Word](docs/Entregable4/Manual_Usuario_SAML.docx)), actualizados a la versión final.

---

## Backlog

Ver [aquí](https://github.com/users/svillav4/projects/9), construido a partir de la planeación realizada mediante los OKRs y la EDT, este backlog será utilizado como base para la planificación y asignación de trabajo en los diferentes sprints del proyecto.

---

## Propósito del repositorio

El repositorio sirve como espacio central para documentar, organizar y dar seguimiento al proyecto SAML, manteniendo trazabilidad entre la formulación del proyecto, su planificación y el trabajo que será ejecutado durante los sprints.
