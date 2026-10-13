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
│   ├── Entregable 1/
│   └── Entregable 2/
├── Ingesta/          # extraccion desde la API de licitaciones.info -> Ingesta/output/*.csv
├── Transformacion/   # limpieza y Excel intermedio -> datos/intermedio/*.xlsx
├── n8n/              # n8n local en Docker + workflow (disparadores manuales del proceso local)
└── README.md
```

### Ingesta y Transformacion:

- Ingesta: configuración en `Ingesta/.env` (plantilla en `Ingesta/.env.example`); se ejecuta con `python extract.py` desde `Ingesta/`.
- Transformacion: genera el Excel intermedio de licitaciones adjudicadas (versión de demostración). Instrucciones en [Transformacion/README.md](Transformacion/README.md).
  - Demo con los CSV ya guardados (no consulta el portal): `.venv/bin/python Transformacion/ejecutar_proceso.py`
  - Actualizar desde Licitaciones.info y generar el Excel: `.venv/bin/python Transformacion/ejecutar_proceso.py --actualizar`
- n8n: instancia local en Docker (http://localhost:5678) con el workflow "Proceso completo" (ingesta desde Licitaciones.info y luego Excel) y otro de demo con CSV locales. Instrucciones en [n8n/README.md](n8n/README.md).
- `datos/` y `.env` no se versionan: contienen credenciales, archivos del cliente y resultados generados.

### docs:

Contiene la documentación y los entregables desarrollados durante las diferentes etapas del proyecto.

- Entregable 1: Acta de constitución, evidencias de reuniones con el Product Owner (PO) y retrospectiva.
- Entregable 2: Documentación correspondiente a la etapa de planificación, incluyendo el documento de formulación y sus anexos, evidencias de reuniones y enlace al backlog.

---

## Backlog

Ver [aquí](https://github.com/users/svillav4/projects/9), construido a partir de la planeación realizada mediante los OKRs y la EDT, este backlog será utilizado como base para la planificación y asignación de trabajo en los diferentes sprints del proyecto.

---

## Propósito del repositorio

El repositorio sirve como espacio central para documentar, organizar y dar seguimiento al proyecto SAML, manteniendo trazabilidad entre la formulación del proyecto, su planificación y el trabajo que será ejecutado durante los sprints.
