@echo off
REM Ejecuta SAML completo: consulta Licitaciones.info, reparte las obras por asesor y
REM genera en datos\salida las copias actualizadas de los maestros. Uso: doble clic.
cd /d "%~dp0"
".venv\Scripts\python.exe" Transformacion\ejecutar_proceso.py --actualizar %*
pause
