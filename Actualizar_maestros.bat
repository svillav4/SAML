@echo off
REM Genera las copias actualizadas de los maestros (con Excel) a partir de la salida mas reciente
REM de SAML en datos\salida (por ejemplo, la que dejo una ejecucion de n8n). Uso: doble clic.
cd /d "%~dp0"
".venv\Scripts\python.exe" Transformacion\actualizar_maestros.py %*
pause
