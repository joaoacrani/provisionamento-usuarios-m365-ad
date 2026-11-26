@echo off
REM Abre a interface web do provisionamento (duplo clique).
REM Para testar sem criar nada: iniciar-web.cmd -Demo
cd /d "%~dp0"
pwsh -NoProfile -ExecutionPolicy Bypass -File "%~dp0web.ps1" %*
pause
