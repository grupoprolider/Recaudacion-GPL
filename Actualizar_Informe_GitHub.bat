@echo off
chcp 65001 >nul
echo ===================================================
echo   ACTUALIZANDO INFORME MENSUAL - GRUPO PROLIDER
echo ===================================================
cd /d "%~dp0\app"
call npm run sync
cd /d "%~dp0"
git add .
git commit -m "Actualizacion mensual de planillas de recaudacion"
git push origin main
echo.
echo ¡Listo! Los datos fueron subidos a GitHub correctamente.
pause
