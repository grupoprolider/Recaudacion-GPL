@echo off
echo ===================================================
echo   ACTUALIZANDO INFORME MENSUAL - GRUPO PROLIDER
echo ===================================================
cd /d "%~dp0\app"
call npm run sync
cd /d "%~dp0"
git add .
git commit -m "Actualizacion mensual de planillas de recaudacion"
git pull --rebase -X ours origin main
git push origin main
echo.
echo Listo! Los datos de la carpeta fueron subidos a GitHub correctamente.
echo En 1 minuto estaran visibles en: https://grupoprolider.github.io/Recaudacion-GPL/
pause
