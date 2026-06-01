@echo off
title Actualizar Dashboard NPS - UDLA
color 1f
echo ==========================================================
echo       ACTUALIZADOR DE DASHBOARD NPS Y SATISFACCION
echo                  Universidad de Las Americas
echo ==========================================================
echo.
echo Procesando archivos Excel en las carpetas de periodos...
echo Por favor, espere...
echo.

python "%~dp0consolidate_data.py"

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Ha ocurrido un error al consolidar los datos.
    echo Por favor, verifique que Python este instalado y los archivos Excel esten cerrados.
    echo.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo ==========================================================
echo  [EXITO] ¡Los datos se han consolidado correctamente!
echo.
echo  Se han generado los siguientes archivos actualizados:
echo   - NPS_Consolidado.xlsx (Para Power BI)
echo   - NPS_Consolidado.db   (Base de datos SQLite)
echo   - dashboard_data.js    (Datos para el Dashboard Web)
echo ==========================================================
echo.

:: Detectar repositorio Git y subir cambios a GitHub Pages de forma automatizada
if exist "%~dp0.git" (
    echo.
    echo ==========================================================
    echo  Detectado repositorio Git. Subiendo cambios a GitHub...
    echo ==========================================================
    echo.
    
    :: Agregar los archivos generados y necesarios
    git add "%~dp0dashboard.html" "%~dp0dashboard_data.js" "%~dp0dashboard_data.json" "%~dp0NPS_Consolidado.xlsx" "%~dp0consolidate_data.py" "%~dp0matriculas_historicas.json" "%~dp0Detalle Estudiantes Postgrado_Tipo Ingreso.xlsx" "%~dp0Actualizar_Dashboard.bat"
    
    :: Crear el commit con fecha y hora actual de la actualizacion
    git commit -m "Actualizacion automatica de datos - %DATE% %TIME%"
    
    :: Hacer el push al repositorio remoto (GitHub)
    echo Subiendo cambios a GitHub Pages...
    git push
    
    if %ERRORLEVEL% NEQ 0 (
        echo.
        echo [ADVERTENCIA] No se pudo realizar el "git push" a GitHub.
        echo Por favor, verifique su conexion a Internet y que sus credenciales esten vigentes.
    ) else (
        echo.
        echo [EXITO] ¡Cambios publicados con exito en GitHub Pages!
        echo Su dashboard en linea estara actualizado en breves minutos en:
        echo https://romanser777.github.io/nps-dashboard-udla/
    )
) else (
    echo.
    echo [INFO] Carpeta local no inicializada con Git.
    echo Si desea que las actualizaciones se suban automaticamente a GitHub Pages,
    echo inicialice Git en esta carpeta ejecutando estos comandos en la terminal:
    echo.
    echo   git init
    echo   git remote add origin https://github.com/romanser777/nps-dashboard-udla.git
    echo   git add .
    echo   git commit -m "Initial commit"
    echo   git push -u origin main
)

echo.
echo ==========================================================
echo  Para ver los cambios locales:
echo   1. Dashboard Web: Abre dashboard.html o refrescalo (F5).
echo   2. Power BI: Abre tu reporte y haz clic en "Actualizar".
echo ==========================================================
echo.
echo Presione cualquier tecla para cerrar esta ventana...
pause > nul
