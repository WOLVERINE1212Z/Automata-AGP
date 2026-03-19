@echo off
title AGP - Instalador de Planos Tecnicos
color 0A

echo ============================================================
echo   AGP - EXTRACTOR DE PLANOS TECNICOS
echo   Instalador y Ejecutor
echo ============================================================
echo.

REM Verificar Python
echo [1/4] Verificando Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo   ERROR: Python no esta instalado
    echo   Por favor instala Python 3.8 o superior desde python.org
    echo.
    pause
    exit /b 1
)

python --version
echo.

REM Instalar dependencias
echo [2/4] Instalando dependencias...
echo   - ezdxf (lector de archivos DXF/DWG)
pip install ezdxf --quiet
if errorlevel 1 (
    echo   ERROR: No se pudo instalar ezdxf
    pause
    exit /b 1
)

echo   - openpyxl (generador de Excel)
pip install openpyxl --quiet
if errorlevel 1 (
    echo   ERROR: No se pudo instalar openpyxl
    pause
    exit /b 1
)

echo   [OK] Dependencias instaladas
echo.

REM Ejecutar extractor
echo [3/4] Ejecutando extractor de planos...
echo   Esto puede tardar varios minutos dependiendo de la cantidad de archivos
echo.

python extractor_prueba.py

if errorlevel 1 (
    echo.
    echo   ERROR: El extractor tuvo problemas
    pause
    exit /b 1
)

echo.
echo [4/4] Abriendo visor de planos...
echo.

REM Abrir visor en Chrome
start chrome "visor_planos.html"

echo ============================================================
echo   INSTALACION COMPLETADA
echo ============================================================
echo.
echo   El visor de planos se ha abierto en tu navegador
echo.
echo   Para actualizar los datos en el futuro, vuelve a ejecutar
echo   este archivo o ejecuta: python extractor_prueba.py
echo.
pause
