@echo off
chcp 65001 >nul
title AGP - Extractor de Planos

echo ========================================
echo   AGP - Extractor de Planos Tecnicos
echo ========================================
echo.

REM Verificar Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python no esta instalado
    pause
    exit /b 1
)

echo [1/3] Verificando dependencias...
pip show ezdxl >nul 2>&1
if errorlevel 1 (
    echo        Instalando ezdxf...
    pip install ezdxf -q
)

pip show openpyxl >nul 2>&1
if errorlevel 1 (
    echo        Instalando openpyxl...
    pip install openpyxl -q
)

echo [2/3] Ejecutando extractor...
echo.

python extractor_planos.py %*

echo.
echo [3/3] Listo!
echo.
echo Para abrir el visor:
echo   - Abre visor_planos.html en tu navegador
echo.
pause
