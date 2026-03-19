@echo off
REM 
REM 

set ODA="C:\Program Files\ODA\ODAFileConverter 27.1.0\ODAFileConverter.exe"

REM Ejecutar sin mostrar ventana
start /b "" %ODA% %1 %2 ACAD2018 DXF 0 0

REM Esperar a que termine
timeout /t 2 /nobreak >nul

REM Verificar si se creó el archivo
if exist %2 (
    echo OK
) else (
    echo ERROR
)
