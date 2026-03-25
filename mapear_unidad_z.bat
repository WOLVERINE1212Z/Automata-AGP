@echo off
echo ================================================
echo   MAPEANDO UNIDAD DE RED Z:
echo ================================================
echo.

echo Verificando si la unidad Z: ya esta conectada...
if exist Z:\ (
    echo La unidad Z: ya esta conectada.
) else (
    echo Conectando a \\192.168.2.2\Sapfiles\PlanosSapProduccion
    net use Z: "\\192.168.2.2\Sapfiles\PlanosSapProduccion" /persistent:yes
    echo.
    echo Unidad Z: conectada correctamente!
)

echo.
echo Presiona cualquier tecla para abrir el visor...
pause > nul

start "" "visor_planos.html"
