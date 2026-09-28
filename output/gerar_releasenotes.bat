@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"
title GMUD - Gerar PPTX

:pedir
echo ==========================================
echo   GMUD - Gerar PPTX
echo ==========================================
echo.
echo   Digite o numero do mes (1 a 12).
echo   Ex: 9 = Setembro. O ano e o atual
echo   (ou o anterior, se o mes ainda nao chegou).
echo.
set "MES="
set /p "MES=Mes: "

echo %MES%| findstr /r /x "[1-9] 1[0-2] 0[1-9]" >nul
if errorlevel 1 (
  echo.
  echo Valor invalido. Use um numero de 1 a 12.
  echo.
  goto pedir
)
if "%MES:~0,1%"=="0" set "MES=%MES:~1%"

set "GMUD_MES=%MES%"
echo.
echo Rodando para o mes %GMUD_MES%...
echo.
node index.js

echo.
pause
endlocal
