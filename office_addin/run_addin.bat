@echo off
title Excel IA Chat - Office Add-in Server
echo.
echo  ========================================
echo   Excel IA Chat - Office Add-in
echo  ========================================
echo.

:: Verificar Python
python --version >nul 2>&1
if errorlevel 1 (
    echo  ERROR: Python no encontrado.
    echo  Instala Python desde python.org
    pause
    exit /b 1
)

:: Instalar dependencias si es necesario
pip show flask >nul 2>&1
if errorlevel 1 (
    echo  Instalando dependencias...
    pip install -r requirements.txt
    echo.
)

:: Verificar certificados SSL
if not exist "certs\server.crt" (
    echo  No se encontraron certificados SSL.
    echo  Generando certificados...
    echo.
    python generate_certs.py
    echo.
)

:: Iniciar servidor
echo  Iniciando servidor...
echo.
python server.py

pause
