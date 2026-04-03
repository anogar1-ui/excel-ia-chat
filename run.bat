@echo off
chcp 65001 >nul
title 📊 Gestor de Excel con IA

echo ============================================
echo    📊 Gestor de Excel con IA - Streamlit
echo ============================================
echo.

:: Verificar si existe el entorno virtual
if exist ".venv\Scripts\activate.bat" (
    echo [✓] Activando entorno virtual...
    call .venv\Scripts\activate.bat
) else (
    echo [!] No se encontro entorno virtual.
    echo     Creando uno nuevo...
    python -m venv .venv
    call .venv\Scripts\activate.bat
    echo [✓] Entorno virtual creado.
)

echo.
echo [*] Verificando dependencias...
pip install -q -r requirements.txt 2>nul
echo [✓] Dependencias verificadas.

echo.

:: Verificar si Ollama está corriendo
echo [*] Verificando Ollama...
curl -s http://localhost:11434/api/tags >nul 2>&1
if errorlevel 1 (
    echo [!] Ollama no esta corriendo. Iniciando...
    start "" /B ollama serve >nul 2>&1
    timeout /t 3 /nobreak >nul
    echo [✓] Ollama iniciado.
) else (
    echo [✓] Ollama ya esta corriendo.
)

echo.
echo [*] Iniciando aplicacion...
echo [*] La aplicacion se abrira en: http://localhost:8501
echo.
echo    Presiona Ctrl+C para detener el servidor
echo ============================================
echo.

streamlit run app.py

pause
