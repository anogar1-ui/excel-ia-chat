@echo off
chcp 65001 >nul
title 🤖 Excel + Ollama API Server

echo ============================================
echo    🤖 Excel + Ollama - API Server
echo ============================================
echo.

:: Ir al directorio del proyecto
cd /d "%~dp0\.."

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
pip install -q flask flask-cors ollama python-dotenv 2>nul
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
echo ============================================
echo    INSTRUCCIONES DE USO:
echo ============================================
echo.
echo    1. Deja esta ventana abierta
echo    2. Abre Excel y tu archivo .xlsm
echo    3. Presiona Alt+F8 y ejecuta "MostrarChatOllama"
echo    4. Escribe instrucciones en lenguaje natural
echo.
echo    API corriendo en: http://localhost:5050
echo    Presiona Ctrl+C para detener el servidor
echo ============================================
echo.

python excel_extension\api_server.py

pause
