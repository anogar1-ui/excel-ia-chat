@echo off
:: Solicitar permisos de administrador
net session >nul 2>&1
if %errorlevel% neq 0 (
    echo Solicitando permisos de administrador...
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

echo ==========================================
echo  Configurando carpeta compartida
echo ==========================================
echo.

:: Eliminar share existente si hay
net share ExcelAddin /delete >nul 2>&1

:: Crear el share
net share ExcelAddin="C:\Users\Propietario\Desktop\PROYECTOS\excellll\office_addin" /grant:todos,read
if %errorlevel% neq 0 (
    net share ExcelAddin="C:\Users\Propietario\Desktop\PROYECTOS\excellll\office_addin" /grant:everyone,read
)

echo.
echo  Share creado: \\localhost\ExcelAddin
echo.
echo  Ahora configura el catalogo en Excel:
echo  1. Archivo ^> Opciones ^> Centro de confianza
echo  2. Configuracion del Centro de confianza
echo  3. Catalogos de complementos de confianza
echo  4. URL del catalogo: \\localhost\ExcelAddin
echo  5. Agregar catalogo
echo  6. Marcar "Mostrar en el menu"
echo  7. Aceptar y reiniciar Excel
echo.
pause
