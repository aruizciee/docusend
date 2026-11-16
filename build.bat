@echo off
rem Compila DocuSend en local. Requiere: pip install -r requirements.txt
echo Limpiando builds anteriores...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"

echo Construyendo el ejecutable...
pyinstaller --noconfirm --onefile --windowed --icon="assets\icon.ico" --add-data "assets;assets/" --collect-all customtkinter "docusend.py"
if errorlevel 1 (
    echo.
    echo ERROR: la compilacion ha fallado.
    pause
    exit /b 1
)

echo.
echo Moviendo el ejecutable a la carpeta principal...
move /y "dist\docusend.exe" "DocuSend_App.exe"

echo Limpiando carpetas temporales de compilacion...
rmdir /s /q "build"
rmdir /s /q "dist"
del /f /q "docusend.spec"

echo.
echo Proceso finalizado. El ejecutable 'DocuSend_App.exe' ya esta listo en la carpeta principal.
pause
