@echo off
echo Limpiando builds anteriores...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"

echo Construyendo el ejecutable...
pyinstaller --noconfirm --onefile --windowed --icon="assets\icon.ico" --add-data "assets;assets/" --add-data "C:\Users\ARuiz\AppData\Local\Programs\Python\Python313\Lib\site-packages\customtkinter;customtkinter/" "docusend.py"

echo.
echo Movimiendo el ejecutable a la carpeta principal...
move /y "dist\docusend.exe" "DocuSend_App.exe"

echo Limpiando carpetas temporales de compilación...
rmdir /s /q "build"
rmdir /s /q "dist"
del /f /q "docusend.spec"

echo.
echo Proceso finalizado. El ejecutable 'DocuSend_App.exe' ya está listo en la carpeta principal.
pause
