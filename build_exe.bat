@echo off
echo.
echo [1/2] Compilando ejecutable con PyInstaller (CARPETA DIST)...
echo.
"C:\Users\fioni\Documents\practica profecional\.venv\Scripts\pyinstaller.exe" CAIT_Informes.spec --clean --noconfirm

echo.
echo [2/2] PyInstaller finalizado con exito.
echo.
echo AHORA: Abre el programa 'Inno Setup Compiler' y carga el archivo:
echo "C:\Users\fioni\Documents\Cait panama nuevo\installer\cait_installer.iss"
echo.
echo Presiona Compilar (Ctrl+F9) para generar el instalador final (.exe).
echo.
pause
