@echo off
setlocal
set "ROOT=%~dp0.."
pushd "%ROOT%"

set "APP_VERSION="
if exist "%ROOT%\VERSION" (
  set /p APP_VERSION=<"%ROOT%\VERSION"
)
if "%APP_VERSION%"=="" (
  set "APP_VERSION=1.0.4"
)

set "PYTHON_EXE=%ROOT%\.venv\Scripts\python.exe"
if not exist "%PYTHON_EXE%" (
  set "PYTHON_EXE=python"
)

echo [1/2] Generando ejecutable con PyInstaller...
"%PYTHON_EXE%" -m PyInstaller -y CAIT_Informes.spec
if errorlevel 1 (
  echo Error: fallo al construir el ejecutable con PyInstaller.
  echo Sugerencia: instala dependencias con ""%PYTHON_EXE%" -m pip install -r requirements.txt".
  popd
  exit /b 1
)

set "ISCC_CMD=iscc"
where /q iscc
if errorlevel 1 (
  if exist "%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe" (
    set "ISCC_CMD=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"
  ) else (
    if exist "%ProgramFiles%\Inno Setup 6\ISCC.exe" (
      set "ISCC_CMD=%ProgramFiles%\Inno Setup 6\ISCC.exe"
    ) else (
      if exist "%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe" (
        set "ISCC_CMD=%LOCALAPPDATA%\Programs\Inno Setup 6\ISCC.exe"
      )
    )
  )
)

echo [2/2] Generando instalador con Inno Setup...
"%ISCC_CMD%" /DMyAppVersion=%APP_VERSION% installer\CAIT_Informes.iss
if errorlevel 1 (
  echo Error: fallo al generar el instalador.
  echo Verifica que Inno Setup 6 este instalado (ISCC.exe).
  popd
  exit /b 1
)

echo Instalador listo en installer\output\CAIT_Informes_Setup.exe
popd
endlocal
