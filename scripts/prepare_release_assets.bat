@echo off
setlocal
set "ROOT=%~dp0.."
pushd "%ROOT%"

set "APP_VERSION="
if exist "%ROOT%\VERSION" (
  set /p APP_VERSION=<"%ROOT%\VERSION"
)
if "%APP_VERSION%"=="" (
  echo Error: no se pudo leer VERSION.
  popd
  exit /b 1
)

set "DIST_DIR=%ROOT%\dist\CAIT_Informes"
set "SETUP_EXE=%ROOT%\installer\output\CAIT_Informes_Setup.exe"
set "RELEASE_DIR=%ROOT%\release\v%APP_VERSION%"
set "PORTABLE_ZIP=%RELEASE_DIR%\CAIT_Informes_Portable_v%APP_VERSION%.zip"
set "SETUP_COPY=%RELEASE_DIR%\CAIT_Informes_Setup_v%APP_VERSION%.exe"
set "CHECKSUMS=%RELEASE_DIR%\SHA256SUMS.txt"

if not exist "%DIST_DIR%" (
  echo Error: no existe "%DIST_DIR%".
  echo Primero ejecuta: pyinstaller CAIT_Informes.spec
  popd
  exit /b 1
)

if not exist "%SETUP_EXE%" (
  echo Error: no existe "%SETUP_EXE%".
  echo Primero ejecuta: scripts\build_installer.bat
  popd
  exit /b 1
)

if not exist "%RELEASE_DIR%" mkdir "%RELEASE_DIR%"

echo Copiando instalador...
copy /Y "%SETUP_EXE%" "%SETUP_COPY%" >nul

echo Creando ZIP portable...
powershell -NoProfile -Command "if (Test-Path '%PORTABLE_ZIP%') { Remove-Item -Force '%PORTABLE_ZIP%' }; Compress-Archive -Path '%DIST_DIR%\*' -DestinationPath '%PORTABLE_ZIP%' -CompressionLevel Optimal"
if errorlevel 1 (
  echo Error: no se pudo crear el ZIP portable.
  popd
  exit /b 1
)

echo Generando checksums SHA-256...
powershell -NoProfile -Command "$files = @('%SETUP_COPY%', '%PORTABLE_ZIP%'); $out = @(); foreach ($f in $files) { $h = (Get-FileHash -Algorithm SHA256 -Path $f).Hash.ToLower(); $name = Split-Path -Leaf $f; $out += ($h + '  ' + $name) }; $out | Set-Content -Encoding ASCII '%CHECKSUMS%'"
if errorlevel 1 (
  echo Error: no se pudieron generar checksums.
  popd
  exit /b 1
)

echo.
echo Assets listos en: %RELEASE_DIR%
echo - %SETUP_COPY%
echo - %PORTABLE_ZIP%
echo - %CHECKSUMS%
echo.
echo Sugerencia de publicacion en GitHub Release:
echo gh release create v%APP_VERSION% "%SETUP_COPY%" "%PORTABLE_ZIP%" "%CHECKSUMS%" --title "v%APP_VERSION%" --notes-file docs\release_v1.0.2_notes.md

popd
endlocal
