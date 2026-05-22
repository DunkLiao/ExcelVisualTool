@echo off
setlocal

cd /d "%~dp0"

echo Building Tauri release executable...
call npm.cmd run tauri build -- --no-bundle
if errorlevel 1 goto fail

if not exist "portableapps\" mkdir "portableapps"

copy /Y "src-tauri\target\release\excel-visual-tool.exe" "portableapps\excel-visual-tool.exe" >nul
if errorlevel 1 goto fail

echo.
echo Build completed.
echo Portable exe:
echo   %CD%\portableapps\excel-visual-tool.exe
echo.
echo Installer bundles are skipped by default because this command is for portable exe testing.
exit /b 0

:fail
echo.
echo Build failed. Check the error output above.
exit /b 1
