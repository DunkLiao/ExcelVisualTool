@echo off
setlocal

cd /d "%~dp0"

if /i "%~1"=="dev" goto dev
if /i "%~1"=="tauri" goto tauri
if /i "%~1"=="build" goto build
if /i "%~1"=="help" goto help
if /i "%~1"=="-h" goto help
if /i "%~1"=="--help" goto help

echo [1/3] Checking npm dependencies...
if not exist "node_modules\" (
  echo node_modules not found. Running npm install...
  call npm.cmd install
  if errorlevel 1 goto fail
)

echo [2/3] Building React frontend...
call npm.cmd run build
if errorlevel 1 goto fail

echo [3/3] Checking Tauri/Rust project...
pushd src-tauri
cargo check
set CARGO_EXIT=%ERRORLEVEL%
popd
if not "%CARGO_EXIT%"=="0" goto fail

echo.
echo All checks passed.
exit /b 0

:dev
echo Starting Vite dev server at http://127.0.0.1:1420/
call npm.cmd run dev -- --host 127.0.0.1
exit /b %ERRORLEVEL%

:tauri
echo Starting Tauri desktop app...
call npm.cmd run tauri dev
exit /b %ERRORLEVEL%

:build
call build.bat
exit /b %ERRORLEVEL%

:help
echo Usage:
echo   test.bat         Run npm build and cargo check
echo   test.bat dev     Start Vite dev server
echo   test.bat tauri   Start Tauri desktop app
echo   test.bat build   Build release executable via build.bat
exit /b 0

:fail
echo.
echo Test failed. Check the error output above.
exit /b 1
