@echo off
setlocal EnableDelayedExpansion
title Invoice Extractor - Updating...
color 0A

echo.
echo  ============================================
echo   Invoice Extractor - Updating
echo  ============================================
echo.

REM --- Paths ---
set "UPDATE_DIR=%~dp0"
REM Remove trailing backslash from UPDATE_DIR
if "%UPDATE_DIR:~-1%"=="\" set "UPDATE_DIR=%UPDATE_DIR:~0,-1%"
set "APP_DIR=%UPDATE_DIR%\.."
set "ROOT_DIR=%APP_DIR%\.."
set "REQUIRED_DIR=%APP_DIR%\required"
set "CONFIG_FILE=%REQUIRED_DIR%\install_config.json"
set "EMBEDDED_PYTHON=%UPDATE_DIR%\python\python.exe"
set "PYTHON_EXE="

REM --- Wait for main app to close ---
echo  [*] Waiting for app to close...
timeout /t 2 /nobreak >nul

REM --- Check for our embedded/isolated Python first ---
if exist "%EMBEDDED_PYTHON%" (
    set "PYTHON_EXE=%EMBEDDED_PYTHON%"
    echo  [OK] Using embedded Python: %EMBEDDED_PYTHON%
    goto :INSTALL_DEPS
)

REM --- Check config for a previously saved path ---
if exist "%CONFIG_FILE%" (
    for /f "tokens=2 delims=:, " %%A in ('findstr /i "python_exe" "%CONFIG_FILE%"') do (
        set "CANDIDATE=%%~A"
        set "CANDIDATE=!CANDIDATE:"=!"
        if exist "!CANDIDATE!" set "PYTHON_EXE=!CANDIDATE!"
    )
)

if defined PYTHON_EXE (
    echo  [OK] Using saved Python: %PYTHON_EXE%
    goto :INSTALL_DEPS
)

REM --- No Python found — download and install isolated Python 3.14 ---
echo  [*] Setting up isolated Python 3.14 (one-time, ~25MB)...
echo      This will only happen once.
echo.

set "PY_INSTALLER_URL=https://www.python.org/ftp/python/3.14.0/python-3.14.0-amd64.exe"
set "PY_INSTALLER_TMP=%TEMP%\python_installer_3.14.0.exe"
set "PY_INSTALL_DIR=%UPDATE_DIR%\python"

echo  [*] Downloading Python 3.14.0...
powershell -Command "Invoke-WebRequest -Uri '%PY_INSTALLER_URL%' -OutFile '%PY_INSTALLER_TMP%'" >nul 2>&1
if !errorlevel! neq 0 (
    powershell -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Failed to download Python. Please check your internet connection and try again.', 'Download Failed', 'OK', 'Error')" >nul 2>&1
    exit /b 1
)
echo  [OK] Downloaded.

echo  [*] Installing Python to update\python\ (no admin needed)...
"%PY_INSTALLER_TMP%" /quiet InstallAllUsers=0 TargetDir="%PY_INSTALL_DIR%" ^
    Include_pip=1 Include_launcher=0 Include_test=0 Include_doc=0 ^
    SimpleInstall=1 SimpleInstallDescription="Invoice Extractor Python"
if !errorlevel! neq 0 (
    powershell -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Python installation failed. See update\build_log.txt for details.', 'Install Failed', 'OK', 'Error')" >nul 2>&1
    exit /b 1
)
del "%PY_INSTALLER_TMP%" >nul 2>&1

if not exist "%PY_INSTALL_DIR%\python.exe" (
    powershell -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Python installed but python.exe not found. Please contact support.', 'Install Failed', 'OK', 'Error')" >nul 2>&1
    exit /b 1
)

set "PYTHON_EXE=%PY_INSTALL_DIR%\python.exe"
echo  [OK] Python installed at: %PYTHON_EXE%

REM --- Save to config ---
if not exist "%REQUIRED_DIR%" mkdir "%REQUIRED_DIR%"
(
    echo {
    echo   "python_exe": "%PYTHON_EXE:\=\\%",
    echo   "installed": "%DATE% %TIME%",
    echo   "repo": "https://github.com/DieselMikeK/InvoiceExtractor"
    echo }
) > "%CONFIG_FILE%"

:INSTALL_DEPS
echo.
echo  [*] Installing/updating dependencies...
"%PYTHON_EXE%" -m pip install --upgrade pip --quiet
"%PYTHON_EXE%" -m pip install -r "%APP_DIR%\requirements.txt" --quiet
"%PYTHON_EXE%" -m pip install pyinstaller --quiet
echo  [OK] Dependencies ready.

REM --- Pull latest source ---
echo.
echo  [*] Downloading latest source from GitHub...
cd /d "%APP_DIR%"

where git >nul 2>&1
if !errorlevel! == 0 (
    git pull origin main >nul 2>&1
    if !errorlevel! == 0 (
        echo  [OK] Source updated via git.
        goto :BUILD
    )
)

REM No git or pull failed — download zip from GitHub
set "ZIP_URL=https://github.com/DieselMikeK/InvoiceExtractor/archive/refs/heads/main.zip"
set "ZIP_TMP=%TEMP%\InvoiceExtractor_src.zip"
set "EXTRACT_TMP=%TEMP%\InvoiceExtractor_src"

echo  [*] Downloading source zip...
powershell -Command "Invoke-WebRequest -Uri '%ZIP_URL%' -OutFile '%ZIP_TMP%'" >nul 2>&1
if !errorlevel! neq 0 (
    powershell -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Source download failed. Please check your internet connection.', 'Update Failed', 'OK', 'Error')" >nul 2>&1
    exit /b 1
)
if exist "%EXTRACT_TMP%" rd /s /q "%EXTRACT_TMP%"
powershell -Command "Expand-Archive -Path '%ZIP_TMP%' -DestinationPath '%EXTRACT_TMP%' -Force" >nul 2>&1
xcopy /e /y /q "%EXTRACT_TMP%\InvoiceExtractor-main\app\*" "%APP_DIR%\" >nul
rd /s /q "%EXTRACT_TMP%" >nul 2>&1
del "%ZIP_TMP%" >nul 2>&1
echo  [OK] Source downloaded and extracted.

:BUILD
echo.
echo  [*] Building new version (1-2 minutes)...
echo.
cd /d "%APP_DIR%"
"%PYTHON_EXE%" -m PyInstaller InvoiceExtractor.spec --noconfirm >"%UPDATE_DIR%\build_log.txt" 2>&1
if !errorlevel! neq 0 (
    powershell -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Build failed. See app\update\build_log.txt for details.', 'Update Failed', 'OK', 'Error')" >nul 2>&1
    exit /b 1
)

REM --- Overwrite exe at root ---
copy /y "%APP_DIR%\dist\InvoiceExtractor.exe" "%ROOT_DIR%\InvoiceExtractor.exe" >nul

REM --- Clean up PyInstaller build artifacts ---
if exist "%APP_DIR%\dist" rd /s /q "%APP_DIR%\dist" >nul 2>&1
if exist "%APP_DIR%\build" rd /s /q "%APP_DIR%\build" >nul 2>&1

REM --- Clean up any leftover zip files in root ---
del /q "%ROOT_DIR%\*.zip" >nul 2>&1

echo  [OK] Update complete!

REM --- Relaunch ---
echo  [*] Relaunching Invoice Extractor...
start "" "%ROOT_DIR%\InvoiceExtractor.exe"
exit
