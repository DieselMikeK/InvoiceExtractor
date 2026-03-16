@echo off
setlocal EnableDelayedExpansion
title Invoice Extractor - Updating...
color 0A

echo.
echo  ============================================
echo   Invoice Extractor - Updating
echo  ============================================
echo.

REM --- Resolve all paths to absolute ---
pushd "%~dp0"
set "UPDATE_DIR=%CD%"
popd
pushd "%~dp0\.."
set "APP_DIR=%CD%"
popd
pushd "%~dp0\..\.."
set "ROOT_DIR=%CD%"
popd
set "REQUIRED_DIR=%APP_DIR%\required"
set "CONFIG_FILE=%REQUIRED_DIR%\install_config.json"
set "PY_INSTALL_DIR=%UPDATE_DIR%\python"
set "EMBEDDED_PYTHON=%UPDATE_DIR%\python\python.exe"
set "PYTHON_EXE="
echo  [*] ROOT_DIR resolved to: %ROOT_DIR%

REM --- Wait for main app process to fully exit ---
echo  [*] Waiting for app to close...
:WAIT_LOOP
tasklist /fi "imagename eq InvoiceExtractor.exe" 2>nul | find /i "InvoiceExtractor.exe" >nul
if !errorlevel! == 0 (
    timeout /t 1 /nobreak >nul
    goto :WAIT_LOOP
)
echo  [OK] App closed.

REM --- Delete old exe so build can write directly to root ---
if exist "%ROOT_DIR%\InvoiceExtractor.exe" (
    del /f /q "%ROOT_DIR%\InvoiceExtractor.exe"
    echo  [OK] Old exe removed.
)

REM --- Check for our embedded/isolated Python first ---
if exist "%EMBEDDED_PYTHON%" (
    REM Verify tkinter is available — if not, wipe and reinstall
    "%EMBEDDED_PYTHON%" -c "import tkinter" >nul 2>&1
    if !errorlevel! neq 0 (
        echo  [!] Embedded Python missing tkinter — reinstalling...
        rd /s /q "%PY_INSTALL_DIR%" >nul 2>&1
    ) else (
        set "PYTHON_EXE=%EMBEDDED_PYTHON%"
        echo  [OK] Using embedded Python: %EMBEDDED_PYTHON%
        goto :INSTALL_DEPS
    )
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

REM --- No Python found — download and install isolated Python 3.12 (PyInstaller stable) ---
echo  [*] Setting up isolated Python 3.12 (one-time setup)...
echo      A Python installer window will appear — please wait for it to finish.
echo.

set "PY_INSTALLER_URL=https://www.python.org/ftp/python/3.12.9/python-3.12.9-amd64.exe"
set "PY_INSTALLER_TMP=%TEMP%\python_installer_3.12.9.exe"
set "PY_INSTALL_DIR=%UPDATE_DIR%\python"

echo  [*] Downloading Python 3.12.9...
powershell -Command "Invoke-WebRequest -Uri '%PY_INSTALLER_URL%' -OutFile '%PY_INSTALLER_TMP%'" >nul 2>&1
if !errorlevel! neq 0 (
    powershell -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Failed to download Python. Please check your internet connection and try again.', 'Download Failed', 'OK', 'Error')" >nul 2>&1
    exit /b 1
)
echo  [OK] Downloaded.

echo  [*] Installing Python to update\python\ (no admin needed)...
"%PY_INSTALLER_TMP%" /passive InstallAllUsers=0 TargetDir="%PY_INSTALL_DIR%" ^
    Include_pip=1 Include_launcher=0 Include_test=0 Include_doc=0 ^
    Include_tcltk=1
if !errorlevel! neq 0 (
    powershell -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Python installation failed. See app\update\build_log.txt for details.', 'Install Failed', 'OK', 'Error')" >nul 2>&1
    exit /b 1
)
del "%PY_INSTALLER_TMP%" >nul 2>&1

if not exist "%PY_INSTALL_DIR%\python.exe" (
    REM Installer may have placed it one level deeper — search for it
    for /r "%PY_INSTALL_DIR%" %%F in (python.exe) do (
        if not defined PYTHON_EXE set "PYTHON_EXE=%%F"
    )
    if not defined PYTHON_EXE (
        powershell -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Python installed but python.exe not found. Please delete app\update\python\ and try again, or contact support.', 'Install Failed', 'OK', 'Error')" >nul 2>&1
        exit /b 1
    )
    echo  [OK] Found python.exe at: !PYTHON_EXE!
) else (
    set "PYTHON_EXE=%PY_INSTALL_DIR%\python.exe"
)
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
"!PYTHON_EXE!" -m pip install --upgrade pip --quiet
"!PYTHON_EXE!" -m pip install -r "%APP_DIR%\requirements.txt" --quiet
"!PYTHON_EXE!" -m pip install pyinstaller --quiet
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
"!PYTHON_EXE!" -m PyInstaller InvoiceExtractor.spec --noconfirm --distpath "%ROOT_DIR%" >"%UPDATE_DIR%\build_log.txt" 2>&1
if !errorlevel! neq 0 (
    powershell -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Build failed. See app\update\build_log.txt for details.', 'Update Failed', 'OK', 'Error')" >nul 2>&1
    exit /b 1
)

REM --- Clean up PyInstaller build folder ---
if exist "%APP_DIR%\build" rd /s /q "%APP_DIR%\build" >nul 2>&1

REM --- Clean up any leftover zip files in root ---
del /q "%ROOT_DIR%\*.zip" >nul 2>&1

echo  [OK] Update complete!

REM --- Relaunch ---
echo  [*] Relaunching Invoice Extractor...
start "" "!ROOT_DIR!\InvoiceExtractor.exe"
exit
