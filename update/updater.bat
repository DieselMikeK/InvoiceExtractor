@echo off
setlocal EnableDelayedExpansion
title Invoice Extractor - Updating...
color 0A

echo.
echo  ============================================
echo   Invoice Extractor - Updating
echo  ============================================
echo.

REM ROOT_DIR = two levels up from app/update/ (where InvoiceExtractor.exe lives)
set "UPDATE_DIR=%~dp0"
set "APP_DIR=%UPDATE_DIR%.."
set "ROOT_DIR=%APP_DIR%\.."
set "CONFIG_FILE=%APP_DIR%\required\install_config.json"
set "PYTHON_EXE="

REM --- Wait for main app to close ---
echo  [*] Waiting for app to close...
timeout /t 2 /nobreak >nul

REM --- Try to read Python path from config ---
if exist "%CONFIG_FILE%" (
    for /f "tokens=2 delims=:, " %%A in ('findstr /i "python_exe" "%CONFIG_FILE%"') do (
        set "CANDIDATE=%%~A"
        set "CANDIDATE=!CANDIDATE:"=!"
        if exist "!CANDIDATE!" set "PYTHON_EXE=!CANDIDATE!"
    )
)

REM --- Auto-detect Python if not configured ---
if not defined PYTHON_EXE (
    echo  [*] Auto-detecting Python...

    for /d %%D in ("%LOCALAPPDATA%\Python\Python3*") do (
        if exist "%%D\python.exe" if not defined PYTHON_EXE set "PYTHON_EXE=%%D\python.exe"
    )
    for /d %%D in ("%LOCALAPPDATA%\Python\pythoncore*") do (
        if exist "%%D\python.exe" if not defined PYTHON_EXE set "PYTHON_EXE=%%D\python.exe"
    )
    if not defined PYTHON_EXE (
        for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*") do (
            if exist "%%D\python.exe" if not defined PYTHON_EXE set "PYTHON_EXE=%%D\python.exe"
        )
    )
    if not defined PYTHON_EXE (
        for /d %%D in ("C:\Python3*") do (
            if exist "%%D\python.exe" if not defined PYTHON_EXE set "PYTHON_EXE=%%D\python.exe"
        )
    )
    if not defined PYTHON_EXE (
        where python >nul 2>&1
        if !errorlevel! == 0 (
            for /f "delims=" %%P in ('where python 2^>nul') do (
                if not defined PYTHON_EXE set "PYTHON_EXE=%%P"
            )
        )
    )

    if not defined PYTHON_EXE (
        powershell -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Python was not found on this machine.`n`nPlease install Python 3.10 or newer from python.org, then run the app again.', 'Python Required', 'OK', 'Warning')" >nul 2>&1
        exit /b 1
    )

    REM Save config for next time
    if not exist "%APP_DIR%\required" mkdir "%APP_DIR%\required"
    (
        echo {
        echo   "python_exe": "%PYTHON_EXE:\=\\%",
        echo   "installed": "%DATE% %TIME%",
        echo   "repo": "https://github.com/DieselMikeK/InvoiceExtractor"
        echo }
    ) > "%CONFIG_FILE%"
    echo  [OK] Detected and saved Python: %PYTHON_EXE%
)

echo  [OK] Using Python: %PYTHON_EXE%

REM --- Install/update dependencies ---
echo.
echo  [*] Installing any new dependencies...
"%PYTHON_EXE%" -m pip install -r "%APP_DIR%\requirements.txt" --quiet
"%PYTHON_EXE%" -m pip install pyinstaller --quiet

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

powershell -Command "Invoke-WebRequest -Uri '%ZIP_URL%' -OutFile '%ZIP_TMP%'" >nul 2>&1
if !errorlevel! neq 0 (
    powershell -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Download failed. Please check your internet connection and try again.', 'Update Failed', 'OK', 'Error')" >nul 2>&1
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
"%PYTHON_EXE%" -m PyInstaller InvoiceExtractor.spec --noconfirm >"%APP_DIR%\update\build_log.txt" 2>&1
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
