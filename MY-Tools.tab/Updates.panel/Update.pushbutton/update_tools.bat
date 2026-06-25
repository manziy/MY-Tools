@echo off
setlocal EnableExtensions EnableDelayedExpansion

REM ==================================================
REM SETTINGS - EDIT THESE THREE LINES ONLY
REM ==================================================

set "EXT_NAME=MY-Tools.extension"
set "GITHUB_OWNER=manziy"
set "GITHUB_REPO=MY-Tools"

REM ==================================================
REM PATHS
REM ==================================================

set "PYREVIT_EXT_DIR=%APPDATA%\pyRevit\Extensions"
set "TEMP_DIR=%TEMP%\pyRevit_Extension_Install_%RANDOM%%RANDOM%"
set "ZIP_FILE=%TEMP_DIR%\repo.zip"
set "EXTRACT_DIR=%TEMP_DIR%\extracted"
set "TARGET_DIR=%PYREVIT_EXT_DIR%\%EXT_NAME%"
set "LOG_FILE=%TEMP%\pyRevit_Install_Log.txt"

echo ================================================== > "%LOG_FILE%"
echo pyRevit Extension Install Log >> "%LOG_FILE%"
echo ================================================== >> "%LOG_FILE%"
echo Extension name: %EXT_NAME% >> "%LOG_FILE%"
echo GitHub repo: https://github.com/%GITHUB_OWNER%/%GITHUB_REPO% >> "%LOG_FILE%"
echo Target folder: %TARGET_DIR% >> "%LOG_FILE%"
echo. >> "%LOG_FILE%"

call :main
set "FINAL_RESULT=%ERRORLEVEL%"

if "%FINAL_RESULT%"=="0" (
    echo. >> "%LOG_FILE%"
    echo SUCCESS: Installation complete. >> "%LOG_FILE%"
) else (
    echo. >> "%LOG_FILE%"
    echo FAILED: Installation failed. Error code: %FINAL_RESULT% >> "%LOG_FILE%"
)

exit /b %FINAL_RESULT%


:main

REM ==================================================
REM CREATE FOLDERS
REM ==================================================

echo Creating folders... >> "%LOG_FILE%"

if not exist "%PYREVIT_EXT_DIR%" (
    mkdir "%PYREVIT_EXT_DIR%" >> "%LOG_FILE%" 2>&1
)

if exist "%TEMP_DIR%" (
    rmdir /s /q "%TEMP_DIR%" >> "%LOG_FILE%" 2>&1
)

mkdir "%TEMP_DIR%" >> "%LOG_FILE%" 2>&1
mkdir "%EXTRACT_DIR%" >> "%LOG_FILE%" 2>&1

if not exist "%TEMP_DIR%" (
    echo ERROR: Could not create temp folder. >> "%LOG_FILE%"
    exit /b 1
)

if not exist "%EXTRACT_DIR%" (
    echo ERROR: Could not create extract folder. >> "%LOG_FILE%"
    exit /b 1
)

REM ==================================================
REM GET DEFAULT BRANCH FROM GITHUB
REM ==================================================

echo Getting default branch from GitHub... >> "%LOG_FILE%"

set "DEFAULT_BRANCH="

powershell -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; $repo='https://api.github.com/repos/%GITHUB_OWNER%/%GITHUB_REPO%'; $data=Invoke-RestMethod -Uri $repo; $data.default_branch" > "%TEMP_DIR%\branch.txt" 2>> "%LOG_FILE%"

if errorlevel 1 (
    echo ERROR: Failed to contact GitHub API. >> "%LOG_FILE%"
    exit /b 2
)

for /f "usebackq delims=" %%B in ("%TEMP_DIR%\branch.txt") do (
    set "DEFAULT_BRANCH=%%B"
)

if "%DEFAULT_BRANCH%"=="" (
    echo ERROR: Could not detect default branch. >> "%LOG_FILE%"
    exit /b 3
)

set "REPO_ZIP_URL=https://github.com/%GITHUB_OWNER%/%GITHUB_REPO%/archive/refs/heads/%DEFAULT_BRANCH%.zip"

echo Default branch detected: %DEFAULT_BRANCH% >> "%LOG_FILE%"
echo Download URL: %REPO_ZIP_URL% >> "%LOG_FILE%"
echo. >> "%LOG_FILE%"

REM ==================================================
REM DOWNLOAD ZIP FROM GITHUB
REM ==================================================

echo Downloading latest version from GitHub... >> "%LOG_FILE%"

powershell -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '%REPO_ZIP_URL%' -OutFile '%ZIP_FILE%'" >> "%LOG_FILE%" 2>&1

if errorlevel 1 (
    echo ERROR: Failed to download repo zip. >> "%LOG_FILE%"
    exit /b 4
)

if not exist "%ZIP_FILE%" (
    echo ERROR: Zip file was not downloaded. >> "%LOG_FILE%"
    exit /b 5
)

REM ==================================================
REM EXTRACT ZIP
REM ==================================================

echo Extracting files... >> "%LOG_FILE%"

powershell -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; Expand-Archive -Path '%ZIP_FILE%' -DestinationPath '%EXTRACT_DIR%' -Force" >> "%LOG_FILE%" 2>&1

if errorlevel 1 (
    echo ERROR: Failed to extract zip file. >> "%LOG_FILE%"
    exit /b 6
)

REM ==================================================
REM FIND SOURCE EXTENSION FOLDER
REM ==================================================
REM Supports:
REM 1. Repo contains MY-Tools.extension
REM 2. Repo root itself is the extension folder

echo Finding extension folder... >> "%LOG_FILE%"

set "SOURCE_EXT_DIR="

powershell -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; $extName='%EXT_NAME%'; $extract='%EXTRACT_DIR%'; $found=Get-ChildItem -Path $extract -Directory -Recurse | Where-Object { $_.Name -ieq $extName } | Select-Object -First 1; if ($found) { $found.FullName } else { $root=Get-ChildItem -Path $extract -Directory | Select-Object -First 1; if ($root) { $root.FullName } }" > "%TEMP_DIR%\source.txt" 2>> "%LOG_FILE%"

if errorlevel 1 (
    echo ERROR: Failed while searching extension folder. >> "%LOG_FILE%"
    exit /b 7
)

for /f "usebackq delims=" %%D in ("%TEMP_DIR%\source.txt") do (
    set "SOURCE_EXT_DIR=%%D"
)

if "%SOURCE_EXT_DIR%"=="" (
    echo ERROR: Could not find source extension folder. >> "%LOG_FILE%"
    exit /b 8
)

echo Source folder found: %SOURCE_EXT_DIR% >> "%LOG_FILE%"
echo. >> "%LOG_FILE%"

REM ==================================================
REM INSTALL / UPDATE EXTENSION
REM ==================================================

echo Removing old extension folder if it exists... >> "%LOG_FILE%"

if exist "%TARGET_DIR%" (
    rmdir /s /q "%TARGET_DIR%" >> "%LOG_FILE%" 2>&1
)

echo Copying new extension files... >> "%LOG_FILE%"

mkdir "%TARGET_DIR%" >> "%LOG_FILE%" 2>&1

robocopy "%SOURCE_EXT_DIR%" "%TARGET_DIR%" /E /NFL /NDL /NJH /NJS /NC /NS >> "%LOG_FILE%" 2>&1

set "ROBOCOPY_RESULT=%ERRORLEVEL%"

REM Robocopy exit codes 0-7 are okay.
if %ROBOCOPY_RESULT% GEQ 8 (
    echo ERROR: Robocopy failed. Robocopy code: %ROBOCOPY_RESULT% >> "%LOG_FILE%"
    exit /b 9
)

REM ==================================================
REM VERIFY INSTALLATION
REM ==================================================

if not exist "%TARGET_DIR%" (
    echo ERROR: Target folder does not exist after copy. >> "%LOG_FILE%"
    exit /b 10
)

echo Installed to: %TARGET_DIR% >> "%LOG_FILE%"

REM ==================================================
REM CLEAN UP
REM ==================================================

echo Cleaning temporary files... >> "%LOG_FILE%"

rmdir /s /q "%TEMP_DIR%" >> "%LOG_FILE%" 2>&1

exit /b 0