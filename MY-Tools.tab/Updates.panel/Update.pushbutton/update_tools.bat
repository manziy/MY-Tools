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
set "TEMP_DIR=%TEMP%\pyRevit_Extension_Install_%RANDOM%"
set "ZIP_FILE=%TEMP_DIR%\repo.zip"
set "EXTRACT_DIR=%TEMP_DIR%\extracted"
set "TARGET_DIR=%PYREVIT_EXT_DIR%\%EXT_NAME%"
set "LOG_FILE=%TEMP%\pyRevit_Install_Log.txt"

echo ================================================== > "%LOG_FILE%"
echo pyRevit Extension Install Log >> "%LOG_FILE%"
echo ================================================== >> "%LOG_FILE%"
echo. >> "%LOG_FILE%"

echo.
echo ==================================================
echo Installing / Updating pyRevit Extension
echo ==================================================
echo.
echo Extension name:
echo %EXT_NAME%
echo.
echo GitHub repo:
echo https://github.com/%GITHUB_OWNER%/%GITHUB_REPO%
echo.
echo Target folder:
echo %TARGET_DIR%
echo.
echo Log file:
echo %LOG_FILE%
echo.

call :main
set "FINAL_RESULT=%ERRORLEVEL%"

echo.
echo ==================================================
if "%FINAL_RESULT%"=="0" (
    echo.
    echo SUCCESS!
    echo The pyRevit extension has been installed / updated successfully.
    echo.
    echo Installed to:
    echo %TARGET_DIR%
    echo.
    echo Please restart Revit, or click Reload in pyRevit.
    echo.
    echo SUCCESS: The pyRevit extension has been installed / updated successfully. >> "%LOG_FILE%"
    echo Installed to: %TARGET_DIR% >> "%LOG_FILE%"

    powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('pyRevit extension installed / updated successfully. Please restart Revit or click Reload in pyRevit.', 'Installation Complete')"
) else (
    echo.
    echo FAILED.
    echo Installation failed. Error code: %FINAL_RESULT%
    echo.
    echo Please check the log file:
    echo %LOG_FILE%
    echo.
    echo Installation failed. Error code: %FINAL_RESULT% >> "%LOG_FILE%"

    powershell -NoProfile -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName PresentationFramework; [System.Windows.MessageBox]::Show('Installation failed. Please check the log file for details.', 'Installation Failed')"
)
echo ==================================================
echo.
echo Log file:
echo %LOG_FILE%
echo.

pause
exit /b %FINAL_RESULT%


:main

echo Extension name: %EXT_NAME% >> "%LOG_FILE%"
echo GitHub repo: https://github.com/%GITHUB_OWNER%/%GITHUB_REPO% >> "%LOG_FILE%"
echo Target folder: %TARGET_DIR% >> "%LOG_FILE%"
echo. >> "%LOG_FILE%"

REM ==================================================
REM CREATE FOLDERS
REM ==================================================

echo Creating folders if needed...
echo Creating folders if needed... >> "%LOG_FILE%"

if not exist "%PYREVIT_EXT_DIR%" (
    mkdir "%PYREVIT_EXT_DIR%" >> "%LOG_FILE%" 2>&1
)

if exist "%TEMP_DIR%" (
    rmdir /s /q "%TEMP_DIR%" >> "%LOG_FILE%" 2>&1
)

mkdir "%TEMP_DIR%" >> "%LOG_FILE%" 2>&1
mkdir "%EXTRACT_DIR%" >> "%LOG_FILE%" 2>&1

if not exist "%TEMP_DIR%" (
    echo ERROR: Could not create temp folder.
    echo ERROR: Could not create temp folder. >> "%LOG_FILE%"
    exit /b 1
)

if not exist "%EXTRACT_DIR%" (
    echo ERROR: Could not create extract folder.
    echo ERROR: Could not create extract folder. >> "%LOG_FILE%"
    exit /b 1
)

REM ==================================================
REM GET DEFAULT BRANCH FROM GITHUB
REM ==================================================

echo.
echo Getting default branch from GitHub...
echo Getting default branch from GitHub... >> "%LOG_FILE%"

set "DEFAULT_BRANCH="

powershell -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; $repo='https://api.github.com/repos/%GITHUB_OWNER%/%GITHUB_REPO%'; $data=Invoke-RestMethod -Uri $repo; $data.default_branch" > "%TEMP_DIR%\branch.txt" 2>> "%LOG_FILE%"

if errorlevel 1 (
    echo ERROR: Failed to contact GitHub API.
    echo ERROR: Failed to contact GitHub API. >> "%LOG_FILE%"
    exit /b 2
)

for /f "usebackq delims=" %%B in ("%TEMP_DIR%\branch.txt") do (
    set "DEFAULT_BRANCH=%%B"
)

if "%DEFAULT_BRANCH%"=="" (
    echo ERROR: Could not detect default branch.
    echo ERROR: Could not detect default branch. >> "%LOG_FILE%"
    exit /b 3
)

set "REPO_ZIP_URL=https://github.com/%GITHUB_OWNER%/%GITHUB_REPO%/archive/refs/heads/%DEFAULT_BRANCH%.zip"

echo Default branch detected:
echo %DEFAULT_BRANCH%
echo.
echo Download URL:
echo %REPO_ZIP_URL%
echo.

echo Default branch detected: %DEFAULT_BRANCH% >> "%LOG_FILE%"
echo Download URL: %REPO_ZIP_URL% >> "%LOG_FILE%"
echo. >> "%LOG_FILE%"

REM ==================================================
REM DOWNLOAD ZIP FROM GITHUB
REM ==================================================

echo Downloading latest version from GitHub...
echo Downloading latest version from GitHub... >> "%LOG_FILE%"

powershell -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; [Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12; Invoke-WebRequest -Uri '%REPO_ZIP_URL%' -OutFile '%ZIP_FILE%'" >> "%LOG_FILE%" 2>&1

if errorlevel 1 (
    echo ERROR: Failed to download repo zip.
    echo ERROR: Failed to download repo zip. >> "%LOG_FILE%"
    exit /b 4
)

if not exist "%ZIP_FILE%" (
    echo ERROR: Zip file was not downloaded.
    echo ERROR: Zip file was not downloaded. >> "%LOG_FILE%"
    exit /b 5
)

REM ==================================================
REM EXTRACT ZIP
REM ==================================================

echo.
echo Extracting files...
echo Extracting files... >> "%LOG_FILE%"

powershell -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; Expand-Archive -Path '%ZIP_FILE%' -DestinationPath '%EXTRACT_DIR%' -Force" >> "%LOG_FILE%" 2>&1

if errorlevel 1 (
    echo ERROR: Failed to extract zip file.
    echo ERROR: Failed to extract zip file. >> "%LOG_FILE%"
    exit /b 6
)

REM ==================================================
REM FIND SOURCE EXTENSION FOLDER
REM ==================================================
REM This supports two repo structures:
REM
REM Option 1:
REM Repo contains:
REM     MY-Tools.extension
REM         MY-Tools.tab
REM
REM Option 2:
REM Repo root itself is the extension folder and contains:
REM     MY-Tools.tab
REM

echo.
echo Finding extension folder...
echo Finding extension folder... >> "%LOG_FILE%"

set "SOURCE_EXT_DIR="

powershell -NoProfile -ExecutionPolicy Bypass -Command "$ErrorActionPreference='Stop'; $extName='%EXT_NAME%'; $extract='%EXTRACT_DIR%'; $found=Get-ChildItem -Path $extract -Directory -Recurse | Where-Object { $_.Name -ieq $extName } | Select-Object -First 1; if ($found) { $found.FullName } else { $root=Get-ChildItem -Path $extract -Directory | Select-Object -First 1; if ($root) { $root.FullName } }" > "%TEMP_DIR%\source.txt" 2>> "%LOG_FILE%"

if errorlevel 1 (
    echo ERROR: Failed while searching extension folder.
    echo ERROR: Failed while searching extension folder. >> "%LOG_FILE%"
    exit /b 7
)

for /f "usebackq delims=" %%D in ("%TEMP_DIR%\source.txt") do (
    set "SOURCE_EXT_DIR=%%D"
)

if "%SOURCE_EXT_DIR%"=="" (
    echo ERROR: Could not find source extension folder.
    echo ERROR: Could not find source extension folder. >> "%LOG_FILE%"
    exit /b 8
)

echo Source folder found:
echo %SOURCE_EXT_DIR%
echo.

echo Source folder found: %SOURCE_EXT_DIR% >> "%LOG_FILE%"
echo. >> "%LOG_FILE%"

REM ==================================================
REM INSTALL / UPDATE EXTENSION
REM ==================================================

echo Removing old extension folder if it exists...
echo Removing old extension folder if it exists... >> "%LOG_FILE%"

if exist "%TARGET_DIR%" (
    rmdir /s /q "%TARGET_DIR%" >> "%LOG_FILE%" 2>&1
)

echo Copying new extension files...
echo Copying new extension files... >> "%LOG_FILE%"

mkdir "%TARGET_DIR%" >> "%LOG_FILE%" 2>&1

robocopy "%SOURCE_EXT_DIR%" "%TARGET_DIR%" /E /NFL /NDL /NJH /NJS /NC /NS >> "%LOG_FILE%" 2>&1

set "ROBOCOPY_RESULT=%ERRORLEVEL%"

REM Robocopy exit codes 0-7 are okay.
if %ROBOCOPY_RESULT% GEQ 8 (
    echo ERROR: Robocopy failed.
    echo ERROR: Robocopy failed. >> "%LOG_FILE%"
    exit /b 9
)

REM ==================================================
REM VERIFY INSTALLATION
REM ==================================================

if not exist "%TARGET_DIR%" (
    echo ERROR: Target folder does not exist after copy.
    echo ERROR: Target folder does not exist after copy. >> "%LOG_FILE%"
    exit /b 10
)

echo.
echo Installed to:
echo %TARGET_DIR%
echo.

echo Installed to: %TARGET_DIR% >> "%LOG_FILE%"

REM ==================================================
REM CLEAN UP
REM ==================================================

echo Cleaning temporary files...
echo Cleaning temporary files... >> "%LOG_FILE%"

rmdir /s /q "%TEMP_DIR%" >> "%LOG_FILE%" 2>&1

exit /b 0