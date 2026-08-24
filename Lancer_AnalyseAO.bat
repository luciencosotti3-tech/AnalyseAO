@echo off
setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"
title AnalyseAO - Demarrage

rem =======================================================
rem CONFIGURATION
rem =======================================================
set "PYTHON_VERSION=3.14"
set "PYTHON_WINGET_ID=Python.Python.3.14"
set "APP_URL=http://127.0.0.1:8008"
set "HEALTH_URL=http://127.0.0.1:8008/api/health"
set "ROOT=%CD%"
set "BACKEND_DIR=%ROOT%\backend"
set "RUNTIME_DIR=%BACKEND_DIR%\runtime"
set "UPLOADS_DIR=%RUNTIME_DIR%\uploads"
set "OUTPUTS_DIR=%RUNTIME_DIR%\outputs"
set "REQUIREMENTS=%BACKEND_DIR%\requirements.txt"
set "VENV_DIR=%ROOT%\.venv"
set "PYTHON_EXE=%VENV_DIR%\Scripts\python.exe"
set "REQUIREMENTS_STAMP=%VENV_DIR%\.analyseao_requirements.sha256"
set "PYTHON_STAMP=%VENV_DIR%\.analyseao_python_version"
set "LOG_DIR=%RUNTIME_DIR%\logs"
set "FLASK_LOG=%LOG_DIR%\flask_runtime.log"
set "FLASK_ERR=%LOG_DIR%\flask_runtime.err.log"
set "FLASK_PID_FILE=%RUNTIME_DIR%\flask.pid"

echo ========================================
echo Lancement de AnalyseAO
echo ========================================
echo [INFO] Racine : %ROOT%

rem =======================================================
rem VERIFICATION DE LA STRUCTURE DU PROJET
rem =======================================================
if not exist "%BACKEND_DIR%\app.py" goto :fatal_app_missing
if not exist "%BACKEND_DIR%\side_by_side.py" goto :fatal_side_missing
if not exist "%BACKEND_DIR%\offer_analysis.py" goto :fatal_offer_missing
if not exist "%BACKEND_DIR%\direct_anomaly_engine.py" goto :fatal_direct_missing
if not exist "%ROOT%\frontend\index.html" goto :fatal_frontend_missing
if not exist "%REQUIREMENTS%" goto :fatal_requirements_missing

rem =======================================================
rem INSTANCE DEJA ACTIVE : NE RIEN NETTOYER
rem =======================================================
call :health_ready
if not errorlevel 1 goto :open_browser

rem Un autre programme ne doit pas occuper le port 8008.
powershell -NoProfile -ExecutionPolicy Bypass -Command "$c=Get-NetTCPConnection -LocalPort 8008 -State Listen -ErrorAction SilentlyContinue; if($c){Write-Host ('[ERREUR] Port 8008 deja occupe par le PID ' + $c[0].OwningProcess); exit 1}else{exit 0}"
if errorlevel 1 goto :fatal

rem =======================================================
rem NETTOYAGE DU RUNTIME
rem Uniquement lorsque le serveur AnalyseAO n'est pas actif.
rem =======================================================
echo [INFO] Nettoyage des fichiers temporaires...
call :reset_runtime_folder "%UPLOADS_DIR%"
if errorlevel 1 goto :fatal_cleanup_uploads
call :reset_runtime_folder "%OUTPUTS_DIR%"
if errorlevel 1 goto :fatal_cleanup_outputs
if not exist "%LOG_DIR%" mkdir "%LOG_DIR%" >nul 2>&1
if exist "%FLASK_LOG%" del /f /q "%FLASK_LOG%" >nul 2>&1
if exist "%FLASK_ERR%" del /f /q "%FLASK_ERR%" >nul 2>&1
if exist "%FLASK_PID_FILE%" del /f /q "%FLASK_PID_FILE%" >nul 2>&1
echo [OK] Runtime nettoye.

rem =======================================================
rem PYTHON 3.14 : DETECTION / INSTALLATION
rem =======================================================
set "BOOTSTRAP_PYTHON="
set "BOOTSTRAP_PYTHON_ARGS="
call :detect_python_314
if not defined BOOTSTRAP_PYTHON (
    echo [SETUP] Python %PYTHON_VERSION% est absent. Installation automatique...
    call :require_winget
    if errorlevel 1 goto :fatal_winget
    winget install --exact --id "%PYTHON_WINGET_ID%" --source winget --scope user --accept-package-agreements --accept-source-agreements --disable-interactivity
    if errorlevel 1 (
        winget list --exact --id "%PYTHON_WINGET_ID%" >nul 2>&1
        if errorlevel 1 goto :fatal_python_install
    )
    call :refresh_path
    call :detect_python_314
)
if not defined BOOTSTRAP_PYTHON goto :fatal_python_path
echo [OK] Python %PYTHON_VERSION% detecte.

rem =======================================================
rem ENVIRONNEMENT VIRTUEL : CREATION OU RECONSTRUCTION
rem =======================================================
set "REBUILD_VENV=0"
if not exist "%PYTHON_EXE%" set "REBUILD_VENV=1"
if exist "%PYTHON_EXE%" (
    "%PYTHON_EXE%" -c "import sys; raise SystemExit(0 if sys.version_info[:2] == (3,14) else 1)" >nul 2>&1
    if errorlevel 1 set "REBUILD_VENV=1"
)
if "!REBUILD_VENV!"=="1" (
    if exist "%VENV_DIR%" (
        echo [SETUP] Ancien environnement incompatible ou endommage. Reconstruction...
        rmdir /s /q "%VENV_DIR%"
        if exist "%VENV_DIR%" goto :fatal_venv_remove
    ) else (
        echo [SETUP] Creation de l'environnement Python local...
    )
    if defined BOOTSTRAP_PYTHON_ARGS (
        "%BOOTSTRAP_PYTHON%" %BOOTSTRAP_PYTHON_ARGS% -m venv "%VENV_DIR%"
    ) else (
        "%BOOTSTRAP_PYTHON%" -m venv "%VENV_DIR%"
    )
    if errorlevel 1 goto :fatal_venv_create
)
if not exist "%PYTHON_EXE%" goto :fatal_venv_create
"%PYTHON_EXE%" -c "import sys; raise SystemExit(0 if sys.version_info[:2] == (3,14) else 1)" >nul 2>&1
if errorlevel 1 goto :fatal_venv_version
for /f "delims=" %%V in ('"%PYTHON_EXE%" -c "import sys; print('.'.join(map(str,sys.version_info[:3])))"') do set "ACTIVE_PYTHON_VERSION=%%V"
>"%PYTHON_STAMP%" echo !ACTIVE_PYTHON_VERSION!
echo [OK] Environnement local Python !ACTIVE_PYTHON_VERSION! disponible.

rem =======================================================
rem DEPENDANCES : INSTALLATION UNIQUEMENT SI NECESSAIRE
rem =======================================================
set "CURRENT_REQUIREMENTS_HASH="
set "INSTALLED_REQUIREMENTS_HASH="
set "INSTALL_REQUIREMENTS=0"
for /f "usebackq delims=" %%H in (`powershell -NoProfile -ExecutionPolicy Bypass -Command "(Get-FileHash -Algorithm SHA256 -LiteralPath $env:REQUIREMENTS).Hash"`) do (
    if not defined CURRENT_REQUIREMENTS_HASH set "CURRENT_REQUIREMENTS_HASH=%%H"
)
if not defined CURRENT_REQUIREMENTS_HASH goto :fatal_requirements_hash
if not exist "%REQUIREMENTS_STAMP%" (
    set "INSTALL_REQUIREMENTS=1"
) else (
    set /p INSTALLED_REQUIREMENTS_HASH=<"%REQUIREMENTS_STAMP%"
    if /i not "!CURRENT_REQUIREMENTS_HASH!"=="!INSTALLED_REQUIREMENTS_HASH!" set "INSTALL_REQUIREMENTS=1"
)
"%PYTHON_EXE%" -c "import flask, openpyxl, werkzeug" >nul 2>&1
if errorlevel 1 set "INSTALL_REQUIREMENTS=1"
"%PYTHON_EXE%" -m pip check >nul 2>&1
if errorlevel 1 set "INSTALL_REQUIREMENTS=1"

if "!INSTALL_REQUIREMENTS!"=="1" (
    echo [SETUP] Installation / mise a jour des dependances Python...
    "%PYTHON_EXE%" -m ensurepip --upgrade >nul 2>&1
    "%PYTHON_EXE%" -m pip install --disable-pip-version-check -r "%REQUIREMENTS%"
    if errorlevel 1 goto :fatal_dependencies
    "%PYTHON_EXE%" -m pip check
    if errorlevel 1 goto :fatal_pip_check
    >"%REQUIREMENTS_STAMP%" echo !CURRENT_REQUIREMENTS_HASH!
    echo [OK] Dependances Python installees.
) else (
    echo [OK] Dependances Python deja a jour.
)

rem =======================================================
rem VALIDATION DU BACKEND
rem =======================================================
echo [INFO] Verification syntaxique du backend...
"%PYTHON_EXE%" -m compileall -q "%BACKEND_DIR%"
if errorlevel 1 goto :fatal_compile
"%PYTHON_EXE%" -c "import sys; sys.path.insert(0, r'%BACKEND_DIR%'); import app; print('[OK] Import backend valide')"
if errorlevel 1 goto :fatal_import

rem =======================================================
rem DEMARRAGE DE FLASK EN ARRIERE-PLAN
rem =======================================================
echo [INFO] Demarrage de Flask sur le port 8008...
set "ANALYSEAO_PYTHON_EXE=%PYTHON_EXE%"
set "ANALYSEAO_APP=%BACKEND_DIR%\app.py"
set "ANALYSEAO_WORKDIR=%ROOT%"
set "ANALYSEAO_FLASK_LOG=%FLASK_LOG%"
set "ANALYSEAO_FLASK_ERR=%FLASK_ERR%"
set "ANALYSEAO_PID_FILE=%FLASK_PID_FILE%"
powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command "$p=Start-Process -FilePath $env:ANALYSEAO_PYTHON_EXE -ArgumentList @($env:ANALYSEAO_APP) -WorkingDirectory $env:ANALYSEAO_WORKDIR -WindowStyle Hidden -RedirectStandardOutput $env:ANALYSEAO_FLASK_LOG -RedirectStandardError $env:ANALYSEAO_FLASK_ERR -PassThru; Set-Content -LiteralPath $env:ANALYSEAO_PID_FILE -Value $p.Id -Encoding ASCII"
if errorlevel 1 goto :fatal_flask_start

powershell -NoProfile -ExecutionPolicy Bypass -Command "$ok=$false; for($i=0;$i -lt 60;$i++){try{$r=Invoke-RestMethod -Uri $env:HEALTH_URL -TimeoutSec 1; if($r.status -eq 'OK'){$ok=$true;break}}catch{}; Start-Sleep -Milliseconds 500}; if($ok){exit 0}else{exit 1}" >nul 2>&1
if errorlevel 1 goto :fatal_flask_service

:open_browser
echo [INFO] Ouverture de l'interface...
start "" "%APP_URL%"
echo [OK] AnalyseAO est disponible.
exit /b 0

rem =======================================================
rem SOUS-ROUTINES
rem =======================================================
:health_ready
powershell -NoProfile -ExecutionPolicy Bypass -Command "try{$r=Invoke-RestMethod -Uri $env:HEALTH_URL -TimeoutSec 1; if($r.status -eq 'OK'){exit 0}}catch{}; exit 1" >nul 2>&1
exit /b %ERRORLEVEL%

:reset_runtime_folder
set "TARGET_DIR=%~1"
if not defined TARGET_DIR exit /b 1
if /i "%TARGET_DIR%"=="%ROOT%" exit /b 1
if /i "%TARGET_DIR%"=="%BACKEND_DIR%" exit /b 1
if /i "%TARGET_DIR%"=="%RUNTIME_DIR%" exit /b 1
if exist "%TARGET_DIR%" rmdir /s /q "%TARGET_DIR%"
if exist "%TARGET_DIR%" exit /b 1
mkdir "%TARGET_DIR%" >nul 2>&1
if not exist "%TARGET_DIR%" exit /b 1
exit /b 0

:detect_python_314
set "BOOTSTRAP_PYTHON="
set "BOOTSTRAP_PYTHON_ARGS="
where py.exe >nul 2>&1
if not errorlevel 1 (
    py -3.14 -c "import sys; raise SystemExit(0 if sys.version_info[:2] == (3,14) else 1)" >nul 2>&1
    if not errorlevel 1 (
        for /f "delims=" %%P in ('where py.exe 2^>nul') do if not defined BOOTSTRAP_PYTHON set "BOOTSTRAP_PYTHON=%%~fP"
        set "BOOTSTRAP_PYTHON_ARGS=-3.14"
        exit /b 0
    )
)
for %%P in (
    "%LocalAppData%\Python\pythoncore-3.14-64\python.exe"
    "%LocalAppData%\Programs\Python\Python314\python.exe"
    "%ProgramFiles%\Python314\python.exe"
) do (
    if not defined BOOTSTRAP_PYTHON if exist "%%~fP" (
        "%%~fP" -c "import sys; raise SystemExit(0 if sys.version_info[:2] == (3,14) else 1)" >nul 2>&1
        if not errorlevel 1 set "BOOTSTRAP_PYTHON=%%~fP"
    )
)
if defined BOOTSTRAP_PYTHON exit /b 0
for /f "delims=" %%P in ('where python.exe 2^>nul') do (
    if not defined BOOTSTRAP_PYTHON (
        "%%~fP" -c "import sys; raise SystemExit(0 if sys.version_info[:2] == (3,14) else 1)" >nul 2>&1
        if not errorlevel 1 set "BOOTSTRAP_PYTHON=%%~fP"
    )
)
exit /b 0

:require_winget
where winget.exe >nul 2>&1
exit /b %ERRORLEVEL%

:refresh_path
for /f "usebackq delims=" %%P in (`powershell -NoProfile -ExecutionPolicy Bypass -Command "[Environment]::GetEnvironmentVariable('Path','Machine') + ';' + [Environment]::GetEnvironmentVariable('Path','User')"`) do set "PATH=%%P"
exit /b 0

:show_flask_error
if exist "%FLASK_ERR%" (
    echo.
    echo [INFO] Erreur Flask :
    type "%FLASK_ERR%"
)
exit /b 0

rem =======================================================
rem ERREURS
rem =======================================================
:fatal_app_missing
echo [ERREUR] backend\app.py introuvable.
goto :fatal
:fatal_side_missing
echo [ERREUR] backend\side_by_side.py introuvable.
goto :fatal
:fatal_offer_missing
echo [ERREUR] backend\offer_analysis.py introuvable.
goto :fatal
:fatal_direct_missing
echo [ERREUR] backend\direct_anomaly_engine.py introuvable.
goto :fatal
:fatal_frontend_missing
echo [ERREUR] frontend\index.html introuvable.
goto :fatal
:fatal_requirements_missing
echo [ERREUR] backend\requirements.txt introuvable.
goto :fatal
:fatal_cleanup_uploads
echo [ERREUR] Impossible de nettoyer backend\runtime\uploads.
goto :fatal
:fatal_cleanup_outputs
echo [ERREUR] Impossible de nettoyer backend\runtime\outputs.
goto :fatal
:fatal_winget
echo [ERREUR] WinGet est absent. Installez Microsoft App Installer puis relancez.
goto :fatal
:fatal_python_install
echo [ERREUR] L'installation automatique de Python 3.14 a echoue.
goto :fatal
:fatal_python_path
echo [ERREUR] Python 3.14 reste introuvable apres installation.
goto :fatal
:fatal_venv_remove
echo [ERREUR] Impossible de supprimer l'ancien environnement .venv.
goto :fatal
:fatal_venv_create
echo [ERREUR] Impossible de creer l'environnement .venv.
goto :fatal
:fatal_venv_version
echo [ERREUR] Le Python du .venv n'est pas une version 3.14.x valide.
goto :fatal
:fatal_requirements_hash
echo [ERREUR] Impossible de calculer l'empreinte de backend\requirements.txt.
goto :fatal
:fatal_dependencies
echo [ERREUR] L'installation des dependances Python a echoue.
goto :fatal
:fatal_pip_check
echo [ERREUR] Les dependances Python installees sont incoherentes.
goto :fatal
:fatal_compile
echo [ERREUR] Erreur de syntaxe dans le backend.
goto :fatal
:fatal_import
echo [ERREUR] Import du backend impossible. L'erreur exacte est affichee ci-dessus.
goto :fatal
:fatal_flask_start
echo [ERREUR] Impossible de lancer Flask.
call :show_flask_error
goto :fatal
:fatal_flask_service
echo [ERREUR] Flask a demarre mais ne repond pas sur le port 8008.
call :show_flask_error
goto :fatal
:fatal
echo.
echo [ECHEC] AnalyseAO n'a pas pu demarrer.
echo Appuyez sur une touche pour fermer cette fenetre.
pause >nul
exit /b 1
