NET SESSION >nul 2>&1

IF %ERRORLEVEL% EQU 0 (
    echo Administrator privileges detected.
    goto :ADMIN_CODE
) ELSE (
    echo Attempting to elevate privileges...
    powershell -Command "Start-Process -FilePath '%~dpnx0' -Verb RunAs"
    exit /b

    goto :CRIA_VENV
)

:ADMIN_CODE
REM Put your administrator-level commands here
winget install -e --id Python.Python.3.13 --scope machine


:CRIA_VENV
cd '%~dp0'
python -m venv .venv
call .\.venv\activate
pip install -r requirements.txt -U
.\.venv\deactivate