rem @ECHO OFF
cd %~dp0
set venv_path=%~dp0\.venv

rem START /WAIT "" python_install.bat

python -m venv %venv_path%

call %venv_path%\Scripts\activate

pip install -r requirements.txt -U

call %venv_path%\Scripts\deactivate

pause