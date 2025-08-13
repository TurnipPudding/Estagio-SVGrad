
@ECHO OFF
cd %~dp0
set venv_path=%~dp0\..\.venv

call %venv_path%\Scripts\activate
python interface_final.py

call %venv_path%\Scripts\deactivate