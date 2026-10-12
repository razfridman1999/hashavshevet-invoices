@echo off
cd /d "%~dp0"
rem py/pyw first: on some PCs "python" is only the Microsoft Store shortcut and pythonw does not exist
where pyw >nul 2>nul && (start "" pyw run.py journal & exit /b)
start "" pythonw run.py journal
