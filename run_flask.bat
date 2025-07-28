@echo off
echo ================================
echo  Starting Flask App on Static IP...
echo ================================

:: Optional - activate virtual environment
call venv\Scripts\activate

:: Set Flask variables
set FLASK_APP=app.py
set FLASK_ENV=development

:: Run Flask on your static IP and port 5000
flask run --host=0.0.0.0 --port=5000

pause
