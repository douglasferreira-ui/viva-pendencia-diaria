@echo off
cd /d "%~dp0"
where python >nul 2>&1
if %errorlevel%==0 (
    python -m pip install -r requirements.txt
    python -m streamlit run app.py
) else (
    py -m pip install -r requirements.txt
    py -m streamlit run app.py
)
pause
