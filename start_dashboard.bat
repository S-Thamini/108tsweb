@echo off
title 108TS Premium Dashboard

echo =========================================
echo Starting 108TS Premium Web Environment
echo =========================================

IF NOT EXIST venv (
    echo Creating an isolated Python virtual environment...
    python -m venv venv
)

echo Activating environment...
call venv\Scripts\activate.bat

echo Verifying dependencies...
pip install -r requirements.txt -q

echo.
echo =========================================
echo Server is launching. 
echo Open your browser to: http://127.0.0.1:8008
echo =========================================
echo.

python main.py
pause
