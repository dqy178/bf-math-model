@echo off
chcp 65001 >nul
title Blast Furnace Calculator
cd /d "%~dp0"

set KMP_DUPLICATE_LIB_OK=TRUE

where python >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python not found.
    echo Please install Python and check "Add Python to PATH" during installation.
    pause
    exit /b
)

echo ======================================
echo Installing dependencies (Tsinghua mirror)
echo ======================================
python -m pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple

echo.
echo Starting streamlit...
echo Close this window to stop service.
echo.
python -m streamlit run app.py --server.port 8501

pause