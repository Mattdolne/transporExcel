@echo off
cls
echo ==============================================================
echo      Iniciando Interface Grafica transporExcel...
echo ==============================================================
cd /d "%~dp0"

if not exist .venv (
    echo Erro: O ambiente virtual .venv nao foi encontrado.
    pause
    exit /b
)

call .venv\Scripts\activate.bat
python transporExcel.py
if %errorlevel% neq 0 (
    echo.
    echo O programa encerrou com erros.
    pause
)
