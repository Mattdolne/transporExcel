@echo off
chcp 65001 > nul
echo ==============================================================
echo        Iniciando Sistema de Gestão de Cartões...
echo ==============================================================
cd /d "%~dp0"

if not exist .venv (
    echo Erro: O ambiente virtual .venv não foi encontrado na pasta atual.
    echo Certifique-se de que a pasta .venv existe.
    pause
    exit /b
)

call .venv\Scripts\activate.bat
python gestao_cartao.py
pause
