@echo off
REM ==========================================================================
REM Atalho para iniciar o painel REDEB2B com um duplo-clique.
REM Ativa o ambiente virtual (venv) e inicia o servidor Flask — o proprio
REM app.py ja abre o navegador automaticamente em http://localhost:5000
REM ==========================================================================

cd /d "%~dp0"

if not exist "venv\Scripts\activate.bat" (
    echo O ambiente virtual "venv" nao foi encontrado nesta pasta.
    echo Rode primeiro: python -m venv venv
    echo E depois: venv\Scripts\activate ^&^& python -m pip install -r requirements.txt
    pause
    exit /b 1
)

call venv\Scripts\activate.bat

python app.py

pause
