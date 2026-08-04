@echo off
REM ==========================================================================
REM Gera um executavel standalone (.exe) do painel REDEB2B usando PyInstaller.
REM O resultado fica em dist\REDEB2B\REDEB2B.exe — essa pasta inteira pode
REM ser copiada para outro computador Windows, MESMO SEM PYTHON INSTALADO.
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

echo Instalando o PyInstaller (se ainda nao estiver instalado)...
python -m pip install pyinstaller --quiet

echo.
echo Gerando o executavel (isso pode levar alguns minutos)...
echo.

pyinstaller --name REDEB2B --noconfirm --onefile ^
  --add-data "templates;templates" ^
  --add-data "static;static" ^
  --collect-all pandas ^
  --collect-all openpyxl ^
  --collect-all flask ^
  app.py

echo.
echo ==========================================================================
echo Pronto! O executavel esta em: dist\REDEB2B.exe
echo (um unico arquivo, sem precisar de pasta ao redor)
echo.
echo IMPORTANTE antes de usar:
echo   1. Copie o arquivo .env (criado a partir do .env.example) para a
echo      mesma pasta onde ficar o REDEB2B.exe
echo   2. No .env copiado, deixe FLASK_DEBUG=false
echo   3. De um duplo-clique em REDEB2B.exe e acesse http://localhost:5000
echo      (a primeira abertura pode demorar alguns segundos a mais, e isso
echo      e normal no modo --onefile)
echo ==========================================================================
echo.
pause
