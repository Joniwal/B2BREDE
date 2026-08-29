@echo off
REM ==========================================================================
REM Gera um executavel standalone (.exe) do painel REDEB2B usando PyInstaller.
REM O resultado fica em dist\REDEB2B\REDEB2B.exe — essa pasta inteira pode
REM ser copiada para outro computador Windows, MESMO SEM PYTHON INSTALADO.
REM ==========================================================================

cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo Python nao foi encontrado. Instale Python 3.10 ou superior e marque
    echo a opcao "Add Python to PATH" durante a instalacao.
    pause
    exit /b 1
)

if not exist "venv\Scripts\activate.bat" (
    echo Criando o ambiente virtual...
    python -m venv venv
    if errorlevel 1 goto :erro
)

call venv\Scripts\activate.bat

echo Instalando as dependencias do projeto...
python -m pip install --upgrade pip --quiet
if errorlevel 1 goto :erro
python -m pip install -r requirements.txt --quiet
if errorlevel 1 goto :erro

echo Instalando o PyInstaller (se ainda nao estiver instalado)...
python -m pip install pyinstaller --quiet
if errorlevel 1 goto :erro

echo.
echo Gerando o executavel (isso pode levar alguns minutos)...
echo.

python -m PyInstaller --name REDEB2B --noconfirm --onefile --icon static\favicon.ico ^
  --add-data "templates;templates" ^
  --add-data "static;static" ^
  --collect-all pandas ^
  --collect-all openpyxl ^
  --collect-all flask ^
  app.py
if errorlevel 1 goto :erro

if not exist "dist\.env" (
    copy /Y ".env.example" "dist\.env" >nul
)

echo.
echo ==========================================================================
echo Pronto! O executavel esta em: dist\REDEB2B.exe
echo A configuracao portatil esta em: dist\.env
echo.
echo IMPORTANTE antes de usar:
echo   1. Distribua juntos dist\REDEB2B.exe e dist\.env
echo   2. No .env, deixe EXCEL_PATH e EXCEL_SEARCH_ROOTS vazios para que cada
echo      usuario procure REDE_B2B.xlsx no proprio OneDrive.
echo   3. De um duplo-clique em REDEB2B.exe; o navegador abre automaticamente.
echo      (a primeira abertura pode demorar alguns segundos a mais, e isso
echo      e normal no modo --onefile)
echo ==========================================================================
echo.
pause
exit /b 0

:erro
echo.
echo ERRO: nao foi possivel gerar o executavel.
echo Verifique as mensagens acima, a conexao com a internet e as permissoes.
pause
exit /b 1
