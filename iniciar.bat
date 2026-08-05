@echo off
setlocal EnableExtensions
chcp 65001 >nul
title Automacao de Retencoes

rem ---------------------------------------------------------------------------
rem  Sobe o sistema num duplo clique: garante a venv, garante as dependencias
rem  e abre o servidor. Roda a partir da pasta do proprio arquivo, entao
rem  funciona de qualquer lugar (inclusive de um atalho na area de trabalho).
rem ---------------------------------------------------------------------------

cd /d "%~dp0"

set "VENV=%~dp0venv"
set "PY=%VENV%\Scripts\python.exe"
set "CARIMBO=%VENV%\.dependencias.txt"

rem --- 1. Python do sistema ---------------------------------------------------
rem  O launcher (py) e o caminho oficial no Windows; python.exe fica de reserva
rem  porque nem toda instalacao registra o launcher.
set "PYBASE="
py -3 --version >nul 2>&1 && set "PYBASE=py -3"
if not defined PYBASE (
  python --version >nul 2>&1 && set "PYBASE=python"
)
if not defined PYBASE (
  echo.
  echo  [ERRO] Python 3.11 ou superior nao foi encontrado neste computador.
  echo         Instale em https://www.python.org/downloads/ marcando
  echo         "Add python.exe to PATH" e rode este arquivo de novo.
  echo.
  pause
  exit /b 1
)

rem --- 2. Ambiente virtual ----------------------------------------------------
if not exist "%PY%" (
  echo  Criando o ambiente virtual pela primeira vez...
  %PYBASE% -m venv "%VENV%"
  if errorlevel 1 goto :falhou_venv
  rem  Forca a instalacao das dependencias: venv novo nao tem nada dentro.
  if exist "%CARIMBO%" del /q "%CARIMBO%" >nul 2>&1
)

if not exist "%PY%" goto :falhou_venv

rem --- 3. Dependencias --------------------------------------------------------
rem  Reinstalar a cada partida custa segundos toda vez. O carimbo guarda o
rem  requirements.txt que ja foi instalado: se o arquivo mudou (ou nunca foi
rem  instalado), instala; se nao, sobe direto.
set "PRECISA=1"
if exist "%CARIMBO%" (
  fc /b "%CARIMBO%" "%~dp0requirements.txt" >nul 2>&1 && set "PRECISA=0"
)

if "%PRECISA%"=="1" (
  echo  Instalando as dependencias na venv...
  "%PY%" -m pip install --upgrade pip --quiet
  "%PY%" -m pip install -r "%~dp0requirements.txt"
  if errorlevel 1 goto :falhou_deps
  copy /y "%~dp0requirements.txt" "%CARIMBO%" >nul
)

rem --- 4. Servidor ------------------------------------------------------------
echo.
echo  ============================================================
echo   Automacao de Retencoes
echo   Endereco: http://127.0.0.1:5000
echo   Para encerrar, feche esta janela ou pressione Ctrl+C.
echo  ============================================================
echo.

rem  Abre o navegador num processo separado, que espera alguns segundos antes.
rem  Abrir na hora e uma corrida perdida: o Flask leva ~1s para atender e o
rem  usuario veria "nao foi possivel conectar" na primeira tentativa.
start "" /min cmd /c "timeout /t 3 /nobreak >nul & start "" http://127.0.0.1:5000"

"%PY%" "%~dp0app.py"
set "SAIDA=%ERRORLEVEL%"

if not "%SAIDA%"=="0" (
  echo.
  echo  [ERRO] O servidor encerrou com codigo %SAIDA%.
  echo         Detalhes tecnicos em logs\app.log
  echo.
  pause
)
exit /b %SAIDA%

:falhou_venv
echo.
echo  [ERRO] Nao foi possivel criar o ambiente virtual em:
echo         %VENV%
echo         Verifique se ha espaco em disco e permissao de escrita nesta pasta.
echo.
pause
exit /b 1

:falhou_deps
echo.
echo  [ERRO] A instalacao das dependencias falhou.
echo         Verifique a conexao com a internet e rode este arquivo de novo.
echo.
pause
exit /b 1
