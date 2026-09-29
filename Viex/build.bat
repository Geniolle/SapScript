@echo off
REM ============================================================
REM Script de build - Validador VIES
REM ============================================================
REM
REM Este script:
REM 1. Cria um ambiente virtual Python
REM 2. Instala as dependências
REM 3. Gera o executável com PyInstaller
REM
REM Uso:
REM   build.bat
REM
REM O executável resultante estará em:
REM   ValidadorVIES\ValidadorVIES.exe
REM ============================================================

setlocal enabledelayedexpansion

cd /d "%~dp0"

echo.
echo ============================================================
echo Validador VIES - Build Script
echo ============================================================
echo.

REM Verificar se Python está instalado
python --version > nul 2>&1
if errorlevel 1 (
    echo ERRO: Python não encontrado no PATH
    echo Instale Python 3.9+ e adicione ao PATH do Windows
    pause
    exit /b 1
)

echo [1/4] Criando ambiente virtual...
if not exist venv (
    python -m venv venv
    if errorlevel 1 (
        echo ERRO: Falha ao criar ambiente virtual
        pause
        exit /b 1
    )
) else (
    echo     Ambiente virtual já existe
)

echo.
echo [2/4] Ativando ambiente virtual...
call venv\Scripts\activate.bat
if errorlevel 1 (
    echo ERRO: Falha ao ativar ambiente virtual
    pause
    exit /b 1
)

echo.
echo [3/4] Instalando dependências...
pip install -r requirements.txt > build.log 2>&1
if errorlevel 1 (
    echo ERRO: Falha ao instalar dependências
    echo Consulte build.log para detalhes
    pause
    exit /b 1
)

echo.
echo [4/4] Gerando executável com PyInstaller...
pyinstaller --onedir --windowed --name ValidadorVIES validador_vies_gui.py > build.log 2>&1
if errorlevel 1 (
    echo ERRO: Falha ao gerar executável
    echo Consulte build.log para detalhes
    pause
    exit /b 1
)

echo.
echo ============================================================
echo BUILD CONCLUÍDO COM SUCESSO
echo ============================================================
echo.
echo Executável gerado em:
echo   %cd%\dist\ValidadorVIES\ValidadorVIES.exe
echo.
echo Para usar:
echo   1. Copie a pasta "dist\ValidadorVIES" para a máquina destino
echo   2. Dê duplo clique em ValidadorVIES.exe
echo.
echo Nota: A pasta dist\ contém todas as dependências necessárias.
echo Não é necessário instalar Python ou pacotes adicionais.
echo.
pause
