@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

title Processador de Dados Cadastrais Vivo

echo ============================================
echo   PROCESSADOR DE DADOS CADASTRAIS VIVO
echo ============================================
echo.

:: --- Verifica se Python esta instalado ---
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo  ERRO: Python nao foi encontrado neste computador.
    echo.
    echo  Para usar este programa, voce precisa instalar o Python.
    echo.
    echo  Passos:
    echo    1. Acesse  https://www.python.org/downloads/
    echo    2. Baixe a ultima versao do Python 3
    echo    3. Na instalacao, MARQUE a opcao "Add Python to PATH"
    echo    4. Clique em "Install Now"
    echo    5. Depois de instalado, execute este arquivo novamente.
    echo.
    pause
    exit /b 1
)

echo [OK] Python encontrado.

:: --- Verifica/instala dependencias ---
echo.
echo Verificando dependencias...
python -c "import pandas, openpyxl, requests" >nul 2>&1
if errorlevel 1 (
    echo Instalando bibliotecas necessarias (pandas, openpyxl, requests)...
    echo Isso pode levar alguns minutos na primeira vez...
    python -m pip install -r requirements.txt -q
    if errorlevel 1 (
        echo ERRO: Nao foi possivel instalar as dependencias.
        echo Tente executar manualmente:  pip install pandas openpyxl requests
        pause
        exit /b 1
    )
    echo [OK] Dependencias instaladas.
) else (
    echo [OK] Todas as dependencias estao instaladas.
)

:: --- Seleciona pasta de entrada ---
echo.
echo Selecione a PASTA onde estao os arquivos .txt da Vivo...
for /f "delims=" %%I in ('powershell -NoProfile -Command "Add-Type -AssemblyName System.Windows.Forms; $f = New-Object System.Windows.Forms.FolderBrowserDialog; $f.Description = 'Selecione a pasta com os arquivos .txt da Vivo'; $f.ShowDialog() | Out-Null; $f.SelectedPath"') do set "INPUT_DIR=%%I"

if "%INPUT_DIR%"=="" (
    echo.
    echo Nenhuma pasta selecionada. Encerrando.
    pause
    exit /b 1
)

echo [OK] Pasta de entrada: %INPUT_DIR%

:: --- Seleciona pasta de saida ---
echo.
echo Agora selecione a PASTA onde deseja salvar o resultado...
for /f "delims=" %%I in ('powershell -NoProfile -Command "Add-Type -AssemblyName System.Windows.Forms; $f = New-Object System.Windows.Forms.FolderBrowserDialog; $f.Description = 'Selecione a pasta para salvar o resultado'; $f.ShowDialog() | Out-Null; $f.SelectedPath"') do set "OUTPUT_DIR=%%I"

if "%OUTPUT_DIR%"=="" (
    echo.
    echo Nenhuma pasta selecionada. Encerrando.
    pause
    exit /b 1
)

echo [OK] Pasta de saida: %OUTPUT_DIR%

:: --- Executa o parser ---
echo.
echo ============================================
echo   PROCESSANDO... Aguarde.
echo ============================================
echo.

python parse_vivo.py --input "%INPUT_DIR%" --output "%OUTPUT_DIR%"

if errorlevel 1 (
    echo.
    echo ============================================
    echo   ERRO durante o processamento.
    echo   Verifique a mensagem acima.
    echo ============================================
    pause
    exit /b 1
)

echo.
echo ============================================
echo   CONCLUIDO COM SUCESSO!
echo ============================================
echo.
echo Os arquivos foram gerados em:
echo   %OUTPUT_DIR%
echo.
echo Voce encontrara:
echo   - dados_cadastrais_vivo.xlsx  (planilha Excel)
echo   - dados_cadastrais.db         (banco de dados)
echo.
echo  ATENCAO: Esses arquivos contem dados pessoais (CPF, nome, endereco).
echo  Proteja-os: nao compartilhe sem criptografia e apague quando nao precisar mais.
echo.

:: --- Tenta abrir a pasta de saida ---
start explorer "%OUTPUT_DIR%"

pause
