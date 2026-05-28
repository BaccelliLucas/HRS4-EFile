@echo off
setlocal
cd /d "%~dp0"

set "PYTHON_EXE=python"
if exist "venv\Scripts\python.exe" (
  set "PYTHON_EXE=venv\Scripts\python.exe"
)

echo Usando interpretador: %PYTHON_EXE%

echo [1/4] Instalando/atualizando PyInstaller...
"%PYTHON_EXE%" -m pip install --upgrade pyinstaller
if errorlevel 1 (
  echo Falha ao instalar PyInstaller.
  exit /b 1
)

echo [2/4] Limpando builds anteriores...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist EFileDownloader.spec del /f /q EFileDownloader.spec

if not exist ".env" (
  echo Arquivo .env nao encontrado na raiz do projeto.
  echo Para gerar um unico .exe com configuracao embutida, o .env e obrigatorio no build.
  exit /b 1
)

echo [3/4] Gerando executavel...
"%PYTHON_EXE%" -m PyInstaller --noconfirm --clean EFileDownloader.spec
if errorlevel 1 (
  echo Falha ao gerar executavel.
  exit /b 1
)

echo [4/4] Finalizando build...

echo.
echo Build concluido com sucesso.
echo Executavel: dist\EFileDownloader.exe
echo O .env foi embutido dentro do executavel; distribuicao: apenas o .exe.
echo Log de execucao: sera criado na mesma pasta do executavel (dist\efile_log_YYYYMMDD_HHMMSS.log)
endlocal
