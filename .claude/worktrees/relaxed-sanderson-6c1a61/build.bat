@echo off
title Build — CCTV Control Panel
echo.
echo  ==============================================
echo   CCTV Control Panel  —  Build com PyInstaller
echo  ==============================================
echo.

:: Ativa o ambiente virtual se existir
if exist venv\Scripts\activate.bat (
    echo [1/4] Ativando ambiente virtual...
    call venv\Scripts\activate.bat
) else (
    echo [AVISO] Ambiente virtual nao encontrado. Usando Python do sistema.
)

:: Instala/atualiza PyInstaller
echo.
echo [2/4] Instalando PyInstaller...
pip install --quiet --upgrade pyinstaller

:: Limpa builds anteriores
echo.
echo [3/4] Limpando builds anteriores...
if exist build  rmdir /s /q build
if exist dist   rmdir /s /q dist

:: Executa o build
echo.
echo [4/4] Gerando executavel...
pyinstaller cctv_panel.spec

echo.
if exist "dist\CCTV_ControlPanel\CCTV_ControlPanel.exe" (
    echo  ============================================
    echo   BUILD CONCLUIDO COM SUCESSO!
    echo   Executavel: dist\CCTV_ControlPanel\CCTV_ControlPanel.exe
    echo  ============================================
    echo.
    explorer dist\CCTV_ControlPanel
) else (
    echo  [ERRO] Build falhou. Verifique as mensagens acima.
)

pause
