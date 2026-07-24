@echo off
chcp 65001 >nul
echo.
echo =====================================================
echo  OBSERVATORIO DE PARCERIAS SP — Criar Checkpoint
echo =====================================================
echo.
echo Selecione o arquivo Backup_Temp_obs6.xlsx quando a janela abrir.
echo.
python "%~dp0..\src\obs6_criar_checkpoint.py"
echo.
pause
