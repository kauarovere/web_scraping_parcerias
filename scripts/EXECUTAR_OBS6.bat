@echo off
chcp 65001 >nul
echo.
echo =====================================================
echo  OBSERVATORIO DE PARCERIAS SP — Executar obs6
echo =====================================================
echo.
echo Selecione o arquivo de processos quando a janela abrir.
echo.
python "%~dp0src\obs6.py"
echo.
pause
