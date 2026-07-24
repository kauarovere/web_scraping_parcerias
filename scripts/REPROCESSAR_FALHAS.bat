@echo off
chcp 65001 >nul
echo.
echo =====================================================
echo  OBSERVATORIO DE PARCERIAS SP — Reprocessar falhas
echo =====================================================
echo.
echo Selecione a planilha gerada pelo obs6 quando a janela abrir.
echo.
python "%~dp0..\src\obs6_retry.py"
echo.
pause
