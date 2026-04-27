@echo off
title Gerador de Relatorio de Vendas
cd /d "%~dp0"

echo.
echo ===============================================
echo  GERADOR AUTOMATICO DE RELATORIO DE VENDAS
echo ===============================================
echo.
echo Procurando planilha na pasta entrada...
echo.

py -3 src\analise_vendas.py

echo.
if %errorlevel% neq 0 (
    echo Nao foi possivel gerar o relatorio.
    echo Confira a mensagem acima e ajuste a planilha de entrada.
) else (
    echo Relatorio gerado com sucesso.
    echo Abra a pasta saida para ver os arquivos finais.
)
echo.
pause
