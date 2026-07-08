# Script para iniciar a aplicação Flask
Write-Host "================================" -ForegroundColor Cyan
Write-Host "Comparador de Inventário" -ForegroundColor Green
Write-Host "================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Iniciando aplicação..." -ForegroundColor Yellow
Write-Host ""

# Mudar para o diretório do script
Set-Location $PSScriptRoot

# Verificar se Python está instalado
try {
    $pythonVersion = python --version
    Write-Host "Python encontrado: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "ERRO: Python não está instalado!" -ForegroundColor Red
    Write-Host "Baixe em: https://www.python.org/" -ForegroundColor Yellow
    Read-Host "Pressione ENTER para sair"
    exit 1
}

Write-Host ""
Write-Host "Iniciando servidor Flask..." -ForegroundColor Cyan
Write-Host ""

# Executar a aplicação (modo debug só para desenvolvimento local)
$env:FLASK_DEBUG = "1"
python app.py

Write-Host ""
Write-Host "Aplicação encerrada." -ForegroundColor Yellow
Read-Host "Pressione ENTER para fechar"
