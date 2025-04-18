# Script para ativar o Windows usando servidor KMS personalizado
# Requer privilégios de administrador

# Verificar se está rodando como administrador
$isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "Este script precisa ser executado como administrador!" -ForegroundColor Red
    Write-Host "Por favor, execute o PowerShell como administrador e tente novamente."
    exit
}

# Configurar o servidor KMS
$kmsServer = "carlosandradepy-39464.portmap.host:44258"

Write-Host "Configurando servidor KMS..." -ForegroundColor Yellow
slmgr /skms $kmsServer

if ($LASTEXITCODE -ne 0) {
    Write-Host "Erro ao configurar o servidor KMS!" -ForegroundColor Red
    exit
}

Write-Host "Ativando Windows..." -ForegroundColor Yellow
slmgr /ato

if ($LASTEXITCODE -ne 0) {
    Write-Host "Erro ao ativar o Windows!" -ForegroundColor Red
    exit
}

# Verificar status da ativação
Write-Host "Verificando status da ativação..." -ForegroundColor Yellow
slmgr /dli

Write-Host "`nProcesso concluído!" -ForegroundColor Green
Write-Host "Se você encontrar algum problema, verifique:"
Write-Host "1. Se o servidor KMS está acessível"
Write-Host "2. Se a porta 44258 está aberta"
Write-Host "3. Se o Windows está configurado para aceitar ativação KMS" 