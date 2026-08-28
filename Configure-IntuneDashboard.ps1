[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

function Read-RequiredValue {
    param([string]$Prompt)

    do {
        $value = (Read-Host $Prompt).Trim()
    } while ([string]::IsNullOrWhiteSpace($value))

    return $value
}

$dashboardConfigDir = Join-Path $env:LOCALAPPDATA 'IntuneDashboard'
$configPath = Join-Path $dashboardConfigDir 'config.json'

Write-Host 'Configuracao do Dashboard de Inventario Intune' -ForegroundColor Cyan
Write-Host 'Informe a URL do SITE do SharePoint'
Write-Host 'https://empresa.sharepoint.com/sites/TI'

$config = [ordered]@{
    TenantId       = Read-RequiredValue 'ID do diretorio (locatario)'
    ClientId       = Read-RequiredValue 'ID do aplicativo (cliente)'
    SharePointSite = (Read-RequiredValue 'URL do site SharePoint').TrimEnd('/')
    ListUrl        = (Read-RequiredValue 'URL completa da lista SharePoint').TrimEnd('/')
}

$secureSecret = Read-Host 'VALOR do segredo do aplicativo' -AsSecureString
if ($null -eq $secureSecret) {
    throw 
}

# ConvertFrom-SecureString protege o valor usando DPAPI. Apenas este usuario,
# nesta maquina, podera descriptografa-lo para executar a rotina agendada.
$config.ClientSecretProtected = ConvertFrom-SecureString -SecureString $secureSecret

New-Item -ItemType Directory -Path $dashboardConfigDir -Force | Out-Null
$config | ConvertTo-Json | Set-Content -Path $configPath -Encoding UTF8

Write-Host "Configuracao salva com seguranca em: $configPath" -ForegroundColor Green
Write-Host 'Nao envie nem copie o arquivo de configuracao.' -ForegroundColor Yellow
