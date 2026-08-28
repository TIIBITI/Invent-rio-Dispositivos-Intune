[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

function Get-GraphToken {
    param($Config)

    $secureSecret = ConvertTo-SecureString $Config.ClientSecretProtected
    $secretPointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureSecret)
    try {
        $clientSecret = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($secretPointer)
        $tokenBody = @{
            client_id     = $Config.ClientId
            client_secret = $clientSecret
            scope         = 'https://graph.microsoft.com/.default'
            grant_type    = 'client_credentials'
        }

        return (Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$($Config.TenantId)/oauth2/v2.0/token" -Body $tokenBody).access_token
    }
    finally {
        if ($secretPointer -ne [IntPtr]::Zero) {
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($secretPointer)
        }
    }
}

function Get-AllGraphPages {
    param(
        [string]$Uri,
        [hashtable]$Headers
    )

    $results = @()
    do {
        try {
            $response = Invoke-RestMethod -Method Get -Uri $Uri -Headers $Headers
        }
        catch {
            throw "Falha ao consultar o Microsoft Graph. URL: $Uri. Detalhe: $($_.Exception.Message)"
        }
        $results += @($response.value)
        $Uri = $response.'@odata.nextLink'
    } while (-not [string]::IsNullOrWhiteSpace($Uri))

    return $results
}

function Invoke-GraphRequest {
    param(
        [string]$Method,
        [string]$Uri,
        [hashtable]$Headers,
        [object]$Body
    )

    $params = @{ Method = $Method; Uri = $Uri; Headers = $Headers }
    if ($null -ne $Body) {
        # Envia JSON explicitamente em UTF-8 para preservar nomes com acentos.
        $jsonBody = $Body | ConvertTo-Json -Depth 8
        $params.Body = [System.Text.Encoding]::UTF8.GetBytes($jsonBody)
        $params.ContentType = 'application/json; charset=utf-8'
    }
    return Invoke-RestMethod @params
}

$dashboardConfigDir = Join-Path $env:LOCALAPPDATA 'IntuneDashboard'
$configPath = Join-Path $dashboardConfigDir 'config.json'
if (-not (Test-Path $configPath)) {
    throw "Configuracao nao encontrada. Execute Configure-IntuneDashboard.ps1 primeiro."
}

$config = Get-Content -Path $configPath -Raw | ConvertFrom-Json
$accessToken = Get-GraphToken -Config $config
$headers = @{ Authorization = "Bearer $accessToken" }

$siteUri = [Uri]$config.SharePointSite
$sitePath = $siteUri.AbsolutePath.Trim('/')

# Aceita tambem uma URL copiada diretamente da lista, por exemplo
# https://empresa.sharepoint.com/sites/TI/Lists/Inventario/AllItems.aspx.
if ($sitePath -match '^(.*?)/Lists(?:/|$)') {
    $sitePath = $matches[1]
}

if ([string]::IsNullOrWhiteSpace($sitePath)) {
    $siteLookupUri = "https://graph.microsoft.com/v1.0/sites/$($siteUri.Host)"
}
else {
    $siteLookupUri = "https://graph.microsoft.com/v1.0/sites/$($siteUri.Host):/$sitePath"
}
$site = Invoke-GraphRequest -Method Get -Uri $siteLookupUri -Headers $headers

$listsUri = "https://graph.microsoft.com/v1.0/sites/$($site.id)/lists?`$select=id,displayName,webUrl&`$top=100"
$availableLists = Get-AllGraphPages -Uri $listsUri -Headers $headers

if (-not [string]::IsNullOrWhiteSpace($config.ListUrl)) {
    $configuredListPath = ([Uri]::UnescapeDataString(([Uri]$config.ListUrl).AbsolutePath) -replace '/AllItems\.aspx$','').TrimEnd('/')
    $list = $availableLists | Where-Object {
        $availableListPath = [Uri]::UnescapeDataString(([Uri]$_.webUrl).AbsolutePath).TrimEnd('/')
        $availableListPath -ieq $configuredListPath
    } | Select-Object -First 1
}
else {
    # Compatibilidade com configuracoes criadas antes da inclusao de ListUrl.
    $list = $availableLists | Where-Object { $_.displayName -eq $config.ListName } | Select-Object -First 1
}
if ($null -eq $list) {
    $listNames = ($availableLists | ForEach-Object { $_.displayName }) -join ', '
    throw "A lista configurada nao foi encontrada no site '$($site.webUrl)'. Listas disponiveis: $listNames"
}

$deviceFields = 'id,deviceName,manufacturer,model,serialNumber,userDisplayName,userPrincipalName,operatingSystem,osVersion,complianceState,managedDeviceOwnerType,lastSyncDateTime'
$devicesUri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$select=$deviceFields&`$top=100"
$devices = Get-AllGraphPages -Uri $devicesUri -Headers $headers

$itemsUri = "https://graph.microsoft.com/v1.0/sites/$($site.id)/lists/$($list.id)/items?`$expand=fields&`$top=100"
$items = Get-AllGraphPages -Uri $itemsUri -Headers $headers
$existingItems = @{}
foreach ($item in $items) {
    if (-not [string]::IsNullOrWhiteSpace($item.fields.IntuneDeviceId)) {
        $existingItems[$item.fields.IntuneDeviceId] = $item.id
    }
}

$updatedAt = (Get-Date).ToUniversalTime().ToString('o')
$created = 0
$updated = 0

foreach ($device in $devices) {
    $fields = @{
        Title                = $device.deviceName
        IntuneDeviceId       = $device.id
        Fabricante           = $device.manufacturer
        Modelo               = $device.model
        NumeroSerie          = $device.serialNumber
        UsuarioPrincipal     = $device.userPrincipalName
        UsuarioNome          = $device.userDisplayName
        SistemaOperacional   = $device.operatingSystem
        VersaoSO             = $device.osVersion
        Conformidade         = $device.complianceState
        Propriedade          = $device.managedDeviceOwnerType
        UltimaSincronizacao  = $device.lastSyncDateTime
        UltimaAtualizacao    = $updatedAt
    }

    if ($existingItems.ContainsKey($device.id)) {
        $itemUri = "https://graph.microsoft.com/v1.0/sites/$($site.id)/lists/$($list.id)/items/$($existingItems[$device.id])/fields"
        Invoke-GraphRequest -Method Patch -Uri $itemUri -Headers $headers -Body $fields | Out-Null
        $updated++
    }
    else {
        $itemUri = "https://graph.microsoft.com/v1.0/sites/$($site.id)/lists/$($list.id)/items"
        Invoke-GraphRequest -Method Post -Uri $itemUri -Headers $headers -Body @{ fields = $fields } | Out-Null
        $created++
    }
}

Write-Host "Sincronizacao concluida: $created criado(s), $updated atualizado(s), $($devices.Count) dispositivo(s) processado(s)." -ForegroundColor Green
