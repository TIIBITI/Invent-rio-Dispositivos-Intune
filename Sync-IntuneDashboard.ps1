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
    try {
        return Invoke-RestMethod @params
    }
    catch {
        $graphDetail = $_.ErrorDetails.Message
        if ([string]::IsNullOrWhiteSpace($graphDetail) -and $null -ne $_.Exception.Response) {
            try {
                $responseStream = $_.Exception.Response.GetResponseStream()
                $reader = [System.IO.StreamReader]::new($responseStream)
                $graphDetail = $reader.ReadToEnd()
                $reader.Dispose()
            }
            catch {
                # Mantem a mensagem padrao caso o corpo da resposta nao possa ser lido.
            }
        }
        if ([string]::IsNullOrWhiteSpace($graphDetail)) {
            $graphDetail = $_.Exception.Message
        }
        throw "Falha na requisicao Graph [$Method] $Uri. Detalhe: $graphDetail"
    }
}

function Convert-BytesToGB {
    param([Nullable[Int64]]$Bytes)

    if ($null -eq $Bytes -or $Bytes -le 0) {
        return $null
    }

    return [Math]::Round(($Bytes / 1GB), 2)
}

function Get-SecurityAlert {
    param($Device)

    $alerts = [System.Collections.Generic.List[string]]::new()

    if ($Device.complianceState -and $Device.complianceState -ne 'compliant') {
        $alerts.Add("Conformidade: $($Device.complianceState)")
    }

    if ($Device.operatingSystem -match '^Windows' -and $Device.isEncrypted -eq $false) {
        $alerts.Add('Criptografia desativada')
    }

    $riskStates = @('lowSeverity', 'mediumSeverity', 'highSeverity', 'unresponsive', 'compromised', 'misconfigured')
    if ($riskStates -contains $Device.partnerReportedThreatState) {
        $alerts.Add("Ameaça reportada: $($Device.partnerReportedThreatState)")
    }

    if ($Device.jailBroken -and $Device.jailBroken -ne 'false' -and $Device.jailBroken -ne 'unknown') {
        $alerts.Add('Dispositivo comprometido (jailbreak/root)')
    }

    if ($alerts.Count -eq 0) {
        return 'Sem alerta identificado'
    }

    return ($alerts -join '; ')
}

function Set-SharePointField {
    param(
        [hashtable]$Fields,
        [hashtable]$ColumnNames,
        [string]$DisplayName,
        [object]$Value
    )

    if (-not $ColumnNames.ContainsKey($DisplayName)) {
        throw "A coluna '$DisplayName' nao foi encontrada na Lista do SharePoint."
    }

    $Fields[$ColumnNames[$DisplayName]] = $Value
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

# Alguns campos de hardware não são retornados na listagem. O Graph exige uma
# consulta detalhada por equipamento para apresentar RAM, MAC Ethernet e outros
# dados de inventário. Com um parque de porte normal, as chamadas sequenciais
# evitam atingir o limite de requisições da API.
$detailedDeviceFields = 'totalStorageSpaceInBytes,freeStorageSpaceInBytes,physicalMemoryInBytes,isEncrypted,wiFiMacAddress,ethernetMacAddress,partnerReportedThreatState,managementState,enrolledDateTime,complianceGracePeriodExpirationDateTime,jailBroken'

$itemsUri = "https://graph.microsoft.com/v1.0/sites/$($site.id)/lists/$($list.id)/items?`$expand=fields&`$top=100"
$items = Get-AllGraphPages -Uri $itemsUri -Headers $headers
$columnsUri = "https://graph.microsoft.com/v1.0/sites/$($site.id)/lists/$($list.id)/columns?`$select=name,displayName&`$top=100"
$listColumns = Get-AllGraphPages -Uri $columnsUri -Headers $headers
$columnNames = @{}
foreach ($column in $listColumns) {
    if (-not [string]::IsNullOrWhiteSpace($column.displayName)) {
        $columnNames[$column.displayName] = $column.name
    }
}
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
    $detailsUri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$($device.id)?`$select=$detailedDeviceFields"
    $details = Invoke-GraphRequest -Method Get -Uri $detailsUri -Headers $headers

    $totalStorageGB = Convert-BytesToGB -Bytes $details.totalStorageSpaceInBytes
    $freeStorageGB = Convert-BytesToGB -Bytes $details.freeStorageSpaceInBytes
    $memoryGB = Convert-BytesToGB -Bytes $details.physicalMemoryInBytes
    $usedStoragePercent = $null
    if ($null -ne $totalStorageGB -and $totalStorageGB -gt 0 -and $null -ne $freeStorageGB) {
        $usedStoragePercent = [Math]::Round((($totalStorageGB - $freeStorageGB) / $totalStorageGB) * 100, 2)
    }

    $securityDevice = [PSCustomObject]@{
        complianceState            = $device.complianceState
        operatingSystem            = $device.operatingSystem
        isEncrypted                = $details.isEncrypted
        partnerReportedThreatState = $details.partnerReportedThreatState
        jailBroken                 = $details.jailBroken
    }

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

    $extendedFieldValues = [ordered]@{
        ArmazenamentoTotalGB                = $totalStorageGB
        ArmazenamentoLivreGB                = $freeStorageGB
        UsoArmazenamentoPercentual          = $usedStoragePercent
        MemoriaRAMGB                        = $memoryGB
        Criptografado                       = $details.isEncrypted
        WifiMac                             = $details.wiFiMacAddress
        EthernetMac                         = $details.ethernetMacAddress
        EstadoAmeaca                        = $details.partnerReportedThreatState
        AlertaSeguranca                     = Get-SecurityAlert -Device $securityDevice
        EstadoGerenciamento                 = $details.managementState
        DataInscricao                       = $details.enrolledDateTime
    }
    foreach ($field in $extendedFieldValues.GetEnumerator()) {
        Set-SharePointField -Fields $fields -ColumnNames $columnNames -DisplayName $field.Key -Value $field.Value
    }

    # ipAddressV4 é um campo beta que o Intune frequentemente retorna vazio.
    # Mantemos o campo sem sobrescrever valores existentes até que o serviço o
    # disponibilize de modo confiável no Graph.
    if (-not [string]::IsNullOrWhiteSpace($details.ipAddressV4)) {
        Set-SharePointField -Fields $fields -ColumnNames $columnNames -DisplayName 'IPV4' -Value $details.ipAddressV4
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
