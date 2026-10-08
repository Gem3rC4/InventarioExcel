#requires -Version 5.1
<#
.SYNOPSIS
    Registra el inventario del equipo local en Excel Online y crea un respaldo CSV.
#>

# ==============================================================================
# VARIABLES DE CONFIGURACIÓN
# ==============================================================================
$DirectorioSalida = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'InventarioEquipos'$NombreTabla      = 'InventarioEquipos'

$urlLibro = 'https://grupopinulitogt-my.sharepoint.com/:x:/r/personal/horacio_sauce_corporacionalisa_com/_layouts/15/Doc.aspx?sourcedoc=%7B7F53AF0E-F9E7-493F-97A8-15FFAFB5E5E4%7D&file=INVENTARIO%20DE%20EQUIPOS.xlsx&fromShare=true&action=default&mobileredirect=true'
$tenantId = '0675a017-358d-4fb1-85c3-368320881e85'$clientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'
# ==============================================================================

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Get-CimData {
    param([Parameter(Mandatory = $true)][string]$ClassName, [string]$Filter)
    try {
        $parameters = @{ ClassName =$ClassName; ErrorAction = 'Stop' }
        if ($Filter) { $parameters.Filter =$Filter }
        return Get-CimInstance @parameters
    } catch { return $null }
}

function ConvertTo-Gigabytes {
    param([object]$Bytes)
    if ($null -eq $Bytes) { return$null }
    return [math]::Round(([double]$Bytes / 1GB), 2)
}

function Normalize-Header {
    param([Parameter(Mandatory = $true)][string]$Text)
    $decomposed =$Text.Normalize([System.Text.NormalizationForm]::FormD)
    $plain = [regex]::Replace($decomposed, '\p{Mn}', '')
    return [regex]::Replace($plain.ToLowerInvariant(), '[^a-z0-9]', '')
}

function Get-GraphCollection {
    param([Parameter(Mandatory = $true)][string]$Uri)$items = @()
    do {
        $response = Invoke-GraphRequest -Method GET -Uri$Uri
        if ($response -is [System.Collections.IDictionary]) {
            $items += @($response['value'])
            $Uri = [string]$response['@odata.nextLink']
        }
        else {
            $items += @($response.value)$nextLink = $response.PSObject.Properties['@odata.nextLink']$Uri = if ($nextLink) { [string]$nextLink.Value } else { '' }
        }
    } while ($Uri)
    return $items | Where-Object { $null -ne$_ }
}

function Get-HttpErrorBody {
    param([Parameter(Mandatory = $true)][System.Management.Automation.ErrorRecord]$ErrorRecord)
    try {
        if (-not $ErrorRecord.Exception.PSObject.Properties['Response']) { return$null }
        $response =$ErrorRecord.Exception.Response
        if ($null -eq $response) { return$null }
        $stream =$response.GetResponseStream()
        if ($null -eq$stream -or -not $stream.CanRead) { return$null }
        if ($stream.CanSeek) { $stream.Position = 0 }$reader = New-Object System.IO.StreamReader($stream)$body = $reader.ReadToEnd()$reader.Dispose()
        return $body
    } catch {
        return $null
    }
}

function Get-GraphAccessToken {
    $script:etapa = 'solicitar el código de inicio de sesión a Microsoft'
    $authority = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0"
    $device = Invoke-RestMethod -Method POST -Uri "$authority/devicecode" -Body @{
        client_id = $clientId
        scope = 'https://graph.microsoft.com/Files.ReadWrite'
    } -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop

    Write-Host "`n=================================================================" -ForegroundColor Cyan
    Write-Host " AUTENTICACIÓN REQUERIDA PARA GUARDAR EN EXCEL" -ForegroundColor Cyan
    Write-Host " 1. Abra su navegador en: " -NoNewline; Write-Host $($device.verification_uri) -ForegroundColor Yellow
    Write-Host " 2. Ingrese este código:  " -NoNewline; Write-Host $($device.user_code) -ForegroundColor Green
    Write-Host " (Use la cuenta de horacio_sauce_corporacionalisa_com u otra con acceso)"
    Write-Host "=================================================================`n" -ForegroundColor Cyan

    $interval = [math]::Max(5, [int]$device.interval)$deadline = (Get-Date).AddSeconds([int]$device.expires_in)$script:etapa = 'esperar la autorización de Microsoft'
    
    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds $interval
        try {
            $token = Invoke-RestMethod -Method POST -Uri "$authority/token" -Body @{
                grant_type = 'urn:ietf:params:oauth:grant-type:device_code'
                client_id = $clientId
                device_code = $device.device_code
            } -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
            
            if ($token.access_token) { 
                Write-Host "¡Autenticación exitosa! Conectando con SharePoint..." -ForegroundColor Green
                return [string]$token.access_token 
            }
        }
        catch {
            # Atrapamos los errores de forma segura. Mientras no ingreses el código, 
            # Microsoft devuelve un "Bad Request" que ignoramos para que siga consultando.
            continue
        }
    }
    throw 'El código de inicio de sesión expiró antes de completar la autorización.'
}

function Invoke-GraphRequest {
    param(
        [Parameter(Mandatory = $true)][string]$Method,
        [Parameter(Mandatory = $true)][string]$Uri,
        [string]$Body
    )
    $parameters = @{
        Method = $Method
        Uri = $Uri
        Headers = @{ Authorization = "Bearer $script:graphAccessToken" }
        ErrorAction = 'Stop'
    }
    if ($PSBoundParameters.ContainsKey('Body')) {$parameters.Body = $Body$parameters.ContentType = 'application/json; charset=utf-8'
    }
    return Invoke-RestMethod @parameters
}

function Get-ColumnMapping {
    param([Parameter(Mandatory = $true)][object[]]$Columns)$aliases = @{
        Hostname        = @('Hostname', 'Nombre de equipo', 'Nombre del equipo', 'Equipo')
        UsuarioLogueado = @('UsuarioLogueado', 'Usuario logueado', 'Usuario conectado', 'Usuario')
        Marca           = @('Marca', 'Fabricante')
        Modelo          = @('Modelo')
        NumeroSerie     = @('NumeroSerie', 'Número de serie', 'No. de serie', 'Serial', 'Serie')
        Procesador      = @('Procesador', 'CPU')
        TipoEquipo      = @('TipoEquipo', 'Tipo de equipo')
        RAM_GB          = @('RAM_GB', 'RAM (GB)', 'Cantidad de memoria RAM', 'Memoria RAM')
        DiscoTotal_GB   = @('DiscoTotal_GB', 'Disco total', 'Disco total (GB)', 'Total disco', 'Total')
        DiscoUsado_GB   = @('DiscoUsado_GB', 'Disco usado', 'Disco usado (GB)', 'Usado')
        DiscoLibre_GB   = @('DiscoLibre_GB', 'Disco libre', 'Disco libre (GB)', 'Libre')
        TipoDisco       = @('TipoDisco', 'Tipo de disco')
        DireccionIP     = @('DireccionIP', 'Dirección IP', 'IP', 'Tipo IP', 'TipoIP')
        WindowsEdicion  = @('WindowsEdicion', 'Windows', 'Edición de Windows', 'Sistema operativo')
        WindowsVersion  = @('WindowsVersion', 'Versión de Windows', 'Version Windows', 'Version')
        LicenciaWindows = @('LicenciaWindows', 'Estado de licencia de Windows', 'Licencia de Windows', 'Estado de licencia', 'Activado')
        FechaEscaneo    = @('FechaEscaneo', 'Fecha actual del escaneo', 'Fecha de escaneo', 'Fecha')
        Area            = @('Area', 'Área', 'Departamento', 'Division')
    }
    $lookup = @{}
    foreach ($field in$aliases.Keys) {
        foreach ($alias in $aliases[$field]) {
            $key = Normalize-Header$alias
            if ($lookup.ContainsKey($key) -and$lookup[$key] -ne$field) { throw "Alias de columna ambiguo: $alias" }
            $lookup[$key] =$field
        }
    }
    $ordered = @($Columns \vert{} Sort-Object { [int]$_.index })
    $fields = @()$unknown = @()
    foreach ($column in$ordered) {
        $key = Normalize-Header ([string]$column.name)
        if (-not $lookup.ContainsKey($key)) {
            $unknown += [string]$column.name
            continue
        }
        $fields += $lookup[$key]
    }
    $missing = @($aliases.Keys | Where-Object { $_ -notin$fields } | Sort-Object)
    $duplicates = @($fields | Group-Object | Where-Object { $_.Count -gt 1 } \vert{} ForEach-Object {$_.Name })
    return [pscustomobject]@{
        Fields     = $fields
        Missing    = $missing
        Unknown    = $unknown
        Duplicates = $duplicates
        IsComplete = ($missing.Count -eq 0 -and $unknown.Count -eq 0 -and$duplicates.Count -eq 0 -and $fields.Count -eq$Columns.Count)
    }
}

do {
    $area = Read-Host 'Ingrese el área o departamento de este equipo (obligatorio)'
    if ($null -eq$area) { throw 'No se recibió el área.' }
    $area =$area.Trim()
} while (-not $area)

$computer = Get-CimData -ClassName 'Win32_ComputerSystem'
$bios = Get-CimData -ClassName 'Win32_BIOS'$processors = @(Get-CimData -ClassName 'Win32_Processor')
$enclosure = Get-CimData -ClassName 'Win32_SystemEnclosure'$os = Get-CimData -ClassName 'Win32_OperatingSystem'
$volumes = @(Get-CimData -ClassName 'Win32_LogicalDisk' -Filter 'DriveType = 3' \vert{} Where-Object {$null -ne $_ -and$null -ne $_.Size })$adapters = @(Get-CimData -ClassName 'Win32_NetworkAdapterConfiguration' -Filter 'IPEnabled = TRUE')

if (-not $computer -or -not$os) { throw 'No se pudieron consultar los datos del sistema.' }

$hostname = if ($computer.Name) {$computer.Name } else { $env:COMPUTERNAME }$userName = if ($computer.UserName) {$computer.UserName } else { 'Sin sesión interactiva' }

$equipmentType = 'Desconocido'
if ($enclosure) {$chassisTypes = @($enclosure \vert{} ForEach-Object {$_.ChassisTypes })
    if (@($chassisTypes | Where-Object { $_ -in @(8, 9, 10, 11, 12, 14, 18, 21, 30, 31, 32) }).Count -gt 0) { $equipmentType = 'Laptop' }
    elseif (@($chassisTypes | Where-Object { $_ -in @(3, 4, 5, 6, 7, 15, 16, 35, 36) }).Count -gt 0) { $equipmentType = 'Desktop' }
}

$ramGB = if ($computer.TotalPhysicalMemory) { ConvertTo-Gigabytes $computer.TotalPhysicalMemory } else {$null }

$diskTotal = [double]0; $diskFree = [double]0
foreach ($volume in$volumes) {
    $diskTotal += [double]$volume.Size
    if ($null -ne$volume.FreeSpace) { $diskFree += [double]$volume.FreeSpace }
}
$diskTotalGB =$null; $diskUsedGB =$null; $diskFreeGB =$null
if ($volumes.Count -gt 0) {$diskTotalGB = ConvertTo-Gigabytes $diskTotal$diskUsedGB = ConvertTo-Gigabytes ($diskTotal -$diskFree)
    $diskFreeGB = ConvertTo-Gigabytes$diskFree
}

$diskType = 'Desconocido'
if (Get-Command -Name Get-PhysicalDisk -ErrorAction SilentlyContinue) {
    try {
        $physicalDisks = @(Get-PhysicalDisk -ErrorAction Stop \vert{} Where-Object { [string]$_.BusType -notin @('USB', 'SD', 'MMC') })
        $types = @($physicalDisks | ForEach-Object { [string]$_.MediaType } \vert{} Where-Object {$_ -in @('SSD', 'HDD') } | Sort-Object -Unique)
        if ($types.Count -gt 0) { $diskType =$types -join '; ' }
    } catch {}
}

$ipAddresses = @($adapters \vert{} Where-Object {$null -ne $_ } \vert{} ForEach-Object {$_.IPAddress } | Where-Object { $_ -match '^\d{1,3}(\.\d{1,3}){3}$' -and $_ -notmatch '^(127\.\vert{}169\.254\.)' } \vert{} Sort-Object -Unique)$ip = if ($ipAddresses.Count -gt 0) {$ipAddresses -join '; ' } else { 'Desconocido' }

$windowsVersion = 'Desconocido'
try {
    $windowsInfo = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop
    if ($windowsInfo.DisplayVersion) { $windowsVersion =$windowsInfo.DisplayVersion }
    elseif ($windowsInfo.ReleaseId) { $windowsVersion =$windowsInfo.ReleaseId }
} catch { if ($os.Version) { $windowsVersion =$os.Version } }

$activation = 'Desconocido'$licenseProducts = @(Get-CimData -ClassName 'SoftwareLicensingProduct' -Filter "ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' AND PartialProductKey IS NOT NULL" | Where-Object { $null -ne$_ })
if ($licenseProducts.Count -gt 0) {$activation = if (@($licenseProducts \vert{} Where-Object {$_.LicenseStatus -eq 1 }).Count -gt 0) { 'Activado' } else { 'No activado' }
}

$processorNames = @($processors \vert{} Where-Object {$null -ne $_ -and$_.Name } | ForEach-Object { $_.Name.Trim() } \vert{} Sort-Object -Unique)$processor = if ($processorNames.Count -gt 0) {$processorNames -join '; ' } else { 'Desconocido' }

$manufacturer = if ($computer.Manufacturer) { $computer.Manufacturer.Trim() } else { 'Desconocido' }$model = if ($computer.Model) {$computer.Model.Trim() } else { 'Desconocido' }
$serial = if ($bios.SerialNumber) { $bios.SerialNumber.Trim() } else { 'Desconocido' }$edition = if ($os.Caption) {$os.Caption.Trim() } else { 'Desconocido' }

$fechaEscaneo = Get-Date$record = [pscustomobject][ordered]@{
    Hostname            = $hostname
    UsuarioLogueado     = $userName
    Marca               = $manufacturer
    Modelo              = $model
    NumeroSerie         = $serial
    Procesador          = $processor
    TipoEquipo          = $equipmentType
    RAM_GB              = $ramGB
    DiscoTotal_GB       = $diskTotalGB
    DiscoUsado_GB       = $diskUsedGB
    DiscoLibre_GB       = $diskFreeGB
    TipoDisco           = $diskType
    DireccionIP         = $ip
    WindowsEdicion      = $edition
    WindowsVersion      = $windowsVersion
    LicenciaWindows     = $activation
    FechaEscaneo        = $fechaEscaneo.ToString('yyyy-MM-ddTHH:mm:sszzz')
    Area                = $area
}

$null = New-Item -ItemType Directory -Path $DirectorioSalida -Force$safeHostname = $hostname -replace '[^A-Za-z0-9._-]', '_'$fileName = '{0}_{1}_{2}.csv' -f $safeHostname, (Get-Date -Format 'yyyyMMdd_HHmmss_fff'), ([guid]::NewGuid().ToString('N').Substring(0, 8))$csvPath = Join-Path $DirectorioSalida$fileName
$record \vert{} Export-Csv -LiteralPath$csvPath -NoTypeInformation -Encoding UTF8
Write-Host "Respaldo CSV guardado en: $csvPath"

$etapa = 'autenticar con Microsoft'
try {
    $script:graphAccessToken = Get-GraphAccessToken

    $etapa = 'resolver el enlace del libro de SharePoint'$shareBytes = [System.Text.Encoding]::UTF8.GetBytes($urlLibro)$shareToken = 'u!' + [Convert]::ToBase64String($shareBytes).TrimEnd('=').Replace('/', '_').Replace('+', '-')$item = Invoke-GraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/shares/$shareToken/driveItem"
    
    if (-not $item.id -or -not$item.parentReference.driveId) { throw 'Graph no devolvió el identificador.' }

    $driveId = [uri]::EscapeDataString([string]$item.parentReference.driveId)$itemId = [uri]::EscapeDataString([string]$item.id)$baseUri = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$itemId/workbook"
    
    $etapa = 'consultar las tablas del libro'
    $tables = @(Get-GraphCollection -Uri "$baseUri/tables")
    $table = @($tables | Where-Object { $_.name -eq$NombreTabla })
    if ($table.Count -ne 1) { throw "No se encontró la tabla '$NombreTabla'." }

    $tableId = [uri]::EscapeDataString([string]$table[0].id)$tableUri = "$baseUri/tables/$tableId"
    
    $etapa = 'consultar las columnas de la tabla'
    $columns = @(Get-GraphCollection -Uri "$tableUri/columns")
    $mapping = Get-ColumnMapping -Columns$columns
    if (-not $mapping.IsComplete) { throw "Las columnas de Excel no coinciden." }

    $rowValues = @()
    foreach ($field in$mapping.Fields) {
        if ($field -eq 'FechaEscaneo') { $rowValues +=$fechaEscaneo.ToOADate() }
        else { $rowValues += $record.PSObject.Properties[$field].Value }
    }
    $body = @{ values = @(,$rowValues) } | ConvertTo-Json -Depth 4 -Compress
    
    $etapa = 'agregar la fila al libro'$added = Invoke-GraphRequest -Method POST -Uri "$tableUri/rows/add" -Body $body
    if (-not $added) { throw 'Graph no confirmó la creación de la fila.' }
    Write-Host "¡ÉXITO! Equipo registrado en SharePoint: $hostname" -ForegroundColor Green
}
catch {
    $detalle =$_.Exception.Message
    Write-Host "`n[ERROR] No se pudo registrar el equipo en la etapa: '$etapa'." -ForegroundColor Red
    Write-Host "Detalle: $detalle" -ForegroundColor Red
    Write-Host "De todas formas, el respaldo local está a salvo en: $csvPath`n" -ForegroundColor Yellow
}
