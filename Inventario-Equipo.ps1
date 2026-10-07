#requires -Version 5.1
<#
.SYNOPSIS
    Registra el inventario del equipo local en Excel Online y crea un respaldo CSV.
.DESCRIPTION
    Se ejecuta en cada computadora. Pide el área al técnico mediante Read-Host
    y añade una fila a la tabla del libro de inventario en SharePoint mediante
    Microsoft Graph. También guarda un CSV local independiente por escaneo.
.PARAMETER DirectorioSalida
    Directorio local o ruta UNC para los respaldos CSV.
.PARAMETER NombreTabla
    Nombre de la Tabla de Excel, si el libro contiene más de una.
.EXAMPLE
    .\Inventario-Equipo.ps1
#>
[CmdletBinding()]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$DirectorioSalida = (Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'InventarioEquipos'),

    [Parameter()]
    [string]$NombreTabla = 'InventarioEquipos'
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$urlLibro = 'https://grupopinulitogt-my.sharepoint.com/:x:/r/personal/horacio_sauce_corporacionalisa_com/_layouts/15/Doc.aspx?sourcedoc=%7B7F53AF0E-F9E7-493F-97A8-15FFAFB5E5E4%7D&file=INVENTARIO%20DE%20EQUIPOS.xlsx&fromShare=true&action=default&mobileredirect=true'
$tenantId = '0675a017-358d-4fb1-85c3-368320881e85'

function Get-CimData {
    param(
        [Parameter(Mandatory = $true)][string]$ClassName,
        [string]$Filter
    )

    try {
        $parameters = @{ ClassName = $ClassName; ErrorAction = 'Stop' }
        if ($Filter) { $parameters.Filter = $Filter }
        return Get-CimInstance @parameters
    }
    catch {
        Write-Warning "No se pudo consultar $ClassName`: $($_.Exception.Message)"
        return $null
    }
}

function ConvertTo-Gigabytes {
    param([object]$Bytes)
    if ($null -eq $Bytes) { return $null }
    return [math]::Round(([double]$Bytes / 1GB), 2)
}

function Normalize-Header {
    param([Parameter(Mandatory = $true)][string]$Text)
    $decomposed = $Text.Normalize([System.Text.NormalizationForm]::FormD)
    $plain = [regex]::Replace($decomposed, '\p{Mn}', '')
    return [regex]::Replace($plain.ToLowerInvariant(), '[^a-z0-9]', '')
}

function Get-GraphCollection {
    param([Parameter(Mandatory = $true)][string]$Uri)
    $items = @()
    do {
        $response = Invoke-MgGraphRequest -Method GET -Uri $Uri -ErrorAction Stop
        if ($response -is [System.Collections.IDictionary]) {
            $items += @($response['value'])
            $Uri = [string]$response['@odata.nextLink']
        }
        else {
            $items += @($response.value)
            $nextLink = $response.PSObject.Properties['@odata.nextLink']
            $Uri = if ($nextLink) { [string]$nextLink.Value } else { '' }
        }
    } while ($Uri)
    return $items | Where-Object { $null -ne $_ }
}

function Get-ColumnMapping {
    param([Parameter(Mandatory = $true)][object[]]$Columns)

    $aliases = @{
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
        DireccionIP     = @('DireccionIP', 'Dirección IP', 'IP')
        WindowsEdicion  = @('WindowsEdicion', 'Windows', 'Edición de Windows', 'Sistema operativo')
        WindowsVersion  = @('WindowsVersion', 'Versión de Windows', 'Version Windows', 'Version')
        LicenciaWindows = @('LicenciaWindows', 'Estado de licencia de Windows', 'Licencia de Windows', 'Estado de licencia', 'Activado')
        FechaEscaneo    = @('FechaEscaneo', 'Fecha actual del escaneo', 'Fecha de escaneo', 'Fecha')
        Area            = @('Area', 'Área', 'Departamento', 'Division')
    }
    $lookup = @{}
    foreach ($field in $aliases.Keys) {
        foreach ($alias in $aliases[$field]) {
            $key = Normalize-Header $alias
            if ($lookup.ContainsKey($key) -and $lookup[$key] -ne $field) {
                throw "Alias de columna ambiguo: $alias"
            }
            $lookup[$key] = $field
        }
    }

    $ordered = @($Columns | Sort-Object { [int]$_.index })
    $fields = @()
    $unknown = @()
    foreach ($column in $ordered) {
        $key = Normalize-Header ([string]$column.name)
        if (-not $lookup.ContainsKey($key)) {
            $unknown += [string]$column.name
            continue
        }
        $fields += $lookup[$key]
    }
    $missing = @($aliases.Keys | Where-Object { $_ -notin $fields } | Sort-Object)
    $duplicates = @($fields | Group-Object | Where-Object { $_.Count -gt 1 } | ForEach-Object { $_.Name })
    return [pscustomobject]@{
        Fields     = $fields
        Missing    = $missing
        Unknown    = $unknown
        Duplicates = $duplicates
        IsComplete = ($missing.Count -eq 0 -and $unknown.Count -eq 0 -and $duplicates.Count -eq 0 -and $fields.Count -eq $Columns.Count)
    }
}

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    throw 'Falta el módulo Microsoft.Graph.Authentication. Instálelo una vez con: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser'
}

# El área nunca se deduce ni se toma de un parámetro: siempre la escribe el técnico.
do {
    $area = Read-Host 'Ingrese el área o departamento de este equipo (obligatorio)'
    if ($null -eq $area) {
        throw 'No se recibió el área. Ejecute el script en una consola interactiva.'
    }
    $area = $area.Trim()
    if (-not $area) { Write-Warning 'El área no puede quedar vacía.' }
} while (-not $area)

$computer = Get-CimData -ClassName 'Win32_ComputerSystem'
$bios = Get-CimData -ClassName 'Win32_BIOS'
$processors = @(Get-CimData -ClassName 'Win32_Processor')
$enclosure = Get-CimData -ClassName 'Win32_SystemEnclosure'
$os = Get-CimData -ClassName 'Win32_OperatingSystem'
$volumes = @(Get-CimData -ClassName 'Win32_LogicalDisk' -Filter 'DriveType = 3' |
    Where-Object { $null -ne $_ -and $null -ne $_.Size })
$adapters = @(Get-CimData -ClassName 'Win32_NetworkAdapterConfiguration' -Filter 'IPEnabled = TRUE')
if (-not $computer -or -not $os) {
    throw 'No se pudieron consultar los datos básicos del sistema mediante CIM. No se generó el CSV.'
}

$hostname = $env:COMPUTERNAME
if ($computer -and $computer.Name) { $hostname = $computer.Name }

$userName = 'Sin sesión interactiva'
if ($computer -and $computer.UserName) { $userName = $computer.UserName }

$equipmentType = 'Desconocido'
if ($enclosure) {
    $chassisTypes = @($enclosure | ForEach-Object { $_.ChassisTypes })
    if (@($chassisTypes | Where-Object { $_ -in @(8, 9, 10, 11, 12, 14, 18, 21, 30, 31, 32) }).Count -gt 0) {
        $equipmentType = 'Laptop'
    }
    elseif (@($chassisTypes | Where-Object { $_ -in @(3, 4, 5, 6, 7, 15, 16, 35, 36) }).Count -gt 0) {
        $equipmentType = 'Desktop'
    }
}
if ($equipmentType -eq 'Desconocido' -and $computer) {
    if ($computer.PCSystemType -eq 2) { $equipmentType = 'Laptop' }
    elseif ($computer.PCSystemType -in @(1, 3)) { $equipmentType = 'Desktop' }
}

$ramGB = $null
if ($computer -and $null -ne $computer.TotalPhysicalMemory) {
    $ramGB = ConvertTo-Gigabytes $computer.TotalPhysicalMemory
}

$diskTotal = [double]0
$diskFree = [double]0
foreach ($volume in $volumes) {
    $diskTotal += [double]$volume.Size
    if ($null -ne $volume.FreeSpace) { $diskFree += [double]$volume.FreeSpace }
}
$diskTotalGB = $null
$diskUsedGB = $null
$diskFreeGB = $null
if ($volumes.Count -gt 0) {
    $diskTotalGB = ConvertTo-Gigabytes $diskTotal
    if (@($volumes | Where-Object { $null -eq $_.FreeSpace }).Count -eq 0) {
        $diskUsedGB = ConvertTo-Gigabytes ($diskTotal - $diskFree)
        $diskFreeGB = ConvertTo-Gigabytes $diskFree
    }
}

$diskType = 'Desconocido'
if (Get-Command -Name Get-PhysicalDisk -ErrorAction SilentlyContinue) {
    try {
        $physicalDisks = @(Get-PhysicalDisk -ErrorAction Stop |
            Where-Object { [string]$_.BusType -notin @('USB', 'SD', 'MMC') })
        $types = @($physicalDisks |
            ForEach-Object { [string]$_.MediaType } |
            Where-Object { $_ -in @('SSD', 'HDD') } |
            Sort-Object -Unique)
        if ($types.Count -gt 0) { $diskType = $types -join '; ' }
    }
    catch {
        Write-Warning "No se pudo consultar el tipo de disco: $($_.Exception.Message)"
    }
}

$ipAddresses = @($adapters |
    Where-Object { $null -ne $_ } |
    ForEach-Object { $_.IPAddress } |
    Where-Object { $_ -match '^\d{1,3}(\.\d{1,3}){3}$' -and $_ -notmatch '^(127\.|169\.254\.)' } |
    Sort-Object -Unique)
$ip = 'Desconocido'
if ($ipAddresses.Count -gt 0) { $ip = $ipAddresses -join '; ' }

$windowsVersion = 'Desconocido'
try {
    $windowsInfo = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop
    $displayVersion = $windowsInfo.PSObject.Properties['DisplayVersion']
    $releaseId = $windowsInfo.PSObject.Properties['ReleaseId']
    if ($displayVersion -and $displayVersion.Value) { $windowsVersion = $displayVersion.Value }
    elseif ($releaseId -and $releaseId.Value) { $windowsVersion = $releaseId.Value }
    elseif ($os -and $os.Version) { $windowsVersion = $os.Version }
}
catch {
    if ($os -and $os.Version) { $windowsVersion = $os.Version }
    Write-Warning "No se pudo leer la versión comercial de Windows: $($_.Exception.Message)"
}

$activation = 'Desconocido'
$licenseProducts = @(Get-CimData -ClassName 'SoftwareLicensingProduct' -Filter "ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' AND PartialProductKey IS NOT NULL" |
    Where-Object { $null -ne $_ })
if ($licenseProducts.Count -gt 0) {
    $activation = 'No activado'
    if (@($licenseProducts | Where-Object { $_.LicenseStatus -eq 1 }).Count -gt 0) {
        $activation = 'Activado'
    }
}

$processorNames = @($processors |
    Where-Object { $null -ne $_ -and $_.Name } |
    ForEach-Object { $_.Name.Trim() } |
    Sort-Object -Unique)
$processor = 'Desconocido'
if ($processorNames.Count -gt 0) { $processor = $processorNames -join '; ' }

$manufacturer = 'Desconocido'
$model = 'Desconocido'
$serial = 'Desconocido'
$edition = 'Desconocido'
if ($computer -and $computer.Manufacturer) { $manufacturer = $computer.Manufacturer.Trim() }
if ($computer -and $computer.Model) { $model = $computer.Model.Trim() }
if ($bios -and $bios.SerialNumber) { $serial = $bios.SerialNumber.Trim() }
if ($os -and $os.Caption) { $edition = $os.Caption.Trim() }

$fechaEscaneo = Get-Date
$record = [pscustomobject][ordered]@{
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

$null = New-Item -ItemType Directory -Path $DirectorioSalida -Force
$safeHostname = $hostname -replace '[^A-Za-z0-9._-]', '_'
$fileName = '{0}_{1}_{2}.csv' -f $safeHostname, (Get-Date -Format 'yyyyMMdd_HHmmss_fff'), ([guid]::NewGuid().ToString('N').Substring(0, 8))
$csvPath = Join-Path $DirectorioSalida $fileName
$record | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
Write-Host "Respaldo CSV guardado en: $csvPath"

try {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    Connect-MgGraph -TenantId $tenantId -Scopes 'Files.ReadWrite' -UseDeviceAuthentication -ContextScope Process -ErrorAction Stop | Out-Null

    # El enlace compartido se transforma en el identificador que acepta Graph.
    $shareBytes = [System.Text.Encoding]::UTF8.GetBytes($urlLibro)
    $shareToken = 'u!' + [Convert]::ToBase64String($shareBytes).TrimEnd('=').Replace('/', '_').Replace('+', '-')
    $item = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/shares/$shareToken/driveItem" -ErrorAction Stop
    if (-not $item.id -or -not $item.parentReference.driveId) {
        throw 'Graph no devolvió el identificador del libro compartido.'
    }
    if ($item.name -ne 'INVENTARIO DE EQUIPOS.xlsx') {
        throw "El enlace apunta a otro archivo: $($item.name)"
    }

    $driveId = [uri]::EscapeDataString([string]$item.parentReference.driveId)
    $itemId = [uri]::EscapeDataString([string]$item.id)
    $baseUri = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$itemId/workbook"
    $tables = @(Get-GraphCollection -Uri "$baseUri/tables")
    $table = @($tables | Where-Object { $_.name -eq $NombreTabla })
    if ($table.Count -ne 1) {
        $available = @($tables | ForEach-Object { $_.name }) -join ', '
        throw "No se encontró exactamente una tabla '$NombreTabla'. Tablas disponibles: $available"
    }

    $tableId = [uri]::EscapeDataString([string]$table[0].id)
    $tableUri = "$baseUri/tables/$tableId"
    $columns = @(Get-GraphCollection -Uri "$tableUri/columns")
    $mapping = Get-ColumnMapping -Columns $columns
    if (-not $mapping.IsComplete) {
        throw ("Las columnas de Excel no coinciden. Faltan: {0}. Sin asignar: {1}. Duplicadas: {2}." -f
            ($mapping.Missing -join ', '), ($mapping.Unknown -join ', '), ($mapping.Duplicates -join ', '))
    }

    $rowValues = @()
    foreach ($field in $mapping.Fields) {
        if ($field -eq 'FechaEscaneo') {
            # Excel recibe una fecha numérica, compatible con la columna Fecha.
            $rowValues += $fechaEscaneo.ToOADate()
        }
        else {
            $rowValues += $record.PSObject.Properties[$field].Value
        }
    }
    $body = @{ values = @(,$rowValues) } | ConvertTo-Json -Depth 4 -Compress
    $added = Invoke-MgGraphRequest -Method POST -Uri "$tableUri/rows" -Body $body -ContentType 'application/json; charset=utf-8' -ErrorAction Stop
    if (-not $added) {
        throw 'Graph no confirmó la creación de la fila. Revise la tabla antes de reintentar para evitar duplicados.'
    }
    Write-Host "Equipo registrado en SharePoint: $hostname"
}
catch {
    throw "No se pudo registrar el equipo en SharePoint: $($_.Exception.Message) Respaldo disponible en: $csvPath"
}

$record
