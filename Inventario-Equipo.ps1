#requires -Version 5.1
<#
.SYNOPSIS
    Registra el inventario del equipo local en un Google Sheet mediante Google Apps Script.
#>

$DirectorioSalida = Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'InventarioEquipos'

# URL de la Web App de Google Apps Script
$webAppUrl = 'https://script.google.com/macros/s/AKfycbxDlLb51gb_sNcpstJmB8py8WHzl7qteRyTlgQUBQjwGs8JWGviJGngscUXJaM7JIIX/exec'

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Get-CimData {
    param([Parameter(Mandatory = $true)][string]$ClassName, [string]$Filter)
    try {
        $parameters = @{ ClassName = $ClassName; ErrorAction = 'Stop' }
        if ($Filter) { $parameters.Filter = $Filter }
        return Get-CimInstance @parameters
    } catch { return $null }
}

function ConvertTo-Gigabytes {
    param([object]$Bytes)
    if ($null -eq $Bytes) { return $null }
    return [math]::Round(([double]$Bytes / 1GB), 2)
}

do {
    $area = Read-Host 'Ingrese el area o departamento de este equipo (obligatorio)'
    if ($null -eq $area) { throw 'No se recibio el area.' }
    $area = $area.Trim()
} while (-not $area)

$computer = Get-CimData -ClassName 'Win32_ComputerSystem'
$bios = Get-CimData -ClassName 'Win32_BIOS'
$processors = @(Get-CimData -ClassName 'Win32_Processor')
$enclosure = Get-CimData -ClassName 'Win32_SystemEnclosure'
$os = Get-CimData -ClassName 'Win32_OperatingSystem'
$volumes = @(Get-CimData -ClassName 'Win32_LogicalDisk' -Filter 'DriveType = 3' | Where-Object { $null -ne $_ -and $null -ne $_.Size })
$adapters = @(Get-CimData -ClassName 'Win32_NetworkAdapterConfiguration' -Filter 'IPEnabled = TRUE')

if (-not $computer -or -not $os) { throw 'No se pudieron consultar los datos del sistema.' }

$hostname = if ($computer.Name) { $computer.Name } else { $env:COMPUTERNAME }

# === CORTE DEL NOMBRE DE USUARIO ===
$userName = if ($computer.UserName) { ($computer.UserName -split '\\')[-1] } else { 'Sin sesion interactiva' }

$equipmentType = 'Desconocido'
if ($enclosure) {
    $chassisTypes = @($enclosure | ForEach-Object { $_.ChassisTypes })
    if (@($chassisTypes | Where-Object { $_ -in @(8, 9, 10, 11, 12, 14, 18, 21, 30, 31, 32) }).Count -gt 0) { $equipmentType = 'Laptop' }
    elseif (@($chassisTypes | Where-Object { $_ -in @(3, 4, 5, 6, 7, 15, 16, 35, 36) }).Count -gt 0) { $equipmentType = 'Desktop' }
}

$ramGB = if ($computer.TotalPhysicalMemory) { ConvertTo-Gigabytes $computer.TotalPhysicalMemory } else { $null }

$diskTotal = [double]0; $diskFree = [double]0
foreach ($volume in $volumes) {
    $diskTotal += [double]$volume.Size
    if ($null -ne $volume.FreeSpace) { $diskFree += [double]$volume.FreeSpace }
}
$diskTotalGB = $null; $diskUsedGB = $null; $diskFreeGB = $null
if ($volumes.Count -gt 0) {
    $diskTotalGB = ConvertTo-Gigabytes $diskTotal
    $diskUsedGB = ConvertTo-Gigabytes ($diskTotal - $diskFree)
    $diskFreeGB = ConvertTo-Gigabytes $diskFree
}

$diskType = 'Desconocido'
if (Get-Command -Name Get-PhysicalDisk -ErrorAction SilentlyContinue) {
    try {
        $physicalDisks = @(Get-PhysicalDisk -ErrorAction Stop | Where-Object { [string]$_.BusType -notin @('USB', 'SD', 'MMC') })
        $types = @($physicalDisks | ForEach-Object { [string]$_.MediaType } | Where-Object { $_ -in @('SSD', 'HDD') } | Sort-Object -Unique)
        if ($types.Count -gt 0) { $diskType = $types -join '; ' }
    } catch {}
}

$ipAddresses = @($adapters | Where-Object { $null -ne $_ } | ForEach-Object { $_.IPAddress } | Where-Object { $_ -match '^\d{1,3}(\.\d{1,3}){3}$' -and $_ -notmatch '^(127\.|169\.254\.)' } | Sort-Object -Unique)
$ip = if ($ipAddresses.Count -gt 0) { $ipAddresses -join '; ' } else { 'Desconocido' }

$windowsVersion = 'Desconocido'
try {
    $windowsInfo = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop
    if ($windowsInfo.DisplayVersion) { $windowsVersion = $windowsInfo.DisplayVersion }
    elseif ($windowsInfo.ReleaseId) { $windowsVersion = $windowsInfo.ReleaseId }
} catch { if ($os.Version) { $windowsVersion = $os.Version } }

$activation = 'Desconocido'
$licenseProducts = @(Get-CimData -ClassName 'SoftwareLicensingProduct' -Filter "ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' AND PartialProductKey IS NOT NULL" | Where-Object { $null -ne $_ })
if ($licenseProducts.Count -gt 0) {
    $activation = if (@($licenseProducts | Where-Object { $_.LicenseStatus -eq 1 }).Count -gt 0) { 'Activado' } else { 'No activado' }
}

$processorNames = @($processors | Where-Object { $null -ne $_ -and $_.Name } | ForEach-Object { $_.Name.Trim() } | Sort-Object -Unique)
$processor = if ($processorNames.Count -gt 0) { $processorNames -join '; ' } else { 'Desconocido' }

$manufacturer = if ($computer.Manufacturer) { $computer.Manufacturer.Trim() } else { 'Desconocido' }
$model = if ($computer.Model) { $computer.Model.Trim() } else { 'Desconocido' }
$serial = if ($bios.SerialNumber) { $bios.SerialNumber.Trim() } else { 'Desconocido' }
$edition = if ($os.Caption) { $os.Caption.Trim() } else { 'Desconocido' }

$fechaEscaneo = Get-Date

# Crear el objeto con los datos
$record = [pscustomobject][ordered]@{
    Equipo              = $hostname
    Usuario             = $userName
    Marca               = $manufacturer
    Modelo              = $model
    Serie               = $serial
    Procesador          = $processor
    TipoEquipo          = $equipmentType
    RAM_GB              = $ramGB
    DiscoTotal_GB       = $diskTotalGB
    DiscoUsado_GB       = $diskUsedGB
    DiscoLibre_GB       = $diskFreeGB
    TipoDisco           = $diskType
    IP                  = $ip
    Windows             = $edition
    Version             = $windowsVersion
    Activado            = $activation
    Fecha               = $fechaEscaneo.ToString('yyyy-MM-dd HH:mm:ss')
    Area                = $area
}

# 1. Guardar copia local CSV
$null = New-Item -ItemType Directory -Path $DirectorioSalida -Force
$safeHostname = $hostname -replace '[^A-Za-z0-9._-]', '_'
$fileName = '{0}_{1}_{2}.csv' -f $safeHostname, (Get-Date -Format 'yyyyMMdd_HHmmss_fff'), ([guid]::NewGuid().ToString('N').Substring(0, 8))
$csvPath = Join-Path $DirectorioSalida $fileName
$record | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
Write-Host "Respaldo CSV guardado en: $csvPath"

# 2. Enviar a Google Sheets
Write-Host "`nEnviando datos a Google Sheets..." -ForegroundColor Cyan

# Convertir el registro a JSON
$jsonPayload =$record | ConvertTo-Json -Depth 3

try {
    $response = Invoke-RestMethod -Uri $webAppUrl -Method Post -Body$jsonPayload -ContentType 'application/json' -ErrorAction Stop
    
    if ($response.status -eq 'success') {
        Write-Host "¡EXITO! Equipo registrado correctamente en Google Sheets." -ForegroundColor Green
    } else {
        throw "El servidor devolvio un estado no exitoso: $($response.message)"
    }
}
catch {
    Write-Host "`n[ERROR] No se pudo enviar a Google Sheets." -ForegroundColor Red
    Write-Host "Detalle: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "De todas formas, el respaldo local esta a salvo en: $csvPath`n" -ForegroundColor Yellow
}
