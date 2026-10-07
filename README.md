# Inventario de equipos

`Inventario-Equipo.ps1` obtiene los datos del equipo donde se ejecuta, pide al técnico el área o división y agrega una fila a la tabla `InventarioEquipos` del libro de Excel de SharePoint. También guarda un CSV de respaldo en `%LOCALAPPDATA%\InventarioEquipos`.

## Ejecutar sin copiar el archivo manualmente

Pega este bloque en **Windows PowerShell 5.1** de cada equipo. Descarga la versión actual de `main` a un archivo temporal, la ejecuta y borra ese archivo al terminar. La instalación del módulo solo se hace si falta.

```powershell
$ErrorActionPreference = 'Stop'
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Repository PSGallery
}
$scriptTemporal = Join-Path $env:TEMP ("Inventario-Equipo-{0}.ps1" -f [guid]::NewGuid().ToString('N'))
try {
    Invoke-WebRequest -UseBasicParsing -Uri 'https://raw.githubusercontent.com/Gem3rC4/InventarioExcel/main/Inventario-Equipo.ps1' -OutFile $scriptTemporal
    & $scriptTemporal
}
finally {
    if (Test-Path -LiteralPath $scriptTemporal) {
        Remove-Item -LiteralPath $scriptTemporal -Force
    }
}
```

El técnico debe escribir el área y completar el inicio de sesión de Microsoft Graph con una cuenta que pueda editar el libro. La autorización de Graph requiere el permiso delegado `Files.ReadWrite`; la organización puede exigir consentimiento del administrador. La sesión de Graph se limita al proceso actual.

## Si falla antes de mostrar el código de inicio de sesión

El script indica la etapa y la versión de `Microsoft.Graph.Authentication` que falló; el CSV queda guardado. En el equipo afectado, comprueba la versión instalada:

```powershell
Get-Module -ListAvailable Microsoft.Graph.Authentication | Select-Object Name, Version
```

Para actualizar el módulo desde PowerShell Gallery, ejecuta lo siguiente y luego abre una consola nueva de Windows PowerShell:

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Repository PSGallery -Force
```

Comprueba el inicio de sesión por separado antes de repetir el inventario:

```powershell
Connect-MgGraph -TenantId '0675a017-358d-4fb1-85c3-368320881e85' -Scopes 'Files.ReadWrite' -UseDeviceAuthentication -ContextScope Process -ErrorAction Stop
```

Si aparece un código, complétalo en `https://microsoft.com/devicelogin` con la cuenta que tiene acceso al libro. Si continúa el error, informa la etapa, la versión del módulo y la causa interna que muestra el script; no envíes contraseñas, códigos ni tokens. Si la falla sucede durante `agregar la fila al libro`, comprueba antes en Excel si ya existe la fila para evitar duplicados.

## Requisitos

- Acceso HTTPS a `raw.githubusercontent.com`, `graph.microsoft.com` y SharePoint.
- Acceso a PowerShell Gallery la primera vez que se instale `Microsoft.Graph.Authentication`.
- Una política de ejecución de PowerShell que permita ejecutar el script. Si la organización exige scripts firmados, el equipo de TI debe firmar este archivo con un certificado de confianza antes de desplegarlo.
- El libro de SharePoint debe conservar la tabla `InventarioEquipos` y sus 18 encabezados. El script comprueba los nombres antes de escribir.

El repositorio contiene código, no credenciales. Se debe proteger la cuenta de GitHub y revisar los cambios a `main` porque el bloque anterior ejecuta la versión publicada allí. Si se requiere inmovilizar una versión, usa la URL de un commit específico y comprueba su SHA-256 antes de ejecutarla.
