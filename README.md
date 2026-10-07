# Inventario de equipos

`Inventario-Equipo.ps1` obtiene los datos del equipo donde se ejecuta, pide al técnico el área o división y agrega una fila a la tabla `InventarioEquipos` del libro de Excel de SharePoint. También guarda un CSV de respaldo en `%LOCALAPPDATA%\InventarioEquipos`.

## Ejecutar sin copiar el archivo manualmente

Pega este comando en **Windows PowerShell 5.1** de cada equipo. Descarga y ejecuta la versión actual de `main` sin copiar el archivo manualmente:

```powershell
iex (irm 'https://raw.githubusercontent.com/Gem3rC4/InventarioExcel/main/Inventario-Equipo.ps1')
```

El técnico debe escribir el Área y completar el inicio de sesión de Microsoft con una cuenta que pueda editar el libro. El script usa llamadas HTTP directas a Microsoft Graph y **no necesita instalar `Microsoft.Graph.Authentication`**. Microsoft puede pedir consentimiento para el permiso delegado `Files.ReadWrite`; la organización puede exigir aprobación del administrador. El token solo se mantiene en memoria durante esta ejecución.

## Relación con el script anterior de Google Sheets

El script anterior enviaba JSON a un **Google Apps Script** mediante `Invoke-RestMethod -Method Post`. Ese Apps Script escribía en Google Sheets. El enlace público de un libro de Excel en SharePoint no es un endpoint de escritura equivalente: Microsoft exige una identidad autorizada. Aquí PowerShell obtiene un token mediante el flujo de código de dispositivo y hace el `POST` a Microsoft Graph. Así se evita el fallo de `Connect-MgGraph` observado con el módulo 2.41.1.

Para eliminar también el inicio de sesión en cada computadora se necesitaría un servicio intermediario, por ejemplo un flujo de Power Automate con el disparador HTTP y la acción **Excel Online (Business) → Add a row into a table**. El disparador HTTP requiere licencia Premium. Su URL o clave de acceso no se debe publicar en este repositorio público.

## Si falla un envío

El script conserva el CSV e indica la etapa. Si falla durante `agregar la fila al libro`, comprueba si ya existe la fila en Excel antes de repetir para evitar duplicados. No envíes contraseñas, códigos de inicio de sesión ni tokens al pedir soporte.

## Requisitos

- Acceso HTTPS a `raw.githubusercontent.com`, `login.microsoftonline.com`, `graph.microsoft.com` y SharePoint.
- Una política de ejecución de PowerShell que permita ejecutar el script. Si la organización exige scripts firmados, el equipo de TI debe firmar este archivo con un certificado de confianza antes de desplegarlo.
- El libro de SharePoint debe conservar la tabla `InventarioEquipos` y sus 18 encabezados. El script comprueba los nombres antes de escribir.

El repositorio contiene código, no credenciales. Se debe proteger la cuenta de GitHub y revisar los cambios a `main` porque el bloque anterior ejecuta la versión publicada allí. Si se requiere inmovilizar una versión, usa la URL de un commit específico y comprueba su SHA-256 antes de ejecutarla.
