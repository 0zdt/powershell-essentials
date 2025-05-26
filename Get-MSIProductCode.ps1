<#
.SYNOPSIS
    Retrieves the ProductCode from MSI files or lists installed software from the Windows registry
    including architecture (x86/x64) and ProductCode information.

.DESCRIPTION
    This script serves two main purposes:

    1. Extracts the ProductCode (GUID) from one or more MSI files using the Windows Installer COM object.
    2. Lists all installed applications on the system by querying the following registry locations:
       - 64-bit: HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\
       - 32-bit: HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\

    The output includes properties such as:
      - DisplayName
      - ProductCode (GUID or registry key)
      - Publisher
      - Version
      - InstallDate
      - Architecture (x86 or x64)

    Optionally, the output can be filtered by software name and exported to a CSV file.

.PARAMETER Path
    Optional. One or more paths to MSI files to extract the ProductCode from.

.PARAMETER Filter
    Optional. A string filter to limit the displayed installed software by DisplayName.

.PARAMETER ExportCsvPath
    Optional. The file path where the software list will be exported as a CSV file.

.EXAMPLE
    .\Get-MSIProductCode.ps1
    Lists all installed software (both 32-bit and 64-bit) from the registry.

.EXAMPLE
    .\Get-MSIProductCode.ps1 -Filter "Microsoft"
    Displays all installed Microsoft products.

.EXAMPLE
    .\Get-MSIProductCode.ps1 -Filter "Adobe" -ExportCsvPath "C:\Temp\AdobeList.csv"
    Exports all found Adobe software entries to a CSV file.

.EXAMPLE
    .\Get-MSIProductCode.ps1 -Path "C:\Install\MySetup.msi"
    Outputs the ProductCode from the specified MSI file.

.EXAMPLE
    .\Get-MSIProductCode.ps1 -Path "C:\App1.msi", "D:\App2.msi"
    Outputs ProductCodes from multiple MSI files.

.NOTES
    File      : Get-MSIProductCode.ps1
    Author    : 0zdt
    Version   : 1.0
    Requires  : PowerShell 5.1 or higher, Windows Installer COM component
    Platform  : Windows
#>

[CmdletBinding()]
param (
    [Parameter(ValueFromPipeline = $true, Position = 0)]
    [string[]]$Path,

    [Parameter(Position = 1)]
    [string]$Filter,

    [Parameter(Position = 2)]
    [string]$ExportCsvPath
)

function Get-MsiProductCode {
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [ValidateScript({
            if ($_ -and $_.EndsWith('.msi')) {
                $true
            } else {
                throw "'$_' ist keine gültige MSI-Datei (.msi)."
            }
        })]
        [string]$Path
    )

    process {
        try {
            $installer = New-Object -ComObject WindowsInstaller.Installer
            $db = $installer.OpenDatabase((Get-Item $Path).FullName, 0)
            $view = $db.OpenView("SELECT Value FROM Property WHERE Property = 'ProductCode'")
            $view.Execute()
            $record = $view.Fetch()
            if ($record) {
                $productCode = $record.StringData(1)
                [PSCustomObject]@{
                    File        = $Path
                    ProductCode = $productCode
                }
            } else {
                Write-Warning "Kein ProductCode in Datei gefunden: $Path"
            }
            $view.Close()
        } catch {
            Write-Error "Fehler beim Auslesen von '$Path': $_"
        } finally {
            if ($installer) {
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($installer)
            }
        }
    }
}

function Get-InstalledSoftware {
    [CmdletBinding()]
    param ()

    $registrySources = @(
        @{ Path = "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*"; Architecture = "x64" },
        @{ Path = "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"; Architecture = "x86" }
    )

    foreach ($source in $registrySources) {
        Get-ItemProperty -Path $source.Path -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName } |
            Select-Object @{
                Name = 'DisplayName'; Expression = { $_.DisplayName }
            }, @{
                Name = 'ProductCode'; Expression = { $_.PSChildName }
            }, Publisher, InstallDate, Version, UninstallString, @{
                Name = 'Architecture'; Expression = { $source.Architecture }
            }
    }
}

# Hauptlogik
if ($Path) {
    foreach ($p in $Path) {
        Get-MsiProductCode -Path $p
    }
} else {
    $softwareList = Get-InstalledSoftware

    if ($Filter) {
        $softwareList = $softwareList | Where-Object { $_.DisplayName -like "*$Filter*" }
    }

    if ($ExportCsvPath) {
        try {
            $softwareList | Export-Csv -Path $ExportCsvPath -NoTypeInformation -Encoding UTF8
            Write-Host "Export erfolgreich: '$ExportCsvPath'" -ForegroundColor Green
        } catch {
            Write-Error "Export fehlgeschlagen: $_"
        }
    } else {
        $softwareList | Sort-Object DisplayName |
            Format-Table DisplayName, ProductCode, Publisher, Version, InstallDate, Architecture -AutoSize
    }
}
