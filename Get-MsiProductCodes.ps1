<#
.SYNOPSIS
    Finds MSI Product Codes (GUIDs) for installed applications.
.DESCRIPTION
    Searches Registry (HKLM/HKCU, 32bit/64bit) for MSI products and extracts GUIDs.
    Simple version: Console output only.
#>

param(
    [string]$Name,
    [switch]$UseWin32Product
)

function Get-UninstallEntries {
    $paths = @(
        @{ Hive = 'HKLM'; View = 'Registry64' },
        @{ Hive = 'HKLM'; View = 'Registry32' },
        @{ Hive = 'HKCU'; View = 'Registry64' }
    )

    $entries = @()
    foreach ($p in $paths) {
        $regView = if ($p.View -eq 'Registry64') { [Microsoft.Win32.RegistryView]::Registry64 } else { [Microsoft.Win32.RegistryView]::Registry32 }
        $baseKey = if ($p.Hive -eq 'HKLM') {
            [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, $regView)
        } else {
            [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::CurrentUser, $regView)
        }
        $uninstKey = $baseKey.OpenSubKey('SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall')
        if (-not $uninstKey) { continue }

        foreach ($subKeyName in $uninstKey.GetSubKeyNames()) {
            if (-not $subKeyName) { continue }
            $subKey = $uninstKey.OpenSubKey($subKeyName)
            if (-not $subKey) { continue }

            $props = @{}
            $subKey.GetValueNames() | ForEach-Object { $props[$_] = $subKey.GetValue($_) }

            $displayName = $props['DisplayName']
            if ($Name -and $displayName -and $displayName -notlike "*$Name*") { continue }

            $entries += [PSCustomObject]@{
                Hive = $p.Hive
                View = $p.View
                KeyName = $subKeyName
                Values = $props
            }
        }
        $baseKey.Close()
        $uninstKey.Close()
    }
    return $entries
}

function Extract-ProductCode {
    param([string]$InputString)
    if (-not $InputString) { return $null }
    $match = [regex]::Match($InputString, '(?i)(\{?[0-9a-f]{8}-(?:[0-9a-f]{4}-){3}[0-9a-f]{12}\}?)')
    if ($match.Success) {
        $guid = $match.Value.ToUpper()
        if ($guid -notmatch '^\{') { $guid = "{ $guid }" }
        if ($guid -notmatch '\}$') { $guid = "$guid}" }
        return $guid
    }
    return $null
}

$entries = Get-UninstallEntries
$result = @()

foreach ($entry in $entries) {
    $values = $entry.Values
    $displayName = $values['DisplayName']
    if (-not $displayName) { continue }

    $productCode = $values['ProductCode']
    if (-not $productCode) {
        @('UninstallString', 'QuietUninstallString', 'InstallSource') | ForEach-Object {
            if (-not $productCode -and $values.ContainsKey($_)) {
                $productCode = Extract-ProductCode -InputString $values[$_]
            }
        }
    }
    if (-not $productCode -and $entry.KeyName -match '^[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}') {
        $productCode = "{ $($entry.KeyName.ToUpper()) }"
    }

    if ($productCode) {
        $result += [PSCustomObject]@{
            DisplayName = $displayName
            Version = $values['DisplayVersion']
            ProductCode = $productCode
            Hive = $entry.Hive
            View = $entry.View
        }
    }
}

if ($UseWin32Product) {
    Write-Warning "Win32_Product is slow and may trigger reconfigures - use for testing only!"
    try {
        Get-CimInstance Win32_Product | ForEach-Object {
            if ($Name -and $_.Name -notlike "*$Name*") { return }
            $result += [PSCustomObject]@{
                DisplayName = $_.Name
                Version = $_.Version
                ProductCode = $_.IdentifyingNumber
                Hive = 'WMI'
                View = ''
            }
        }
    } catch {
        Write-Warning "WMI query failed: $_"
    }
}

$final = $result | Sort-Object DisplayName, ProductCode -Unique

if ($final) {
    $final | Format-Table -AutoSize
    Write-Host "Found $($final.Count) entries." -ForegroundColor Green
} else {
    Write-Host "No matching entries found." -ForegroundColor Yellow
}