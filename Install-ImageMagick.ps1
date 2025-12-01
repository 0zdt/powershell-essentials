# Install ImageMagick 7.1.2-9 silently + enforce secure policy.xml
# The .exe must be in the same folder as this script

$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$installer  = Join-Path $scriptPath "ImageMagick-7.1.2-9-Q16-HDRI-x64-dll.exe"

if (-not (Test-Path $installer)) {
    Write-Error "Installer not found: $installer"
    pause
    exit 1
}

Write-Host "Starting silent installation of ImageMagick 7.1.2-9..."
Start-Process -FilePath $installer -ArgumentList "/VERYSILENT /NORESTART /MERGETASKS=!runcode" -Wait

# Wait until installation directory exists
Start-Sleep -Seconds 10

# Find the installed ImageMagick folder
$imDir = Get-Item "$env:ProgramFiles\ImageMagick-*" | Sort-Object Name -Descending | Select-Object -First 1
if (-not $imDir) {
    Write-Error "ImageMagick folder not found!"
    pause
    exit 1
}

Write-Host "ImageMagick installed to: $($imDir.FullName)"

# Write a strict secure policy.xml
$policyFile = "$($imDir.FullName)\policy.xml"

$safePolicy = @'
<policymap>
  <!-- Deny everything by default, then explicitly allow only safe formats -->
  <policy domain="coder" rights="none" pattern="*" />
  <policy domain="coder" rights="read|write" pattern="{JPG,JPEG,PNG,GIF,WEBP,BMP,TIF,TIFF,ICO,SVG}" />

  <!-- Block dangerous coders and protocols completely -->
  <policy domain="coder" rights="none" pattern="EPHEMERAL" />
  <policy domain="coder" rights="none" pattern="URL" />
  <policy domain="coder" rights="none" pattern="HTTP" />
  <policy domain="coder" rights="none" pattern="HTTPS" />
  <policy domain="coder" rights="none" pattern="MVG" />
  <policy domain="coder" rights="none" pattern="MSL" />
  <policy domain="coder" rights="none" pattern="TEXT" />
  <policy domain="coder" rights="none" pattern="LABEL" />
  <policy domain="coder" rights="none" pattern="SHOW" />
  <policy domain="coder" rights="none" pattern="WIN" />

  <!-- Resource limits (DoS protection) -->
  <policy domain="resource" name="time" value="30"/>
  <policy domain="resource" name="memory" value="512MiB"/>
  <policy domain="resource" name="map" value="1GiB"/>
  <policy domain="resource" name="area" value="100MP"/>
  <policy domain="resource" name="disk" value="2GiB"/>
  <policy domain="resource" name="thread" value="4"/>
</policymap>
'@

$safePolicy | Out-File -FilePath $policyFile -Encoding UTF8 -Force
Write-Host "Secure policy.xml created: $policyFile"

# Add to PATH (current session + permanently for all users)
# $env:Path += ";$($imDir.FullName)"
# [Environment]::SetEnvironmentVariable("Path", $env:Path, "Machine")
