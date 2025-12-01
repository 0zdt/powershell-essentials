# center-to-512x512.ps1
# Usage: .\center-to-512x512.ps1 image.jpg
#        or drag & drop an image onto the script

param (
    [Parameter(Mandatory=$true)][string]$InputFile
)

# Check if the file exists
if (-not (Test-Path $InputFile)) {
    Write-Host "File not found: $InputFile"
    exit 1
}

# Resolve full path and prepare output filename
$InputFile  = Resolve-Path $InputFile
$Folder     = Split-Path $InputFile -Parent
$BaseName   = [IO.Path]::GetFileNameWithoutExtension($InputFile)
$OutputFile = Join-Path $Folder "$BaseName`_centered.png"

# Run ImageMagick: resize to 256x256 → center on 512x512 transparent canvas → save as PNG
magick "$InputFile" `
    -resize 256x256 `
    -background none `
    -gravity center `
    -extent 512x512 `
    "$OutputFile"

# Final feedback
if (Test-Path $OutputFile) {
    Write-Host "Done: $OutputFile"
} else {
    Write-Host "Error: Failed to create the output file."
}