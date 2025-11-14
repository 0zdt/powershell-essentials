# Pfad zu deinem Logo
$inputLogo = "C:\Pfad\zum\Logo.png"

# Zielpfad für die fertige Datei
$outputLogo = "C:\Pfad\zum\fertigen\Logo.png"

# Zwischenwerte
$resizedLogo = "C:\Pfad\zum\resizedLogo.png"

# ImageMagick Pfad (wenn 'magick' global verfügbar ist, kann man es so nutzen)
$magick = "magick"

# Schritt 1: Logo auf 256x256 skalieren, proportional
& $magick convert $inputLogo -resize 256x256 $resizedLogo

# Schritt 2: In 516x516 transparenten Hintergrund zentrieren
& $magick convert -size 516x516 xc:none $resizedLogo -gravity center -composite $outputLogo

Write-Host "Logo wurde erfolgreich verarbeitet und gespeichert: $outputLogo"
