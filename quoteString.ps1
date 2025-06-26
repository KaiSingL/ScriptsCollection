# Prompt for multiline text input
Write-Host "Paste your text (press Enter twice to finish):"
$lines = @()
while ($true) {
    $line = Read-Host
    if ($line -eq "") { break }
    $lines += $line
}

# Prompt for left and right wrapping characters
$leftChar = Read-Host "Enter the left wrapping character (e.g., ', "", [)"
$rightChar = Read-Host "Enter the right wrapping character (e.g., ', "", ])"

# Prompt for delimiter, defaulting to space
$delimiter = Read-Host "Enter the delimiter (default is space)"
if ($delimiter -eq "") { $delimiter = " " }

# Process the lines: wrap each with left/right characters
$wrappedLines = $lines | ForEach-Object { "$leftChar$_$rightChar" }

# Join the wrapped lines with the delimiter
$result = $wrappedLines -join $delimiter

# Output to console
Write-Host "`nResult:"
Write-Host $result

# Copy to clipboard
$result | Set-Clipboard
Write-Host "`nResult has been copied to the clipboard."