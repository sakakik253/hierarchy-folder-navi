# Fix encoding to UTF-8 with BOM
$filePath = Join-Path $PSScriptRoot "FolderNavigator_Phase3_Complete.ps1"
Write-Host "Converting encoding: $filePath" -ForegroundColor Cyan

# Read file
$content = Get-Content -Path $filePath -Raw -Encoding UTF8

# Save as UTF-8 with BOM
$utf8WithBom = New-Object System.Text.UTF8Encoding $true
[System.IO.File]::WriteAllText($filePath, $content, $utf8WithBom)

Write-Host "Done! Saved as UTF-8 with BOM." -ForegroundColor Green
