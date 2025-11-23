param([string]$FilePath = "FolderNavigator_Phase3.5_Complete.ps1")

Write-Host "Converting encoding: $FilePath"

try {
    $content = Get-Content -Path $FilePath -Raw -Encoding UTF8
    [System.IO.File]::WriteAllText($FilePath, $content, [System.Text.UTF8Encoding]::new($true))
    Write-Host "Done! Saved as UTF-8 with BOM." -ForegroundColor Green
}
catch {
    Write-Host "Error converting encoding: $_" -ForegroundColor Red
}

