$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

$manifest = Get-Content -Raw -LiteralPath "manifest.json" | ConvertFrom-Json
$output = "mc2zotero-$($manifest.version).xpi"
$temporary = [System.IO.Path]::ChangeExtension($output, ".zip")

Remove-Item -LiteralPath $output, $temporary -Force -ErrorAction SilentlyContinue
Compress-Archive -Path "manifest.json", "bootstrap.js", "prefs.js", "content" -DestinationPath $temporary -Force
Move-Item -LiteralPath $temporary -Destination $output
Write-Output "Built $output"
