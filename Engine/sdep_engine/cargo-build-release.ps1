Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$projectRoot = $PSScriptRoot
$repoRoot = Split-Path -Parent (Split-Path -Parent $projectRoot)
Push-Location $projectRoot
try {
    cargo build --release

    $metadata = cargo metadata --format-version 1 --no-deps | ConvertFrom-Json
    $targetDir = $metadata.target_directory
    $dllName = "sdep_engine.dll"
    $sourceDll = Join-Path $targetDir "release\$dllName"
    $destinationDll = Join-Path $repoRoot $dllName

    if (-not (Test-Path -LiteralPath $sourceDll)) {
        throw "Release DLL not found at $sourceDll"
    }

    Copy-Item -LiteralPath $sourceDll -Destination $destinationDll -Force
    Write-Host "Copied $dllName to repo root."
}
finally {
    Pop-Location
}
