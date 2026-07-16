Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$projectRoot = $PSScriptRoot
$repoRoot = Split-Path -Parent (Split-Path -Parent $projectRoot)

Push-Location $projectRoot
try {
    cargo build --release

    $metadata = cargo metadata --format-version 1 --no-deps | ConvertFrom-Json
    $targetDir = $metadata.target_directory
    $exeName = "sdep_converter.exe"
    $sourceExe = Join-Path $targetDir "release\$exeName"
    $destinationExe = Join-Path $repoRoot $exeName

    if (-not (Test-Path -LiteralPath $sourceExe)) {
        throw "Release executable not found at $sourceExe"
    }

    Copy-Item -LiteralPath $sourceExe -Destination $destinationExe -Force
    Write-Host "Copied $exeName to repo root."
}
finally {
    Pop-Location
}
