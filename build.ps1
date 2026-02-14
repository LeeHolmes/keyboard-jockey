# build.ps1 - Build script for Keyboard Jockey
# Usage: .\build.ps1 [-Configuration Debug|Release] [-Platform x64|Win32] [-Clean]

param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",
    
    [ValidateSet("x64", "Win32")]
    [string]$Platform = "x64",
    
    [switch]$Clean
)

$ErrorActionPreference = "Stop"

# Find MSBuild
$msbuildPaths = @(
    "${env:ProgramFiles}\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe"
    "${env:ProgramFiles}\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe"
    "${env:ProgramFiles}\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
    "${env:ProgramFiles}\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
    "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Enterprise\MSBuild\Current\Bin\MSBuild.exe"
    "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe"
    "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\Community\MSBuild\Current\Bin\MSBuild.exe"
    "${env:ProgramFiles(x86)}\Microsoft Visual Studio\2019\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
)

$msbuild = $null
foreach ($path in $msbuildPaths) {
    if (Test-Path $path) {
        $msbuild = $path
        break
    }
}

# Try vswhere as fallback
if (-not $msbuild) {
    $vswhere = "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path $vswhere) {
        $vsPath = & $vswhere -latest -requires Microsoft.Component.MSBuild -property installationPath
        if ($vsPath) {
            $msbuild = Join-Path $vsPath "MSBuild\Current\Bin\MSBuild.exe"
        }
    }
}

if (-not $msbuild -or -not (Test-Path $msbuild)) {
    Write-Host "ERROR: MSBuild not found. Please install Visual Studio or Build Tools." -ForegroundColor Red
    exit 1
}

Write-Host "Using MSBuild: $msbuild" -ForegroundColor Cyan

# Project paths
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$solution = Join-Path $scriptDir "KeyboardJockey.sln"
$outputDir = Join-Path $scriptDir "$Platform\$Configuration"

# Clean if requested
if ($Clean) {
    Write-Host "`nCleaning..." -ForegroundColor Yellow
    & $msbuild $solution /t:Clean /p:Configuration=$Configuration /p:Platform=$Platform /v:minimal
    
    # Remove output directories
    $dirsToClean = @("x64", "Win32", "Debug", "Release")
    foreach ($dir in $dirsToClean) {
        $fullPath = Join-Path $scriptDir $dir
        if (Test-Path $fullPath) {
            Remove-Item -Recurse -Force $fullPath
            Write-Host "Removed: $dir" -ForegroundColor Gray
        }
    }
}

# Build
Write-Host "`nBuilding Keyboard Jockey..." -ForegroundColor Green
Write-Host "Configuration: $Configuration" -ForegroundColor Gray
Write-Host "Platform: $Platform" -ForegroundColor Gray
Write-Host ""

& $msbuild $solution /p:Configuration=$Configuration /p:Platform=$Platform /v:minimal /m

if ($LASTEXITCODE -eq 0) {
    $exe = Join-Path $outputDir "KeyboardJockey.exe"
    Write-Host "`n========================================" -ForegroundColor Green
    Write-Host "BUILD SUCCEEDED" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "Output: $exe" -ForegroundColor Cyan
    
    if (Test-Path $exe) {
        $fileInfo = Get-Item $exe
        Write-Host "Size: $([math]::Round($fileInfo.Length / 1KB, 2)) KB" -ForegroundColor Gray
    }
} else {
    Write-Host "`n========================================" -ForegroundColor Red
    Write-Host "BUILD FAILED" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    exit 1
}
