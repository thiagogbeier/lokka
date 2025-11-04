#!/usr/bin/env pwsh
# Build Docker image locally

param(
    [string]$Tag = "latest"
)

Write-Host "Building Docker image..." -ForegroundColor Cyan

$ErrorActionPreference = "Stop"

try {
    Push-Location "$PSScriptRoot"
    
    # Build TypeScript first
    Write-Host "Building TypeScript code first..." -ForegroundColor Yellow
    & .\build.ps1
    if ($LASTEXITCODE -ne 0) {
        throw "TypeScript build failed"
    }
    
    # Build Docker image
    Write-Host "Building Docker image: lokka-rest-wrapper:$Tag" -ForegroundColor Yellow
    docker build -t "lokka-rest-wrapper:$Tag" -f Dockerfile .
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✓ Docker image built successfully!" -ForegroundColor Green
        Write-Host ""
        Write-Host "To run locally:" -ForegroundColor Cyan
        Write-Host "  docker run --rm -p 3000:3000 -e AUTH_MODE=token lokka-rest-wrapper:$Tag" -ForegroundColor Gray
    }
    else {
        throw "Docker build failed"
    }
}
catch {
    Write-Host "✗ Docker build failed: $_" -ForegroundColor Red
    exit 1
}
finally {
    Pop-Location
}
