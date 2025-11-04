#!/usr/bin/env pwsh
# Build TypeScript code

Write-Host "Building TypeScript code..." -ForegroundColor Cyan

$ErrorActionPreference = "Stop"

try {
    Push-Location "$PSScriptRoot"
    
    # Install dependencies if needed
    if (-not (Test-Path "node_modules")) {
        Write-Host "Installing dependencies..." -ForegroundColor Yellow
        npm install
    }
    
    # Build TypeScript
    Write-Host "Running TypeScript compiler..." -ForegroundColor Yellow
    npm run build
    
    Write-Host "✓ Build completed successfully!" -ForegroundColor Green
}
catch {
    Write-Host "✗ Build failed: $_" -ForegroundColor Red
    exit 1
}
finally {
    Pop-Location
}
