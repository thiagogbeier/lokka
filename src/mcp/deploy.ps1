#!/usr/bin/env pwsh
# Complete deployment: Build + Deploy to Azure

param(
    [string]$ResourceGroup = "rg-lokka-mcp",
    [string]$ContainerAppName = "lokka-mcp",
    [string]$RegistryName = "lokkamcp",
    [string]$ImageTag = "latest"
)

Write-Host "═══════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Lokka MCP - Complete Deployment" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

$ErrorActionPreference = "Stop"

try {
    & "$PSScriptRoot\azure-deploy.ps1" `
        -ResourceGroup $ResourceGroup `
        -ContainerAppName $ContainerAppName `
        -RegistryName $RegistryName `
        -ImageTag $ImageTag
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host ""
        Write-Host "═══════════════════════════════════════" -ForegroundColor Green
        Write-Host "  ✓ Deployment Successful!" -ForegroundColor Green
        Write-Host "═══════════════════════════════════════" -ForegroundColor Green
    }
}
catch {
    Write-Host ""
    Write-Host "═══════════════════════════════════════" -ForegroundColor Red
    Write-Host "  ✗ Deployment Failed!" -ForegroundColor Red
    Write-Host "═══════════════════════════════════════" -ForegroundColor Red
    Write-Host ""
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}
