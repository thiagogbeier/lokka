#!/usr/bin/env pwsh
# Deploy to Azure Container Apps

param(
    [string]$ResourceGroup = "rg-lokka-mcp",
    [string]$ContainerAppName = "lokka-mcp",
    [string]$RegistryName = "acrlokka1761746532",
    [string]$ImageTag = "latest"
)

Write-Host "Deploying to Azure Container Apps..." -ForegroundColor Cyan
Write-Host "Resource Group: $ResourceGroup" -ForegroundColor Gray
Write-Host "Container App: $ContainerAppName" -ForegroundColor Gray
Write-Host "Registry: $RegistryName" -ForegroundColor Gray
Write-Host "Image Tag: $ImageTag" -ForegroundColor Gray
Write-Host ""

$ErrorActionPreference = "Stop"

try {
    Push-Location "$PSScriptRoot"
    
    # Check if logged in to Azure
    Write-Host "Checking Azure login status..." -ForegroundColor Yellow
    $account = az account show 2>$null | ConvertFrom-Json
    if (-not $account) {
        Write-Host "Not logged in to Azure. Logging in..." -ForegroundColor Yellow
        az login
    }
    else {
        Write-Host "✓ Logged in as: $($account.user.name)" -ForegroundColor Green
    }
    
    # Build TypeScript
    Write-Host "`nBuilding TypeScript code..." -ForegroundColor Yellow
    & .\build.ps1
    if ($LASTEXITCODE -ne 0) {
        throw "TypeScript build failed"
    }
    
    # Build and push to Azure Container Registry
    Write-Host "`nBuilding and pushing image to ACR..." -ForegroundColor Yellow
    az acr build `
        --registry $RegistryName `
        --image "lokka-rest-wrapper:$ImageTag" `
        --file Dockerfile `
        .
    
    if ($LASTEXITCODE -ne 0) {
        throw "ACR build failed"
    }
    
    Write-Host "✓ Image built and pushed to ACR" -ForegroundColor Green
    
    # Update Container App
    Write-Host "`nUpdating Container App..." -ForegroundColor Yellow
    az containerapp update `
        --name $ContainerAppName `
        --resource-group $ResourceGroup `
        --image "$RegistryName.azurecr.io/lokka-rest-wrapper:$ImageTag" `
        --target-port 3000 `
        --set-env-vars `
        AUTH_MODE=token `
        PORT=3000 `
        MCP_SERVER_PATH=./build/main.js
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✓ Container App updated successfully!" -ForegroundColor Green
        
        # Get the FQDN
        $fqdn = az containerapp show `
            --name $ContainerAppName `
            --resource-group $ResourceGroup `
            --query "properties.configuration.ingress.fqdn" `
            -o tsv
        
        Write-Host ""
        Write-Host "Deployment complete!" -ForegroundColor Cyan
        Write-Host "API URL: https://$fqdn" -ForegroundColor Green
        Write-Host ""
        Write-Host "Test the API:" -ForegroundColor Cyan
        Write-Host "  Invoke-RestMethod -Uri `"https://$fqdn/health`"" -ForegroundColor Gray
    }
    else {
        throw "Container App update failed"
    }
}
catch {
    Write-Host "✗ Deployment failed: $_" -ForegroundColor Red
    exit 1
}
finally {
    Pop-Location
}
