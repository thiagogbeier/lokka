#!/usr/bin/env pwsh
# Quick login script for admin@letsintune.com

$TenantId = "b41f1ee6-0ebd-4439-bbbc-07b635f451e0"
$AdminAccount = "admin@letsintune.com"

Write-Host "Logging in to Azure..." -ForegroundColor Cyan
Write-Host "Tenant: letsintune.com ($TenantId)" -ForegroundColor Gray
Write-Host "Admin: $AdminAccount" -ForegroundColor Gray
Write-Host ""

try {
    # Check current login
    $currentAccount = az account show 2>$null | ConvertFrom-Json
    
    if ($currentAccount -and $currentAccount.user.name -eq $AdminAccount) {
        Write-Host "✓ Already logged in as $AdminAccount" -ForegroundColor Green
        Write-Host "Subscription: $($currentAccount.name)" -ForegroundColor Gray
        Write-Host "Tenant: $($currentAccount.tenantId)" -ForegroundColor Gray
    }
    else {
        Write-Host "Logging in..." -ForegroundColor Yellow
        az login --tenant $TenantId
        
        if ($LASTEXITCODE -eq 0) {
            $account = az account show | ConvertFrom-Json
            Write-Host "✓ Logged in successfully as $($account.user.name)" -ForegroundColor Green
        }
        else {
            throw "Login failed"
        }
    }
    
    Write-Host ""
    Write-Host "Ready to deploy!" -ForegroundColor Green
    Write-Host "Run: npm run deploy" -ForegroundColor Cyan
}
catch {
    Write-Host "✗ Login failed: $_" -ForegroundColor Red
    exit 1
}
