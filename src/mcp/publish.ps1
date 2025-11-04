#!/usr/bin/env pwsh
# Publish Lokka to NPM

param(
    [ValidateSet("patch", "minor", "major")]
    [string]$VersionBump = "minor",
    [switch]$DryRun = $false
)

Write-Host "═══════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Lokka NPM Publishing" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

$ErrorActionPreference = "Stop"

try {
    Push-Location "$PSScriptRoot"
    
    # Check if git working directory is clean
    Write-Host "Checking git status..." -ForegroundColor Yellow
    $gitStatus = git status --porcelain
    if ($gitStatus) {
        Write-Host "⚠️  Warning: You have uncommitted changes:" -ForegroundColor Yellow
        Write-Host $gitStatus -ForegroundColor Gray
        $response = Read-Host "Continue anyway? (y/N)"
        if ($response -ne "y") {
            Write-Host "❌ Publishing cancelled" -ForegroundColor Red
            exit 1
        }
    }
    else {
        Write-Host "✓ Working directory is clean" -ForegroundColor Green
    }
    
    # Check npm login
    Write-Host "`nChecking npm login..." -ForegroundColor Yellow
    try {
        $npmUser = npm whoami 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host "✓ Logged in as: $npmUser" -ForegroundColor Green
        }
        else {
            throw "Not logged in"
        }
    }
    catch {
        Write-Host "❌ Not logged in to npm" -ForegroundColor Red
        Write-Host "Run: npm login" -ForegroundColor Yellow
        exit 1
    }
    
    # Clean and build
    Write-Host "`nCleaning previous build..." -ForegroundColor Yellow
    if (Test-Path "build") {
        Remove-Item -Recurse -Force build
    }
    
    Write-Host "Building TypeScript..." -ForegroundColor Yellow
    npm run build
    if ($LASTEXITCODE -ne 0) {
        throw "Build failed"
    }
    Write-Host "✓ Build completed" -ForegroundColor Green
    
    # Get current version
    $packageJson = Get-Content "package.json" | ConvertFrom-Json
    $currentVersion = $packageJson.version
    Write-Host "`nCurrent version: $currentVersion" -ForegroundColor Cyan
    
    # Version bump
    if ($DryRun) {
        Write-Host "`nDRY RUN MODE - No changes will be made" -ForegroundColor Yellow
        Write-Host "Would bump version: $VersionBump" -ForegroundColor Gray
        
        # Simulate version bump
        $parts = $currentVersion.Split('.')
        switch ($VersionBump) {
            "major" { $newVersion = "$([int]$parts[0] + 1).0.0" }
            "minor" { $newVersion = "$($parts[0]).$([int]$parts[1] + 1).0" }
            "patch" { $newVersion = "$($parts[0]).$($parts[1]).$([int]$parts[2] + 1)" }
        }
        Write-Host "New version would be: $newVersion" -ForegroundColor Cyan
        
        # Dry run publish
        Write-Host "`nRunning npm publish --dry-run..." -ForegroundColor Yellow
        npm publish --dry-run --access public
        
        Write-Host "`n✓ Dry run completed successfully!" -ForegroundColor Green
        Write-Host "To actually publish, run: .\publish.ps1 -VersionBump $VersionBump" -ForegroundColor Cyan
    }
    else {
        Write-Host "`nBumping version: $VersionBump" -ForegroundColor Yellow
        npm version $VersionBump --no-git-tag-version
        
        $packageJson = Get-Content "package.json" | ConvertFrom-Json
        $newVersion = $packageJson.version
        Write-Host "✓ Version bumped to: $newVersion" -ForegroundColor Green
        
        # Commit version change
        Write-Host "`nCommitting version change..." -ForegroundColor Yellow
        git add package.json package-lock.json
        git commit -m "Bump version to $newVersion"
        git tag "v$newVersion"
        Write-Host "✓ Version committed and tagged" -ForegroundColor Green
        
        # Publish
        Write-Host "`nPublishing to npm..." -ForegroundColor Yellow
        npm publish --access public
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "✓ Published successfully!" -ForegroundColor Green
            
            # Push to GitHub
            Write-Host "`nPushing to GitHub..." -ForegroundColor Yellow
            git push origin main
            git push origin "v$newVersion"
            
            if ($LASTEXITCODE -eq 0) {
                Write-Host "✓ Pushed to GitHub" -ForegroundColor Green
                
                Write-Host ""
                Write-Host "═══════════════════════════════════════" -ForegroundColor Green
                Write-Host "  ✓ Publishing Complete!" -ForegroundColor Green
                Write-Host "═══════════════════════════════════════" -ForegroundColor Green
                Write-Host ""
                Write-Host "Package: @thiagobeier/lokka@$newVersion" -ForegroundColor Cyan
                Write-Host "View on npm: https://www.npmjs.com/package/@thiagobeier/lokka" -ForegroundColor Cyan
                Write-Host ""
                Write-Host "Install with:" -ForegroundColor Yellow
                Write-Host "  npm install -g @thiagobeier/lokka" -ForegroundColor Gray
                Write-Host ""
                Write-Host "Create GitHub Release:" -ForegroundColor Yellow
                Write-Host "  https://github.com/thiagogbeier/lokka/releases/new?tag=v$newVersion" -ForegroundColor Gray
            }
            else {
                Write-Host "⚠️  Published but failed to push to GitHub" -ForegroundColor Yellow
                Write-Host "Run manually: git push origin main --tags" -ForegroundColor Gray
            }
        }
        else {
            throw "npm publish failed"
        }
    }
}
catch {
    Write-Host ""
    Write-Host "═══════════════════════════════════════" -ForegroundColor Red
    Write-Host "  ✗ Publishing Failed!" -ForegroundColor Red
    Write-Host "═══════════════════════════════════════" -ForegroundColor Red
    Write-Host ""
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}
finally {
    Pop-Location
}
