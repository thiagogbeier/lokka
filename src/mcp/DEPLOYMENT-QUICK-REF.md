# Quick Deployment Reference

## 🚀 Quick Commands

### Build and Run Locally

```powershell
# Build TypeScript
npm run build

# Start REST API locally
npm run start:rest
```

### Deploy to Azure (All Methods)

#### Method 1: VS Code Tasks (EASIEST)

1. Press `Ctrl+Shift+B` to build
2. Press `Ctrl+Shift+P` → "Tasks: Run Task" → "Azure: Full Deployment"

#### Method 2: NPM Scripts

```powershell
npm run deploy
```

#### Method 3: PowerShell Scripts

```powershell
.\deploy.ps1
```

#### Method 4: Step by Step

```powershell
.\build.ps1                    # Build TypeScript
.\docker-build.ps1             # Build Docker image locally (optional)
.\azure-deploy.ps1             # Deploy to Azure
```

#### Method 5: Manual Azure CLI

```powershell
# Build
npm run build

# Deploy to ACR and Container App
az acr build --registry lokkamcp --image lokka-rest-wrapper:latest --file Dockerfile .

az containerapp update `
  --name lokka-mcp `
  --resource-group rg-lokka-mcp `
  --image lokkamcp.azurecr.io/lokka-rest-wrapper:latest `
  --set-env-vars AUTH_MODE=token PORT=3000 MCP_SERVER_PATH=./build/main.js
```

## 📋 Prerequisites Checklist

- [ ] Azure CLI installed (`az --version`)
- [ ] Logged in to Azure (`az login`)
- [ ] Node.js 20+ installed
- [ ] TypeScript compiled (`npm run build`)
- [ ] Docker installed (optional - only for local testing)

## 🔧 Configuration

### Your Azure Resources

- **Resource Group**: `rg-lokka-mcp`
- **Container App**: `lokka-mcp`
- **Container Registry**: `lokkamcp`
- **URL**: `https://lokka-mcp.jollygrass-0d6fb706.canadacentral.azurecontainerapps.io`

### Environment Variables (Set automatically)

```
PORT=3000
AUTH_MODE=token
MCP_SERVER_PATH=./build/main.js
NODE_ENV=production
```

## 🧪 Testing After Deployment

```powershell
# Test health endpoint
Invoke-RestMethod -Uri "https://lokka-mcp.jollygrass-0d6fb706.canadacentral.azurecontainerapps.io/health"

# Test with your tenant
$token = az account get-access-token --resource https://graph.microsoft.com --query accessToken -o tsv

Invoke-RestMethod `
  -Uri "https://lokka-mcp.jollygrass-0d6fb706.canadacentral.azurecontainerapps.io/api/graph/groups/recent?daysBack=7" `
  -Method Get `
  -Headers @{"Authorization" = "Bearer $token"}
```

## 📊 Monitoring

```powershell
# View logs
az containerapp logs show --name lokka-mcp --resource-group rg-lokka-mcp --follow

# Check status
az containerapp show --name lokka-mcp --resource-group rg-lokka-mcp
```

## 🐛 Troubleshooting

### Build fails

```powershell
# Clean and rebuild
Remove-Item -Recurse -Force build, node_modules
npm install
npm run build
```

### Deployment fails

```powershell
# Check Azure login
az account show

# Re-login if needed
az login

# Try deployment again
npm run deploy
```

### Container App not responding

```powershell
# Restart
az containerapp revision restart --name lokka-mcp --resource-group rg-lokka-mcp

# Check logs for errors
az containerapp logs show --name lokka-mcp --resource-group rg-lokka-mcp --tail 50
```

## 📁 Files Created

- `Dockerfile` - Container image definition
- `.dockerignore` - Files to exclude from Docker build
- `.vscode/tasks.json` - VS Code build/deploy tasks
- `build.ps1` - Build TypeScript
- `docker-build.ps1` - Build Docker image locally
- `azure-deploy.ps1` - Deploy to Azure
- `deploy.ps1` - Complete deployment
- `AZURE-DEPLOYMENT.md` - Full deployment guide
- `DEPLOYMENT-QUICK-REF.md` - This file

## 🎯 Common Workflows

### Daily Development

```powershell
npm run build
npm run start:rest
# Test locally, then when ready:
npm run deploy
```

### Production Deployment

```powershell
# Option 1: Quick
npm run deploy

# Option 2: With custom tag
.\azure-deploy.ps1 -ImageTag "v1.0.0"
```

### Local Docker Testing

```powershell
npm run docker:build
docker run --rm -p 3000:3000 -e AUTH_MODE=token lokka-rest-wrapper:latest
```
