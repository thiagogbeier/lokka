# Azure Container Apps Deployment Guide

## Prerequisites

1. Azure CLI installed and logged in
2. Docker installed (optional - for local testing)
3. Azure Container Registry created
4. Azure Container App created

## Quick Deployment

### Option 1: Using VS Code Tasks (Recommended)

1. Press `Ctrl+Shift+P` (or `Cmd+Shift+P` on Mac)
2. Type "Tasks: Run Task"
3. Select one of these tasks:
   - **Build TypeScript** - Build the TypeScript code
   - **Build and Start** - Build and start the REST wrapper locally
   - **Azure: Full Deployment** - Complete deployment to Azure

### Option 2: Using PowerShell Scripts

```powershell
# Build and deploy
.\deploy.ps1

# Or step by step:
.\build.ps1          # Build TypeScript
.\docker-build.ps1   # Build Docker image
.\azure-deploy.ps1   # Deploy to Azure
```

### Option 3: Manual Deployment

#### Step 1: Build TypeScript

```powershell
cd c:\temp\work\lokka\src\mcp
npm run build
```

#### Step 2: Build Docker Image Locally (Optional - for testing)

```powershell
docker build -t lokka-rest-wrapper:latest -f Dockerfile .
docker run --rm -p 3000:3000 -e AUTH_MODE=token lokka-rest-wrapper:latest
```

#### Step 3: Build and Push to Azure Container Registry

```powershell
# Login to Azure
az login

# Build and push image to ACR
az acr build `
  --registry lokkamcp `
  --image lokka-rest-wrapper:latest `
  --file Dockerfile `
  .
```

#### Step 4: Update Container App

```powershell
az containerapp update `
  --name lokka-mcp `
  --resource-group rg-lokka-mcp `
  --image lokkamcp.azurecr.io/lokka-rest-wrapper:latest `
  --set-env-vars `
    AUTH_MODE=token `
    PORT=3000 `
    MCP_SERVER_PATH=./build/main.js
```

## Environment Variables

The following environment variables are set in the container:

- `PORT=3000` - The port the REST API listens on
- `AUTH_MODE=token` - Use token-based authentication (no interactive browser)
- `MCP_SERVER_PATH=./build/main.js` - Path to the MCP server
- `NODE_ENV=production` - Node environment

## Testing the Deployment

After deployment, test the API:

```powershell
# Get your access token
$token = az account get-access-token --resource https://graph.microsoft.com --query accessToken -o tsv

# Test the health endpoint
Invoke-RestMethod -Uri "https://lokka-mcp.jollygrass-0d6fb706.canadacentral.azurecontainerapps.io/health"

# Test listing groups created in the past week
Invoke-RestMethod `
  -Uri "https://lokka-mcp.jollygrass-0d6fb706.canadacentral.azurecontainerapps.io/api/graph/groups/recent?daysBack=7" `
  -Method Get `
  -Headers @{"Authorization" = "Bearer $token"}
```

## Troubleshooting

### View Container App Logs

```powershell
az containerapp logs show `
  --name lokka-mcp `
  --resource-group rg-lokka-mcp `
  --follow
```

### Check Container App Status

```powershell
az containerapp show `
  --name lokka-mcp `
  --resource-group rg-lokka-mcp `
  --query "properties.{Fqdn:configuration.ingress.fqdn,Replicas:template.scale.minReplicas,Status:provisioningState}"
```

### Restart Container App

```powershell
az containerapp revision restart `
  --name lokka-mcp `
  --resource-group rg-lokka-mcp
```

## Continuous Deployment

For automated deployment, see the GitHub Actions workflow in `.github/workflows/deploy.yml` (if configured).

## Rollback

To rollback to a previous version:

```powershell
# List all revisions
az containerapp revision list `
  --name lokka-mcp `
  --resource-group rg-lokka-mcp `
  --query "[].{Name:name,Active:properties.active,Created:properties.createdTime}" `
  -o table

# Activate a specific revision
az containerapp revision activate `
  --name lokka-mcp `
  --resource-group rg-lokka-mcp `
  --revision <revision-name>
```

## Cost Optimization

- Container Apps scale to zero when not in use
- Minimum replicas: 0 (for development)
- Maximum replicas: 10 (adjust based on load)
- Auto-scale rules based on HTTP requests

## Security Notes

1. **Never commit access tokens** - Always use environment variables
2. **Use Managed Identity** - Configure Azure Managed Identity for production
3. **Enable authentication** - Use Azure AD authentication at the Container App level
4. **HTTPS only** - Container Apps enforce HTTPS by default
5. **Secrets management** - Use Azure Key Vault for sensitive data
