# Testing the Lokka REST API Wrapper

This guide shows you how to test the REST API wrapper locally and then deploy it to Copilot Studio.

## Prerequisites

1. A valid Microsoft Graph access token
2. Node.js installed
3. The Lokka MCP project built

## Step 1: Get a Microsoft Graph Access Token

You need a valid access token to authenticate with Microsoft Graph. Here are a few ways to get one:

### Option A: Using Azure CLI

```powershell
# Login to Azure
az login

# Get an access token for Microsoft Graph
$token = az account get-access-token --resource https://graph.microsoft.com --query accessToken -o tsv
Write-Host "Token: $token"
```

### Option B: Using PowerShell with MSAL

```powershell
# Install the MSAL.PS module if you haven't already
# Install-Module MSAL.PS -Scope CurrentUser

$clientId = "your-client-id"  # Your app registration client ID
$tenantId = "your-tenant-id"   # Your tenant ID

$token = Get-MsalToken -ClientId $clientId -TenantId $tenantId -Interactive
$accessToken = $token.AccessToken
Write-Host "Token: $accessToken"
```

### Option C: Using the Azure Portal (for testing)

1. Go to https://developer.microsoft.com/en-us/graph/graph-explorer
2. Sign in with your account
3. Click "Access token" to view and copy your token

## Step 2: Start the REST API Server Locally

Open a terminal and navigate to the MCP directory:

```powershell
cd c:\temp\work\lokka\src\mcp

# Start the REST API wrapper
npm run start:rest
```

You should see output like:

```
[2025-11-04T...] [INFO] Starting MCP server from: ./build/main.js
[2025-11-04T...] [INFO] MCP client started successfully
[2025-11-04T...] [INFO] MCP client initialized successfully
[2025-11-04T...] [INFO] Lokka REST API Wrapper listening on port 3000
```

## Step 3: Test the API Endpoints

Open a new PowerShell terminal to test the endpoints:

### Test 1: Health Check (No auth required)

```powershell
Invoke-RestMethod -Uri "http://localhost:3000/health" -Method Get | ConvertTo-Json
```

Expected response:

```json
{
	"status": "ok",
	"service": "Lokka REST API Wrapper",
	"version": "1.0.0",
	"timestamp": "2025-11-04T..."
}
```

### Test 2: Set Access Token

```powershell
$token = "your-access-token-here"

$body = @{
    accessToken = $token
} | ConvertTo-Json

Invoke-RestMethod -Uri "http://localhost:3000/api/auth/token" `
    -Method Post `
    -ContentType "application/json" `
    -Body $body | ConvertTo-Json
```

### Test 3: List Groups Created in the Past Week

```powershell
$token = "your-access-token-here"

$headers = @{
    "Authorization" = "Bearer $token"
}

Invoke-RestMethod -Uri "http://localhost:3000/api/graph/groups/recent?daysBack=7" `
    -Method Get `
    -Headers $headers | ConvertTo-Json -Depth 10
```

### Test 4: List All Users

```powershell
$token = "your-access-token-here"

$headers = @{
    "Authorization" = "Bearer $token"
}

Invoke-RestMethod -Uri "http://localhost:3000/api/graph/users?`$select=id,displayName,mail&`$top=10" `
    -Method Get `
    -Headers $headers | ConvertTo-Json -Depth 10
```

### Test 5: Generic Microsoft Graph Call

```powershell
$token = "your-access-token-here"

$headers = @{
    "Authorization" = "Bearer $token"
    "Content-Type" = "application/json"
}

$body = @{
    name = "Lokka-Microsoft"
    arguments = @{
        apiType = "graph"
        path = "/users"
        method = "get"
        queryParams = @{
            "`$top" = "5"
            "`$select" = "id,displayName,mail"
        }
    }
} | ConvertTo-Json -Depth 10

Invoke-RestMethod -Uri "http://localhost:3000/api/mcp/tools/call" `
    -Method Post `
    -Headers $headers `
    -Body $body | ConvertTo-Json -Depth 10
```

## Step 4: Deploy to Azure Container Apps

Update your Dockerfile to start the REST wrapper:

```dockerfile
CMD ["node", "build/rest-wrapper.js"]
```

Set environment variables:

- `PORT`: 3000
- `MCP_SERVER_PATH`: ./build/main.js
- `AUTH_MODE`: token

## Step 5: Configure Copilot Studio

1. Import the `openapi.yaml` file
2. Update the server URL to your deployed endpoint
3. Configure OAuth 2.0 authentication
4. Create actions for your endpoints

## Common Issues

### "Graph client not initialized"

Set the access token using the `Authorization: Bearer <token>` header or `X-Access-Token` header.

### "Permission denied"

Ensure your token has the required permissions:

- `Group.Read.All` for reading groups
- `User.Read.All` for reading users
- `Directory.Read.All` for advanced queries
