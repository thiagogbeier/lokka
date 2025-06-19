# How to Start Lokka MCP Server Locally and Test Microsoft Graph

This guide shows you how to start the Lokka MCP Server locally and test it with real Microsoft Graph API requests.

## Prerequisites

1. **Node.js** installed (v16 or later)
2. **Valid Microsoft Graph access token** (see below for how to get one)
3. **Build the project**: `npm run build`

## Getting an Access Token

### Option 1: Azure CLI (Easiest)
```bash
# Login to Azure CLI
az login

# Get a token for Microsoft Graph
az account get-access-token --resource https://graph.microsoft.com --query accessToken -o tsv
```

### Option 2: Graph Explorer (Quick Testing)
1. Go to https://developer.microsoft.com/en-us/graph/graph-explorer
2. Sign in with your Microsoft account
3. Open browser developer tools (F12)
4. Go to Network tab
5. Make any Graph request (like GET /me)
6. Find the request in Network tab
7. Copy the Authorization header value (remove "Bearer " prefix)

### Option 3: Interactive Demo (Built-in)
```bash
npm run demo:token
```

## Method 1: Manual Testing (Simplest)

### Step 1: Set your access token
```powershell
# PowerShell
$env:ACCESS_TOKEN = "your-access-token-here"

# Command Prompt
set ACCESS_TOKEN=your-access-token-here
```

### Step 2: Start the MCP Server in client token mode
```bash
# In one terminal window
$env:USE_CLIENT_TOKEN = "true"
npm start
```

### Step 3: Test with MCP Client
In another terminal, you can now send MCP requests to the server using stdin/stdout.

## Method 2: Automated Live Test (Recommended)

### Step 1: Set your access token
```powershell
$env:ACCESS_TOKEN = "your-access-token-here"
```

### Step 2: Run the live test
```bash
npm run test:live
```

This will:
1. ✅ Start the MCP Server automatically
2. ✅ Initialize the MCP protocol
3. ✅ Set your access token
4. ✅ Test authentication status
5. ✅ Make a real `/me` request to Microsoft Graph
6. ✅ Test additional endpoints like `/me/memberOf`
7. ✅ Clean up and stop the server

## Method 3: Manual MCP Protocol Testing

### Step 1: Start the server
```bash
$env:USE_CLIENT_TOKEN = "true"
node build/main.js
```

### Step 2: Send MCP messages via stdin

**Initialize the connection:**
```json
{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{"tools":{}},"clientInfo":{"name":"test-client","version":"1.0.0"}}}
```

**Set access token:**
```json
{"jsonrpc":"2.0","id":2,"method":"tools/call","params":{"name":"set-access-token","arguments":{"accessToken":"your-token-here"}}}
```

**Get current user profile:**
```json
{"jsonrpc":"2.0","id":3,"method":"tools/call","params":{"name":"Lokka-Microsoft","arguments":{"apiType":"graph","path":"/me","method":"get"}}}
```

## Testing Different Scenarios

### Test with User Context (Delegated Permissions)
Use a token obtained through interactive login or device code flow:
```bash
# This should work for /me endpoint
npm run test:live
```

### Test with App Context (Application Permissions)
Use a token from client credentials flow:
```bash
# This will show error for /me but work for other endpoints like /users
npm run test:live
```

## Common Endpoints to Test

### User Profile Endpoints (Require Delegated Permissions)
- `/me` - Current user profile
- `/me/memberOf` - User's group memberships
- `/me/ownedObjects` - Objects owned by user

### Directory Endpoints (Work with App Permissions)
- `/users` - List all users
- `/groups` - List all groups
- `/applications` - List applications

### Example Test Commands

```bash
# Test current user
{"jsonrpc":"2.0","id":1,"method":"tools/call","params":{"name":"Lokka-Microsoft","arguments":{"apiType":"graph","path":"/me","method":"get"}}}

# Test list users (with pagination)
{"jsonrpc":"2.0","id":2,"method":"tools/call","params":{"name":"Lokka-Microsoft","arguments":{"apiType":"graph","path":"/users","method":"get","fetchAll":true,"consistencyLevel":"eventual"}}}

# Test create group
{"jsonrpc":"2.0","id":3,"method":"tools/call","params":{"name":"Lokka-Microsoft","arguments":{"apiType":"graph","path":"/groups","method":"post","body":{"displayName":"Test Group","mailEnabled":false,"mailNickname":"testgroup","securityEnabled":true}}}}
```

## Troubleshooting

### ❌ "/me request is only valid with delegated authentication flow"
- Your token has application permissions, not user permissions
- Get a token through interactive login instead of client credentials

### ❌ "Token has expired"
- Get a fresh token (Graph tokens typically expire in 1 hour)
- Use `npm run demo:token` to get a new one

### ❌ "Failed to acquire access token"
- Check that your token is valid JWT format
- Ensure token has proper Microsoft Graph audience

### ❌ "Insufficient privileges"
- Your token doesn't have the required permissions for the endpoint
- Check your app registration or user permissions

## Success Output Example

When everything works correctly, you'll see:
```
✅ Access token set successfully
👤 User: John Doe (john.doe@company.com)
📧 Email: john.doe@company.com
🏢 Company: Contoso Ltd
✅ Groups request successful
🎉 Live test completed successfully!
```

This confirms that:
- ✅ Token authentication is working
- ✅ MCP Server is properly configured
- ✅ Microsoft Graph API calls are successful
- ✅ You can now use the server with any MCP Client!
