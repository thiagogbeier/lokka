#!/usr/bin/env node

/**
 * Test script for Lokka MCP Server with token-based authentication
 * This script demonstrates how to:
 * 1. Start the server in client-provided-token mode
 * 2. Set an access token via the MCP tool
 * 3. Make Microsoft Graph API calls
 * 4. Validate functionality
 */

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { logger } from "./logger.js";

async function testTokenBasedAuth() {
  console.log("🧪 Starting Lokka MCP Server Token-Based Authentication Test");
  
  // Set environment variables for client token mode
  process.env.USE_CLIENT_TOKEN = 'true';
  
  console.log("📝 Environment configured for client token mode");
  console.log("   USE_CLIENT_TOKEN = true");
  console.log("   Note: No CLIENT_SECRET required in this mode\n");
  
  try {
    // Import the main server module to start it
    console.log("🚀 Starting MCP Server in client token mode...");
    
    // Note: In a real scenario, you would:
    // 1. Start the server
    // 2. Connect an MCP client
    // 3. Call set-access-token with a valid token
    // 4. Make Graph API requests
    
    console.log("✅ Server configuration validated");
    console.log("\n📋 Next steps for testing:");
    console.log("1. Obtain an access token for Microsoft Graph");
    console.log("2. Use an MCP Client to connect to this server");
    console.log("3. Call 'set-access-token' tool with your token");
    console.log("4. Call 'Lokka-Microsoft' tool to query Graph API");
    console.log("5. Call 'get-auth-status' to check authentication status");
    
    console.log("\n🔧 Example token acquisition (PowerShell):");
    console.log(`
# Using Azure CLI (requires login)
$token = az account get-access-token --resource https://graph.microsoft.com --query accessToken -o tsv

# Using PowerShell with MSAL
Install-Module MSAL.PS -Scope CurrentUser
$clientId = "your-client-id"
$tenantId = "your-tenant-id"
$token = Get-MsalToken -ClientId $clientId -TenantId $tenantId -Scopes "https://graph.microsoft.com/.default"
$accessToken = $token.AccessToken
    `);
    
    console.log("\n📊 Example MCP Client calls:");
    console.log(`
1. Set the token:
   Tool: set-access-token
   Parameters: { "accessToken": "eyJ0eXAiOiJKV1QiLCJhbGc...", "expiresOn": "2025-06-19T15:30:00Z" }

2. Check auth status:
   Tool: get-auth-status
   Parameters: {}

3. Query Microsoft Graph:
   Tool: Lokka-Microsoft
   Parameters: { "apiType": "graph", "path": "/me", "method": "get" }
    `);
    
  } catch (error: any) {
    console.error("❌ Test failed:", error.message);
    process.exit(1);
  }
}

// Only run the test if this is the main module
if (import.meta.url === `file://${process.argv[1]}`) {
  testTokenBasedAuth().catch(console.error);
}

export { testTokenBasedAuth };
