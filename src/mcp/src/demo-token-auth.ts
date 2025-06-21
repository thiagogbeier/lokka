#!/usr/bin/env node

/**
 * Interactive demo for Lokka MCP Server with token-based authentication
 * This script will:
 * 1. Help obtain an access token via device code flow
 * 2. Test the token with the MCP server
 * 3. Demonstrate Graph API functionality
 */

import { DeviceCodeCredential } from "@azure/identity";
import { readFileSync } from 'fs';
import { logger } from "./logger.js";

interface TokenInfo {
  accessToken: string;
  expiresOn: Date;
}

async function getAccessTokenInteractively(tenantId?: string, clientId?: string): Promise<TokenInfo> {
  // Use default test app registration if not provided
  const defaultClientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e"; // Microsoft Graph Command Line Tools
  const defaultTenantId = "common";
  
  const actualClientId = clientId || defaultClientId;
  const actualTenantId = tenantId || defaultTenantId;
  
  console.log(`🔐 Obtaining access token for Microsoft Graph`);
  console.log(`   Client ID: ${actualClientId}`);
  console.log(`   Tenant ID: ${actualTenantId}`);
  
  const credential = new DeviceCodeCredential({
    tenantId: actualTenantId,
    clientId: actualClientId,
    userPromptCallback: (info) => {
      console.log(`\n📱 Please authenticate:`);
      console.log(`   Visit: ${info.verificationUri}`);
      console.log(`   Code: ${info.userCode}\n`);
      return Promise.resolve();
    },
  });

  const token = await credential.getToken("https://graph.microsoft.com/.default");
  if (!token) {
    throw new Error("Failed to acquire access token");
  }

  return {
    accessToken: token.token,
    expiresOn: new Date(token.expiresOnTimestamp)
  };
}

async function testGraphApiCall(accessToken: string): Promise<any> {
  console.log("🌐 Testing Microsoft Graph API call...");
  
  const response = await fetch("https://graph.microsoft.com/v1.0/me", {
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Content-Type': 'application/json'
    }
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Graph API call failed: ${response.status} ${errorText}`);
  }

  return await response.json();
}

async function demonstrateTokenAuth() {
  console.log("🧪 Lokka MCP Server - Token-Based Authentication Demo");
  console.log("=".repeat(60));
  
  try {
    // Step 1: Get access token
    console.log("\n📋 Step 1: Obtain Access Token");
    const tokenInfo = await getAccessTokenInteractively();
    
    console.log("✅ Token acquired successfully");
    console.log(`   Expires: ${tokenInfo.expiresOn.toISOString()}`);
    console.log(`   Token preview: ${tokenInfo.accessToken.substring(0, 50)}...`);
    
    // Step 2: Test the token directly
    console.log("\n📋 Step 2: Test Token with Microsoft Graph");
    const userInfo = await testGraphApiCall(tokenInfo.accessToken);
    console.log("✅ Direct Graph API call successful");
    console.log(`   User: ${userInfo.displayName} (${userInfo.userPrincipalName})`);
    
    // Step 3: Show how to use with MCP Server
    console.log("\n📋 Step 3: MCP Server Integration");
    console.log("✅ Token is ready for use with Lokka MCP Server");
    
    console.log("\n🔧 Environment setup for client token mode:");
    console.log(`export USE_CLIENT_TOKEN=true`);
    console.log(`# Note: No CLIENT_SECRET needed`);
    
    console.log("\n📊 Example MCP tool calls:");
    console.log(`
1. Set access token:
   {
     "tool": "set-access-token",
     "arguments": {
       "accessToken": "${tokenInfo.accessToken.substring(0, 50)}...",
       "expiresOn": "${tokenInfo.expiresOn.toISOString()}"
     }
   }

2. Query user profile:
   {
     "tool": "Lokka-Microsoft", 
     "arguments": {
       "apiType": "graph",
       "path": "/me",
       "method": "get"
     }
   }

3. List users (with pagination):
   {
     "tool": "Lokka-Microsoft",
     "arguments": {
       "apiType": "graph", 
       "path": "/users",
       "method": "get",
       "fetchAll": true,
       "consistencyLevel": "eventual"
     }
   }
    `);
    
    console.log("\n🎉 Demo completed successfully!");
    console.log("The token can now be used with any MCP Client that supports Lokka.");
    
  } catch (error: any) {
    console.error("❌ Demo failed:", error.message);
    process.exit(1);
  }
}

// Run the demo if this is the main module
if (import.meta.url === `file://${process.argv[1]}`) {
  demonstrateTokenAuth().catch(console.error);
}

export { demonstrateTokenAuth, getAccessTokenInteractively };
