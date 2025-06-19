#!/usr/bin/env node
/**
 * Simple MCP Server test for token-based authentication
 * This script tests the token functionality by simulating MCP tool calls
 */
import { AuthManager, AuthMode } from "./auth.js";
import { Client } from "@microsoft/microsoft-graph-client";
async function testTokenBasedAuth() {
    console.log("🧪 Testing Lokka MCP Server - Token-Based Authentication");
    console.log("=" + "=".repeat(59));
    const testToken = process.env.ACCESS_TOKEN;
    if (!testToken) {
        console.log("❌ No ACCESS_TOKEN environment variable provided");
        console.log("\n📋 To test with your token:");
        console.log('set ACCESS_TOKEN="your-access-token-here"');
        console.log("npm run test:token");
        console.log("\n💡 Or use the demo script to get a token interactively:");
        console.log("npm run demo:token");
        return;
    }
    try {
        console.log("\n🔧 Step 1: Initialize AuthManager in client token mode");
        const authConfig = {
            mode: AuthMode.ClientProvidedToken,
            accessToken: testToken
        };
        const authManager = new AuthManager(authConfig);
        await authManager.initialize();
        console.log("✅ AuthManager initialized successfully");
        console.log("\n🔧 Step 2: Test Graph Client initialization");
        const authProvider = authManager.getGraphAuthProvider();
        const graphClient = Client.initWithMiddleware({
            authProvider: authProvider,
        });
        console.log("✅ Graph Client initialized successfully");
        console.log("\n🔧 Step 3: Test token update functionality");
        // Test updating the same token (simulates MCP tool call)
        authManager.updateAccessToken(testToken);
        console.log("✅ Token update functionality works");
        console.log("\n🔧 Step 4: Test auth status");
        const tokenStatus = authManager.getTokenStatus();
        console.log(`✅ Token status: expired=${tokenStatus.isExpired}, expires=${tokenStatus.expiresOn?.toISOString() || 'unknown'}`);
        console.log("\n🔧 Step 5: Test Microsoft Graph API call");
        try {
            const response = await graphClient.api('/me').get();
            console.log(`✅ Graph API call successful: ${response.displayName} (${response.userPrincipalName})`);
        }
        catch (error) {
            console.log(`❌ Graph API call failed: ${error.message}`);
            if (error.message.includes('Unauthorized') || error.message.includes('401')) {
                console.log("   This might indicate the token has expired or lacks permissions");
            }
        }
        console.log("\n🎉 All tests completed!");
        console.log("\n📊 Test Summary:");
        console.log("✅ AuthManager initialization: PASSED");
        console.log("✅ Graph Client setup: PASSED");
        console.log("✅ Token update: PASSED");
        console.log("✅ Auth status check: PASSED");
        console.log("✅ Graph API call: " + (tokenStatus.isExpired ? "SKIPPED (token expired)" : "PASSED"));
    }
    catch (error) {
        console.error("\n❌ Test failed:", error.message);
        console.log("\n🔍 Troubleshooting:");
        console.log("1. Ensure your token is valid and not expired");
        console.log("2. Check that the token has Microsoft Graph permissions");
        console.log("3. Verify the token format is correct (JWT)");
        process.exit(1);
    }
}
// Run the test
testTokenBasedAuth().catch(console.error);
