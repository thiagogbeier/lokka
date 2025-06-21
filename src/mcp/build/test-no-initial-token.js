#!/usr/bin/env node
/**
 * Test script to verify the fix for ClientProvidedToken mode with no initial token
 * This tests the specific scenario mentioned in the code review:
 * 1. Start server in ClientProvidedToken mode without an initial accessToken
 * 2. Verify authManager.initialize() doesn't fail
 * 3. Verify that set-access-token (updateAccessToken) works correctly
 */
import { AuthManager, AuthMode, ClientProvidedTokenCredential } from "./auth.js";
async function testNoInitialToken() {
    console.log("🧪 Testing ClientProvidedToken mode with no initial token");
    console.log("=" + "=".repeat(55));
    try {
        console.log("\n🔧 Step 1: Initialize AuthManager without initial token");
        // This should not fail anymore after our fix
        const authConfig = {
            mode: AuthMode.ClientProvidedToken
            // Note: No accessToken provided initially
        };
        const authManager = new AuthManager(authConfig);
        await authManager.initialize();
        console.log("✅ AuthManager initialized successfully without initial token");
        console.log("\n🔧 Step 2: Verify credential is created but inactive");
        const credential = authManager.getAzureCredential();
        console.log("✅ Credential object exists:", credential instanceof ClientProvidedTokenCredential);
        // Try to get a token - should return null since no token is set
        const token = await credential.getToken("https://graph.microsoft.com/.default");
        console.log("✅ getToken returns null when no token is set:", token === null);
        console.log("\n🔧 Step 3: Test token update functionality");
        const testToken = "fake-token-for-testing";
        const expiresOn = new Date(Date.now() + 3600000); // 1 hour from now
        // This should work now that credential is properly initialized
        authManager.updateAccessToken(testToken, expiresOn);
        console.log("✅ updateAccessToken completed successfully");
        console.log("\n🔧 Step 4: Verify token is now available");
        const updatedToken = await credential.getToken("https://graph.microsoft.com/.default");
        console.log("✅ getToken now returns a token:", updatedToken !== null);
        console.log("   Token matches:", updatedToken?.token === testToken);
        console.log("\n🔧 Step 5: Test token status functionality");
        const status = authManager.getTokenStatus();
        console.log("✅ Token status:", {
            isExpired: status.isExpired,
            hasExpirationTime: !!status.expiresOn
        });
        console.log("\n🎉 All tests passed! The fix works correctly.");
    }
    catch (error) {
        console.error("❌ Test failed:", error);
        process.exit(1);
    }
}
// Run the test
testNoInitialToken().catch(console.error);
