#!/usr/bin/env node
/**
 * Live MCP Server Test - Test real Microsoft Graph requests
 * This script demonstrates how to start the MCP server and make live Graph API calls
 */
import { spawn } from 'child_process';
// Simple MCP client simulator for testing
class SimpleMCPClient {
    serverProcess;
    messageId;
    constructor(serverProcess) {
        this.serverProcess = serverProcess;
        this.messageId = 1;
    }
    async sendMessage(method, params = {}) {
        const message = {
            jsonrpc: "2.0",
            id: this.messageId++,
            method: method,
            params: params
        };
        console.log(`📤 Sending: ${JSON.stringify(message, null, 2)}`);
        return new Promise((resolve, reject) => {
            const timeout = setTimeout(() => {
                reject(new Error('Request timeout'));
            }, 30000);
            const onData = (data) => {
                try {
                    const response = JSON.parse(data.toString());
                    clearTimeout(timeout);
                    this.serverProcess.stdout?.off('data', onData);
                    console.log(`📥 Received: ${JSON.stringify(response, null, 2)}`);
                    resolve(response);
                }
                catch (error) {
                    // Ignore parse errors, might be partial data
                }
            };
            this.serverProcess.stdout?.on('data', onData);
            this.serverProcess.stdin?.write(JSON.stringify(message) + '\n');
        });
    }
    async callTool(toolName, arguments_) {
        return this.sendMessage('tools/call', {
            name: toolName,
            arguments: arguments_
        });
    }
    async listTools() {
        return this.sendMessage('tools/list');
    }
    async initialize() {
        return this.sendMessage('initialize', {
            protocolVersion: "2024-11-05",
            capabilities: {
                tools: {}
            },
            clientInfo: {
                name: "simple-test-client",
                version: "1.0.0"
            }
        });
    }
}
async function testLiveGraphRequest() {
    console.log("🚀 Starting Live MCP Server Test");
    console.log("=".repeat(50));
    // Check if we have an access token
    const accessToken = process.env.ACCESS_TOKEN;
    if (!accessToken) {
        console.log("❌ No ACCESS_TOKEN environment variable found");
        console.log("\n📋 To get an access token, you can:");
        console.log("1. Use Azure CLI: az account get-access-token --resource https://graph.microsoft.com --query accessToken -o tsv");
        console.log("2. Use the demo script: npm run demo:token");
        console.log("3. Get one from https://developer.microsoft.com/en-us/graph/graph-explorer");
        console.log("\nThen set it: $env:ACCESS_TOKEN = \"your-token-here\"");
        return;
    }
    console.log("✅ Access token found");
    console.log(`   Token preview: ${accessToken.substring(0, 50)}...`);
    // Start the MCP server in client token mode
    console.log("\n🔧 Starting MCP Server in client token mode...");
    const serverProcess = spawn('node', ['build/main.js'], {
        env: {
            ...process.env,
            USE_CLIENT_TOKEN: 'true'
        },
        stdio: ['pipe', 'pipe', 'pipe']
    });
    serverProcess.stderr.on('data', (data) => {
        console.log(`🔍 Server log: ${data.toString().trim()}`);
    });
    // Wait a moment for server to start
    await new Promise(resolve => setTimeout(resolve, 2000));
    try {
        const client = new SimpleMCPClient(serverProcess);
        console.log("\n📡 Step 1: Initialize MCP connection");
        await client.initialize();
        console.log("\n📡 Step 2: List available tools");
        const toolsResponse = await client.listTools();
        console.log(`✅ Found ${toolsResponse.result?.tools?.length || 0} tools`);
        console.log("\n📡 Step 3: Set access token");
        const tokenResponse = await client.callTool('set-access-token', {
            accessToken: accessToken
        });
        if (tokenResponse.result?.isError) {
            throw new Error(`Failed to set token: ${tokenResponse.result?.content?.[0]?.text}`);
        }
        console.log("✅ Access token set successfully");
        console.log("\n📡 Step 4: Check auth status");
        const statusResponse = await client.callTool('get-auth-status', {});
        console.log("✅ Auth status checked");
        console.log("\n📡 Step 5: Make live Microsoft Graph request (/me)");
        const graphResponse = await client.callTool('Lokka-Microsoft', {
            apiType: 'graph',
            path: '/me',
            method: 'get'
        });
        if (graphResponse.result?.isError) {
            console.log("❌ Graph request failed:");
            console.log(graphResponse.result?.content?.[0]?.text);
        }
        else {
            console.log("✅ Graph request successful!");
            const responseText = graphResponse.result?.content?.[0]?.text;
            if (responseText) {
                try {
                    const userData = JSON.parse(responseText.split('\n\n')[1]);
                    console.log(`👤 User: ${userData.displayName} (${userData.userPrincipalName})`);
                    console.log(`📧 Email: ${userData.mail || 'N/A'}`);
                    console.log(`🏢 Company: ${userData.companyName || 'N/A'}`);
                }
                catch (e) {
                    console.log("Raw response:", responseText);
                }
            }
        }
        console.log("\n📡 Step 6: Test additional Graph endpoints");
        // Test groups
        console.log("\n📡 Testing /me/memberOf (groups)");
        const groupsResponse = await client.callTool('Lokka-Microsoft', {
            apiType: 'graph',
            path: '/me/memberOf',
            method: 'get',
            queryParams: {
                '$select': 'displayName,id,groupTypes'
            }
        });
        if (!groupsResponse.result?.isError) {
            console.log("✅ Groups request successful");
        }
        console.log("\n🎉 Live test completed successfully!");
    }
    catch (error) {
        console.error("❌ Test failed:", error?.message || error);
    }
    finally {
        console.log("\n🔧 Stopping MCP Server...");
        serverProcess.kill();
    }
}
// Run the test
testLiveGraphRequest().catch(console.error);
