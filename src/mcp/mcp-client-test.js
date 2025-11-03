#!/usr/bin/env node
const { spawn } = require('child_process');
const readline = require('readline');

class SimpleMCPClient {
    constructor() {
        this.serverProcess = null;
        this.requestId = 1;
        this.pendingRequests = new Map();
    }

    async startServer() {
        console.log('🚀 Starting Lokka MCP Server...');
        
        // Start the MCP server as a child process
        this.serverProcess = spawn('node', ['build/main_backup.js'], {
            cwd: process.cwd(),
            stdio: ['pipe', 'pipe', 'pipe'],
            env: {
                ...process.env,
                TENANT_ID: 'your-tenant-id',
                CLIENT_ID: 'your-client-id', 
                CLIENT_SECRET: 'your-client-secret'
            }
        });

        // Handle server output
        this.serverProcess.stdout.on('data', (data) => {
            try {
                const response = JSON.parse(data.toString());
                this.handleResponse(response);
            } catch (error) {
                console.log('Server output:', data.toString());
            }
        });

        this.serverProcess.stderr.on('data', (data) => {
            console.log('Server log:', data.toString());
        });

        this.serverProcess.on('close', (code) => {
            console.log(`Server process exited with code ${code}`);
        });

        // Wait a moment for server to start
        await new Promise(resolve => setTimeout(resolve, 2000));
        
        // Initialize the MCP connection
        await this.initialize();
    }

    async initialize() {
        console.log('🔧 Initializing MCP connection...');
        
        const initRequest = {
            jsonrpc: "2.0",
            id: this.requestId++,
            method: "initialize",
            params: {
                protocolVersion: "2024-11-05",
                clientInfo: {
                    name: "Simple MCP Client",
                    version: "1.0.0"
                },
                capabilities: {}
            }
        };

        await this.sendRequest(initRequest);
    }

    async listTools() {
        console.log('📋 Getting available tools...');
        
        const toolsRequest = {
            jsonrpc: "2.0",
            id: this.requestId++,
            method: "tools/list",
            params: {}
        };

        return await this.sendRequest(toolsRequest);
    }

    async callTool(toolName, params = {}) {
        console.log(`🔨 Calling tool: ${toolName}`);
        console.log('Parameters:', JSON.stringify(params, null, 2));
        
        const toolRequest = {
            jsonrpc: "2.0",
            id: this.requestId++,
            method: "tools/call",
            params: {
                name: toolName,
                arguments: params
            }
        };

        return await this.sendRequest(toolRequest);
    }

    sendRequest(request) {
        return new Promise((resolve, reject) => {
            const id = request.id;
            this.pendingRequests.set(id, { resolve, reject });
            
            const message = JSON.stringify(request) + '\n';
            this.serverProcess.stdin.write(message);
            
            // Set timeout
            setTimeout(() => {
                if (this.pendingRequests.has(id)) {
                    this.pendingRequests.delete(id);
                    reject(new Error('Request timeout'));
                }
            }, 30000); // 30 second timeout
        });
    }

    handleResponse(response) {
        console.log('📨 Received response:', JSON.stringify(response, null, 2));
        
        if (response.id && this.pendingRequests.has(response.id)) {
            const { resolve } = this.pendingRequests.get(response.id);
            this.pendingRequests.delete(response.id);
            resolve(response);
        }
    }

    async cleanup() {
        if (this.serverProcess) {
            this.serverProcess.kill();
        }
    }
}

// Main test function
async function runTests() {
    const client = new SimpleMCPClient();
    
    try {
        // Start server
        await client.startServer();
        
        // List available tools
        const toolsResponse = await client.listTools();
        console.log('✅ Available tools:', toolsResponse.result?.tools?.map(t => t.name) || []);
        
        // Test get-auth-status
        console.log('\n=== Testing get-auth-status ===');
        const authStatus = await client.callTool('get-auth-status');
        console.log('Auth Status Result:', authStatus.result);
        
        // Test create-office-file
        console.log('\n=== Testing create-office-file ===');
        const fileResult = await client.callTool('create-office-file', {
            fileType: 'word',
            fileName: 'MyTestDocument',
            location: 'onedrive'
        });
        console.log('File Creation Result:', fileResult.result);
        
    } catch (error) {
        console.error('❌ Test failed:', error.message);
    } finally {
        await client.cleanup();
        process.exit(0);
    }
}

// Handle Ctrl+C
process.on('SIGINT', async () => {
    console.log('\n🛑 Shutting down...');
    process.exit(0);
});

// Run the tests
console.log('🧪 Starting MCP Client Test Suite...');
runTests();
