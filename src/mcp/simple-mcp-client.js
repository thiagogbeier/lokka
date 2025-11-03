#!/usr/bin/env node

// Test Contoso SharePoint site with Interactive Authentication
import { spawn } from 'child_process';

const CONFIG = {
    TENANT_ID: 'b41f1ee6-0ebd-4439-bbbc-07b635f451e0',
    CLIENT_ID: 'e4c27331-3fee-4cb0-8932-3c9bda313025',
    USE_INTERACTIVE: 'true'  // This will use interactive auth instead of client credentials
};

console.log('🧪 Testing Contoso SharePoint with Interactive Auth');
console.log('📝 This will prompt you to sign in as admin@letsintune.com');
console.log('');

async function testContosoInteractive() {
    const server = spawn('node', ['build/main.js'], {
        stdio: ['pipe', 'pipe', 'pipe'], // Capture all output
        env: { ...process.env, ...CONFIG }
    });

    // Show all server output (including auth prompts)
    server.stderr.on('data', (data) => {
        console.log('🔐 Server:', data.toString());
    });

    let requestId = 1;

    function sendRequest(method, params = {}) {
        return new Promise((resolve, reject) => {
            const request = { jsonrpc: "2.0", id: requestId++, method, params };
            console.log(`📤 Sending: ${method}`);
            server.stdin.write(JSON.stringify(request) + '\n');

            const timeout = setTimeout(() => reject(new Error('Timeout')), 60000); // Longer timeout for interactive auth
            server.stdout.once('data', (data) => {
                clearTimeout(timeout);
                try {
                    const response = JSON.parse(data.toString().trim());
                    console.log(`📥 Response: ${JSON.stringify(response, null, 2)}`);
                    resolve(response);
                } catch (error) {
                    console.log(`📥 Raw: ${data.toString()}`);
                    resolve({ raw: data.toString() });
                }
            });
        });
    }

    try {
        console.log('⏳ Starting server (this may take a moment for interactive auth)...');
        await new Promise(resolve => setTimeout(resolve, 5000));
        
        console.log('🔧 Initializing MCP connection...');
        await sendRequest('initialize', {
            protocolVersion: "2024-11-05",
            clientInfo: { name: "Contoso Interactive Test", version: "1.0.0" },
            capabilities: {}
        });

        console.log('🔨 Checking auth status...');
        await sendRequest('tools/call', {
            name: 'get-auth-status',
            arguments: {}
        });

        console.log('🏢 Creating Excel file in Contoso SharePoint site...');
        await sendRequest('tools/call', {
            name: 'create-office-file',
            arguments: {
                fileType: 'excel',
                fileName: 'ContosoTestFile_Interactive',
                location: 'sharepoint',
                sitePath: '/sites/Contoso'
            }
        });

        console.log('✅ Test completed successfully!');
    } catch (error) {
        console.error('❌ Error:', error.message);
    } finally {
        server.kill();
        process.exit(0);
    }
}

testContosoInteractive();
