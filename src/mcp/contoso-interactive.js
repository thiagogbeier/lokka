#!/usr/bin/env node

// Direct Interactive Auth Test for Contoso
import { spawn } from 'child_process';

const CONFIG = {
    TENANT_ID: 'b41f1ee6-0ebd-4439-bbbc-07b635f451e0',
    CLIENT_ID: '04b07795-8ddb-461a-bbee-02f9e1bf7b46', // Azure CLI - supports device code
    USE_INTERACTIVE: 'true'
};

console.log('🔐 Starting Interactive Authentication for Contoso...');
console.log('⏳ This will show you a device code to authenticate with');
console.log('');

const server = spawn('node', ['build/main.js'], {
    stdio: ['pipe', 'inherit', 'inherit'], // Let server output show directly
    env: { ...process.env, ...CONFIG }
});

// Give server time to start and show auth prompt
setTimeout(() => {
    console.log('');
    console.log('📋 Once you see the device code above:');
    console.log('1. Visit https://microsoft.com/devicelogin in your browser');
    console.log('2. Enter the code shown above');
    console.log('3. Sign in as admin@letsintune.com');
    console.log('4. Press Enter here to continue the test...');
    console.log('');
    
    process.stdin.once('data', async () => {
        console.log('🧪 Testing Contoso file creation...');
        
        let requestId = 1;
        
        function sendRequest(method, params = {}) {
            return new Promise((resolve) => {
                const request = { jsonrpc: "2.0", id: requestId++, method, params };
                server.stdin.write(JSON.stringify(request) + '\n');
                
                if (server.stdout) {
                    server.stdout.once('data', (data) => {
                        try {
                            const response = JSON.parse(data.toString().trim());
                            resolve(response);
                        } catch (error) {
                            resolve({ raw: data.toString() });
                        }
                    });
                } else {
                    resolve({ error: 'Server not available' });
                }
            });
        }
        
        try {
            // Initialize
            await sendRequest('initialize', {
                protocolVersion: "2024-11-05",
                clientInfo: { name: "Contoso Test", version: "1.0.0" },
                capabilities: {}
            });
            
            // Create file in Contoso
            const result = await sendRequest('tools/call', {
                name: 'create-office-file',
                arguments: {
                    fileType: 'excel',
                    fileName: 'ContosoSuccess',
                    location: 'sharepoint',
                    sitePath: '/sites/Contoso'
                }
            });
            
            console.log('📊 Result:', JSON.stringify(result, null, 2));
            
        } catch (error) {
            console.error('❌ Error:', error);
        } finally {
            server.kill();
            process.exit(0);
        }
    });
}, 3000);

// Handle server exit
server.on('exit', (code) => {
    if (code !== 0) {
        console.log(`Server exited with code ${code}`);
        process.exit(code);
    }
});
