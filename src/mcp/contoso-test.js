#!/usr/bin/env node

// Simple Contoso Site Test
import { spawn } from 'child_process';

const CONFIG = {
    TENANT_ID: 'b41f1ee6-0ebd-4439-bbbc-07b635f451e0',
    CLIENT_ID: '04b07795-8ddb-461a-bbee-02f9e1bf7b46', // Azure CLI - supports device code
    USE_INTERACTIVE: 'true'
};

console.log('Testing Contoso SharePoint site file creation...');

async function testContosoSite() {
    const server = spawn('node', ['build/main.js'], {
        stdio: ['pipe', 'pipe', 'inherit'],
        env: { ...process.env, ...CONFIG }
    });

    let requestId = 1;

    function sendRequest(method, params = {}) {
        return new Promise((resolve, reject) => {
            const request = { jsonrpc: "2.0", id: requestId++, method, params };
            console.log(`Sending: ${method}`);
            server.stdin.write(JSON.stringify(request) + '\n');

            const timeout = setTimeout(() => reject(new Error('Timeout')), 30000);
            server.stdout.once('data', (data) => {
                clearTimeout(timeout);
                try {
                    const response = JSON.parse(data.toString().trim());
                    console.log(`Response: ${JSON.stringify(response, null, 2)}`);
                    resolve(response);
                } catch (error) {
                    console.log(`Raw: ${data.toString()}`);
                    resolve({ raw: data.toString() });
                }
            });
        });
    }

    try {
        await new Promise(resolve => setTimeout(resolve, 3000));
        
        await sendRequest('initialize', {
            protocolVersion: "2024-11-05",
            clientInfo: { name: "Contoso Test", version: "1.0.0" },
            capabilities: {}
        });

        // Create Excel file in Contoso site
        await sendRequest('tools/call', {
            name: 'create-office-file',
            arguments: {
                fileType: 'excel',
                fileName: 'ContosoTestFile',
                location: 'sharepoint',
                sitePath: '/sites/Contoso'
            }
        });

        console.log('Test completed!');
    } catch (error) {
        console.error('Error:', error.message);
    } finally {
        server.kill();
        process.exit(0);
    }
}

testContosoSite();
