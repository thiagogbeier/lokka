#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { ConfidentialClientApplication } from "@azure/msal-node";
import { logger } from "./logger.js";
// Create server instance
const server = new McpServer({
    name: "Lokka",
    version: "0.1.6",
});
logger.info("Starting Lokka MCP Server");
// Initialize MSAL application outside the tool function
let msalApp = null;
server.tool("Lokka-MicrosoftGraph", "A tool to call Microsoft Graph API. It supports querying a Microsoft 365 tenant using the Graph API. Updates are also supported if permissions are provided.", {
    path: z.string().describe("The Graph API URL path to call (e.g. '/me', '/users')"),
    method: z.enum(["get", "post", "put", "patch", "delete"]).describe("HTTP method to use"),
    queryParams: z.record(z.string()).optional().describe("Query parameters like $filter, $select, etc. All parameters are strings."),
    body: z.any().optional().describe("The request body (for POST, PUT, PATCH)"),
}, async ({ path, method, queryParams, body }) => {
    try {
        if (!msalApp) {
            throw new Error("MSAL application not initialized");
        }
        // Acquire token using the initialized MSAL application
        const tokenResponse = await msalApp.acquireTokenByClientCredential({
            scopes: ["https://graph.microsoft.com/.default"]
        });
        if (!tokenResponse || !tokenResponse.accessToken) {
            throw new Error("Failed to acquire access token");
        }
        // Build URL with query parameters
        let url = `https://graph.microsoft.com/v1.0${path}`;
        if (queryParams && Object.keys(queryParams).length > 0) {
            const searchParams = new URLSearchParams();
            for (const [key, value] of Object.entries(queryParams)) {
                searchParams.append(key, value);
            }
            url += `?${searchParams.toString()}`;
        }
        // Prepare headers
        const headers = {
            'Authorization': `Bearer ${tokenResponse.accessToken}`,
        };
        // For methods that send body data, add Content-Type header and ensure body is properly formatted
        const requestOptions = {
            method: method.toUpperCase(),
            headers: headers
        };
        // Only add Content-Type and body if we're using a method that supports sending data
        // and if body is provided
        if (["POST", "PUT", "PATCH"].includes(method.toUpperCase())) {
            if (body) {
                // Add Content-Type header
                headers['Content-Type'] = 'application/json';
                // Ensure body is properly stringified
                requestOptions.body = typeof body === 'string' ? body : JSON.stringify(body);
                // Log the request body for debugging
                logger.info(`Request body for ${method} ${path}: ${requestOptions.body}`);
            }
            else {
                // If no body is provided for methods that require it, send an empty object
                headers['Content-Type'] = 'application/json';
                requestOptions.body = JSON.stringify({});
                logger.info(`No body provided for ${method} ${path}. Using empty object instead.`);
            }
        }
        else if ("GET" === method.toUpperCase()) {
            headers['ConsistencyLevel'] = 'eventual';
        }
        // Make Graph API request
        const graphResponse = await fetch(url, requestOptions);
        // Handle response
        let responseData;
        const responseText = await graphResponse.text();
        try {
            // Try to parse as JSON
            responseData = responseText ? JSON.parse(responseText) : {};
        }
        catch (e) {
            // If not JSON, use the raw text
            responseData = { rawResponse: responseText };
        }
        if (!graphResponse.ok) {
            logger.error(`Graph API error for ${method} ${path}:`, responseData);
            throw new Error(`Graph API error (${graphResponse.status}): ${JSON.stringify(responseData)}`);
        }
        let resultText = `Result for ${method} ${path}:\n\n`;
        resultText += JSON.stringify(responseData, null, 2);
        return {
            content: [
                {
                    type: "text",
                    text: resultText,
                },
            ],
        };
    }
    catch (error) {
        logger.error("Error in microsoftGraph tool:", error);
        return {
            content: [
                {
                    type: "text",
                    text: JSON.stringify({
                        error: error instanceof Error ? error.message : String(error),
                    }),
                },
            ],
            isError: true
        };
    }
});
// Start the server with stdio transport
async function main() {
    // Check for required environment variables
    const tenantId = process.env.TENANT_ID;
    const clientId = process.env.CLIENT_ID;
    const clientSecret = process.env.CLIENT_SECRET;
    if (!tenantId || !clientId || !clientSecret) {
        throw new Error("Missing required environment variables: TENANT_ID, CLIENT_ID, or CLIENT_SECRET");
    }
    // Initialize MSAL application once
    const msalConfig = {
        auth: {
            clientId,
            clientSecret,
            authority: `https://login.microsoftonline.com/${tenantId}`,
        }
    };
    msalApp = new ConfidentialClientApplication(msalConfig);
    const transport = new StdioServerTransport();
    await server.connect(transport);
}
main().catch((error) => {
    console.error("Fatal error in main():", error);
    logger.error("Fatal error in main()", error);
    process.exit(1);
});
