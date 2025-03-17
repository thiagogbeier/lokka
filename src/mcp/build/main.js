import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { ConfidentialClientApplication } from "@azure/msal-node";
import { logger } from "./logger.js";
// Create server instance
const server = new McpServer({
    name: "Lokka",
    version: "0.1.0",
});
logger.info("Starting Lokka MCP Server");
// Initialize MSAL application outside the tool function
let msalApp = null;
server.tool("microsoftGraph", {
    path: z.string().describe("The Graph API URL path to call (e.g. '/me', '/users')"),
    method: z.enum(["get", "post", "put", "patch", "delete"]).describe("HTTP method to use"),
    queryParams: z.record(z.string()).optional().describe("Query parameters like $filter, $select, etc. All paremeters are strings."),
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
        let processedBody = body;
        // Fix for handling stringified JSON body
        if (body) {
            if (typeof body === 'string') {
                try {
                    // If it's a string, try to parse it as JSON
                    const parsedBody = JSON.parse(body);
                    // Use the parsed object as the processed body
                    processedBody = parsedBody;
                }
                catch (e) {
                    // If parsing fails, keep the original string
                    console.error('Failed to parse body as JSON, using as is:', e);
                }
            }
        }
        // Make Graph API request
        const graphResponse = await fetch(url, {
            method: method.toUpperCase(),
            headers: {
                'Authorization': `Bearer ${tokenResponse.accessToken}`,
                'Accept': 'application/json',
                'Content-Type': 'application/json',
                'ConsistencyLevel': 'eventual' // Include consistency level in all requests
            },
            ...(["POST", "PUT", "PATCH"].includes(method.toUpperCase()) && body ? { body: processedBody } : {}),
        });
        // Parse response
        const responseData = await graphResponse.json();
        if (!graphResponse.ok) {
            throw new Error(`Graph API error: ${JSON.stringify(responseData)}`);
        }
        let resultText = `Result for ${method} ${path} ${queryParams}:\n\n`;
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
        };
    }
});
// Start the server with stdio transport
async function main() {
    // Check for required environment variables
    const tenantId = process.env.MS_GRAPH_TENANT_ID;
    const clientId = process.env.MS_GRAPH_CLIENT_ID;
    const clientSecret = process.env.MS_GRAPH_CLIENT_SECRET;
    if (!tenantId || !clientId || !clientSecret) {
        throw new Error("Missing required environment variables: MS_GRAPH_TENANT_ID, MS_GRAPH_CLIENT_ID, or MS_GRAPH_CLIENT_SECRET");
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
    console.error("Lokka MCP Server running on stdio");
}
main().catch((error) => {
    console.error("Fatal error in main():", error);
    process.exit(1);
});
