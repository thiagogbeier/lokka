import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { ConfidentialClientApplication } from "@azure/msal-node";
// Create server instance
const server = new McpServer({
    name: "Lokka",
    version: "0.1.0",
});
server.tool("microsoftGraph", {
    path: z.string().describe("The Graph API URL path to call (e.g. '/me', '/users')"),
    method: z.enum(["get", "post", "put", "patch", "delete"]).describe("HTTP method to use"),
    queryParams: z.record(z.string()).optional().describe("Query parameters like $filter, $select, etc."),
    body: z.any().optional().describe("The request body (for POST, PUT, PATCH)"),
}, async ({ path, method, queryParams, body }) => {
    try {
        const tenantId = process.env.MS_GRAPH_TENANT_ID;
        const clientId = process.env.MS_GRAPH_CLIENT_ID;
        const clientSecret = process.env.MS_GRAPH_CLIENT_SECRET;
        if (!tenantId || !clientId || !clientSecret) {
            throw new Error("Missing required environment variables: MS_GRAPH_TENANT_ID, MS_GRAPH_CLIENT_ID, or MS_GRAPH_CLIENT_SECRET");
        }
        // Set up MSAL confidential client application
        const msalConfig = {
            auth: {
                clientId,
                clientSecret,
                authority: `https://login.microsoftonline.com/${tenantId}`,
            }
        };
        const cca = new ConfidentialClientApplication(msalConfig);
        // Acquire token
        const tokenResponse = await cca.acquireTokenByClientCredential({
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
        // Make Graph API request
        const graphResponse = await fetch(url, {
            method: method.toUpperCase(),
            headers: {
                "Authorization": `Bearer ${tokenResponse.accessToken}`,
                "Content-Type": "application/json",
            },
            ...(["POST", "PUT", "PATCH"].includes(method.toUpperCase()) && body ? { body: JSON.stringify(body) } : {}),
        });
        // Parse response
        const responseData = await graphResponse.json();
        if (!graphResponse.ok) {
            throw new Error(`Graph API error: ${JSON.stringify(responseData)}`);
        }
        return responseData;
    }
    catch (error) {
        return {
            error: error instanceof Error ? error.message : String(error),
        };
    }
});
// Start the server with stdio transport
async function main() {
    const transport = new StdioServerTransport();
    await server.connect(transport);
    console.error("Lokka MCP Server running on stdio");
}
main().catch((error) => {
    console.error("Fatal error in main():", error);
    process.exit(1);
});
