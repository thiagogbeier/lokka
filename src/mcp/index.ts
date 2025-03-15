import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";

// Create server instance
const server = new McpServer({
  name: "lokka",
  version: "0.1.0",
});

// Register Microsoft Graph API tool using server.tool
server.tool({
  name: "microsoftGraph",
  description: "Call Microsoft Graph API endpoints",
  parameters: z.object({
    path: z.string().describe("The Graph API URL path to call (e.g. '/me', '/users')"),
    method: z.enum(["get", "post", "put", "patch", "delete"]).describe("HTTP method to use"),
    queryParams: z.record(z.string()).optional().describe("Query parameters like $filter, $select, etc."),
    body: z.any().optional().describe("The request body (for POST, PUT, PATCH)"),
  }),
  configurationParameters: z.object({
    tenantId: z.string().describe("Microsoft tenant ID"),
    clientId: z.string().describe("Microsoft application (client) ID"),
    clientSecret: z.string().describe("Microsoft application client secret"),
  }),
  execute: async (params, context) => {
    try {
      const { path, method, queryParams, body } = params;
      const { tenantId, clientId, clientSecret } = context.configuration;

      // Get access token
      const tokenResponse = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
        method: "POST",
        headers: {
          "Content-Type": "application/x-www-form-urlencoded",
        },
        body: new URLSearchParams({
          client_id: clientId,
          client_secret: clientSecret,
          scope: "https://graph.microsoft.com/.default",
          grant_type: "client_credentials",
        }),
      });
      
      if (!tokenResponse.ok) {
        const errorData = await tokenResponse.json();
        throw new Error(`Failed to get access token: ${JSON.stringify(errorData)}`);
      }
      
      const { access_token } = await tokenResponse.json();
      
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
          "Authorization": `Bearer ${access_token}`,
          "Content-Type": "application/json",
        },
        ...(["POST", "PUT", "PATCH"].includes(method.toUpperCase()) && body ? { body: JSON.stringify(body) } : {}),
      });
      
      const responseData = await graphResponse.json();
      
      if (!graphResponse.ok) {
        throw new Error(`Graph API error: ${JSON.stringify(responseData)}`);
      }
      
      return responseData;
    } catch (error) {
      return {
        error: error instanceof Error ? error.message : String(error),
      };
    }
  },
});

// Start the server with stdio transport
async function main() {
    const transport = new StdioServerTransport();
    await server.connect(transport);
    console.error("Weather MCP Server running on stdio");
  }
  
  main().catch((error) => {
    console.error("Fatal error in main():", error);
    process.exit(1);
  });