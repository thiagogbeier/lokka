#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { ConfidentialClientApplication } from "@azure/msal-node";
import { logger } from "./logger.js";

// Create server instance
const server = new McpServer({
  name: "Lokka-Microsoft",
  version: "0.1.8",
});

logger.info("Starting Lokka Multi-Microsoft API MCP Server");

// Initialize MSAL application outside the tool function
let msalApp: ConfidentialClientApplication | null = null;

server.tool(
  "Lokka-Microsoft",
  "A versatile tool to interact with Microsoft APIs including Microsoft Graph (Entra) and Azure Resource Management.",
  {
    apiType: z.enum(["graph", "azure"]).describe("Type of Microsoft API to query. Options: 'graph' for Microsoft Graph (Entra) or 'azure' for Azure Resource Management."),
    path: z.string().describe("The Azure or Graph API URL path to call (e.g. '/users', '/groups', '/subscriptions')"),
    method: z.enum(["get", "post", "put", "patch", "delete"]).describe("HTTP method to use"),
    apiVersion: z.string().optional().describe("Azure Resource Management API version (required for apiType Azure)"),
    subscriptionId: z.string().optional().describe("Azure Subscription ID (for Azure Resource Management)."),
    queryParams: z.record(z.string()).optional().describe("Query parameters for the request"),
    body: z.any().optional().describe("The request body (for POST, PUT, PATCH)"),
    graphApiVersion: z.enum(["v1.0", "beta"]).optional().default("v1.0").describe("Microsoft Graph API version to use (default: v1.0)"),
  },
  async ({
    apiType,
    path,
    method,
    apiVersion,
    subscriptionId,
    queryParams,
    body,
    graphApiVersion
  }: {
    apiType: "graph" | "azure";
    path: string;
    method: "get" | "post" | "put" | "patch" | "delete";
    apiVersion?: string;
    subscriptionId?: string;
    queryParams?: Record<string, string>;
    body?: any;
    graphApiVersion: "v1.0" | "beta";
  }) => {
    logger.info(`Executing Lokka-Microsoft tool with params: apiType=${apiType}, path=${path}, method=${method}, graphApiVersion=${graphApiVersion}`); // Log input params

    try {
      if (!msalApp) {
        throw new Error("MSAL application not initialized");
      }

      // Determine correct scope and base URL based on API type
      // Declare graphBaseUrl within the try block as its scope is limited here
      const graphBaseUrl = `https://graph.microsoft.com/${graphApiVersion}`;
      logger.info(`Using graphBaseUrl: ${graphBaseUrl} based on graphApiVersion: ${graphApiVersion}`);

      const apiConfig = { // Keep apiConfig local to try block
        graph: {
          scope: "https://graph.microsoft.com/.default",
          baseUrl: graphBaseUrl, // Use the dynamically determined base URL
        },
        azure: {
          scope: "https://management.azure.com/.default",
          baseUrl: "https://management.azure.com",
        }
      };

      const currentApi = apiConfig[apiType];

      // Acquire token using the initialized MSAL application
      const tokenResponse = await msalApp.acquireTokenByClientCredential({
        scopes: [currentApi.scope]
      });

      if (!tokenResponse || !tokenResponse.accessToken) {
        throw new Error("Failed to acquire access token");
      }

      // Construct the URL
      let url = currentApi.baseUrl;

      // Special handling for Azure Resource Management
      if (apiType === 'azure') {
        if (subscriptionId) {
          url += `/subscriptions/${subscriptionId}`;
        }

        // Append path
        url += path;

        // Add API version (required for Azure)
        if (!apiVersion) {
          throw new Error("API version is required for Azure Resource Management queries");
        }

        const urlParams = new URLSearchParams({
          'api-version': apiVersion
        });

        // Add additional query parameters if provided
        if (queryParams) {
          for (const [key, value] of Object.entries(queryParams)) {
            // Ensure key and value are treated as strings
            urlParams.append(String(key), String(value));
          }
        }

        url += `?${urlParams.toString()}`;
      }
      // Handling for Microsoft Graph
      else {
        url += path;

        // Add query parameters for Graph
        if (queryParams && Object.keys(queryParams).length > 0) {
          const searchParams = new URLSearchParams();
          for (const [key, value] of Object.entries(queryParams)) {
             // Ensure key and value are treated as strings
            searchParams.append(String(key), String(value));
          }
          url += `?${searchParams.toString()}`;
        }
      }

      // Prepare request options
      const headers: Record<string, string> = {
        'Authorization': `Bearer ${tokenResponse.accessToken}`,
        'Content-Type': 'application/json'
      };

      // Special headers for Graph consistency
      if (apiType === 'graph') {
        headers['ConsistencyLevel'] = 'eventual';
      }

      const requestOptions: RequestInit = {
        method: method.toUpperCase(),
        headers: headers
      };

      // Add body for methods that support it
      if (["POST", "PUT", "PATCH"].includes(method.toUpperCase())) {
        requestOptions.body = body ?
          (typeof body === 'string' ? body : JSON.stringify(body)) :
          JSON.stringify({});
      }

      // Make API request
      const apiResponse = await fetch(url, requestOptions);

      // Handle response
      let responseData: any;
      const responseText = await apiResponse.text();

      try {
        // Try to parse as JSON
        responseData = responseText ? JSON.parse(responseText) : {};
      } catch (e) {
        // If not JSON, use the raw text
        responseData = { rawResponse: responseText };
      }

      if (!apiResponse.ok) {
        logger.error(`API error for ${method} ${path}:`, responseData);
        throw new Error(`API error (${apiResponse.status}): ${JSON.stringify(responseData)}`);
      }

      // Include the determined base URL and version in the response for debugging
      let resultText = `Result for ${apiType} API (${graphApiVersion}) - ${method} ${path}:\n\n`;
      resultText += JSON.stringify(responseData, null, 2);

      return {
        content: [
          {
            type: "text" as const,
            text: resultText,
          },
        ],
      };

    } catch (error) {
      logger.error("Error in Multi-Microsoft API tool:", error);
      // Determine the URL that was attempted (declare determinedUrl within the catch block)
      // Simplified assignment as suggested
      const determinedUrl = apiType === 'graph'
        ? `https://graph.microsoft.com/${graphApiVersion}` // Use the provided version
        : "https://management.azure.com";
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify({
              error: error instanceof Error ? error.message : String(error),
              attemptedBaseUrl: determinedUrl // Add the URL here
            }),
          },
        ],
        isError: true
      };
    }
  },
);

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
