#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { Client, PageIterator, PageCollection } from "@microsoft/microsoft-graph-client";
import fetch from 'isomorphic-fetch'; // Required polyfill for Graph client
import { logger } from "./logger.js";
import { AuthManager, AuthConfig, AuthMode } from "./auth.js";

// Set up global fetch for the Microsoft Graph client
(global as any).fetch = fetch;

// Create server instance
const server = new McpServer({
  name: "Lokka-Microsoft",
  version: "0.2.0", // Updated version for token-based auth support
});

logger.info("Starting Lokka Multi-Microsoft API MCP Server (v0.2.0 - Token-Based Auth Support)");

// Initialize authentication and clients
let authManager: AuthManager | null = null;
let graphClient: Client | null = null;

server.tool(
  "Lokka-Microsoft",
  "A versatile tool to interact with Microsoft APIs including Microsoft Graph (Entra) and Azure Resource Management. IMPORTANT: For Graph API GET requests using advanced query parameters ($filter, $count, $search, $orderby), you are ADVISED to set 'consistencyLevel: \"eventual\"'.",
  {
    apiType: z.enum(["graph", "azure"]).describe("Type of Microsoft API to query. Options: 'graph' for Microsoft Graph (Entra) or 'azure' for Azure Resource Management."),
    path: z.string().describe("The Azure or Graph API URL path to call (e.g. '/users', '/groups', '/subscriptions')"),
    method: z.enum(["get", "post", "put", "patch", "delete"]).describe("HTTP method to use"),
    apiVersion: z.string().optional().describe("Azure Resource Management API version (required for apiType Azure)"),
    subscriptionId: z.string().optional().describe("Azure Subscription ID (for Azure Resource Management)."),
    queryParams: z.record(z.string()).optional().describe("Query parameters for the request"),
    body: z.record(z.string(), z.any()).optional().describe("The request body (for POST, PUT, PATCH)"),
    graphApiVersion: z.enum(["v1.0", "beta"]).optional().default("v1.0").describe("Microsoft Graph API version to use (default: v1.0)"),
    fetchAll: z.boolean().optional().default(false).describe("Set to true to automatically fetch all pages for list results (e.g., users, groups). Default is false."),
    consistencyLevel: z.string().optional().describe("Graph API ConsistencyLevel header. ADVISED to be set to 'eventual' for Graph GET requests using advanced query parameters ($filter, $count, $search, $orderby)."),
  },
  async ({
    apiType,
    path,
    method,
    apiVersion,
    subscriptionId,
    queryParams,
    body,
    graphApiVersion,
    fetchAll,
    consistencyLevel
  }: {
    apiType: "graph" | "azure";
    path: string;
    method: "get" | "post" | "put" | "patch" | "delete";
    apiVersion?: string;
    subscriptionId?: string;
    queryParams?: Record<string, string>;
    body?: any;
    graphApiVersion: "v1.0" | "beta";
    fetchAll: boolean;
    consistencyLevel?: string;
  }) => {
    logger.info(`Executing Lokka-Microsoft tool with params: apiType=${apiType}, path=${path}, method=${method}, graphApiVersion=${graphApiVersion}, fetchAll=${fetchAll}, consistencyLevel=${consistencyLevel}`);
    let determinedUrl: string | undefined;

    try {
      let responseData: any;

      // --- Microsoft Graph Logic ---
      if (apiType === 'graph') {
        if (!graphClient) {
          throw new Error("Graph client not initialized");
        }
        determinedUrl = `https://graph.microsoft.com/${graphApiVersion}`; // For error reporting

        // Construct the request using the Graph SDK client
        let request = graphClient.api(path).version(graphApiVersion);

        // Add query parameters if provided and not empty
        if (queryParams && Object.keys(queryParams).length > 0) {
          request = request.query(queryParams);
        }

        // Add ConsistencyLevel header if provided
        if (consistencyLevel) {
          request = request.header('ConsistencyLevel', consistencyLevel);
          logger.info(`Added ConsistencyLevel header: ${consistencyLevel}`);
        }

        // Handle different methods
        switch (method.toLowerCase()) {
          case 'get':
            if (fetchAll) {
              logger.info(`Fetching all pages for Graph path: ${path}`);
              // Fetch the first page to get context and initial data
              const firstPageResponse: PageCollection = await request.get();
              const odataContext = firstPageResponse['@odata.context']; // Capture context from first page
              let allItems: any[] = firstPageResponse.value || []; // Initialize with first page's items

              // Callback function to process subsequent pages
              const callback = (item: any) => {
                allItems.push(item);
                return true; // Return true to continue iteration
              };

              // Create a PageIterator starting from the first response
              const pageIterator = new PageIterator(graphClient, firstPageResponse, callback);

              // Iterate over all remaining pages
              await pageIterator.iterate();

              // Construct final response with context and combined values under 'value' key
              responseData = {
                '@odata.context': odataContext,
                value: allItems
              };
              logger.info(`Finished fetching all Graph pages. Total items: ${allItems.length}`);

            } else {
              logger.info(`Fetching single page for Graph path: ${path}`);
              responseData = await request.get();
            }
            break;
          case 'post':
            responseData = await request.post(body ?? {});
            break;
          case 'put':
            responseData = await request.put(body ?? {});
            break;
          case 'patch':
            responseData = await request.patch(body ?? {});
            break;
          case 'delete':
            responseData = await request.delete(); // Delete often returns no body or 204
            // Handle potential 204 No Content response
            if (responseData === undefined || responseData === null) {
              responseData = { status: "Success (No Content)" };
            }
            break;
          default:
            throw new Error(`Unsupported method: ${method}`);
        }
      }      // --- Azure Resource Management Logic (using direct fetch) ---
      else { // apiType === 'azure'
        if (!authManager) {
          throw new Error("Auth manager not initialized");
        }
        determinedUrl = "https://management.azure.com"; // For error reporting

        // Acquire token for Azure RM
        const azureCredential = authManager.getAzureCredential();
        const tokenResponse = await azureCredential.getToken("https://management.azure.com/.default");
        if (!tokenResponse || !tokenResponse.token) {
          throw new Error("Failed to acquire Azure access token");
        }

        // Construct the URL (similar to previous implementation)
        let url = determinedUrl;
        if (subscriptionId) {
          url += `/subscriptions/${subscriptionId}`;
        }
        url += path;

        if (!apiVersion) {
          throw new Error("API version is required for Azure Resource Management queries");
        }
        const urlParams = new URLSearchParams({ 'api-version': apiVersion });
        if (queryParams) {
          for (const [key, value] of Object.entries(queryParams)) {
            urlParams.append(String(key), String(value));
          }
        }
        url += `?${urlParams.toString()}`;

        // Prepare request options
        const headers: Record<string, string> = {
          'Authorization': `Bearer ${tokenResponse.token}`,
          'Content-Type': 'application/json'
        };
        const requestOptions: RequestInit = {
          method: method.toUpperCase(),
          headers: headers
        };
        if (["POST", "PUT", "PATCH"].includes(method.toUpperCase())) {
          requestOptions.body = body ? JSON.stringify(body) : JSON.stringify({});
        }

        // --- Pagination Logic for Azure RM (Manual Fetch) ---
        if (fetchAll && method === 'get') {
          logger.info(`Fetching all pages for Azure RM starting from: ${url}`);
          let allValues: any[] = [];
          let currentUrl: string | null = url;

          while (currentUrl) {            logger.info(`Fetching Azure RM page: ${currentUrl}`);
            // Re-acquire token for each page (Azure tokens might expire)
            const azureCredential = authManager.getAzureCredential();
            const currentPageTokenResponse = await azureCredential.getToken("https://management.azure.com/.default");
            if (!currentPageTokenResponse || !currentPageTokenResponse.token) {
              throw new Error("Failed to acquire Azure access token during pagination");
            }
            const currentPageHeaders = { ...headers, 'Authorization': `Bearer ${currentPageTokenResponse.token}` };
            const currentPageRequestOptions: RequestInit = { method: 'GET', headers: currentPageHeaders };

            const pageResponse = await fetch(currentUrl, currentPageRequestOptions);
            const pageText = await pageResponse.text();
            let pageData: any;
            try {
              pageData = pageText ? JSON.parse(pageText) : {};
            } catch (e) {
              logger.error(`Failed to parse JSON from Azure RM page: ${currentUrl}`, pageText);
              pageData = { rawResponse: pageText };
            }

            if (!pageResponse.ok) {
              logger.error(`API error on Azure RM page ${currentUrl}:`, pageData);
              throw new Error(`API error (${pageResponse.status}) during Azure RM pagination on ${currentUrl}: ${JSON.stringify(pageData)}`);
            }

            if (pageData.value && Array.isArray(pageData.value)) {
              allValues = allValues.concat(pageData.value);
            } else if (currentUrl === url && !pageData.nextLink) {
              allValues.push(pageData);
            } else if (currentUrl !== url) {
              logger.info(`[Warning] Azure RM response from ${currentUrl} did not contain a 'value' array.`);
            }
            currentUrl = pageData.nextLink || null; // Azure uses nextLink
          }
          responseData = { allValues: allValues };
          logger.info(`Finished fetching all Azure RM pages. Total items: ${allValues.length}`);
        } else {
          // Single page fetch for Azure RM
          logger.info(`Fetching single page for Azure RM: ${url}`);
          const apiResponse = await fetch(url, requestOptions);
          const responseText = await apiResponse.text();
          try {
            responseData = responseText ? JSON.parse(responseText) : {};
          } catch (e) {
            logger.error(`Failed to parse JSON from single Azure RM page: ${url}`, responseText);
            responseData = { rawResponse: responseText };
          }
          if (!apiResponse.ok) {
            logger.error(`API error for Azure RM ${method} ${path}:`, responseData);
            throw new Error(`API error (${apiResponse.status}) for Azure RM: ${JSON.stringify(responseData)}`);
          }
        }
      }

      // --- Format and Return Result ---
      // For all requests, format as text
      let resultText = `Result for ${apiType} API (${apiType === 'graph' ? graphApiVersion : apiVersion}) - ${method} ${path}:\n\n`;
      resultText += JSON.stringify(responseData, null, 2); // responseData already contains the correct structure for fetchAll Graph case

      // Add pagination note if applicable (only for single page GET)
      if (!fetchAll && method === 'get') {
         const nextLinkKey = apiType === 'graph' ? '@odata.nextLink' : 'nextLink';
         if (responseData && responseData[nextLinkKey]) { // Added check for responseData existence
             resultText += `\n\nNote: More results are available. To retrieve all pages, add the parameter 'fetchAll: true' to your request.`;
         }
      }

      return {
        content: [{ type: "text" as const, text: resultText }],
      };

    } catch (error: any) {
      logger.error(`Error in Lokka-Microsoft tool (apiType: ${apiType}, path: ${path}, method: ${method}):`, error); // Added more context to error log
      // Try to determine the base URL even in case of error
      if (!determinedUrl) {
         determinedUrl = apiType === 'graph'
           ? `https://graph.microsoft.com/${graphApiVersion}`
           : "https://management.azure.com";
      }
      // Include error body if available from Graph SDK error
      const errorBody = error.body ? (typeof error.body === 'string' ? error.body : JSON.stringify(error.body)) : 'N/A';
      return {
        content: [{
          type: "text",
          text: JSON.stringify({
            error: error instanceof Error ? error.message : String(error),
            statusCode: error.statusCode || 'N/A', // Include status code if available from SDK error
            errorBody: errorBody,
            attemptedBaseUrl: determinedUrl
          }),
        }],
        isError: true
      };
    }
  },
);

// Add token management tools
server.tool(
  "set-access-token",
  "Set or update the access token for Microsoft Graph authentication. Use this when the MCP Client has obtained a fresh token through interactive authentication.",
  {
    accessToken: z.string().describe("The access token obtained from Microsoft Graph authentication"),
    expiresOn: z.string().optional().describe("Token expiration time in ISO format (optional, defaults to 1 hour from now)")
  },
  async ({ accessToken, expiresOn }) => {
    try {
      const expirationDate = expiresOn ? new Date(expiresOn) : undefined;
      
      if (authManager?.getAuthMode() === AuthMode.ClientProvidedToken) {
        authManager.updateAccessToken(accessToken, expirationDate);
        
        // Reinitialize the Graph client with the new token
        const authProvider = authManager.getGraphAuthProvider();
        graphClient = Client.initWithMiddleware({
          authProvider: authProvider,
        });
        
        return {
          content: [{ 
            type: "text" as const, 
            text: "Access token updated successfully. You can now make Microsoft Graph requests on behalf of the authenticated user." 
          }],
        };
      } else {
        return {
          content: [{ 
            type: "text" as const, 
            text: "Error: MCP Server is not configured for client-provided token authentication. Set USE_CLIENT_TOKEN=true in environment variables." 
          }],
          isError: true
        };
      }
    } catch (error: any) {
      logger.error("Error setting access token:", error);
      return {
        content: [{ 
          type: "text" as const, 
          text: `Error setting access token: ${error.message}` 
        }],
        isError: true
      };
    }
  }
);

server.tool(
  "get-auth-status",
  "Check the current authentication status and mode of the MCP Server",
  {},
  async () => {
    try {
      const authMode = authManager?.getAuthMode() || "Not initialized";
      const isReady = authManager !== null;
      const tokenStatus = authManager?.getTokenStatus();
      
      return {
        content: [{ 
          type: "text" as const, 
          text: JSON.stringify({
            authMode,
            isReady,
            supportsTokenUpdates: authMode === AuthMode.ClientProvidedToken,
            tokenStatus: tokenStatus || { isExpired: false },
            timestamp: new Date().toISOString()
          }, null, 2)
        }],
      };
    } catch (error: any) {
      return {
        content: [{ 
          type: "text" as const, 
          text: `Error checking auth status: ${error.message}` 
        }],
        isError: true
      };
    }
  }
);

// Start the server with stdio transport
async function main() {
  // Determine authentication mode based on environment variables
  const tenantId = process.env.TENANT_ID;
  const clientId = process.env.CLIENT_ID;
  const clientSecret = process.env.CLIENT_SECRET;
  const useInteractive = process.env.USE_INTERACTIVE === 'true';
  const useClientToken = process.env.USE_CLIENT_TOKEN === 'true';
  const initialAccessToken = process.env.ACCESS_TOKEN;
  
  let authMode: AuthMode;
  
  if (useClientToken) {
    authMode = AuthMode.ClientProvidedToken;
    if (!initialAccessToken) {
      logger.info("Client token mode enabled but no initial token provided. Token must be set via set-access-token tool.");
    }
  } else if (useInteractive) {
    authMode = AuthMode.Interactive;
  } else {
    authMode = AuthMode.ClientCredentials;
  }

  logger.info(`Starting with authentication mode: ${authMode}`);

  const authConfig: AuthConfig = {
    mode: authMode,
    tenantId,
    clientId,
    clientSecret,
    accessToken: initialAccessToken,
    redirectUri: process.env.REDIRECT_URI
  };

  // Validate required configuration
  if (authMode === AuthMode.ClientCredentials) {
    if (!tenantId || !clientId || !clientSecret) {
      throw new Error("Client credentials mode requires TENANT_ID, CLIENT_ID, and CLIENT_SECRET");
    }
  } else if (authMode === AuthMode.Interactive) {
    if (!tenantId || !clientId) {
      throw new Error("Interactive mode requires TENANT_ID and CLIENT_ID");
    }
  }
  // Note: Client token mode can start without a token and receive it later

  authManager = new AuthManager(authConfig);
  
  // Only initialize if we have required config (for client token mode, we can start without a token)
  if (authMode !== AuthMode.ClientProvidedToken || initialAccessToken) {
    await authManager.initialize();
    
    // Initialize Graph Client
    const authProvider = authManager.getGraphAuthProvider();
    graphClient = Client.initWithMiddleware({
      authProvider: authProvider,
    });
    
    logger.info(`Authentication initialized successfully using ${authMode} mode`);
  } else {
    logger.info("Started in client token mode. Use set-access-token tool to provide authentication token.");
  }

  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch((error) => {
  console.error("Fatal error in main():", error);
  logger.error("Fatal error in main()", error);
  process.exit(1);
});
