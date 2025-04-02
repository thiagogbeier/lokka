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
    fetchAll: z.boolean().optional().default(false).describe("Set to true to automatically fetch all pages for list results (e.g., users, groups). Default is false."),
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
    fetchAll
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
  }) => {
    logger.info(`Executing Lokka-Microsoft tool with params: apiType=${apiType}, path=${path}, method=${method}, graphApiVersion=${graphApiVersion}, fetchAll=${fetchAll}`); // Log input params
    let determinedUrl: string | undefined; // Declare with wider scope
    try {
      if (!msalApp) {
        throw new Error("MSAL application not initialized");
      }

      // Determine correct scope and base URL based on API type
      const graphBaseUrl = `https://graph.microsoft.com/${graphApiVersion}`; // Use the parameter directly
      logger.info(`Using graphBaseUrl: ${graphBaseUrl} based on graphApiVersion: ${graphApiVersion}`); // Log determined URL

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

       // --- Pagination Logic ---
       let responseData: any;
 
       if (fetchAll && method === 'get') { // Only paginate GET requests
         logger.info(`Fetching all pages starting from: ${url}`);
         let allValues: any[] = []; // Initialize results array
         let currentUrl: string | null = url; // URL for the current page
 
         while (currentUrl) {
           logger.info(`Fetching page: ${currentUrl}`);
           
           // Re-acquire token INSIDE the loop for each page request
           const currentTokenResponse = await msalApp.acquireTokenByClientCredential({
             scopes: [currentApi.scope]
           });
           if (!currentTokenResponse || !currentTokenResponse.accessToken) {
             throw new Error("Failed to acquire access token during pagination");
           }
 
           // Prepare headers INSIDE the loop for each page request
           const currentHeaders: Record<string, string> = {
             'Authorization': `Bearer ${currentTokenResponse.accessToken}`,
             'Content-Type': 'application/json'
           };
           if (apiType === 'graph') {
             currentHeaders['ConsistencyLevel'] = 'eventual';
             currentHeaders['Accept-Language'] = 'en-US'; // Ensure this is included
           }
           const currentPageRequestOptions: RequestInit = {
             method: 'GET', // Pagination is always GET
             headers: currentHeaders
           };
 
           // Fetch the current page
           const pageResponse = await fetch(currentUrl, currentPageRequestOptions); 
           const pageText = await pageResponse.text();
           let pageData: any;

           try {
             pageData = pageText ? JSON.parse(pageText) : {};
           } catch (e) {
             logger.error(`Failed to parse JSON from page: ${currentUrl}`, pageText); // Use currentUrl
             pageData = { rawResponse: pageText }; // Store raw text on parse error
           }
 
           if (!pageResponse.ok) {
             logger.error(`API error on page ${currentUrl}:`, pageData); // Use currentUrl
             // Throw error, stopping pagination
             throw new Error(`API error (${pageResponse.status}) during pagination on ${currentUrl}: ${JSON.stringify(pageData)}`); // Use currentUrl
           }
 
           if (pageData.value && Array.isArray(pageData.value)) {
             allValues = allValues.concat(pageData.value);
           } else {
              // If the first page doesn't have 'value', treat it as a single result
              if (currentUrl === url && !pageData['@odata.nextLink']) { // Use currentUrl
                  allValues.push(pageData); // Add the whole response if it's not a collection
              }
               // Otherwise, log a warning if subsequent pages lack 'value'
               else if (currentUrl !== url) { // Use currentUrl
                   logger.info(`[Warning] Response from ${currentUrl} did not contain a 'value' array.`); // Use currentUrl
               }
            }
 
           currentUrl = pageData['@odata.nextLink'] || null; // Update currentUrl for the next iteration
        }
        
         // Final response data structure for fetchAll
         responseData = { allValues: allValues }; 
         logger.info(`Finished fetching all pages. Total items: ${allValues.length}`);
 
       } else {
         // Original single-page fetch logic
         logger.info(`Fetching single page: ${url}`);
         // Use the original requestOptions prepared outside the loop
         const apiResponse = await fetch(url, requestOptions); 
         const responseText = await apiResponse.text();
         
         try {
          responseData = responseText ? JSON.parse(responseText) : {};
        } catch (e) {
          logger.error(`Failed to parse JSON from single page: ${url}`, responseText);
          responseData = { rawResponse: responseText };
        }

        if (!apiResponse.ok) {
          logger.error(`API error for ${method} ${path}:`, responseData);
          throw new Error(`API error (${apiResponse.status}): ${JSON.stringify(responseData)}`);
        }
      }
       // --- End Pagination Logic ---

       // responseData is now populated either with single page or allValues object
 
       // Construct result text based on whether all pages were fetched
       let resultText = `Result for ${apiType} API (${graphApiVersion}) - ${method} ${path}:\n\n`;
 
       if (fetchAll && method === 'get' && responseData.allValues) {
          // Use the structure containing allValues
          resultText += JSON.stringify(responseData, null, 2); // Stringify the object containing allValues
       } else {
          // Original single-page response formatting
          resultText += JSON.stringify(responseData, null, 2);
          // Add note if more pages are available and we didn't fetch all
          if (!fetchAll && responseData['@odata.nextLink']) {
             resultText += `\n\nNote: More results are available. To retrieve all pages, add the parameter 'fetchAll: true' to your request.`;
          }
       }
       
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
      // Determine the URL that was attempted
      determinedUrl = apiType === 'graph' 
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
