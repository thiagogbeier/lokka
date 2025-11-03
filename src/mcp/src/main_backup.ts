#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { Client, PageIterator, PageCollection } from "@microsoft/microsoft-graph-client";
import fetch from 'isomorphic-fetch'; // Required polyfill for Graph client
import JSZip from 'jszip';
import { logger } from "./logger.js";
import { AuthManager, AuthConfig, AuthMode } from "./auth.js";
import { LokkaClientId, LokkaDefaultTenantId, LokkaDefaultRedirectUri, getDefaultGraphApiVersion } from "./constants.js";

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

// Check USE_GRAPH_BETA environment variable
const useGraphBeta = process.env.USE_GRAPH_BETA !== 'false'; // Default to true unless explicitly set to 'false'
const defaultGraphApiVersion = getDefaultGraphApiVersion();

logger.info(`Graph API default version: ${defaultGraphApiVersion} (USE_GRAPH_BETA=${process.env.USE_GRAPH_BETA || 'undefined'})`);

// Helper functions for Office file creation
async function createEmptyOfficeFile(fileType: "word" | "excel" | "powerpoint"): Promise<Buffer> {
  const zip = new JSZip();

  switch (fileType) {
    case "word":
      return createEmptyWordDocument(zip);
    case "excel":
      return createEmptyExcelWorkbook(zip);
    case "powerpoint":
      return createEmptyPowerPointPresentation(zip);
    default:
      throw new Error(`Unsupported file type: ${fileType}`);
  }
}

async function createEmptyWordDocument(zip: JSZip): Promise<Buffer> {
  // Create minimal Word document structure
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`);

  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);

  zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>This is a new Word document created by Lokka MCP.</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`);

  return await zip.generateAsync({ type: "nodebuffer" });
}

async function createEmptyExcelWorkbook(zip: JSZip): Promise<Buffer> {
  // Create minimal Excel workbook structure
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`);

  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`);

  zip.file("xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`);

  zip.file("xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`);

  zip.file("xl/worksheets/sheet1.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr">
        <is><t>Created by Lokka MCP</t></is>
      </c>
    </row>
  </sheetData>
</worksheet>`);

  return await zip.generateAsync({ type: "nodebuffer" });
}

async function createEmptyPowerPointPresentation(zip: JSZip): Promise<Buffer> {
  // Create minimal PowerPoint presentation structure  
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
</Types>`);

  zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`);

  zip.file("ppt/presentation.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId2"/>
  </p:sldIdLst>
</p:presentation>`);

  zip.file("ppt/slides/slide1.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title Placeholder 1"/>
          <p:cNvSpPr>
            <a:spLocks noGrp="1"/>
          </p:cNvSpPr>
          <p:nvPr>
            <p:ph type="ctrTitle"/>
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr/>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r>
              <a:rPr lang="en-US" smtClean="0"/>
              <a:t>Created by Lokka MCP</a:t>
            </a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`);

  return await zip.generateAsync({ type: "nodebuffer" });
}

async function createOfficeFileWithContent(fileType: "word" | "excel" | "powerpoint", templateContent: any): Promise<Buffer> {
  // This function can be extended to handle template content
  // For now, create empty file and add custom logic based on templateContent
  const baseFile = await createEmptyOfficeFile(fileType);
  
  // TODO: Implement template content processing
  // You could use libraries like:
  // - docxtemplater for Word documents
  // - exceljs for Excel files  
  // - officegen for PowerPoint files
  
  return baseFile;
}

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
    graphApiVersion: z.enum(["v1.0", "beta"]).optional().default(defaultGraphApiVersion as "v1.0" | "beta").describe(`Microsoft Graph API version to use (default: ${defaultGraphApiVersion})`),
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
    // Override graphApiVersion if USE_GRAPH_BETA is explicitly set to false
    const effectiveGraphApiVersion = !useGraphBeta ? "v1.0" : graphApiVersion;
    
    logger.info(`Executing Lokka-Microsoft tool with params: apiType=${apiType}, path=${path}, method=${method}, graphApiVersion=${effectiveGraphApiVersion}, fetchAll=${fetchAll}, consistencyLevel=${consistencyLevel}`);
    let determinedUrl: string | undefined;

    try {
      let responseData: any;

      // --- Microsoft Graph Logic ---
      if (apiType === 'graph') {
        // Initialize auth if not already done (for interactive mode)
        if (!graphClient && authManager && authManager.getAuthMode() === AuthMode.Interactive) {
          logger.info("Initializing interactive authentication on first API call");
          try {
            console.log("\n🔐 Authentication Required:");
            console.log("This will open device code authentication...");
            await authManager.initialize();
            const authProvider = authManager.getGraphAuthProvider();
            graphClient = Client.initWithMiddleware({
              authProvider: authProvider,
            });
            logger.info("Authentication initialized successfully");
          } catch (error) {
            logger.error("Interactive authentication failed:", error);
            throw new Error("Interactive authentication failed. Please check the console for device code instructions.");
          }
        }
        
        if (!graphClient) {
          throw new Error("Graph client not initialized");
        }
        determinedUrl = `https://graph.microsoft.com/${effectiveGraphApiVersion}`; // For error reporting

        // Construct the request using the Graph SDK client
        let request = graphClient.api(path).version(effectiveGraphApiVersion);

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
      let resultText = `Result for ${apiType} API (${apiType === 'graph' ? effectiveGraphApiVersion : apiVersion}) - ${method} ${path}:\n\n`;
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
           ? `https://graph.microsoft.com/${effectiveGraphApiVersion}`
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
  "Check the current authentication status and mode of the MCP Server and also returns the current graph permission scopes of the access token for the current session.",
  {},
  async () => {
    try {
      const authMode = authManager?.getAuthMode() || "Not initialized";
      const isReady = authManager !== null;
      const tokenStatus = authManager ? await authManager.getTokenStatus() : { isExpired: false };
      
      return {
        content: [{ 
          type: "text" as const, 
          text: JSON.stringify({
            authMode,
            isReady,
            supportsTokenUpdates: authMode === AuthMode.ClientProvidedToken,
            tokenStatus: tokenStatus,
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

// Add tool for requesting additional Graph permissions
server.tool(
  "add-graph-permission",
  "Request additional Microsoft Graph permission scopes by performing a fresh interactive sign-in. This tool only works in interactive authentication mode and should be used if any Graph API call returns permissions related errors.",
  {
    scopes: z.array(z.string()).describe("Array of Microsoft Graph permission scopes to request (e.g., ['User.Read', 'Mail.ReadWrite', 'Directory.Read.All'])")
  },
  async ({ scopes }) => {
    try {
      // Check if we're in interactive mode
      if (!authManager || authManager.getAuthMode() !== AuthMode.Interactive) {
        const currentMode = authManager?.getAuthMode() || "Not initialized";
        const clientId = process.env.CLIENT_ID;
        
        let errorMessage = `Error: add-graph-permission tool is only available in interactive authentication mode. Current mode: ${currentMode}.\n\n`;
        
        if (currentMode === AuthMode.ClientCredentials) {
          errorMessage += `📋 To add permissions in Client Credentials mode:\n`;
          errorMessage += `1. Open the Microsoft Entra admin center (https://entra.microsoft.com)\n`;
          errorMessage += `2. Navigate to Applications > App registrations\n`;
          errorMessage += `3. Find your application${clientId ? ` (Client ID: ${clientId})` : ''}\n`;
          errorMessage += `4. Go to API permissions\n`;
          errorMessage += `5. Click "Add a permission" and select Microsoft Graph\n`;
          errorMessage += `6. Choose "Application permissions" and add the required scopes:\n`;
          errorMessage += `   ${scopes.map(scope => `• ${scope}`).join('\n   ')}\n`;
          errorMessage += `7. Click "Grant admin consent" to approve the permissions\n`;
          errorMessage += `8. Restart the MCP server to use the new permissions`;
        } else if (currentMode === AuthMode.ClientProvidedToken) {
          errorMessage += `📋 To add permissions in Client Provided Token mode:\n`;
          errorMessage += `1. Obtain a new access token that includes the required scopes:\n`;
          errorMessage += `   ${scopes.map(scope => `• ${scope}`).join('\n   ')}\n`;
          errorMessage += `2. When obtaining the token, ensure these scopes are included in the consent prompt\n`;
          errorMessage += `3. Use the set-access-token tool to update the server with the new token\n`;
          errorMessage += `4. The new token will include the additional permissions`;
        } else {
          errorMessage += `To use interactive permission requests, set USE_INTERACTIVE=true in environment variables and restart the server.`;
        }
        
        return {
          content: [{ 
            type: "text" as const, 
            text: errorMessage
          }],
          isError: true
        };
      }

      // Validate scopes array
      if (!scopes || scopes.length === 0) {
        return {
          content: [{ 
            type: "text" as const, 
            text: "Error: At least one permission scope must be specified." 
          }],
          isError: true
        };
      }

      // Validate scope format (basic validation)
      const invalidScopes = scopes.filter(scope => !scope.includes('.') || scope.trim() !== scope);
      if (invalidScopes.length > 0) {
        return {
          content: [{ 
            type: "text" as const, 
            text: `Error: Invalid scope format detected: ${invalidScopes.join(', ')}. Scopes should be in format like 'User.Read' or 'Mail.ReadWrite'.` 
          }],
          isError: true
        };
      }

      logger.info(`Requesting additional Graph permissions: ${scopes.join(', ')}`);

      // Get current configuration with defaults for interactive auth
      const tenantId = process.env.TENANT_ID || LokkaDefaultTenantId;
      const clientId = process.env.CLIENT_ID || LokkaClientId;
      const redirectUri = process.env.REDIRECT_URI || LokkaDefaultRedirectUri;

      logger.info(`Using tenant ID: ${tenantId}, client ID: ${clientId} for interactive authentication`);

      // Create a new interactive credential with the requested scopes
      const { InteractiveBrowserCredential, DeviceCodeCredential } = await import("@azure/identity");
      
      // Clear any existing auth manager to force fresh authentication
      authManager = null;
      graphClient = null;
      
      // Request token with the new scopes - this will trigger interactive authentication
      const scopeString = scopes.map(scope => `https://graph.microsoft.com/${scope}`).join(' ');
      logger.info(`Requesting fresh token with scopes: ${scopeString}`);
      
      console.log(`\n🔐 Requesting Additional Graph Permissions:`);
      console.log(`Scopes: ${scopes.join(', ')}`);
      console.log(`You will be prompted to sign in to grant these permissions.\n`);

      let newCredential;
      let tokenResponse;
      
      // Use Device Code flow by default to avoid redirect URI issues
      logger.info("Using device code flow to avoid redirect URI configuration issues");
      newCredential = new DeviceCodeCredential({
        tenantId: tenantId,
        clientId: clientId,
        userPromptCallback: (info) => {
          console.log(`\n🔐 Additional Permissions Required:`);
          console.log(`Please visit: ${info.verificationUri}`);
          console.log(`And enter code: ${info.userCode}`);
          console.log(`Requested scopes: ${scopes.join(', ')}\n`);
          return Promise.resolve();
        },
      });
      
      // Request token with device code credential
      tokenResponse = await newCredential.getToken(scopeString);

      if (!tokenResponse) {
        return {
          content: [{ 
            type: "text" as const, 
            text: "Error: Failed to acquire access token with the requested scopes. Please check your permissions and try again." 
          }],
          isError: true
        };
      }

      // Create a completely new auth manager instance with the updated credential
      const authConfig: AuthConfig = {
        mode: AuthMode.Interactive,
        tenantId,
        clientId,
        redirectUri
      };

      // Create a new auth manager instance
      authManager = new AuthManager(authConfig);
      
      // Manually set the credential to our new one with the additional scopes
      (authManager as any).credential = newCredential;

      // DO NOT call initialize() as it might interfere with our fresh token
      // Instead, directly create the Graph client with the new credential
      const authProvider = authManager.getGraphAuthProvider();
      graphClient = Client.initWithMiddleware({
        authProvider: authProvider,
      });

      // Get the token status to show the new scopes
      const tokenStatus = await authManager.getTokenStatus();

      logger.info(`Successfully acquired fresh token with additional scopes: ${scopes.join(', ')}`);

      return {
        content: [{ 
          type: "text" as const, 
          text: JSON.stringify({
            message: "Successfully acquired additional Microsoft Graph permissions with fresh authentication",
            requestedScopes: scopes,
            tokenStatus: tokenStatus,
            note: "A fresh sign-in was performed to ensure the new permissions are properly granted",
            timestamp: new Date().toISOString()
          }, null, 2)
        }],
      };

    } catch (error: any) {
      logger.error("Error requesting additional Graph permissions:", error);
      return {
        content: [{ 
          type: "text" as const, 
          text: `Error requesting additional permissions: ${error.message}` 
        }],
        isError: true
      };
    }
  }
);

// Add this after the existing tools, before the main() function
server.tool(
  "create-office-file",
  "Create Microsoft Office files (Word, Excel, PowerPoint) in SharePoint Online or OneDrive for Business using Microsoft Graph API.",
  {
    fileType: z.enum(["word", "excel", "powerpoint"]).describe("Type of Office file to create"),
    fileName: z.string().describe("Name of the file (without extension, will be added automatically)"),
    location: z.enum(["sharepoint", "onedrive"]).describe("Where to create the file: SharePoint site or user's OneDrive"),
    sitePath: z.string().optional().describe("SharePoint site path (e.g., '/sites/TeamSite') - required for SharePoint location"),
    libraryName: z.string().optional().describe("Document library name (default: 'Shared Documents' for SharePoint, root for OneDrive)"),
    folderPath: z.string().optional().describe("Folder path within the library (e.g., '/Projects/Q4')"),
    templateContent: z.record(z.string(), z.any()).optional().describe("Template content for the file (varies by file type)"),
    userPrincipalName: z.string().optional().describe("User's UPN for OneDrive access (defaults to current user)")
  },
  async ({ fileType, fileName, location, sitePath, libraryName, folderPath, templateContent, userPrincipalName }) => {
    try {
      // Initialize auth if not already done (for interactive mode)
      if (!graphClient && authManager && authManager.getAuthMode() === AuthMode.Interactive) {
        logger.info("Initializing interactive authentication for file creation");
        try {
          console.log("\n🔐 Authentication Required for File Creation:");
          console.log("Please complete device code authentication...");
          await authManager.initialize();
          const authProvider = authManager.getGraphAuthProvider();
          graphClient = Client.initWithMiddleware({
            authProvider: authProvider,
          });
          logger.info("Authentication initialized successfully");
        } catch (error) {
          logger.error("Interactive authentication failed:", error);
          throw new Error("Interactive authentication failed. Please check the console for device code instructions.");
        }
      }
      
      if (!graphClient) {
        throw new Error("Graph client not initialized");
      }

      logger.info(`Creating ${fileType} file '${fileName}' in ${location}`);

      // Determine file extension and content type
      const fileExtensions = {
        word: "docx",
        excel: "xlsx", 
        powerpoint: "pptx"
      };

      const contentTypes = {
        word: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        excel: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        powerpoint: "application/vnd.openxmlformats-officedocument.presentationml.presentation"
      };

      const fullFileName = `${fileName}.${fileExtensions[fileType]}`;
      const contentType = contentTypes[fileType];

      // Build the target path
      let targetPath: string;
      
      if (location === "sharepoint") {
        if (!sitePath) {
          throw new Error("sitePath is required for SharePoint location");
        }
        const library = libraryName || "Shared Documents";
        const folder = folderPath || "";
        targetPath = `/sites/${sitePath.replace(/^\/sites\//, '')}/drive/root:/${library}${folder}/${fullFileName}:/content`;
      } else {
        // OneDrive for Business
        const userPart = userPrincipalName ? `/users/${userPrincipalName}` : "/me";
        const folder = folderPath || "";
        targetPath = `${userPart}/drive/root:${folder}/${fullFileName}:/content`;
      }

      // Create empty Office file with template
      let fileContent: Buffer;
      
      if (templateContent) {
        // If template content provided, create file with content
        fileContent = await createOfficeFileWithContent(fileType, templateContent);
      } else {
        // Create minimal empty file
        fileContent = await createEmptyOfficeFile(fileType);
      }

      // Upload the file using Graph API
      const uploadResponse = await graphClient
        .api(targetPath)
        .header('Content-Type', contentType)
        .put(fileContent);

      logger.info(`Successfully created ${fileType} file: ${fullFileName}`);

      return {
        content: [{
          type: "text" as const,
          text: JSON.stringify({
            message: `Successfully created ${fileType} file`,
            fileName: fullFileName,
            location: location,
            fileId: uploadResponse.id,
            webUrl: uploadResponse.webUrl,
            downloadUrl: uploadResponse["@microsoft.graph.downloadUrl"],
            size: uploadResponse.size,
            createdDateTime: uploadResponse.createdDateTime,
            lastModifiedDateTime: uploadResponse.lastModifiedDateTime
          }, null, 2)
        }]
      };

    } catch (error: any) {
      logger.error("Error creating Office file:", error);
      return {
        content: [{
          type: "text" as const,
          text: `Error creating Office file: ${error.message}`
        }],
        isError: true
      };
    }
  }
);

async function main() {
  // Determine authentication mode
  const useClientToken = process.env.USE_CLIENT_TOKEN === 'true';
  const useInteractive = process.env.USE_INTERACTIVE === 'true';
  const useCertificate = process.env.USE_CERTIFICATE === 'true';
  const initialAccessToken = process.env.ACCESS_TOKEN;

  let authMode: AuthMode;
  if (useClientToken) {
    authMode = AuthMode.ClientProvidedToken;
  } else if (useInteractive) {
    authMode = AuthMode.Interactive;
  } else if (useCertificate) {
    authMode = AuthMode.Certificate;
  } else {
    authMode = AuthMode.ClientCredentials; // Default to client credentials
  }

  logger.info(`Starting with authentication mode: ${authMode}`);

  const tenantId = process.env.TENANT_ID;
  const clientId = process.env.CLIENT_ID;
  const clientSecret = process.env.CLIENT_SECRET;
  const certificatePath = process.env.CERTIFICATE_PATH;
  const certificatePassword = process.env.CERTIFICATE_PASSWORD; // optional

  // Validate required configuration
  if (authMode === AuthMode.ClientCredentials) {
    if (!tenantId || !clientId || !clientSecret) {
      throw new Error("Client credentials mode requires explicit TENANT_ID, CLIENT_ID, and CLIENT_SECRET environment variables");
    }
  } else if (authMode === AuthMode.Certificate) {
    if (!tenantId || !clientId || !certificatePath) {
      throw new Error("Certificate mode requires explicit TENANT_ID, CLIENT_ID, and CERTIFICATE_PATH environment variables");
    }
  }
  // Note: Client token mode can start without a token and receive it later

  const authConfig: AuthConfig = {
    mode: authMode,
    tenantId,
    clientId,
    clientSecret,
    accessToken: initialAccessToken,
    redirectUri: process.env.REDIRECT_URI,
    certificatePath,
    certificatePassword
  };

  authManager = new AuthManager(authConfig);
  
  // Start the MCP server first
  const transport = new StdioServerTransport();
  await server.connect(transport);
  
  // For interactive mode, defer authentication until first API call
  if (authMode === AuthMode.Interactive) {
    logger.info("Started in interactive mode. Authentication will be requested on first API call.");
    logger.info("Note: Device code authentication will be triggered when you make your first Graph API call.");
  } else if (authMode !== AuthMode.ClientProvidedToken || initialAccessToken) {
    // Initialize auth for other modes
    try {
      await authManager.initialize();
      
      // Initialize Graph Client
      const authProvider = authManager.getGraphAuthProvider();
      graphClient = Client.initWithMiddleware({
        authProvider: authProvider,
      });
      
      logger.info(`Authentication initialized successfully using ${authMode} mode`);
    } catch (error) {
      logger.error("Authentication initialization failed:", error);
      // Don't fail the server startup, just log the error
    }
  } else {
    logger.info("Started in client token mode. Use set-access-token tool to provide authentication token.");
  }
}

main().catch((error) => {
  console.error("Fatal error in main():", error);
  logger.error("Fatal error in main()", error);
  process.exit(1);
});
