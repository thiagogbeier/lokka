#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { Client, PageIterator } from "@microsoft/microsoft-graph-client";
import fetch from 'isomorphic-fetch'; // Required polyfill for Graph client
import JSZip from 'jszip';
import { logger } from "./logger.js";
import { AuthManager, AuthMode } from "./auth.js";
import { LokkaClientId, LokkaDefaultTenantId, LokkaDefaultRedirectUri, getDefaultGraphApiVersion } from "./constants.js";
// Set up global fetch for the Microsoft Graph client
global.fetch = fetch;
// Create server instance
const server = new McpServer({
    name: "Lokka-Microsoft",
    version: "0.2.0", // Updated version for token-based auth support
});
logger.info("Starting Lokka Multi-Microsoft API MCP Server (v0.2.0 - Token-Based Auth Support)");
// Initialize authentication and clients
let authManager = null;
let graphClient = null;
// Check USE_GRAPH_BETA environment variable
const useGraphBeta = process.env.USE_GRAPH_BETA !== 'false'; // Default to true unless explicitly set to 'false'
const defaultGraphApiVersion = getDefaultGraphApiVersion();
logger.info(`Graph API default version: ${defaultGraphApiVersion} (USE_GRAPH_BETA=${process.env.USE_GRAPH_BETA || 'undefined'})`);
// Helper functions for Office file creation - Using Graph API built-in templates
async function createEmptyOfficeFile(fileType) {
    // Return minimal content that Graph API can understand and convert
    switch (fileType) {
        case "word":
            return "Created by Lokka MCP"; // Plain text that Graph API will convert to DOCX
        case "excel":
            return "Created by Lokka MCP"; // Plain text that can be put in cell A1
        case "powerpoint":
            return "Created by Lokka MCP"; // Plain text for slide title
        default:
            throw new Error(`Unsupported file type: ${fileType}`);
    }
}
// Use Graph API's copy from template approach
async function createOfficeFileFromTemplate(fileType) {
    // Use Microsoft's online templates or create via Graph API directly
    const templateIds = {
        word: "blank_document_template",
        excel: "blank_workbook_template",
        powerpoint: "blank_presentation_template"
    };
    return {
        "@microsoft.graph.conflictBehavior": "rename",
        name: `template_${fileType}`,
        content: "Created by Lokka MCP"
    };
}
server.tool("Lokka-Microsoft", "A versatile tool to interact with Microsoft APIs including Microsoft Graph (Entra) and Azure Resource Management. IMPORTANT: For Graph API GET requests using advanced query parameters ($filter, $count, $search, $orderby), you are ADVISED to set 'consistencyLevel: \"eventual\"'.", {
    apiType: z.enum(["graph", "azure"]).describe("Type of Microsoft API to query. Options: 'graph' for Microsoft Graph (Entra) or 'azure' for Azure Resource Management."),
    path: z.string().describe("The Azure or Graph API URL path to call (e.g. '/users', '/groups', '/subscriptions')"),
    method: z.enum(["get", "post", "put", "patch", "delete"]).describe("HTTP method to use"),
    apiVersion: z.string().optional().describe("Azure Resource Management API version (required for apiType Azure)"),
    subscriptionId: z.string().optional().describe("Azure Subscription ID (for Azure Resource Management)."),
    queryParams: z.record(z.string()).optional().describe("Query parameters for the request"),
    body: z.record(z.string(), z.any()).optional().describe("The request body (for POST, PUT, PATCH)"),
    graphApiVersion: z.enum(["v1.0", "beta"]).optional().default(defaultGraphApiVersion).describe(`Microsoft Graph API version to use (default: ${defaultGraphApiVersion})`),
    fetchAll: z.boolean().optional().default(false).describe("Set to true to automatically fetch all pages for list results (e.g., users, groups). Default is false."),
    consistencyLevel: z.string().optional().describe("Graph API ConsistencyLevel header. ADVISED to be set to 'eventual' for Graph GET requests using advanced query parameters ($filter, $count, $search, $orderby)."),
}, async ({ apiType, path, method, apiVersion, subscriptionId, queryParams, body, graphApiVersion, fetchAll, consistencyLevel }) => {
    // Override graphApiVersion if USE_GRAPH_BETA is explicitly set to false
    const effectiveGraphApiVersion = !useGraphBeta ? "v1.0" : graphApiVersion;
    logger.info(`Executing Lokka-Microsoft tool with params: apiType=${apiType}, path=${path}, method=${method}, graphApiVersion=${effectiveGraphApiVersion}, fetchAll=${fetchAll}, consistencyLevel=${consistencyLevel}`);
    let determinedUrl;
    try {
        let responseData;
        // --- Microsoft Graph Logic ---
        if (apiType === 'graph') {
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
                        const firstPageResponse = await request.get();
                        const odataContext = firstPageResponse['@odata.context']; // Capture context from first page
                        let allItems = firstPageResponse.value || []; // Initialize with first page's items
                        // Callback function to process subsequent pages
                        const callback = (item) => {
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
                    }
                    else {
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
        } // --- Azure Resource Management Logic (using direct fetch) ---
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
            const headers = {
                'Authorization': `Bearer ${tokenResponse.token}`,
                'Content-Type': 'application/json'
            };
            const requestOptions = {
                method: method.toUpperCase(),
                headers: headers
            };
            if (["POST", "PUT", "PATCH"].includes(method.toUpperCase())) {
                requestOptions.body = body ? JSON.stringify(body) : JSON.stringify({});
            }
            // --- Pagination Logic for Azure RM (Manual Fetch) ---
            if (fetchAll && method === 'get') {
                logger.info(`Fetching all pages for Azure RM starting from: ${url}`);
                let allValues = [];
                let currentUrl = url;
                while (currentUrl) {
                    logger.info(`Fetching Azure RM page: ${currentUrl}`);
                    // Re-acquire token for each page (Azure tokens might expire)
                    const azureCredential = authManager.getAzureCredential();
                    const currentPageTokenResponse = await azureCredential.getToken("https://management.azure.com/.default");
                    if (!currentPageTokenResponse || !currentPageTokenResponse.token) {
                        throw new Error("Failed to acquire Azure access token during pagination");
                    }
                    const currentPageHeaders = { ...headers, 'Authorization': `Bearer ${currentPageTokenResponse.token}` };
                    const currentPageRequestOptions = { method: 'GET', headers: currentPageHeaders };
                    const pageResponse = await fetch(currentUrl, currentPageRequestOptions);
                    const pageText = await pageResponse.text();
                    let pageData;
                    try {
                        pageData = pageText ? JSON.parse(pageText) : {};
                    }
                    catch (e) {
                        logger.error(`Failed to parse JSON from Azure RM page: ${currentUrl}`, pageText);
                        pageData = { rawResponse: pageText };
                    }
                    if (!pageResponse.ok) {
                        logger.error(`API error on Azure RM page ${currentUrl}:`, pageData);
                        throw new Error(`API error (${pageResponse.status}) during Azure RM pagination on ${currentUrl}: ${JSON.stringify(pageData)}`);
                    }
                    if (pageData.value && Array.isArray(pageData.value)) {
                        allValues = allValues.concat(pageData.value);
                    }
                    else if (currentUrl === url && !pageData.nextLink) {
                        allValues.push(pageData);
                    }
                    else if (currentUrl !== url) {
                        logger.info(`[Warning] Azure RM response from ${currentUrl} did not contain a 'value' array.`);
                    }
                    currentUrl = pageData.nextLink || null; // Azure uses nextLink
                }
                responseData = { allValues: allValues };
                logger.info(`Finished fetching all Azure RM pages. Total items: ${allValues.length}`);
            }
            else {
                // Single page fetch for Azure RM
                logger.info(`Fetching single page for Azure RM: ${url}`);
                const apiResponse = await fetch(url, requestOptions);
                const responseText = await apiResponse.text();
                try {
                    responseData = responseText ? JSON.parse(responseText) : {};
                }
                catch (e) {
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
            content: [{ type: "text", text: resultText }],
        };
    }
    catch (error) {
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
});
// Add token management tools
server.tool("set-access-token", "Set or update the access token for Microsoft Graph authentication. Use this when the MCP Client has obtained a fresh token through interactive authentication.", {
    accessToken: z.string().describe("The access token obtained from Microsoft Graph authentication"),
    expiresOn: z.string().optional().describe("Token expiration time in ISO format (optional, defaults to 1 hour from now)")
}, async ({ accessToken, expiresOn }) => {
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
                        type: "text",
                        text: "Access token updated successfully. You can now make Microsoft Graph requests on behalf of the authenticated user."
                    }],
            };
        }
        else {
            return {
                content: [{
                        type: "text",
                        text: "Error: MCP Server is not configured for client-provided token authentication. Set USE_CLIENT_TOKEN=true in environment variables."
                    }],
                isError: true
            };
        }
    }
    catch (error) {
        logger.error("Error setting access token:", error);
        return {
            content: [{
                    type: "text",
                    text: `Error setting access token: ${error.message}`
                }],
            isError: true
        };
    }
});
server.tool("get-auth-status", "Check the current authentication status and mode of the MCP Server and also returns the current graph permission scopes of the access token for the current session.", {}, async () => {
    try {
        const authMode = authManager?.getAuthMode() || "Not initialized";
        const isReady = authManager !== null;
        const tokenStatus = authManager ? await authManager.getTokenStatus() : { isExpired: false };
        return {
            content: [{
                    type: "text",
                    text: JSON.stringify({
                        authMode,
                        isReady,
                        supportsTokenUpdates: authMode === AuthMode.ClientProvidedToken,
                        tokenStatus: tokenStatus,
                        timestamp: new Date().toISOString()
                    }, null, 2)
                }],
        };
    }
    catch (error) {
        return {
            content: [{
                    type: "text",
                    text: `Error checking auth status: ${error.message}`
                }],
            isError: true
        };
    }
});
// Add tool for requesting additional Graph permissions
server.tool("add-graph-permission", "Request additional Microsoft Graph permission scopes by performing a fresh interactive sign-in. This tool only works in interactive authentication mode and should be used if any Graph API call returns permissions related errors.", {
    scopes: z.array(z.string()).describe("Array of Microsoft Graph permission scopes to request (e.g., ['User.Read', 'Mail.ReadWrite', 'Directory.Read.All'])")
}, async ({ scopes }) => {
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
            }
            else if (currentMode === AuthMode.ClientProvidedToken) {
                errorMessage += `📋 To add permissions in Client Provided Token mode:\n`;
                errorMessage += `1. Obtain a new access token that includes the required scopes:\n`;
                errorMessage += `   ${scopes.map(scope => `• ${scope}`).join('\n   ')}\n`;
                errorMessage += `2. When obtaining the token, ensure these scopes are included in the consent prompt\n`;
                errorMessage += `3. Use the set-access-token tool to update the server with the new token\n`;
                errorMessage += `4. The new token will include the additional permissions`;
            }
            else {
                errorMessage += `To use interactive permission requests, set USE_INTERACTIVE=true in environment variables and restart the server.`;
            }
            return {
                content: [{
                        type: "text",
                        text: errorMessage
                    }],
                isError: true
            };
        }
        // Validate scopes array
        if (!scopes || scopes.length === 0) {
            return {
                content: [{
                        type: "text",
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
                        type: "text",
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
        try {
            // Try Interactive Browser first - create fresh instance each time
            newCredential = new InteractiveBrowserCredential({
                tenantId: tenantId,
                clientId: clientId,
                redirectUri: redirectUri,
            });
            // Request token immediately after creating credential
            tokenResponse = await newCredential.getToken(scopeString);
        }
        catch (error) {
            // Fallback to Device Code flow
            logger.info("Interactive browser failed, falling back to device code flow");
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
        }
        if (!tokenResponse) {
            return {
                content: [{
                        type: "text",
                        text: "Error: Failed to acquire access token with the requested scopes. Please check your permissions and try again."
                    }],
                isError: true
            };
        }
        // Create a completely new auth manager instance with the updated credential
        const authConfig = {
            mode: AuthMode.Interactive,
            tenantId,
            clientId,
            redirectUri
        };
        // Create a new auth manager instance
        authManager = new AuthManager(authConfig);
        // Manually set the credential to our new one with the additional scopes
        authManager.credential = newCredential;
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
                    type: "text",
                    text: JSON.stringify({
                        message: "Successfully acquired additional Microsoft Graph permissions with fresh authentication",
                        requestedScopes: scopes,
                        tokenStatus: tokenStatus,
                        note: "A fresh sign-in was performed to ensure the new permissions are properly granted",
                        timestamp: new Date().toISOString()
                    }, null, 2)
                }],
        };
    }
    catch (error) {
        logger.error("Error requesting additional Graph permissions:", error);
        return {
            content: [{
                    type: "text",
                    text: `Error requesting additional permissions: ${error.message}`
                }],
            isError: true
        };
    }
});
// Add this after the existing tools, before the main() function
server.tool("create-office-file", "Create Microsoft Office files (Word, Excel, PowerPoint) using a working method - upload a real template file via Graph API.", {
    fileType: z.enum(["word", "excel", "powerpoint"]).describe("Type of Office file to create"),
    fileName: z.string().describe("Name of the file (without extension, will be added automatically)"),
    location: z.enum(["sharepoint", "onedrive"]).describe("Where to create the file: SharePoint site or user's OneDrive"),
    sitePath: z.string().optional().describe("SharePoint site path (e.g., '/sites/TeamSite') - required for SharePoint location"),
    libraryName: z.string().optional().describe("Document library name (default: 'Shared Documents' for SharePoint, root for OneDrive)"),
    folderPath: z.string().optional().describe("Folder path within the library (e.g., '/Projects/Q4')"),
    initialContent: z.string().optional().default("Created by Lokka MCP").describe("Initial text content for the file"),
    userPrincipalName: z.string().optional().describe("User's UPN for OneDrive access (defaults to current user)")
}, async ({ fileType, fileName, location, sitePath, libraryName, folderPath, initialContent, userPrincipalName }) => {
    try {
        if (!graphClient) {
            throw new Error("Graph client not initialized");
        }
        logger.info(`Creating ${fileType} file '${fileName}' in ${location} using proven template method`);
        const fileExtensions = {
            word: "docx",
            excel: "xlsx",
            powerpoint: "pptx"
        };
        const fullFileName = `${fileName}.${fileExtensions[fileType]}`;
        // Build the target path for direct upload
        let targetPath;
        if (location === "sharepoint") {
            if (!sitePath) {
                throw new Error("sitePath is required for SharePoint location");
            }
            const library = libraryName || "Shared Documents";
            const folder = folderPath || "";
            targetPath = `/sites/${sitePath.replace(/^\/sites\//, '')}/drive/root:/${library}${folder}/${fullFileName}:/content`;
        }
        else {
            // OneDrive for Business
            const userPart = userPrincipalName ? `/users/${userPrincipalName}` : "/me";
            const folder = folderPath || "";
            targetPath = `${userPart}/drive/root:${folder}/${fullFileName}:/content`;
        }
        // CREATE FRESH EMPTY OFFICE FILES FOR ALL FORMATS
        let fileBuffer;
        const zip = new JSZip();
        if (fileType === "word") {
            // Get authenticated user info to avoid "sharepoint app" showing as creator
            let userDisplayName = "Lokka MCP Server";
            try {
                const userInfo = await graphClient.api("/me").select("displayName").get();
                userDisplayName = userInfo.displayName || userDisplayName;
            }
            catch (error) {
                logger.info("Could not retrieve user info for Word document, using default creator name");
            }
            // Create complete Word document structure
            zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>`);
            zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`);
            zip.file("word/_rels/document.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
</Relationships>`);
            zip.file("word/document.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Normal"/>
      </w:pPr>
      <w:r>
        <w:t>${initialContent || 'Created by Lokka MCP'}</w:t>
      </w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>`);
            zip.file("word/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:eastAsia="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
        <w:lang w:val="en-US" w:eastAsia="en-US" w:bidi="ar-SA"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault/>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
</w:styles>`);
            zip.file("word/settings.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:zoom w:percent="100"/>
  <w:defaultTabStop w:val="708"/>
  <w:characterSpacingControl w:val="doNotCompress"/>
</w:settings>`);
            zip.file("word/fontTable.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:font w:name="Calibri">
    <w:panose1 w:val="020F0502020204030204"/>
    <w:charset w:val="00"/>
    <w:family w:val="swiss"/>
    <w:pitch w:val="variable"/>
    <w:sig w:usb0="E00002FF" w:usb1="4000ACFF" w:usb2="00000001" w:usb3="00000000" w:csb0="0000019F" w:csb1="00000000"/>
  </w:font>
</w:fonts>`);
            zip.file("docProps/app.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>Microsoft Office Word</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0000</AppVersion>
</Properties>`);
            zip.file("docProps/core.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>${userDisplayName}</dc:creator>
  <cp:lastModifiedBy>${userDisplayName}</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString()}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${new Date().toISOString()}</dcterms:modified>
</cp:coreProperties>`);
            fileBuffer = await zip.generateAsync({ type: "nodebuffer" });
        }
        else if (fileType === "excel") {
            // Get authenticated user info to avoid "sharepoint app" showing as creator
            let userDisplayName = "Lokka MCP Server";
            try {
                const userInfo = await graphClient.api("/me").select("displayName").get();
                userDisplayName = userInfo.displayName || userDisplayName;
            }
            catch (error) {
                logger.info("Could not retrieve user info for Excel document, using default creator name");
            }
            // Create complete Excel workbook structure
            zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>`);
            zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`);
            zip.file("xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>`);
            zip.file("xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`);
            zip.file("xl/worksheets/sheet1.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetData>
    ${initialContent ? `<row r="1"><c r="A1" t="inlineStr"><is><t>${initialContent}</t></is></c></row>` : ''}
  </sheetData>
</worksheet>`);
            zip.file("xl/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1">
    <font>
      <sz val="11"/>
      <color theme="1"/>
      <name val="Calibri"/>
      <family val="2"/>
      <scheme val="minor"/>
    </font>
  </fonts>
  <fills count="2">
    <fill>
      <patternFill patternType="none"/>
    </fill>
    <fill>
      <patternFill patternType="gray125"/>
    </fill>
  </fills>
  <borders count="1">
    <border>
      <left/>
      <right/>
      <top/>
      <bottom/>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  </cellXfs>
</styleSheet>`);
            zip.file("xl/sharedStrings.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0">
</sst>`);
            zip.file("xl/theme/theme1.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="5B9BD5"/></a:accent1>
      <a:accent2><a:srgbClr val="70AD47"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="4472C4"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Calibri Light"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs>
            <a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:lin ang="5400000" scaled="0"/>
        </a:gradFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs>
            <a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:lin ang="5400000" scaled="0"/>
        </a:gradFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="6350" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
          <a:prstDash val="solid"/>
          <a:miter lim="800000"/>
        </a:ln>
        <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
          <a:prstDash val="solid"/>
          <a:miter lim="800000"/>
        </a:ln>
        <a:ln w="19050" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
          <a:prstDash val="solid"/>
          <a:miter lim="800000"/>
        </a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle>
          <a:effectLst/>
        </a:effectStyle>
        <a:effectStyle>
          <a:effectLst/>
        </a:effectStyle>
        <a:effectStyle>
          <a:effectLst>
            <a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">
              <a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr>
            </a:outerShdw>
          </a:effectLst>
        </a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs>
            <a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:lin ang="5400000" scaled="0"/>
        </a:gradFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>`);
            zip.file("docProps/app.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>Lokka MCP Server</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0000</AppVersion>
</Properties>`);
            zip.file("docProps/core.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>${userDisplayName}</dc:creator>
  <dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString()}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${new Date().toISOString()}</dcterms:modified>
</cp:coreProperties>`);
            fileBuffer = await zip.generateAsync({ type: "nodebuffer" });
        }
        else if (fileType === "powerpoint") {
            // Get authenticated user info to avoid "sharepoint app" showing as creator
            let userDisplayName = "Lokka MCP Server";
            try {
                const userInfo = await graphClient.api("/me").select("displayName").get();
                userDisplayName = userInfo.displayName || userDisplayName;
            }
            catch (error) {
                logger.info("Could not retrieve user info for PowerPoint document, using default creator name");
            }
            // Create complete PowerPoint presentation structure
            zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>`);
            zip.file("_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`);
            zip.file("ppt/_rels/presentation.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`);
            zip.file("ppt/slides/_rels/slide1.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`);
            zip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`);
            zip.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`);
            zip.file("ppt/presentation.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId2"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/>
  </p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000" type="screen4x3"/>
  <p:notesSz cx="6858000" cy="9144000"/>
  <p:defaultTextStyle>
    <a:defPPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <a:defRPr lang="en-US"/>
    </a:defPPr>
  </p:defaultTextStyle>
</p:presentation>`);
            zip.file("ppt/slides/slide1.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
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
      ${initialContent ? `<p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title 1"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="ctrTitle"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="914400" y="1828800"/>
            <a:ext cx="7315200" cy="1143000"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr anchor="ctr" rtlCol="0"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="ctr">
              <a:defRPr sz="4400" b="1" cap="none"/>
            </a:pPr>
            <a:r>
              <a:rPr lang="en-US" dirty="0" smtClean="0"/>
              <a:t>${initialContent}</a:t>
            </a:r>
            <a:endParaRPr lang="en-US" dirty="0"/>
          </a:p>
        </p:txBody>
      </p:sp>` : ''}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sld>`);
            zip.file("ppt/slideMasters/slideMaster1.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:bg>
      <p:bgRef idx="1001">
        <a:schemeClr val="bg1"/>
      </p:bgRef>
    </p:bg>
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
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483649" r:id="rId1"/>
  </p:sldLayoutIdLst>
  <p:txStyles>
    <p:titleStyle>
      <a:lvl1pPr algn="ctr" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
        <a:spcBef>
          <a:spcPct val="0"/>
        </a:spcBef>
        <a:buNone/>
        <a:defRPr sz="4400" kern="1200">
          <a:solidFill>
            <a:schemeClr val="tx1"/>
          </a:solidFill>
          <a:latin typeface="+mj-lt"/>
          <a:ea typeface="+mj-ea"/>
          <a:cs typeface="+mj-cs"/>
        </a:defRPr>
      </a:lvl1pPr>
    </p:titleStyle>
    <p:bodyStyle>
      <a:lvl1pPr marL="342900" indent="-342900" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">
        <a:spcBef>
          <a:spcPct val="20000"/>
        </a:spcBef>
        <a:buFont typeface="Arial" pitchFamily="34" charset="0"/>
        <a:buChar char="•"/>
        <a:defRPr sz="1800" kern="1200">
          <a:solidFill>
            <a:schemeClr val="tx1"/>
          </a:solidFill>
          <a:latin typeface="+mn-lt"/>
          <a:ea typeface="+mn-ea"/>
          <a:cs typeface="+mn-cs"/>
        </a:defRPr>
      </a:lvl1pPr>
    </p:bodyStyle>
    <p:otherStyle>
      <a:defPPr>
        <a:defRPr lang="en-US"/>
      </a:defPPr>
    </p:otherStyle>
  </p:txStyles>
</p:sldMaster>`);
            zip.file("ppt/slideLayouts/slideLayout1.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" type="title" preserve="1">
  <p:cSld name="Title Slide">
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
          <p:cNvPr id="2" name="Title 1"/>
          <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
          <p:nvPr><p:ph type="ctrTitle"/></p:nvPr>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="914400" y="1828800"/>
            <a:ext cx="7315200" cy="1143000"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr anchor="ctr" rtlCol="0"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="ctr">
              <a:defRPr sz="4400" b="1" cap="none"/>
            </a:pPr>
            <a:endParaRPr lang="en-US"/>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping/>
  </p:clrMapOvr>
</p:sldLayout>`);
            zip.file("ppt/theme/theme1.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="44546A"/></a:dk2>
      <a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
      <a:accent1><a:srgbClr val="5B9BD5"/></a:accent1>
      <a:accent2><a:srgbClr val="70AD47"/></a:accent2>
      <a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
      <a:accent4><a:srgbClr val="FFC000"/></a:accent4>
      <a:accent5><a:srgbClr val="4472C4"/></a:accent5>
      <a:accent6><a:srgbClr val="70AD47"/></a:accent6>
      <a:hlink><a:srgbClr val="0563C1"/></a:hlink>
      <a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Calibri Light"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs>
            <a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:lin ang="5400000" scaled="0"/>
        </a:gradFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs>
            <a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:lin ang="5400000" scaled="0"/>
        </a:gradFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="6350" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
          <a:prstDash val="solid"/>
          <a:miter lim="800000"/>
        </a:ln>
        <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
          <a:prstDash val="solid"/>
          <a:miter lim="800000"/>
        </a:ln>
        <a:ln w="19050" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
          <a:prstDash val="solid"/>
          <a:miter lim="800000"/>
        </a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle>
          <a:effectLst/>
        </a:effectStyle>
        <a:effectStyle>
          <a:effectLst/>
        </a:effectStyle>
        <a:effectStyle>
          <a:effectLst>
            <a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0">
              <a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr>
            </a:outerShdw>
          </a:effectLst>
        </a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs>
            <a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs>
            <a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs>
          </a:gsLst>
          <a:lin ang="5400000" scaled="0"/>
        </a:gradFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>`);
            zip.file("docProps/app.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
  <Application>Lokka MCP Server</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0000</AppVersion>
</Properties>`);
            zip.file("docProps/core.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>${userDisplayName}</dc:creator>
  <dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString()}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${new Date().toISOString()}</dcterms:modified>
</cp:coreProperties>`);
            fileBuffer = await zip.generateAsync({ type: "nodebuffer" });
        }
        else {
            throw new Error(`Unsupported file type: ${fileType}`);
        }
        logger.info(`Uploading ${fileType} file to: ${targetPath}`);
        // Upload the actual file content with proper content types
        const contentTypes = {
            word: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            excel: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            powerpoint: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
        };
        const uploadResponse = await graphClient
            .api(targetPath)
            .header('Content-Type', contentTypes[fileType])
            .put(fileBuffer);
        logger.info(`Successfully created and uploaded ${fileType} file: ${fullFileName}`);
        return {
            content: [{
                    type: "text",
                    text: JSON.stringify({
                        message: `Successfully created ${fileType} file using direct upload method`,
                        fileName: fullFileName,
                        location: location,
                        fileId: uploadResponse.id,
                        webUrl: uploadResponse.webUrl,
                        downloadUrl: uploadResponse["@microsoft.graph.downloadUrl"],
                        size: uploadResponse.size,
                        createdDateTime: uploadResponse.createdDateTime,
                        lastModifiedDateTime: uploadResponse.lastModifiedDateTime,
                        method: "Direct file upload with proper content",
                        note: `Fresh empty ${fileType.toUpperCase()} file created with proper OpenXML format`
                    }, null, 2)
                }]
        };
    }
    catch (error) {
        logger.error("Error creating Office file:", error);
        return {
            content: [{
                    type: "text",
                    text: `Error creating Office file: ${error.message}\n\nDetails: ${JSON.stringify(error, null, 2)}`
                }],
            isError: true
        };
    }
});
async function main() {
    // Determine authentication mode based on environment variables
    const useCertificate = process.env.USE_CERTIFICATE === 'true';
    const useInteractive = process.env.USE_INTERACTIVE === 'true';
    const useClientToken = process.env.USE_CLIENT_TOKEN === 'true';
    const initialAccessToken = process.env.ACCESS_TOKEN;
    let authMode;
    // Ensure only one authentication mode is enabled at a time
    const enabledModes = [
        useClientToken,
        useInteractive,
        useCertificate
    ].filter(Boolean);
    if (enabledModes.length > 1) {
        throw new Error("Multiple authentication modes enabled. Please enable only one of USE_CLIENT_TOKEN, USE_INTERACTIVE, or USE_CERTIFICATE.");
    }
    if (useClientToken) {
        authMode = AuthMode.ClientProvidedToken;
        if (!initialAccessToken) {
            logger.info("Client token mode enabled but no initial token provided. Token must be set via set-access-token tool.");
        }
    }
    else if (useInteractive) {
        authMode = AuthMode.Interactive;
    }
    else if (useCertificate) {
        authMode = AuthMode.Certificate;
    }
    else {
        // Check if we have client credentials environment variables
        const hasClientCredentials = process.env.TENANT_ID && process.env.CLIENT_ID && process.env.CLIENT_SECRET;
        if (hasClientCredentials) {
            authMode = AuthMode.ClientCredentials;
        }
        else {
            // Default to interactive mode for better user experience
            authMode = AuthMode.Interactive;
            logger.info("No authentication mode specified and no client credentials found. Defaulting to interactive mode.");
        }
    }
    logger.info(`Starting with authentication mode: ${authMode}`);
    // Get tenant ID and client ID with defaults only for interactive mode
    let tenantId;
    let clientId;
    if (authMode === AuthMode.Interactive) {
        // Interactive mode can use defaults
        tenantId = process.env.TENANT_ID || LokkaDefaultTenantId;
        clientId = process.env.CLIENT_ID || LokkaClientId;
        logger.info(`Interactive mode using tenant ID: ${tenantId}, client ID: ${clientId}`);
    }
    else {
        // All other modes require explicit values from environment variables
        tenantId = process.env.TENANT_ID;
        clientId = process.env.CLIENT_ID;
    }
    const clientSecret = process.env.CLIENT_SECRET;
    const certificatePath = process.env.CERTIFICATE_PATH;
    const certificatePassword = process.env.CERTIFICATE_PASSWORD; // optional
    // Validate required configuration
    if (authMode === AuthMode.ClientCredentials) {
        if (!tenantId || !clientId || !clientSecret) {
            throw new Error("Client credentials mode requires explicit TENANT_ID, CLIENT_ID, and CLIENT_SECRET environment variables");
        }
    }
    else if (authMode === AuthMode.Certificate) {
        if (!tenantId || !clientId || !certificatePath) {
            throw new Error("Certificate mode requires explicit TENANT_ID, CLIENT_ID, and CERTIFICATE_PATH environment variables");
        }
    }
    // Note: Client token mode can start without a token and receive it later
    const authConfig = {
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
    // Only initialize if we have required config (for client token mode, we can start without a token)
    if (authMode !== AuthMode.ClientProvidedToken || initialAccessToken) {
        await authManager.initialize();
        // Initialize Graph Client
        const authProvider = authManager.getGraphAuthProvider();
        graphClient = Client.initWithMiddleware({
            authProvider: authProvider,
        });
        logger.info(`Authentication initialized successfully using ${authMode} mode`);
    }
    else {
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
