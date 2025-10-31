#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import { Client, PageIterator } from "@microsoft/microsoft-graph-client";
import fetch from 'isomorphic-fetch'; // Required polyfill for Graph client
import { logger } from "./logger.js";
import { AuthManager, AuthMode } from "./auth.js";
import { LokkaClientId, LokkaDefaultTenantId, LokkaDefaultRedirectUri, getDefaultGraphApiVersion } from "./constants.js";
// Set up global fetch for the Microsoft Graph client
global.fetch = fetch;
// Create server instance
const server = new McpServer({
    name: "Lokka-Microsoft",
    version: "0.3.0",
});
logger.info("Starting Lokka Multi-Microsoft API MCP Server (v0.3.0)");
// Initialize authentication and clients
let authManager = null;
let graphClient = null;
// Check USE_GRAPH_BETA environment variable
const useGraphBeta = process.env.USE_GRAPH_BETA !== 'false'; // Default to true unless explicitly set to 'false'
const defaultGraphApiVersion = getDefaultGraphApiVersion();
logger.info(`Graph API default version: ${defaultGraphApiVersion} (USE_GRAPH_BETA=${process.env.USE_GRAPH_BETA || 'undefined'})`);
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
// Add tool for querying recently created groups
server.tool("get-recent-groups", "Query groups created within a specified time period (default: past week). This is a convenience tool that uses Microsoft Graph's filter capabilities to find recently created groups.", {
    daysAgo: z.number().optional().default(7).describe("Number of days to look back from today (default: 7 days for past week)"),
    fetchAll: z.boolean().optional().default(true).describe("Set to true to automatically fetch all pages. Default is true."),
    includeAllFields: z.boolean().optional().default(false).describe("Set to true to include all group properties. Default is false (returns id, displayName, createdDateTime, description, groupTypes, mail)."),
}, async ({ daysAgo, fetchAll, includeAllFields }) => {
    try {
        logger.info(`Querying groups created in the past ${daysAgo} days`);
        if (!graphClient) {
            throw new Error("Graph client not initialized");
        }
        // Calculate the date threshold (daysAgo days ago from now)
        const thresholdDate = new Date();
        thresholdDate.setTime(thresholdDate.getTime() - (daysAgo * 24 * 60 * 60 * 1000));
        const isoThresholdDate = thresholdDate.toISOString();
        // Build the filter query
        const filterQuery = `createdDateTime ge ${isoThresholdDate}`;
        // Build query parameters
        const queryParams = {
            '$filter': filterQuery,
            '$orderby': 'createdDateTime desc'
        };
        // Add select clause to limit fields if not requesting all fields
        if (!includeAllFields) {
            queryParams['$select'] = 'id,displayName,createdDateTime,description,groupTypes,mail';
        }
        // Use the Graph SDK to query groups
        let request = graphClient.api('/groups')
            .version(useGraphBeta ? "beta" : "v1.0")
            .query(queryParams)
            .header('ConsistencyLevel', 'eventual');
        let responseData;
        if (fetchAll) {
            logger.info(`Fetching all pages for groups created after ${isoThresholdDate}`);
            // Fetch the first page
            const firstPageResponse = await request.get();
            const odataContext = firstPageResponse['@odata.context'];
            let allItems = firstPageResponse.value || [];
            // Create a PageIterator to get all pages
            const callback = (item) => {
                allItems.push(item);
                return true;
            };
            const pageIterator = new PageIterator(graphClient, firstPageResponse, callback);
            await pageIterator.iterate();
            responseData = {
                '@odata.context': odataContext,
                value: allItems,
                count: allItems.length
            };
            logger.info(`Found ${allItems.length} groups created in the past ${daysAgo} days`);
        }
        else {
            logger.info(`Fetching single page for groups created after ${isoThresholdDate}`);
            responseData = await request.get();
            responseData.count = responseData.value?.length || 0;
        }
        // Format the result
        const dayLabel = daysAgo === 1 ? 'day' : 'days';
        let resultText = `Groups created in the past ${daysAgo} ${dayLabel}:\n\n`;
        resultText += `Query: ${filterQuery}\n`;
        resultText += `Threshold Date: ${isoThresholdDate}\n`;
        resultText += `Total Groups Found: ${responseData.count}\n\n`;
        resultText += JSON.stringify(responseData, null, 2);
        if (!fetchAll && responseData['@odata.nextLink']) {
            resultText += `\n\nNote: More results are available. Set 'fetchAll: true' to retrieve all pages.`;
        }
        return {
            content: [{ type: "text", text: resultText }],
        };
    }
    catch (error) {
        logger.error(`Error querying recent groups:`, error);
        const errorBody = error.body ? (typeof error.body === 'string' ? error.body : JSON.stringify(error.body)) : 'N/A';
        return {
            content: [{
                    type: "text",
                    text: JSON.stringify({
                        error: error instanceof Error ? error.message : String(error),
                        statusCode: error.statusCode || 'N/A',
                        errorBody: errorBody,
                        suggestion: "Ensure the authenticated user/app has the required permissions to read groups (Group.Read.All or Group.ReadWrite.All)"
                    }),
                }],
            isError: true
        };
    }
});
// Start the server with stdio transport
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
