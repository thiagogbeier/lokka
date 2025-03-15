import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";

// Create server instance
const server = new McpServer({
  name: "lokka",
  version: "0.1.0",
});

interface GraphResponse<T> {
  value: T[];
  "@odata.nextLink"?: string;
}

// Helper function to get Microsoft Graph access token
async function getGraphAccessToken(): Promise<string> {
  // Check if environment variables are set
  const tenantId = process.env.MS_GRAPH_TENANT_ID;
  const clientId = process.env.MS_GRAPH_CLIENT_ID;
  const clientSecret = process.env.MS_GRAPH_CLIENT_SECRET;

  if (!tenantId || !clientId || !clientSecret) {
    throw new Error("Microsoft Graph credentials not configured. Set MS_GRAPH_TENANT_ID, MS_GRAPH_CLIENT_ID, and MS_GRAPH_CLIENT_SECRET environment variables.");
  }

  const tokenUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: clientId,
    scope: 'https://graph.microsoft.com/.default',
    client_secret: clientSecret,
    grant_type: 'client_credentials'
  });

  try {
    const response = await fetch(tokenUrl, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      body: body
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`Failed to get access token: ${response.status} ${errorText}`);
    }

    const data = await response.json();
    return data.access_token;
  } catch (error) {
    console.error('Error acquiring Graph token:', error);
    throw new Error('Failed to authenticate with Microsoft Graph');
  }
}

// Specialized helper function for advanced Graph API querying with paging support
async function makeAdvancedGraphRequest<T>(
  path: string,
  method: string = 'GET',
  version: string = 'beta',
  queryParams: Record<string, string> = {},
  body?: any,
  allData: boolean = false
): Promise<T> {
  const token = await getGraphAccessToken();

  // Use the specified API version
  const apiBase = `https://graph.microsoft.com/${version}`;

  // Build query string with consistency level
  const params = { ...queryParams };

  // Build query string
  const queryString = Object.entries(params)
    .map(([key, value]) => `${key}=${encodeURIComponent(value)}`)
    .join('&');

  let url = `${apiBase}${path}${queryString ? '?' + queryString : ''}`;

  const headers: Record<string, string> = {
    'Authorization': `Bearer ${token}`,
    'Accept': 'application/json',
    'ConsistencyLevel': 'eventual'  // Include consistency level in all requests
  };

  // Add Content-Type header for requests with body
  if (body) {
    headers['Content-Type'] = 'application/json';
  }

  // Process body parameter to ensure it's correctly formatted
  let processedBody: any = body;
  let bodyForLogging: any = body;

  // Fix for handling stringified JSON body
  if (body) {
    console.error(`Original body type: ${typeof body}`);
    
    if (typeof body === 'string') {
      try {
        // If it's a string, try to parse it as JSON
        const parsedBody = JSON.parse(body);
        bodyForLogging = parsedBody;
        // Use the parsed object as the processed body
        processedBody = parsedBody;
        console.error('Successfully parsed body string to object');
      } catch (e) {
        // If parsing fails, keep the original string
        console.error('Failed to parse body as JSON, using as is:', e);
        // For logging, show the string
        bodyForLogging = body;
      }
    }
    
    console.error(`Processed body type: ${typeof processedBody}`);
  }

  try {
    // For storing results when allData is true
    let completeResults: any[] = [];
    let response: Response;
    let responseData: any;

    // Create request info object for logging
    const requestInfo = {
      url,
      method,
      version,
      path,
      queryParams,
      headers: { ...headers, 'Authorization': 'REDACTED' }, // Redact auth token in logs
      body: bodyForLogging, // Use the logging-friendly version of the body
      bodyType: typeof processedBody,
      timestamp: new Date().toISOString()
    };

    // First request
    console.error(`Making ${method} request to ${url}`);
    const requestStartTime = Date.now();
    response = await fetch(url, {
      method,
      headers,
      body: body
    });
    const requestDuration = Date.now() - requestStartTime;

    let responseBody: any;
    let responseContentType = response.headers.get('content-type') || '';

    if (responseContentType.includes('application/json')) {
      responseBody = await response.json();
      // Clone the response for further use
      responseData = JSON.parse(JSON.stringify(responseBody));
    } else {
      responseBody = await response.text();
      responseData = responseBody;
    }

    // Create response info object for logging
    const responseInfo = {
      status: response.status,
      statusText: response.statusText,
      headers: Object.fromEntries([...response.headers.entries()]),
      body: responseBody,
      duration: requestDuration,
      timestamp: new Date().toISOString()
    };


    if (!response.ok) {
      throw new Error(`Graph API error: ${response.status} ${response.statusText}`);
    }

    // For methods like PATCH that may not return content
    if ((method === 'PATCH' || method === 'PUT') && response.status === 204) {
      return {} as T;
    }

    // If not collecting all data or response doesn't have a 'value' array, return directly
    if (!allData || !responseData.value) {
      return responseData as T;
    }

    // Start collecting data for paging
    completeResults = [...responseData.value];

    // Continue fetching if there's a next link and allData is true
    let pageCount = 1;
    while (allData && responseData['@odata.nextLink']) {
      pageCount++;
      url = responseData['@odata.nextLink'];

      // Create paging request info
      const pagingRequestInfo = {
        url,
        method: 'GET',  // Next link requests are always GET
        headers: { ...headers, 'Authorization': 'REDACTED' },
        pageNumber: pageCount,
        timestamp: new Date().toISOString()
      };

      // Next page request
      console.error(`Fetching page ${pageCount}: ${url}`);
      const pageRequestStartTime = Date.now();
      response = await fetch(url, {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${token}`,
          'Accept': 'application/json',
          'ConsistencyLevel': 'eventual'
        }
      });
      const pageDuration = Date.now() - pageRequestStartTime;

      if (response.headers.get('content-type')?.includes('application/json')) {
        responseBody = await response.json();
        responseData = JSON.parse(JSON.stringify(responseBody));
      } else {
        responseBody = await response.text();
        responseData = responseBody;
      }

      // Create paging response info
      const pagingResponseInfo = {
        status: response.status,
        statusText: response.statusText,
        headers: Object.fromEntries([...response.headers.entries()]),
        body: responseBody,
        pageNumber: pageCount,
        duration: pageDuration,
        timestamp: new Date().toISOString()
      };

  
      if (!response.ok) {
        throw new Error(`Graph API paging error: ${response.status} ${response.statusText}`);
      }

      if (Array.isArray(responseData.value)) {
        completeResults = [...completeResults, ...responseData.value];
      }
    }

    // Return an object with the same structure but with all collected results
    const finalResult = {
      value: completeResults,
      // Include any other properties from the last response except the nextLink
      ...Object.fromEntries(
        Object.entries(responseData).filter(([key]) => key !== '@odata.nextLink' && key !== 'value')
      ),
      totalPages: pageCount,
      totalRecords: completeResults.length
    } as T;

    return finalResult;

  } catch (error) {
    console.error("Error making advanced Graph request:", error);

    throw error;
  }
}

// Register Microsoft Graph generic query tool
server.tool(
  "graph-query",
  `Execute flexible Microsoft Graph API queries with support for multiple HTTP methods, API versions, and pagination 
  `,
  {
    path: z.string().describe("Graph API path to call (e.g. /users, /groups)"),
    method: z.enum(['GET', 'POST', 'PUT', 'PATCH']).default('GET').describe("HTTP method to use"),
    version: z.enum(['v1.0', 'beta']).default('beta').describe("Microsoft Graph API version"),
    queryParams: z.record(z.string()).optional().describe("Query parameters to include in the request"),
    body: z.any().optional().describe("Request body for POST, PUT, or PATCH requests"),
    allData: z.boolean().default(false).describe("Whether to automatically retrieve all pages of data")
  },
  async ({ path, method, version, queryParams, body, allData }) => {
    try {
      // Ensure path starts with /
      if (!path.startsWith('/')) {
        path = '/' + path;
      }
      
      // Debug log the body content and type
      if (body) {
        console.error(`Body received in graph-query: ${typeof body}`);
        console.error(`Body content: ${JSON.stringify(body).substring(0, 200)}`);
      }

      const response = await makeAdvancedGraphRequest<any>(
        path,
        method,
        version,
        queryParams || {},
        body,
        allData
      );

      // Format the response for display
      let resultText: string;

      if (response.value && Array.isArray(response.value)) {
        // For collection responses
        resultText = `Found ${response.value.length} result${response.value.length === 1 ? '' : 's'} for ${method} ${path}:\n\n`;
        resultText += JSON.stringify(response, null, 2);
      } else {
        // For single entity responses
        resultText = `Result for ${method} ${path}:\n\n`;
        resultText += JSON.stringify(response, null, 2);
      }

      return {
        content: [
          {
            type: "text",
            text: resultText,
          },
        ],
      };
    } catch (error) {
      console.error("Error in graph-query:", error);
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';

      return {
        content: [
          {
            type: "text",
            text: `Error querying Microsoft Graph: ${errorMessage}`,
          },
        ],
      };
    }
  },
);

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error("Weather MCP Server running on stdio");
}

main().catch((error) => {
  console.error("Fatal error in main():", error);
  process.exit(1);
});