#!/usr/bin/env node
import express, { Request, Response, NextFunction } from "express";
import { spawn, ChildProcess } from "child_process";
import { EventEmitter } from "events";
import { logger } from "./logger.js";

/**
 * REST API Wrapper for Lokka MCP Server
 *
 * This server acts as a bridge between HTTP REST clients (like Copilot Studio)
 * and the Lokka MCP server that uses stdio transport.
 */

interface MCPRequest {
	jsonrpc: "2.0";
	id: number | string;
	method: string;
	params?: any;
}

interface MCPResponse {
	jsonrpc: "2.0";
	id: number | string;
	result?: any;
	error?: {
		code: number;
		message: string;
		data?: any;
	};
}

interface MCPToolCallRequest {
	name: string;
	arguments: Record<string, any>;
}

interface GraphAPIRequest {
	path: string;
	method: "get" | "post" | "put" | "patch" | "delete";
	queryParams?: Record<string, string>;
	body?: any;
	graphApiVersion?: "v1.0" | "beta";
	fetchAll?: boolean;
	consistencyLevel?: string;
}

class MCPClient extends EventEmitter {
	private process: ChildProcess | null = null;
	private messageBuffer: string = "";
	private pendingRequests: Map<
		number | string,
		{
			resolve: (value: any) => void;
			reject: (error: any) => void;
		}
	> = new Map();
	private requestId = 0;
	private isInitialized = false;
	private accessToken: string | null = null;

	constructor(private mcpServerPath: string) {
		super();
	}

	async start(): Promise<void> {
		if (this.process) {
			logger.info("MCP client already started");
			return;
		}

		logger.info(`Starting MCP server from: ${this.mcpServerPath}`);

		// Spawn the MCP server process
		this.process = spawn("node", [this.mcpServerPath], {
			stdio: ["pipe", "pipe", "pipe"],
			env: {
				...process.env,
				// AUTH_MODE is inherited from parent process env
			},
		});

		// Handle stdout (MCP protocol messages)
		this.process.stdout!.on("data", (data: Buffer) => {
			this.handleOutput(data.toString());
		});

		// Handle stderr (logs)
		this.process.stderr!.on("data", (data: Buffer) => {
			const logMessage = data.toString().trim();
			if (logMessage) {
				logger.info(`MCP Server: ${logMessage}`);
			}
		});

		// Handle process exit
		this.process.on("exit", (code) => {
			logger.info(`MCP server process exited with code ${code}`);
			this.process = null;
			this.isInitialized = false;

			// Reject all pending requests
			for (const [id, promise] of this.pendingRequests) {
				promise.reject(new Error("MCP server process terminated"));
			}
			this.pendingRequests.clear();
		});

		// Initialize the MCP connection
		await this.initialize();
	}

	private async initialize(): Promise<void> {
		const response = await this.sendRequest({
			jsonrpc: "2.0",
			id: this.requestId++,
			method: "initialize",
			params: {
				protocolVersion: "2024-11-05",
				capabilities: {},
				clientInfo: {
					name: "Lokka-REST-Wrapper",
					version: "1.0.0",
				},
			},
		});

		if (response.error) {
			throw new Error(`Failed to initialize MCP: ${response.error.message}`);
		}

		this.isInitialized = true;
		logger.info("MCP client initialized successfully");
	}

	setAccessToken(token: string): void {
		this.accessToken = token;
		logger.info("Access token set for MCP client");
	}

	async callTool(name: string, args: Record<string, any>): Promise<any> {
		if (!this.isInitialized) {
			throw new Error("MCP client not initialized");
		}

		// If we have an access token and this is the first call, set it
		if (this.accessToken && name !== "set-access-token") {
			try {
				await this.sendRequest({
					jsonrpc: "2.0",
					id: this.requestId++,
					method: "tools/call",
					params: {
						name: "set-access-token",
						arguments: {
							accessToken: this.accessToken,
						},
					},
				});
				logger.info("Access token sent to MCP server");
			} catch (error) {
				logger.error("Failed to set access token in MCP server", error);
			}
		}

		const response = await this.sendRequest({
			jsonrpc: "2.0",
			id: this.requestId++,
			method: "tools/call",
			params: {
				name,
				arguments: args,
			},
		});

		if (response.error) {
			throw new Error(response.error.message);
		}

		return response.result;
	}

	async listTools(): Promise<any> {
		if (!this.isInitialized) {
			throw new Error("MCP client not initialized");
		}

		const response = await this.sendRequest({
			jsonrpc: "2.0",
			id: this.requestId++,
			method: "tools/list",
			params: {},
		});

		if (response.error) {
			throw new Error(response.error.message);
		}

		return response.result;
	}

	private handleOutput(data: string): void {
		this.messageBuffer += data;

		// Try to extract complete JSON-RPC messages
		const lines = this.messageBuffer.split("\n");
		this.messageBuffer = lines.pop() || ""; // Keep incomplete line in buffer

		for (const line of lines) {
			if (!line.trim()) continue;

			try {
				const message: MCPResponse = JSON.parse(line);
				this.handleMessage(message);
			} catch (error) {
				logger.error(`Failed to parse MCP message: ${line}`, error);
			}
		}
	}

	private handleMessage(message: MCPResponse): void {
		const pending = this.pendingRequests.get(message.id);
		if (pending) {
			this.pendingRequests.delete(message.id);
			if (message.error) {
				pending.reject(new Error(message.error.message));
			} else {
				pending.resolve(message);
			}
		} else {
			// This might be a notification or unsolicited message
			logger.info(
				`Received MCP message with no pending request: ${JSON.stringify(
					message
				)}`
			);
		}
	}

	private sendRequest(request: MCPRequest): Promise<MCPResponse> {
		return new Promise((resolve, reject) => {
			if (!this.process || !this.process.stdin) {
				reject(new Error("MCP server process not running"));
				return;
			}

			this.pendingRequests.set(request.id, { resolve, reject });

			const requestStr = JSON.stringify(request) + "\n";
			this.process.stdin.write(requestStr, (error) => {
				if (error) {
					this.pendingRequests.delete(request.id);
					reject(error);
				}
			});

			// Set a timeout for the request
			setTimeout(() => {
				if (this.pendingRequests.has(request.id)) {
					this.pendingRequests.delete(request.id);
					reject(new Error("MCP request timeout"));
				}
			}, 30000); // 30 second timeout
		});
	}

	async stop(): Promise<void> {
		if (this.process) {
			this.process.kill();
			this.process = null;
			this.isInitialized = false;
			logger.info("MCP client stopped");
		}
	}
}

// Create Express app
const app = express();
app.use(express.json());

// CORS middleware
app.use((req: Request, res: Response, next: NextFunction) => {
	res.header("Access-Control-Allow-Origin", "*");
	res.header(
		"Access-Control-Allow-Methods",
		"GET, POST, PUT, PATCH, DELETE, OPTIONS"
	);
	res.header(
		"Access-Control-Allow-Headers",
		"Content-Type, Authorization, X-Access-Token"
	);

	if (req.method === "OPTIONS") {
		res.sendStatus(200);
	} else {
		next();
	}
});

// Logging middleware
app.use((req: Request, res: Response, next: NextFunction) => {
	logger.info(`${req.method} ${req.path}`);
	next();
});

// Initialize MCP client
const mcpServerPath = process.env.MCP_SERVER_PATH || "./build/main.js";
const mcpClient = new MCPClient(mcpServerPath);

// Authentication middleware - extract token from header or body
function extractToken(req: Request): string | null {
	// Try Authorization header first
	const authHeader = req.headers.authorization;
	if (authHeader && authHeader.startsWith("Bearer ")) {
		return authHeader.substring(7);
	}

	// Try X-Access-Token header
	const tokenHeader = req.headers["x-access-token"];
	if (tokenHeader && typeof tokenHeader === "string") {
		return tokenHeader;
	}

	// Try body token
	if (req.body && req.body.accessToken) {
		return req.body.accessToken;
	}

	return null;
}

// Health check endpoint
app.get("/health", (req: Request, res: Response) => {
	res.json({
		status: "ok",
		service: "Lokka REST API Wrapper",
		version: "1.0.0",
		timestamp: new Date().toISOString(),
	});
});

// Manifest endpoint - MCP server capabilities
app.get("/manifest.json", async (req: Request, res: Response) => {
	try {
		const tools = await mcpClient.listTools();

		res.json({
			name: "Lokka Microsoft Graph MCP Server",
			version: "0.4.0",
			description:
				"REST API wrapper for Lokka MCP server - Microsoft Graph and Entra ID integration",
			author: "Thiago Beier",
			capabilities: {
				tools: tools.tools || [],
				transport: "http",
			},
			endpoints: {
				health: "/health",
				manifest: "/manifest.json",
				tools: {
					list: "/api/mcp/tools/list",
					call: "/api/mcp/tools/call",
				},
				graph: {
					recent_groups: "/api/graph/groups/recent",
					generic: "/api/graph/*",
				},
				auth: {
					set_token: "/api/auth/token",
				},
			},
			openapi: "/openapi.yaml",
		});
	} catch (error: any) {
		logger.error("Error generating manifest", error);
		res.status(500).json({
			error: "Failed to generate manifest",
			message: error.message,
		});
	}
});

// List available MCP tools
app.get("/api/mcp/tools/list", async (req: Request, res: Response) => {
	try {
		const tools = await mcpClient.listTools();
		res.json(tools);
	} catch (error: any) {
		logger.error("Error listing MCP tools", error);
		res.status(500).json({
			error: "Failed to list MCP tools",
			message: error.message,
		});
	}
});

// Set access token endpoint
app.post("/api/auth/token", async (req: Request, res: Response) => {
	try {
		const { accessToken } = req.body;

		if (!accessToken) {
			res.status(400).json({ error: "accessToken is required" });
			return;
		}

		mcpClient.setAccessToken(accessToken);

		res.json({
			success: true,
			message: "Access token set successfully",
		});
	} catch (error: any) {
		logger.error("Error setting access token", error);
		res.status(500).json({
			error: "Failed to set access token",
			message: error.message,
		});
	}
});

// Microsoft Graph API endpoints

// List all groups created in the past week
app.get("/api/graph/groups/recent", async (req: Request, res: Response) => {
	try {
		const token = extractToken(req);
		if (token) {
			mcpClient.setAccessToken(token);
		}

		const daysBack = parseInt(req.query.daysBack as string) || 7;
		const now = new Date();
		const thresholdDate = new Date(
			now.getTime() - daysBack * 24 * 60 * 60 * 1000
		);

		// Fetch first page of groups (Graph API doesn't support filtering by createdDateTime)
		// Note: We limit to first page to avoid memory issues with large tenants
		const result = await mcpClient.callTool("Lokka-Microsoft", {
			apiType: "graph",
			path: "/groups",
			method: "get",
			queryParams: {
				$select:
					"id,displayName,createdDateTime,groupTypes,mail,mailEnabled,securityEnabled,description",
				$top: "100",
			},
			graphApiVersion: "beta",
		});

		// Filter client-side for groups created within the specified days
		if (result.content && Array.isArray(result.content)) {
			const textContent = result.content.find((c: any) => c.type === "text");
			if (textContent && textContent.text) {
				try {
					// Extract JSON from "Result for..." message if present
					let jsonText = textContent.text;
					const jsonStart = jsonText.indexOf("{");
					if (jsonStart > 0) {
						jsonText = jsonText.substring(jsonStart);
					}

					// Remove any trailing text after the JSON (like "Note: More results...")
					// Find the last closing brace
					const jsonEnd = jsonText.lastIndexOf("}");
					if (jsonEnd > 0 && jsonEnd < jsonText.length - 1) {
						jsonText = jsonText.substring(0, jsonEnd + 1);
					}

					const data = JSON.parse(jsonText);
					if (data.value && Array.isArray(data.value)) {
						const filteredGroups = data.value.filter((group: any) => {
							if (!group.createdDateTime) return false;
							const createdDate = new Date(group.createdDateTime);
							return createdDate >= thresholdDate;
						});

						// Return in the same format
						result.content[0].text = JSON.stringify({
							"@odata.context": data["@odata.context"],
							value: filteredGroups,
						});
					}
				} catch (parseError: any) {
					logger.error("Failed to parse MCP response", {
						text: textContent.text,
						error: parseError.message,
					});
					// Return the raw response if we can't parse it
					res.json(result);
					return;
				}
			}
		}

		res.json(result);
	} catch (error: any) {
		logger.error("Error fetching recent groups", error);
		res.status(500).json({
			error: "Failed to fetch recent groups",
			message: error.message,
		});
	}
});

// Generic Graph API endpoint
app.all("/api/graph/*", async (req: Request, res: Response) => {
	try {
		const token = extractToken(req);
		if (token) {
			mcpClient.setAccessToken(token);
		}

		const path = "/" + req.params[0];
		const method = req.method.toLowerCase() as
			| "get"
			| "post"
			| "put"
			| "patch"
			| "delete";

		// Filter out internal parameters that shouldn't be sent to Graph API
		const internalParams = new Set([
			"fetchAll",
			"graphApiVersion",
			"consistencyLevel",
		]);
		const queryParams: Record<string, string> = {};
		for (const [key, value] of Object.entries(req.query)) {
			if (!internalParams.has(key) && typeof value === "string") {
				queryParams[key] = value;
			}
		}

		const graphRequest: GraphAPIRequest = {
			path,
			method,
			queryParams,
			graphApiVersion: (req.query.graphApiVersion as "v1.0" | "beta") || "beta",
			fetchAll: req.query.fetchAll === "true",
			consistencyLevel: req.query.consistencyLevel as string,
		};

		if (method !== "get" && req.body) {
			graphRequest.body = req.body;
		}

		const result = await mcpClient.callTool("Lokka-Microsoft", {
			apiType: "graph",
			...graphRequest,
		});

		res.json(result);
	} catch (error: any) {
		logger.error("Error calling Graph API", error);
		res.status(500).json({
			error: "Failed to call Graph API",
			message: error.message,
		});
	}
});

// Generic MCP tool call endpoint
app.post("/api/mcp/tools/call", async (req: Request, res: Response) => {
	try {
		const token = extractToken(req);
		if (token) {
			mcpClient.setAccessToken(token);
		}

		const { name, arguments: args } = req.body as MCPToolCallRequest;

		if (!name) {
			res.status(400).json({ error: "Tool name is required" });
			return;
		}

		const result = await mcpClient.callTool(name, args || {});
		res.json(result);
	} catch (error: any) {
		logger.error("Error calling MCP tool", error);
		res.status(500).json({
			error: "Failed to call MCP tool",
			message: error.message,
		});
	}
});

// Error handling middleware
app.use((err: Error, req: Request, res: Response, next: NextFunction) => {
	logger.error("Unhandled error", err);
	res.status(500).json({
		error: "Internal server error",
		message: err.message,
	});
});

// Start the server
const PORT = process.env.PORT || 3000;

async function startServer() {
	try {
		// Start MCP client
		await mcpClient.start();
		logger.info("MCP client started successfully");

		// Start Express server
		app.listen(PORT, () => {
			logger.info(`Lokka REST API Wrapper listening on port ${PORT}`);
			logger.info(`Health check: http://localhost:${PORT}/health`);
			logger.info(`API endpoints:`);
			logger.info(
				`  - GET  /api/graph/groups/recent - List groups created in the past week`
			);
			logger.info(`  - ALL  /api/graph/* - Generic Graph API endpoint`);
			logger.info(`  - POST /api/mcp/tools/call - Generic MCP tool call`);
			logger.info(`  - POST /api/auth/token - Set access token`);
		});
	} catch (error) {
		logger.error("Failed to start server", error);
		process.exit(1);
	}
}

// Graceful shutdown
process.on("SIGTERM", async () => {
	logger.info("SIGTERM received, shutting down gracefully");
	await mcpClient.stop();
	process.exit(0);
});

process.on("SIGINT", async () => {
	logger.info("SIGINT received, shutting down gracefully");
	await mcpClient.stop();
	process.exit(0);
});

// Handle uncaught errors
process.on("uncaughtException", (error) => {
	logger.error("Uncaught exception:", error);
	process.exit(1);
});

process.on("unhandledRejection", (reason, promise) => {
	logger.error(`Unhandled rejection at: ${promise} reason: ${reason}`);
	process.exit(1);
});

// Log startup
logger.info("Starting Lokka REST API Wrapper...");
logger.info(`Node version: ${process.version}`);
logger.info(`Working directory: ${process.cwd()}`);
logger.info(`PORT: ${process.env.PORT}`);
logger.info(`AUTH_MODE: ${process.env.AUTH_MODE}`);
logger.info(`MCP_SERVER_PATH: ${process.env.MCP_SERVER_PATH}`);

// Start the server
startServer();
