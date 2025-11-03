#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
// Create server instance
const server = new McpServer({
    name: "Lokka-Microsoft-Simple",
    version: "1.0.0",
});
// Add a simple test tool
server.tool("test-connection", "Simple test to verify MCP server is working", {}, async () => {
    return {
        content: [{
                type: "text",
                text: "✅ Lokka MCP Server is working! Connection successful."
            }],
    };
});
async function main() {
    try {
        console.error("Starting Lokka Simple MCP Server...");
        const transport = new StdioServerTransport();
        await server.connect(transport);
        console.error("✅ Server connected successfully!");
    }
    catch (error) {
        console.error("❌ Fatal error:", error);
        process.exit(1);
    }
}
main();
