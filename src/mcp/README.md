# Lokka - MCP Server for Azure and Microsoft Graph

Lokka is an MCP server for querying and managing your Azure and Microsoft 365 tenants using the Microsoft Azure/Graph APIs. It acts as a bridge between the Microsoft APIs and any compatible MCP client, allowing you to interact with your Azure and Microsoft 365 tenant using natural language queries.

## Sample queries

Here are some examples of queries you can use with Lokka.

- `Create a new security group called 'Sales and HR' with a dynamic rule based on the department attribute.`
- `Find all the conditional access policies that haven't excluded the emergency access account`
- `Show me all the device configuration policies assigned to the 'Call center' group`
- `What was the most expensive service in Azure last month?`

## What is Lokka?

Lokka is designed to be used with any compatible MCP client, such as Claude Desktop, Cursor, Goose, or any other AI model and client that support the Model Context Protocol. It provides a simple and intuitive way to manage your Azure and Microsoft 365 tenant using natural language queries.

Follow the guide at [Lokka.dev](https://lokka.dev) to get started with Lokka and learn how to use it with your favorite AI model and chat client.

![How does Lokka work?](https://github.com/merill/lokka/blob/main/website/docs/assets/how-does-lokka-mcp-server-work.png?raw=true)

## Authentication Methods *(Enhanced in v0.2.0)*

Lokka now supports multiple authentication methods to accommodate different deployment scenarios:

### 1. Client Credentials (Service-to-Service)
Traditional app-only authentication using client credentials:

```json
{
  "mcpServers": {
    "Lokka-Microsoft": {
      "command": "npx",
      "args": ["-y", "@merill/lokka"],
      "env": {
        "TENANT_ID": "<tenant-id>",
        "CLIENT_ID": "<client-id>",
        "CLIENT_SECRET": "<client-secret>"
      }
    }
  }
}
```

### 2. Interactive Authentication *(New)*
User-based authentication with interactive login:

```json
{
  "mcpServers": {
    "Lokka-Microsoft": {
      "command": "npx",
      "args": ["-y", "@merill/lokka"],
      "env": {
        "TENANT_ID": "<tenant-id>",
        "CLIENT_ID": "<client-id>",
        "USE_INTERACTIVE": "true",
        "REDIRECT_URI": "http://localhost:3000"
      }
    }
  }
}
```

### 3. Client-Provided Token *(New)*
Token-based authentication where the MCP Client provides access tokens:

```json
{
  "mcpServers": {
    "Lokka-Microsoft": {
      "command": "npx",
      "args": ["-y", "@merill/lokka"],
      "env": {
        "USE_CLIENT_TOKEN": "true"
      }
    }
  }
}
```

When using client-provided token mode:
1. Start the MCP server with `USE_CLIENT_TOKEN=true`
2. Use the `set-access-token` tool to provide a valid Microsoft Graph access token
3. Use the `get-auth-status` tool to verify authentication status
4. Refresh tokens as needed using `set-access-token`

## New Tools *(v0.2.0)*

### Token Management Tools
- **`set-access-token`**: Set or update access tokens for Microsoft Graph authentication
- **`get-auth-status`**: Check current authentication status and capabilities

### Enhanced Microsoft Graph Tool
- **`Lokka-Microsoft`**: Now supports all three authentication modes with improved error handling and token management

## MCP Client Configuration

For backward compatibility, the original client credentials configuration still works:

```json
{
  "mcpServers": {
    "Lokka-Microsoft": {
      "command": "npx",
      "args": ["-y", "@merill/lokka"],
      "env": {
        "TENANT_ID": "<tenant-id>",
        "CLIENT_ID": "<client-id>",
        "CLIENT_SECRET": "<client-secret>"
      }
    }
  }
}
```

## Get started

See the docs for more information on how to install and configure Lokka.

- [Introduction](https://lokka.dev/docs/intro)
- [Install guide](https://lokka.dev/docs/installation)
- [Developer guide](https://lokka.dev/docs/developer-guide)

## Contributors

- Interactive and Token-based Authentication (v0.2.0) - [@darrenjrobinson](https://github.com/darrenjrobinson)
