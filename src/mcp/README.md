# Lokka - MCP Server for Azure and Microsoft Graph

Lokka is an MCP server for querying and managing your Azure and Microsoft 365 tenants using the Microsoft Azure/Graph APIs. It acts as a bridge between the Microsoft APIs and any compatible MCP client, allowing you to interact with your Azure and Microsoft 365 tenant using natural language queries.

## Sample queries

Here are some examples of queries you can use with Lokka.

- `Create a new security group called 'Sales and HR' with a dynamic rule based on the department attribute.`
- `Find all the conditional access policies that haven't excluded the emergency access account`
- `Show me all the device configuration policies assigned to the 'Call center' group`

## What is Lokka?

Lokka is designed to be used with any compatible MCP client, such as Claude Desktop, Cursor, Goose, or any other AI model and client that support the Model Context Protocol. It provides a simple and intuitive way to manage your Azure and Microsoft 365 tenant using natural language queries.

Follow the guide at [Lokka.dev](https://lokka.dev) to get started with Lokka and learn how to use it with your favorite AI model and chat client.

![How does Lokka work?](https://github.com/merill/lokka/blob/main/website/docs/assets/how-does-lokka-mcp-server-work.png?raw=true)

## MCP Client Configuration

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
