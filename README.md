# Lokka

[![npm version](https://badge.fury.io/js/@merill%2Flokka.svg)](https://badge.fury.io/js/@merill%2Flokka)

Lokka is a model-context-protocol server for the Microsoft Graph API that allows you to query and managing your Microsoft tenant with AI.

<img src="https://github.com/merill/lokka/blob/main/assets/lokka-demo-1.gif?raw=true" alt="Lokka Demo - user create demo" width="500"/>

Please see [Lokka.dev](https://lokka.dev) for how to use Lokka with your favorite AI model and chat client.

Lokka lets you use Claude Desktop, or any MCP Client, to use natural language to accomplish things in your Microsoft 365 tenant through the Microsoft Graph API.

e.g.:

- `Create a new security group called 'Sales and HR' with a dynamic rule based on the department attribute.` 
- `Find all the conditional access policies that haven't excluded the emergency access account`
- `Show me all the Intune device configuration policies assigned to the 'Call center' group`

![How does Lokka work?](https://github.com/merill/lokka/blob/main/website/docs/assets/how-does-lokka-mcp-server-work.png?raw=true)

## Getting started

See the docs for more information on how to install and configure Lokka.

- [Introduction](https://lokka.dev/)
- [Install guide](https://lokka.dev/docs/installation)
- [Developer guide](https://lokka.dev/docs/developer-guide)

## Components

### Tools

1. `Lokka-MicrosoftGraph`
   - Call Microsoft Graph API. Supports querying a Microsoft 365 tenant using the Graph API. Updates are also supported if permissions are provided.
   - Input:
     - `path` (string): The Graph API URL path to call (e.g. '/me', '/users', '/groups')
     - `method` (string): HTTP method to use (e.g., get, post, put, patch, delete)
     - `queryParams` (string): Array of query parameters like $filter, $select, etc. All parameters are strings.
     - `body` (JSON): The request body (for POST, PUT, PATCH)
   - Returns: Results from the Graph API call.


### Environment Variables

The configuration of the server is done using environment variables. The following environment variables are required:

| Name | Description |
|------|-------------|
| `TENANT_ID` | The ID of the Microsoft Entra tenant. |
| `CLIENT_ID` | The ID of the application registered in Microsoft Entra. |
| `CLIENT_SECRET` | The client secret of the application registered in Microsoft Entra. |

## Installation

To use this server with the Claude Desktop app, add the following configuration to the "mcpServers" section of your
`claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "Lokka-Microsoft-Graph": {
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

Make sure to replace `<tenant-id>`, `<client-id>`, and `<client-secret>` with the actual values from your Microsoft Entra application. (See [Install Guide](https://lokka.dev/docs/installation) for more details on how to create an Entra app and configure the agent.)
