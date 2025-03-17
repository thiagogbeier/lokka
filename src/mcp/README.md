# Lokka

Lokka is an MCP server for querying and managing your Microsoft 365 tenant using the Microsoft Graph API. It acts as a bridge between the Microsoft Graph API and any compatible MCP client, allowing you to interact with your Microsoft 365 tenant using natural language queries.

Lokka is designed to be used with any compatible MCP client, such as Claude Desktop, Cursor, Goose, or any other AI model and client that support the Model Context Protocol. It provides a simple and intuitive way to manage your Microsoft 365 tenant using natural language queries.

Follow this guide to get started with Lokka.

## Pre-requisites

- Install [Node.js](https://nodejs.org/en/download/)

### Create an Entra Application

- Open [Entra admin center](https://entra.microsoft.com) > **Identity** > **Applications** > **App registrations**
  - Tip: [enappreg.cmd.ms](https://enappreg.cmd.ms) is a shortcut to the App registrations page.
- Select **New registration**
- Enter a name for the application (e.g. `Lokka`)
- Select **Register**

### Grant permissions to Microsoft Graph

- Open the application you created in the previous step
- Select **API permissions** > **Add a permission**
- Select **Microsoft Graph** > **Application permissions**
- Search for each of the permissions and check the box next to each permission you want to allow.
  - The agent will only be able to perform the actions based on the permissions you grant it.
  - We recommend starting with User.Read.All to be able to query information about users in your tenant. To learn more about the other permissions you can grant, see [Microsoft Graph permissions reference](https://learn.microsoft.com/en-us/graph/permissions-reference).
- Select **Add permissions**
- Select **Grant admin consent for [your organization]**
- Select **Yes** to confirm

### Create a client secret

- Select **Certificates & secrets** > **Client secrets** > **New client secret**
- Enter a description for the secret (e.g. `Agent Config`)
- Select **Add**
- Copy the value of the secret, we will use this value in the agent configuration file.

## Configuring the agent tool

Now you can use Lokka with any compatible MCP client. See [MCP clients](https://modelcontextprotocol.io/clients) for a list of compatible clients.

In the example below, we'll use the Claude Desktop client. You can use Claude for free but you will be limited to a certain number of queries per day. If you get the Claude monthly plan you get a larger number of queries that you can use per day.

### Install Claude Desktop

- Download the latest version of Claude Desktop from [https://claude.ai/download](https://claude.ai/download)
- Install the application by following the instructions on the website.
- Open the application and sign in with your account.

### Configure Lokka

- In Claude Desktop, open the settings by clicking on the hamburger icon in the top left corner.
- Select **File** > **Settings** (or press `Ctrl + ,`)
- In the **Developer** tab, click **Edit Config**
- This opens explorer, edit `claude_desktop_config.json` in your favorite text editor.
- Add the following configuration to the file, using the information you in the **Overview** blade of the Entra application you created earlier.

- Note: On Windows the path needs to be escaped with `\\` or use `/` instead of `\`.
  - E.g. `C:\\Users\\<username>\\Documents\\lokka\\src\\mcp\\build\\index.js` or `C:/Users/<username>/Documents/lokka/src/mcp/build/index.js`

```json
{
  "mcpServers": {
    "lokka": {
      "command": "npx",
      "args": [
        "-y",
        "@merill/lokka"
      ],
      "env": {
        "MS_GRAPH_TENANT_ID": "<tenant_id>",
        "MS_GRAPH_CLIENT_ID": "<client_id>",
        "MS_GRAPH_CLIENT_SECRET": "<client_secret>",
      }
    }
  }
}
```

- Exit Claude Desktop and restart it.
  - Every time you make changes to the code or configuration, you need to restart Claude desktop for the changes to take effect.
  - In Windows, Claude doesn't exit when you close the window, it runs in the background. You can find it in the system tray. Right-click on the icon and select **Quit** to exit the application completely.

### Testing the agent

- Open the Claude Desktop application.
- In the chat window on the bottom right you should see a hammer icon if the configuration is correct.
- Now you can start quering your Microsoft tenant using the Lokka agent tool.
- Some sample queries you can try are:
  - `Get all users`
  - `Show me the details for John Doe`
  - `Change John's department to IT` - Needs User.ReadWrite.All permission to be granted
- If the agent is not using graph to query the tenant, you can explicitly tell it to use Lokka or tell it to use microsoft graph.