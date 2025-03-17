---
title: 🧩 Developer guide
sidebar_position: 3
---

Follow this guide if you want to build Lokka from source to contribute to the project.

## Pre-requisites

- Follow the [installation guide](https://lokka.dev/docs/installation) to install Node and create the Entra application.
- Clone the Lokka repository from GitHub [https://github.com/merill/lokka](https://github.com/merill/lokka)

## Building the project

- Open a terminal and navigate to the Lokka project directory.
- Change into the folder `\src\mcp\`
- Run the following command to install the dependencies:

  ```bash
  npm install
  ```

- After the dependencies are installed, run the following command to build the project:

  ```bash
  npm run build
  ```
- When the build is complete, you will see a main.js file find the compiled files in the `\src\mcp\build\` folder.

## Configuring the agent

### Claude Desktop

- In Claude Desktop, open the settings by clicking on the hamburger icon in the top left corner.
- Select **File** > **Settings** (or press `Ctrl + ,`)
- In the **Developer** tab, click **Edit Config**
- This opens explorer, edit `claude_desktop_config.json` in your favorite text editor.
- Add the following configuration to the file, using the information you in the **Overview** blade of the Entra application you created earlier.

- Note: On Windows the path needs to be escaped with `\\` or use `/` instead of `\`.
  - E.g. `C:\\Users\\<username>\\Documents\\lokka\\src\\mcp\\build\\main.js` or `C:/Users/<username>/Documents/lokka/src/mcp/build/main.js`

```json
{
  "mcpServers": {
      "lokka": {
          "command": "node",
          "args": [
              "<absolute-path-to-main.js>/src/mcp/build/main.js"
          ],
          "env": {
            "TENANT_ID": "<tenant-id>",
            "CLIENT_ID": "<client-id>",
            "CLIENT_SECRET": "<client-secret>"
          }
      }
  }
}
```

- Exit Claude Desktop and restart it.
  - Every time you make changes to the code or configuration, you need to restart Claude desktop for the changes to take effect.
  - Note: In Windows, Claude doesn't exit when you close the window, it runs in the background. You can find it in the system tray. Right-click on the icon and select **Quit** to exit the application completely.

### Testing the agent

- Open the Claude Desktop application.
- In the chat window on the bottom right you should see a hammer icon if the configuration is correct.
- Now you can start quering your Microsoft tenant using the Lokka agent tool.
- Some sample queries you can try are:
  - `Get all the users in my tenant`
  - `Show me the details for John Doe`
  - `Change John's department to IT` - Needs User.ReadWrite.All permission to be granted

## Learn about MCP

- [Model Context Protocol Tutorial by Matt Pocock](https://www.aihero.dev/model-context-protocol-tutorial) - This is a great tutorial that explains the Model Context Protocol and how to use it.
- [Model Context Protocol docs](https://modelcontextprotocol.io/introduction) - This is the official docs for the Model Context Protocol.
- [Model Context Protocol Clients](https://modelcontextprotocol.io/clients) - This is a list of all the clients that support the Model Context Protocol.
