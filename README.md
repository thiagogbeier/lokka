# Lokka

Lokka is a model-context-protocol tool for querying and managing your Microsoft tenant with AI.

Follow this guide to get started with Lokka.

## Pre-requisites

- Install [Node.js](https://nodejs.org/en/download/)
- Clone the Lokka repository from GitHub [https://github.com/merill/lokka](https://github.com/merill/lokka)

### Create an Entra Application

- Open [Entra admin center](https://entra.microsoft.com) > **Identity** > **Applications** > **App registrations**
  - Tip: [enappreg.cmd.ms](https://enappreg.cmd.ms) is a shortcut to the App registrations page.
- Select **New registration**
- Enter a name for the application (e.g. `Lokka Agent Tool`)
- Select **Register**

### Grant permissions to Microsoft Graph

- Open the application you created in the previous step
- Select **API permissions** > **Add a permission**
- Select **Microsoft Graph** > **Application permissions**
- Search for each of the permissions and check the box next to each permission you want to allow.
  - The agent will only be able to perform the actions based on the permissions you grant it.
- Select **Add permissions**
- Select **Grant admin consent for [your organization]**
- Select **Yes** to confirm

### Create a client secret

- Select **Certificates & secrets** > **Client secrets** > **New client secret**
- Enter a description for the secret (e.g. `Agent Config`)
- Select **Add**
- Copy the value of the secret, we will use this value in the agent configuration file.

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

## Configuring the agent

Now you can use the Lokka agent tool with any compatible MCP client. See [MCP clients](https://modelcontextprotocol.io/clients) for a list of compatible clients.

In the example below, we'll use the Claude Desktop client.

### Install Claude Desktop

- Download the latest version of Claude Desktop from [https://claude.ai/download](https://claude.ai/download)
- Install the application by following the instructions on the website.
- Open the application and sign in with your account.

## Create a configuration file
### Create a configuration file