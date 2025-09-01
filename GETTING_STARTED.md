# Getting Started with Enhanced Outlook MCP Server

This guide will walk you through the necessary steps to set up and run the Enhanced Outlook MCP Server. The server uses OAuth 2.0 to authenticate with the Microsoft Graph API, which requires you to register an application in the Microsoft Azure portal.

## Prerequisites

- A Microsoft account (e.g., Outlook.com, Office 365, or Azure)
- Node.js and npm installed on your machine

## Steps to Obtain Credentials

To get your `MS_CLIENT_ID` and `MS_CLIENT_SECRET`, you need to register an application with the Microsoft identity platform.

### 1. Register a New Application

1.  Go to the [Azure portal](https://portal.azure.com/) and sign in.
2.  Search for and select **Azure Active Directory**.
3.  Under **Manage**, select **App registrations** > **New registration**.
4.  Enter a **Name** for your application (e.g., `EnhancedOutlookMCP`).
5.  For **Supported account types**, select **Accounts in any organizational directory (Any Azure AD directory - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)**.
6.  Under **Redirect URI (optional)**, select **Web** and enter `http://localhost:3000/auth/callback`. This must match the `redirectUri` in `config.js`.
7.  Click **Register**.

### 2. Get the Client ID

After registration, you'll be taken to the application's **Overview** page.

- The **Application (client) ID** is your `MS_CLIENT_ID`. Copy this value.

### 3. Create a Client Secret

1.  In your app registration, go to **Certificates & secrets**.
2.  Click **New client secret**.
3.  Add a description for your secret and select an expiration period.
4.  Click **Add**.
5.  **Immediately copy the secret's value.** This is your `MS_CLIENT_SECRET`. You will not be able to see it again after you leave this page.

## Project Setup

1.  **Clone the repository**:
    ```bash
    git clone <repository-url>
    cd enhanced-outlook-mcp
    ```

2.  **Install dependencies**:
    ```bash
    npm install
    ```

3.  **Configure your environment variables**:

    Copy the example environment file to a new `.env` file in the project root:

    ```bash
    cp .env.example .env
    ```

    Then, open the new `.env` file and fill in the `MS_CLIENT_ID` and `MS_CLIENT_SECRET` values you obtained from the Azure portal.

## Running the Server with Claude

This server is designed to be started and managed by the Claude desktop application via a StdIO command.

1.  **Configure the Claude Desktop App**:
    - Open your Claude desktop configuration file. On macOS, this is located at `~/Library/Application Support/Claude/claude_desktop_config.json`.
    - Add the following entry to the `mcpServers` object, ensuring you replace `/path/to/project` with the **absolute path** to this project's directory on your machine.

    ```json
    "enhanced-outlook-mcp": {
      "command": "/path/to/project/start.sh"
    }
    ```

    For example:
    ```json
    {
      "mcpServers": {
        "enhanced-outlook-mcp": {
          "command": "/Users/pragone/code/personal/enhanced-outlook-mcp/start.sh"
        }
      }
    }
    ```

2.  **Make the Start Script Executable**:
    For Claude to be able to run the server, the `start.sh` script must be executable. Open your terminal, navigate to the project's root directory, and run:
    ```bash
    chmod +x start.sh
    ```

3.  **Authenticate**:
    You do not need to start the server manually. Claude will launch it automatically when you use one of its tools.

    To authenticate, simply ask Claude to use one of the tools (e.g., "list my emails"). The server will provide an authentication URL. Open this URL in your browser, sign in with your Microsoft account, and grant the requested permissions. The server will then be authenticated and ready to use.
