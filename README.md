[![MseeP.ai Security Assessment Badge](https://mseep.net/pr/aanerud-mcp-microsoft-office-badge.png)](https://mseep.ai/app/aanerud-mcp-microsoft-office)

# MCP Microsoft Office

**Connect Claude (or any AI) to your Microsoft 365 account**

Give AI assistants the ability to read your emails, manage your calendar, access your files, send Teams messages, and more - all through a secure, multi-user server that you control.

---

## What Does This Project Do?

This project creates a bridge between AI assistants (like Claude) and Microsoft 365. When you ask Claude "What meetings do I have tomorrow?" or "Send an email to John about the project update" - this system makes it happen.

**Key Benefits:**

- **70 Tools** - Email, Calendar, Files, Teams, Contacts, To-Do, and more
- **Multi-User** - One server can support your entire team, each with their own data
- **Your Control** - Run locally on your computer or deploy to your own server
- **Secure** - All tokens encrypted, no data stored on third-party servers
- **Works with Any MCP Client** - Claude Desktop, or any other MCP-compatible AI

---

## How It Works (The Simple Version)

```
┌─────────────────┐                    ┌─────────────────┐                    ┌─────────────────┐
│                 │    "Send email"    │                 │   "Here's the     │                 │
│  Claude Desktop │ ◄────────────────► │   MCP Adapter   │   email data"     │   MCP Server    │
│  (Your AI)      │                    │ (On Your PC)    │ ◄───────────────► │   (Local or     │
│                 │                    │                 │                    │    Remote)      │
└─────────────────┘                    └─────────────────┘                    └────────┬────────┘
                                                                                       │
                                                                                       │ Talks to
                                                                                       │ Microsoft
                                                                                       ▼
                                                                             ┌─────────────────┐
                                                                             │  Microsoft 365  │
                                                                             │  (Your Account) │
                                                                             └─────────────────┘
```

**Three Parts:**

1. **Claude Desktop** - The AI you chat with
2. **MCP Adapter** - A small program that runs on your computer (translates what Claude asks into web requests)
3. **MCP Server** - Handles security and talks to Microsoft 365 (can run on your PC or a remote server)

---

## Why This Architecture?

**Q: Why not connect Claude directly to Microsoft?**

A: The Model Context Protocol (MCP) requires a local adapter to translate between Claude and any service. By separating the adapter from the server, you get:

- **Flexibility**: Run the server locally for personal use, or deploy it for your whole team
- **Security**: Your Microsoft credentials never leave your server
- **Multi-User**: Multiple people can authenticate separately and use the same server
- **Any AI Client**: The adapter pattern works with any MCP-compatible AI, not just Claude

---

## Quick Start Guide

### Prerequisites

Before you begin, you'll need:

- **Node.js 18+** - [Download here](https://nodejs.org/)
- **Claude Desktop** - [Download here](https://claude.ai/download)
- **Azure App Registration** - Free, instructions below
- **Microsoft 365 Account** - Work, school, or personal

---

### Step 1: Create Azure App Registration

This tells Microsoft that your server is allowed to access your data.

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to **Microsoft Entra ID** → **App registrations**
3. Click **+ New registration**
4. Fill in:
   - **Name**: `MCP-Microsoft-Office` (or whatever you like)
   - **Supported account types**: Choose based on your needs
   - **Redirect URI**: Leave blank for now
5. Click **Register**
6. **Copy these values** (you'll need them later):
   - Application (client) ID
   - Directory (tenant) ID

#### Add API Permissions

1. Go to **API permissions** → **+ Add a permission**
2. Select **Microsoft Graph** → **Delegated permissions**
3. Add these permissions:

| Permission | What It's For |
|------------|---------------|
| `User.Read` | Read your profile |
| `Mail.ReadWrite` | Read and send emails |
| `Mail.Send` | Send emails |
| `Calendars.ReadWrite` | Manage calendar |
| `Files.ReadWrite` | Access OneDrive files |
| `People.Read` | Find contacts |
| `Tasks.ReadWrite` | Manage To-Do lists |
| `Contacts.ReadWrite` | Manage contacts |
| `Group.Read.All` | Read groups |
| `Chat.ReadWrite` | Teams chat access |
| `ChannelMessage.Send` | Send Teams messages |

4. If you're an admin, click **Grant admin consent**

#### Configure Authentication

1. Go to **Authentication** → **+ Add a platform**
2. Select **Web**
3. Add Redirect URI:
   - For local: `http://localhost:3000/api/auth/callback`
   - For remote: `https://your-server.example.com/api/auth/callback`
4. Under **Advanced settings**, set **Allow public client flows** to **Yes**
5. Click **Save**

---

### Step 2: Set Up the Server

#### Option A: Run Locally (Recommended for Getting Started)

```bash
# Clone the project
git clone https://github.com/Aanerud/MCP-Microsoft-Office.git
cd MCP-Microsoft-Office

# Install dependencies (this also sets up the database)
npm install

# Edit the .env file with your Azure app details
# Open .env and add:
# MICROSOFT_CLIENT_ID=your-client-id-here
# MICROSOFT_TENANT_ID=your-tenant-id-here

# Start the server
npm run dev:web
```

Your server is now running at `http://localhost:3000`

#### Option B: Use a Remote Server

If someone has deployed an MCP server for your team, you just need:
- The server URL (e.g., `https://your-server.example.com`)
- Skip to Step 3

---

### Step 3: Authenticate with Microsoft

1. Open your browser and go to your server:
   - Local: `http://localhost:3000`
   - Remote: `https://your-server.example.com`
2. Click **Login with Microsoft**
3. Sign in with your Microsoft account
4. Grant the requested permissions
5. You'll be redirected back to the server

---

### Step 4: Get Your MCP Token

After logging in:

1. Click **Generate MCP Token** (or find it in the setup section)
2. Copy the token - it looks like a long string starting with `eyJ...`
3. Keep this token safe - it's your key to accessing the server

---

### Step 5: Set Up the MCP Adapter

The adapter is a small file that Claude Desktop uses to communicate with the server.

#### On macOS

```bash
# Create a folder for MCP adapters
mkdir -p ~/mcp-adapters

# Copy the adapter (from the project you cloned, or download from your server)
cp /path/to/MCP-Microsoft-Office/mcp-adapter.cjs ~/mcp-adapters/
```

#### On Windows

```powershell
# Create a folder for MCP adapters
mkdir %USERPROFILE%\mcp-adapters

# Copy the adapter
copy C:\path\to\MCP-Microsoft-Office\mcp-adapter.cjs %USERPROFILE%\mcp-adapters\
```

---

### Step 6: Configure Claude Desktop

Claude Desktop needs to know about your MCP adapter.

#### On macOS

Edit: `~/Library/Application Support/Claude/claude_desktop_config.json`

```json
{
  "mcpServers": {
    "microsoft365": {
      "command": "node",
      "args": ["/Users/YOUR_USERNAME/mcp-adapters/mcp-adapter.cjs"],
      "env": {
        "MCP_SERVER_URL": "http://localhost:3000",
        "MCP_BEARER_TOKEN": "paste-your-token-here"
      }
    }
  }
}
```

**Replace:**
- `YOUR_USERNAME` with your macOS username
- `paste-your-token-here` with the token from Step 4
- Change `MCP_SERVER_URL` if using a remote server

#### On Windows

Edit: `%APPDATA%\Claude\claude_desktop_config.json`

```json
{
  "mcpServers": {
    "microsoft365": {
      "command": "node",
      "args": ["C:\\Users\\YOUR_USERNAME\\mcp-adapters\\mcp-adapter.cjs"],
      "env": {
        "MCP_SERVER_URL": "http://localhost:3000",
        "MCP_BEARER_TOKEN": "paste-your-token-here"
      }
    }
  }
}
```

---

### Step 7: Restart Claude Desktop

1. Quit Claude Desktop completely
2. Start it again
3. You should see the Microsoft 365 tools available

**Test it:** Ask Claude "What emails do I have?" or "What's on my calendar today?"

---

## Understanding the Token System

This project uses **two different tokens** for security:

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                         TOKEN TYPES                                          │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                             │
│  ┌─────────────────────────┐         ┌─────────────────────────┐           │
│  │   MCP Bearer Token      │         │   Microsoft Graph Token │           │
│  │   (You manage this)     │         │   (Server manages this) │           │
│  ├─────────────────────────┤         ├─────────────────────────┤           │
│  │ • Lasts 24h to 30 days  │         │ • Lasts 1 hour          │           │
│  │ • Goes in Claude config │         │ • Auto-refreshed        │           │
│  │ • Identifies YOU        │         │ • Talks to Microsoft    │           │
│  └───────────┬─────────────┘         └───────────┬─────────────┘           │
│              │                                   │                          │
│              ▼                                   ▼                          │
│     Claude ←→ Adapter ←→ Server         Server ←→ Microsoft 365            │
│                                                                             │
└─────────────────────────────────────────────────────────────────────────────┘
```

**MCP Bearer Token** (the one you copied):
- Proves to the server that requests are from you
- You put this in Claude's configuration
- If it expires, generate a new one from the web UI

**Microsoft Graph Token** (handled automatically):
- The server uses this to talk to Microsoft
- Automatically refreshed - you never see it
- Stored encrypted on the server

---

## Multi-User Support

This server can support multiple users at once, each with completely separate data:

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                        ONE SERVER, MANY USERS                                │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                             │
│  Alice (alice@company.com)           │  Bob (bob@company.com)               │
│  ├─ Her own Microsoft tokens         │  ├─ His own Microsoft tokens         │
│  ├─ Her own session                  │  ├─ His own session                  │
│  ├─ Her own activity logs            │  ├─ His own activity logs            │
│  │                                   │  │                                   │
│  └─ Claude Desktop (her laptop)      │  └─ Claude Desktop (his PC)          │
│                                                                             │
│  ═══════════════════════════════════════════════════════════════════════   │
│                      COMPLETE DATA ISOLATION                                 │
│                 Alice can NEVER see Bob's data                              │
│                 Bob can NEVER see Alice's data                              │
└─────────────────────────────────────────────────────────────────────────────┘
```

**How it works:**
- Each user logs in with their own Microsoft account
- Each user gets their own MCP token
- All data is tagged with the user's identity
- The database enforces isolation at every query

---

## Available Tools (70 Total)

### Email (9 tools)
| Tool | Description |
|------|-------------|
| `getInbox` | Read your inbox messages |
| `sendEmail` | Send an email |
| `searchEmails` | Search for specific emails |
| `flagEmail` | Flag/unflag an email |
| `getEmailDetails` | Get full email content |
| `markAsRead` | Mark email as read/unread |
| `getMailAttachments` | Get email attachments |
| `addMailAttachment` | Add attachment to email |
| `removeMailAttachment` | Remove attachment from email |

### Calendar (13 tools)
| Tool | Description |
|------|-------------|
| `getEvents` | Get calendar events |
| `createEvent` | Create a new meeting |
| `updateEvent` | Modify an existing event |
| `cancelEvent` | Cancel/delete an event |
| `getAvailability` | Check free/busy times |
| `findMeetingTimes` | Find optimal meeting slots |
| `acceptEvent` | Accept a meeting invite |
| `declineEvent` | Decline a meeting invite |
| `tentativelyAcceptEvent` | Tentatively accept |
| `getCalendars` | List all calendars |
| `getRooms` | Find meeting rooms |
| `addAttachment` | Add attachment to event |
| `removeAttachment` | Remove event attachment |

### Files (11 tools)
| Tool | Description |
|------|-------------|
| `listFiles` | List OneDrive files |
| `searchFiles` | Search for files |
| `downloadFile` | Download a file |
| `uploadFile` | Upload a new file |
| `getFileMetadata` | Get file info |
| `getFileContent` | Read file contents |
| `setFileContent` | Write file contents |
| `updateFileContent` | Update existing file |
| `createSharingLink` | Create share link |
| `getSharingLinks` | List share links |
| `removeSharingPermission` | Remove sharing |

### Teams (12 tools)
| Tool | Description |
|------|-------------|
| `listChats` | List Teams chats |
| `getChat` | Get chat details |
| `listChatMessages` | Read chat messages |
| `sendChatMessage` | Send a chat message |
| `listTeams` | List your teams |
| `getTeam` | Get team details |
| `listChannels` | List team channels |
| `getChannel` | Get channel details |
| `listChannelMessages` | Read channel messages |
| `sendChannelMessage` | Post to a channel |
| `createOnlineMeeting` | Create Teams meeting |
| `getOnlineMeeting` | Get meeting details |

### People (3 tools)
| Tool | Description |
|------|-------------|
| `find` | Find people by name |
| `search` | Search directory |
| `getRelevantPeople` | Get frequent contacts |

### Search (1 tool)
| Tool | Description |
|------|-------------|
| `search` | Unified search across Microsoft 365 |

### To-Do (11 tools)
| Tool | Description |
|------|-------------|
| `listTaskLists` | List all task lists |
| `getTaskList` | Get a specific list |
| `createTaskList` | Create new list |
| `updateTaskList` | Rename a list |
| `deleteTaskList` | Delete a list |
| `listTasks` | List tasks in a list |
| `getTask` | Get task details |
| `createTask` | Create a new task |
| `updateTask` | Update a task |
| `deleteTask` | Delete a task |
| `completeTask` | Mark task complete |

### Contacts (6 tools)
| Tool | Description |
|------|-------------|
| `listContacts` | List your contacts |
| `getContact` | Get contact details |
| `createContact` | Add new contact |
| `updateContact` | Update contact info |
| `deleteContact` | Remove a contact |
| `searchContacts` | Search contacts |

### Groups (4 tools)
| Tool | Description |
|------|-------------|
| `listGroups` | List Microsoft 365 groups |
| `getGroup` | Get group details |
| `listGroupMembers` | List group members |
| `listMyGroups` | List groups you're in |

---

## Environment Variables

Configure these in your `.env` file:

| Variable | Required | Description | Default |
|----------|----------|-------------|---------|
| `MICROSOFT_CLIENT_ID` | Yes | Azure App Client ID | - |
| `MICROSOFT_TENANT_ID` | Yes | Azure Tenant ID | `common` |
| `MICROSOFT_REDIRECT_URI` | No | OAuth callback URL | `http://localhost:3000/api/auth/callback` |
| `PORT` | No | Server port | `3000` |
| `NODE_ENV` | No | Environment mode | `development` |
| `MCP_TOKEN_SECRET` | No | Secret for MCP tokens | Auto-generated |
| `MCP_TOKEN_EXPIRY` | No | Token expiry in seconds | `2592000` (30 days) |
| `DATABASE_TYPE` | No | Database type | `sqlite` |

---

## Troubleshooting

### "AADSTS7000218: client_assertion or client_secret required"

**Problem:** Azure thinks you need a client secret.

**Fix:**
1. Go to Azure Portal → Your App → Authentication
2. Under "Advanced settings", set **Allow public client flows** to **Yes**
3. Click Save

### "Needs administrator approval"

**Problem:** Your organization requires admin consent for the permissions.

**Fix:**
- Ask your IT admin to grant consent, OR
- Use a personal Microsoft account for testing

### "Invalid redirect URI"

**Problem:** The callback URL doesn't match exactly.

**Fix:**
1. Go to Azure Portal → Your App → Authentication
2. Check that the Redirect URI matches exactly:
   - Local: `http://localhost:3000/api/auth/callback`
   - Remote: `https://your-server.example.com/api/auth/callback`

### "Connection refused" or "ECONNREFUSED"

**Problem:** The server isn't running.

**Fix:**
1. Make sure you started the server: `npm run dev:web`
2. Check the server is on the correct port
3. Check firewall settings

### "401 Unauthorized"

**Problem:** Your MCP token expired.

**Fix:**
1. Go to the web UI
2. Log in again if needed
3. Generate a new MCP token
4. Update Claude Desktop's config with the new token
5. Restart Claude Desktop

### Claude doesn't show Microsoft 365 tools

**Fix:**
1. Make sure the config file is valid JSON (no trailing commas!)
2. Check the adapter path is correct for your OS
3. Make sure Node.js is installed: `node --version`
4. Restart Claude Desktop completely

---

## Security

- **Encrypted Storage**: All Microsoft tokens are encrypted at rest using AES-256
- **No Client Secrets**: Uses public client flow (safer for desktop apps)
- **Token Isolation**: Each user's tokens are stored separately and encrypted with different keys
- **Session Expiry**: Sessions automatically expire after 24 hours
- **HTTPS**: Use HTTPS for production deployments

---

## For Developers

### Project Structure

```
MCP-Microsoft-Office/
├── mcp-adapter.cjs          # The adapter that runs locally
├── src/
│   ├── api/                 # Express routes and controllers
│   ├── auth/                # MSAL authentication
│   ├── core/                # Services (cache, events, storage)
│   ├── graph/               # Microsoft Graph API services
│   └── modules/             # Feature modules (mail, calendar, etc.)
├── public/                  # Web UI
└── data/                    # SQLite database (created on first run)
```

### Running Tests

```bash
npm test
```

### Development Mode

```bash
npm run dev:web    # Start server with hot reload
```

---

## Contributing

Contributions are welcome! Please:

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request

---

## License

MIT License - see [LICENSE](LICENSE) file for details.

---

## Acknowledgments

- [Microsoft Graph API](https://developer.microsoft.com/en-us/graph) - The API that powers this integration
- [Model Context Protocol](https://modelcontextprotocol.io/) - The protocol that enables AI tool integration
- [Claude](https://claude.ai/) - The AI assistant this was built for
