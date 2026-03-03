[![MseeP.ai Security Assessment Badge](https://mseep.net/pr/aanerud-mcp-microsoft-office-badge.png)](https://mseep.ai/app/aanerud-mcp-microsoft-office)

# MCP Microsoft Office

**One MCP server. Multiple users. Real Microsoft 365 traffic on your test tenant.**

---

## The Problem

Test tenants sit empty. Static test data does not exercise real workflows. When you need agents that send real emails, schedule real meetings, and collaborate in real Teams channels, mocks and stubs fall short.

## What This Solves

This project connects any MCP-compatible AI client to Microsoft 365 through the Graph API. Each agent authenticates as a distinct tenant user and performs real operations against real data.

- **78 tools** across 9 modules: Mail, Calendar, Files, Teams, Contacts, To-Do, Groups, People, Search
- **Multi-user**: one server supports your entire team, each with isolated data
- **Real Graph API calls**: every operation hits the actual tenant, not a mock
- **Secure**: tokens encrypted at rest, no credentials stored on third-party servers

---

## Architecture

```
                    ┌──────────────────┐
                    │  MCP Client      │
                    │  (Claude, etc.)  │
                    └────────┬─────────┘
                             │ JSON-RPC (stdin/stdout)
                    ┌────────▼─────────┐
                    │  MCP Adapter     │
                    │  (runs locally)  │
                    └────────┬─────────┘
                             │ HTTP + Bearer Token
                    ┌────────▼─────────┐
                    │  MCP Server      │
                    │  (local or       │
                    │   remote)        │
                    └────────┬─────────┘
                             │ Microsoft Graph API
                    ┌────────▼─────────┐
                    │  Microsoft 365   │
                    │  (your tenant)   │
                    └──────────────────┘
```

**Three parts:**

1. **MCP Client** -- the AI you interact with
2. **MCP Adapter** -- a Node.js process that translates MCP protocol to HTTP requests (runs on the same machine as the client)
3. **MCP Server** -- handles authentication and calls the Microsoft Graph API (runs locally or on a remote server)

---

## Permissions

The server requires 18 Microsoft Graph delegated permissions. Twelve work without admin consent. Six require a tenant administrator to grant consent.

### No Admin Consent Required

| Permission | Tools Unlocked |
|---|---|
| `User.Read` | Authentication, user profile |
| `Mail.ReadWrite` | readMail, readMailDetails, markEmailRead, flagMail, getMailAttachments, addMailAttachment, removeMailAttachment |
| `Mail.Send` | sendMail, replyToMail |
| `Calendars.ReadWrite` | getEvents, createEvent, updateEvent, cancelEvent, acceptEvent, tentativelyAcceptEvent, declineEvent, getAvailability, findMeetingTimes, getRooms, getCalendars, addAttachment, removeAttachment |
| `Files.ReadWrite.All` | listFiles, uploadFile, downloadFile, getFileMetadata, getFileContent, setFileContent, updateFileContent, createSharingLink, getSharingLinks, removeSharingPermission, listChannelFiles, uploadFileToChannel, readChannelFile |
| `Contacts.ReadWrite` | listContacts, getContact, createContact, updateContact, deleteContact, searchContacts |
| `Tasks.ReadWrite` | listTaskLists, getTaskList, createTaskList, updateTaskList, deleteTaskList, listTasks, getTask, createTask, updateTask, deleteTask, completeTask |
| `Chat.ReadWrite` | listChats, createChat, getChatMessages, sendChatMessage |
| `Channel.ReadBasic.All` | listTeamChannels, getChannelMessages |
| `ChannelMessage.Send` | sendChannelMessage, replyToMessage |
| `Channel.Create` | createTeamChannel |
| `OnlineMeetings.ReadWrite` | createOnlineMeeting, getOnlineMeeting, listOnlineMeetings, getMeetingByJoinUrl |

### Requires Admin Consent

| Permission | Additional Tools Unlocked |
|---|---|
| `User.Read.All` | Resolve user IDs across Teams, People search |
| `People.Read.All` | findPeople, getRelevantPeople, getPersonById |
| `Group.Read.All` | listGroups, getGroup, listGroupMembers, listMyGroups |
| `ChannelMember.ReadWrite.All` | addChannelMember |
| `ChannelMessage.Read.All` | Read channel message history |
| `OnlineMeetingTranscript.Read.All` | getMeetingTranscripts, getMeetingTranscriptContent |

**Without admin consent**, you get Mail, Calendar, Files, Contacts, To-Do, Chat, and basic Teams channel operations. **With admin consent**, you add People directory search, Groups, channel member management, and meeting transcripts.

---

## Quick Start

### Prerequisites

- **Node.js 18+** ([download](https://nodejs.org/))
- **Claude Desktop** ([download](https://claude.ai/download)) or another MCP client
- **Microsoft 365 account** (work, school, or personal)

### Step 1: Azure App Registration

1. Go to [Azure Portal](https://portal.azure.com) > **Microsoft Entra ID** > **App registrations** > **New registration**
2. Name it `MCP-Microsoft-Office`, register with your preferred account type
3. Copy the **Application (client) ID** and **Directory (tenant) ID**
4. Go to **API permissions** > **Add a permission** > **Microsoft Graph** > **Delegated permissions**
5. Add the 18 permissions listed above
6. If you are a tenant admin, click **Grant admin consent**
7. Go to **Authentication** > **Add a platform** > **Web**
   - Redirect URI: `http://localhost:3000/api/auth/callback`
   - Enable **Allow public client flows**

### Step 2: Clone and Configure

```bash
git clone https://github.com/Aanerud/MCP-Microsoft-Office.git
cd MCP-Microsoft-Office
npm install
```

Copy `.env.example` to `.env` and fill in your Azure app details:

```
MICROSOFT_CLIENT_ID=your-client-id
MICROSOFT_TENANT_ID=your-tenant-id
```

### Step 3: Start the Server and Authenticate

```bash
npm run dev:web
```

Open `http://localhost:3000` in your browser. Click **Login with Microsoft**, sign in, and grant permissions. Then click **Generate MCP Token** and copy the token.

### Step 4: Configure Claude Desktop

Edit your Claude Desktop config:

**macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`
**Windows**: `%APPDATA%\Claude\claude_desktop_config.json`

```json
{
  "mcpServers": {
    "microsoft365": {
      "command": "node",
      "args": ["/path/to/MCP-Microsoft-Office/mcp-adapter.cjs"],
      "env": {
        "MCP_SERVER_URL": "http://localhost:3000",
        "MCP_BEARER_TOKEN": "paste-your-token-here"
      }
    }
  }
}
```

Restart Claude Desktop. Ask: *"What's on my calendar today?"*

---

## Tools (78)

### Mail (9)

| Tool | Description |
|---|---|
| `readMail` | Read inbox messages |
| `sendMail` | Send an email |
| `replyToMail` | Reply to an email |
| `readMailDetails` | Get full email content |
| `markEmailRead` | Mark email as read/unread |
| `flagMail` | Flag or unflag an email |
| `getMailAttachments` | List email attachments |
| `addMailAttachment` | Add attachment to email |
| `removeMailAttachment` | Remove attachment from email |

### Calendar (13)

| Tool | Description |
|---|---|
| `getEvents` | Get calendar events |
| `createEvent` | Create a meeting or event |
| `updateEvent` | Modify an existing event |
| `cancelEvent` | Cancel an event |
| `acceptEvent` | Accept a meeting invitation |
| `tentativelyAcceptEvent` | Tentatively accept |
| `declineEvent` | Decline a meeting invitation |
| `getAvailability` | Check free/busy times |
| `findMeetingTimes` | Find optimal meeting slots |
| `getRooms` | Find meeting rooms |
| `getCalendars` | List all calendars |
| `addAttachment` | Add attachment to event |
| `removeAttachment` | Remove event attachment |

### Files (10)

| Tool | Description |
|---|---|
| `listFiles` | List OneDrive files |
| `uploadFile` | Upload a file |
| `downloadFile` | Download a file |
| `getFileMetadata` | Get file info |
| `getFileContent` | Read file contents |
| `setFileContent` | Write file contents |
| `updateFileContent` | Update existing file |
| `createSharingLink` | Create a sharing link |
| `getSharingLinks` | List sharing links |
| `removeSharingPermission` | Remove sharing access |

### Teams (21)

| Tool | Description |
|---|---|
| `listChats` | List Teams chats |
| `createChat` | Create a new chat |
| `getChatMessages` | Read chat messages |
| `sendChatMessage` | Send a chat message |
| `listJoinedTeams` | List your teams |
| `listTeamChannels` | List team channels |
| `createTeamChannel` | Create a channel |
| `addChannelMember` | Add member to channel |
| `getChannelMessages` | Read channel messages |
| `sendChannelMessage` | Post to a channel |
| `replyToMessage` | Reply to a channel message |
| `listChannelFiles` | List files in a channel |
| `uploadFileToChannel` | Upload file to channel |
| `readChannelFile` | Read a channel file |
| `createOnlineMeeting` | Create a Teams meeting |
| `getOnlineMeeting` | Get meeting details |
| `listOnlineMeetings` | List online meetings |
| `getMeetingByJoinUrl` | Find meeting by join URL |
| `getMeetingTranscripts` | Get meeting transcripts |
| `getMeetingTranscriptContent` | Read transcript content |

*(Note: `addChannelMember` applies to private channels only. Standard channels auto-include all team members.)*

### Contacts (6)

| Tool | Description |
|---|---|
| `listContacts` | List contacts |
| `getContact` | Get contact details |
| `createContact` | Create a contact |
| `updateContact` | Update contact info |
| `deleteContact` | Delete a contact |
| `searchContacts` | Search contacts |

### To-Do (11)

| Tool | Description |
|---|---|
| `listTaskLists` | List task lists |
| `getTaskList` | Get a task list |
| `createTaskList` | Create a task list |
| `updateTaskList` | Rename a task list |
| `deleteTaskList` | Delete a task list |
| `listTasks` | List tasks |
| `getTask` | Get task details |
| `createTask` | Create a task |
| `updateTask` | Update a task |
| `deleteTask` | Delete a task |
| `completeTask` | Mark task complete |

### Groups (4)

| Tool | Description |
|---|---|
| `listGroups` | List Microsoft 365 groups |
| `getGroup` | Get group details |
| `listGroupMembers` | List group members |
| `listMyGroups` | List your groups |

### People (3)

| Tool | Description |
|---|---|
| `findPeople` | Search the directory |
| `getRelevantPeople` | Get frequent contacts |
| `getPersonById` | Get person details |

### Search (1)

| Tool | Description |
|---|---|
| `search` | Unified search across emails, files, events, and chat messages |

---

## Multi-User

Each user authenticates independently. The server isolates all data by user identity.

```
  Alice (alice@contoso.com)          Bob (bob@contoso.com)
  ├─ Her own Microsoft tokens        ├─ His own Microsoft tokens
  ├─ Her own session                  ├─ His own session
  └─ Claude Desktop (her laptop)     └─ Claude Desktop (his PC)

              Complete data isolation.
         Alice never sees Bob's data.
```

For **automated testing with multiple agents**, use the ROPC (Resource Owner Password Credentials) flow to authenticate programmatically:

```bash
# Start the server
npm run dev:web

# Run the E2E test suite (authenticates 3 users via ROPC)
node tests/run-all.cjs
```

The test suite authenticates multiple users, then exercises all 78 tools across 9 modules plus 5 cross-module workflows. See `tests/` for the full implementation.

---

## E2E Test Suite

The project includes a comprehensive test suite covering all 78 tools.

```bash
# Run all tests (requires server running)
node tests/run-all.cjs

# Run a single module
node tests/run-all.cjs --bucket mail --buckets-only

# Run only workflows
node tests/run-all.cjs --workflows-only
```

**Test structure:**

```
tests/
  lib/           Shared auth, HTTP client, reporter
  buckets/       One file per module (9 files, 78 tools)
  workflows/     Cross-module tests (5 files)
  run-all.cjs    Master runner
```

Tests authenticate via ROPC (no manual token management) and run in ~100 seconds.

---

## Environment Variables

Copy `.env.example` to `.env` and configure:

| Variable | Required | Description |
|---|---|---|
| `MICROSOFT_CLIENT_ID` | Yes | Azure App Client ID |
| `MICROSOFT_TENANT_ID` | Yes | Azure Tenant ID |
| `MICROSOFT_REDIRECT_URI` | No | OAuth callback URL (default: `http://localhost:3000/api/auth/callback`) |
| `DEVICE_REGISTRY_ENCRYPTION_KEY` | Production | 32-byte encryption key for token storage |
| `JWT_SECRET` | Production | Secret for signing JWT tokens |
| `CORS_ALLOWED_ORIGINS` | Production | Comma-separated allowed origins |
| `PORT` | No | Server port (default: `3000`) |
| `NODE_ENV` | No | `development` or `production` |

---

## Deployment

### Local (Recommended for Getting Started)

```bash
npm install
npm run dev:web
```

### Azure App Service

See [docs/azure-deployment.md](docs/azure-deployment.md) for CI/CD deployment with GitHub Actions.

---

## Security

- **Encrypted storage**: all Microsoft tokens encrypted at rest with AES-256
- **No client secrets**: uses public client flow (PKCE) for desktop authentication
- **Token isolation**: each user's tokens stored separately with different encryption keys
- **Rate limiting**: built-in rate limiting protects against abuse
- **CORS protection**: origin allowlist in production
- **Session expiry**: sessions expire after 24 hours

### Production Checklist

- [ ] Set `NODE_ENV=production`
- [ ] Set `DEVICE_REGISTRY_ENCRYPTION_KEY` (32 bytes)
- [ ] Set `JWT_SECRET` (strong random string)
- [ ] Set `CORS_ALLOWED_ORIGINS`
- [ ] Use HTTPS with a valid certificate

---

## Project Structure

```
MCP-Microsoft-Office/
├── mcp-adapter.cjs          MCP protocol adapter (runs locally)
├── src/
│   ├── api/                 Express routes and controllers
│   ├── auth/                MSAL authentication
│   ├── core/                Services (cache, storage, tools)
│   ├── graph/               Microsoft Graph API services
│   └── modules/             Feature modules (mail, calendar, etc.)
├── public/                  Web UI
└── tests/                   E2E test suite (gitignored)
```

---

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request

---

## License

MIT License -- see [LICENSE](LICENSE) file.
