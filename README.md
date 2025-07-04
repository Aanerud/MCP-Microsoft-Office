[![MseeP.ai Security Assessment Badge](https://mseep.net/pr/aanerud-mcp-microsoft-office-badge.png)](https://mseep.ai/app/aanerud-mcp-microsoft-office)

# MCP Microsoft Office Bridge 

**A secure multi-user remote MCP server connecting LLMs to Microsoft 365**

This project is a **complete MCP server** that enables Claude and other LLM clients to interact with Microsoft 365 data through a secure remote API. It features session-based user isolation and dual authentication modes for secure remote operation.

## Quick Setup

1. **Clone & Install**
   ```bash
   git clone https://github.com/Aanerud/MCP-Microsoft-Office.git
   cd MCP-Microsoft-Office
   npm install
   ```

2. **Environment Setup**
   ```bash
   # Create .env file with required variables
   MICROSOFT_CLIENT_ID=your_client_id
   MICROSOFT_TENANT_ID=your_tenant_id
   ```

3. **Start the Server**
   ```bash
   npm run dev  # Development mode with full logging
   ```

4. **Configure Claude Desktop**
   Edit `~/Library/Application Support/Claude/claude_desktop_config.json`:
   ```json
   {
     "mcpServers": {
       "microsoft365": {
         "command": "node",
         "args": ["/path/to/mcp-adapter.cjs"],
         "env": {
           "MCP_SERVER_URL": "http://localhost:3000",
           "MCP_BEARER_TOKEN": "your-token-here"
         },
         "restrictions": {}
       }
     }
   }
   ```

## Key Features

The MCP server provides secure multi-user access to Microsoft 365 data:

### Core Functionality
- **Email**: Read, search, and send emails through Microsoft Graph API
- **Calendar**: View and manage calendar events
- **OneDrive/SharePoint**: Browse and access files
- **Contacts**: Find and manage contacts

### Multi-User Architecture
- **Session isolation**: Complete data separation between users
- **Device authentication**: Secure registry for MCP adapter connections
- **JWT token system**: Short-lived access tokens with user binding
- **Remote operation**: Configurable MCP_SERVER_URL endpoint

## Available Tools

This MCP server provides 37+ tools for comprehensive Microsoft 365 integration:

### Mail Tools
- `getMail` / `readMail` - Fetch messages from inbox
- `sendMail` - Compose and send emails with attachments
- `searchMail` - Find emails with powerful query filters
- `flagMail` - Flag/unflag important emails
- `getMailDetails` - View complete email content
- `markMailRead` / `markEmailRead` - Update read status
- `addMailAttachment` - Add files to draft emails
- `removeMailAttachment` - Remove email attachments

### Calendar Tools
- `getCalendar` / `getEvents` - View upcoming events
- `createEvent` - Schedule new meetings with attendees
- `updateEvent` - Modify existing calendar entries
- `cancelEvent` - Remove events from calendar
- `getAvailability` - Check free/busy times
- `findMeetingTimes` - Find optimal meeting slots
- `addAttachment` - Add files to calendar events
- `removeAttachment` - Remove event attachments

### Files Tools
- `listFiles` - Browse OneDrive and SharePoint files
- `searchFiles` - Find files by name or content
- `downloadFile` - Retrieve file content
- `uploadFile` - Add new files to cloud storage
- `getFileMetadata` - View file properties
- `getFileContent` - Read document contents
- `setFileContent` / `updateFileContent` - Modify files
- `createSharingLink` - Generate sharing URLs
- `getSharingLinks` - View existing sharing permissions
- `removeSharingPermission` - Revoke access to shared content

### People Tools
- `findPeople` - Locate contacts by name or email
- `searchPeople` - Search organization directory
- `getRelevantPeople` - Find frequently contacted people
- `getPersonById` - Look up specific contacts

### AI Query Tool
- `query` - Natural language queries across all Microsoft 365 data

### Example Usage
```text
# Ask Claude to show your calendar
"What meetings do I have today?"

# Ask Claude to send an email
"Send an email to Sarah about the project status"

# Ask Claude to find files
"Find PowerPoint presentations I created last week"

# Ask Claude to add events
"Schedule a team meeting tomorrow at 2pm"
```

## Multi-User Architecture

### Remote Service Design
```
Claude Desktop ←→ MCP Adapter ←→ Remote MCP Server ←→ Microsoft 365
```

The MCP server can be deployed as a remote service, allowing multiple users to connect via MCP adapters:

- **Session-Based User Isolation**: User sessions are managed by `session-service.cjs` with unique session IDs
- **Dual Authentication**: Supports both browser session and JWT bearer token authentication
- **Remote Server Configuration**: MCP adapter connects via `MCP_SERVER_URL` environment variable
- **OAuth 2.0 Compliance**: Supports discovery via `/.well-known/oauth-protected-resource` endpoint
- **Device Registry**: Secure management and authorization of MCP adapter connections

### User Isolation & Session Management

Each user's data is completely isolated through session-based architecture:

```javascript
// From session-service.cjs
async createSession(options = {}) {
    const sessionId = uuid();
    const sessionSecret = crypto.randomBytes(SESSION_SECRET_LENGTH).toString('hex');
    const expiresAt = Date.now() + SESSION_EXPIRY;
    
    const sessionData = {
        session_id: sessionId,
        session_secret: sessionSecret,
        expires_at: expiresAt,
        created_at: Date.now(),
        // User-specific data storage
    };
}
```

This creates a unique context for each user, ensuring complete data separation in the multi-user environment.

### Authentication System

The system implements two modes of authentication:

1. **Simple Bearer Token Mode** - Detected when `MCP_BEARER_TOKEN` is present:
   ```javascript
   // From mcp-adapter.cjs
   const MCP_BEARER_TOKEN = process.env.MCP_BEARER_TOKEN;
   const SIMPLE_MODE = !!MCP_BEARER_TOKEN;
   ```

2. **OAuth Device Flow Mode** - Used when no bearer token is provided:
   ```javascript
   // OAuth 2.0 discovery path in adapter
   const OAUTH_DISCOVERY_PATH = '/.well-known/oauth-protected-resource';
   ```

### Device Registry

The `device-registry.cjs` module manages secure MCP adapter connections:

```javascript
// Device code configuration from device-registry.cjs
const DEVICE_CODE_LENGTH = 8;
const USER_CODE_LENGTH = 6;
const DEVICE_CODE_EXPIRY = 15 * 60 * 1000; // 15 minutes
```

1. **Device Registration**: Each adapter registers with unique IDs
   ```javascript
   // From device-registry.cjs - registerDevice() function
   const deviceId = uuid();
   const deviceSecret = this.generateDeviceSecret();
   const deviceCode = this.generateDeviceCode();
   const userCode = this.generateUserCode();
   ```

2. **Authorization Flow**: Users authorize devices with user codes
   ```javascript
   // From device-registry.cjs - authorizeDevice() function
   async authorizeDevice(userCode, userId) {
       // Links device to specific user account
   }
   ```

3. **Token-User Binding**: Every JWT token is bound to a specific user
   ```javascript
   // Device JWT payload structure
   req.user = {
     deviceId: decoded.deviceId,
     userId: decoded.userId,
     isApiCall: true
   };
   ```

4. **Automatic Expiration**: Unused device registrations automatically expire
   ```javascript
   // From cleanupExpiredDevices() function
   'DELETE FROM devices WHERE expires_at < ? AND is_authorized = FALSE'
   ```

## Security Architecture

### Authentication Layers
- **Browser Session Authentication**: MSAL tokens stored in Express sessions
  ```javascript
  // From auth-middleware.cjs
  if (req.session && req.session.id) {
      const msalService = require('../../auth/msal-service.cjs');
      const isAuthenticated = await msalService.isAuthenticated(req);
      if (isAuthenticated) {
          req.user = { userId: `user:${req.session.id}`, sessionId: req.session.id };
          return next();
      }
  }
  ```

### JWT Authentication for API
- **Token Validation**: Proper signature and expiry checks 
- **User Context Extraction**: Standardized req.user object
- **WWW-Authenticate Headers**: MCP specification compliance
  ```javascript
  // From auth-middleware.cjs
  return res.status(401).json({
      error: 'Authentication required',
      message: 'Missing Authorization header or valid session'
  }).header('WWW-Authenticate', 'Bearer realm="MCP Remote Service"');
  ```

## End-to-End Authentication Flow

This project uses a sophisticated two-layer authentication system to ensure secure, multi-user access. It's crucial for developers to understand how these two layers work together.

### Layer 1: Client-to-Server Authentication (Session Management)

This layer secures the connection between the MCP Adapter (the client) and this MCP Server.

-   **Responsibility**: To identify and authenticate the specific user making a request to the server.
-   **Mechanism**: The server uses a session-based system (`session-service.cjs`). When a user authenticates, a unique session is created.
-   **The `req` Object**: Every incoming request from an authenticated client carries session information, which Express makes available in the `req` object (specifically `req.session`). This object is the key to identifying the user in all subsequent operations.

### Layer 2: Server-to-Microsoft Graph Authentication (MSAL)

This layer secures the connection between our MCP Server and the Microsoft Graph API.

-   **Responsibility**: To authenticate as a specific user when making requests to Microsoft 365.
-   **Mechanism**: The server uses the Microsoft Authentication Library (MSAL) via `msal-service.cjs`. It securely stores OAuth 2.0 access tokens for each user who has granted consent.
-   **Multi-User Token Storage**: The server stores multiple MSAL tokens, one for each authenticated user session.

### The Bridge: Connecting Sessions to Graph Tokens via `req`

The two layers are independent, but they must be connected to function correctly. The `req` object serves as this critical bridge.

**The Problem:** When the server needs to call the Microsoft Graph API, how does it know which user's MSAL token to use?

**The Solution:** The `req.session` object, carried from Layer 1, provides the user's unique identifier. The `msal-service.cjs` uses this identifier to look up the correct MSAL token from secure storage.

This is why the `req` object **must** be passed down from the API controllers all the way to any service that interacts with the Graph API.

**Data Flow Example (`getAvailability`):**

1.  **Client Request**: The MCP Adapter sends a request to `/api/v1/calendar/availability`.
2.  **Controller (`calendar-controller.js`)**: The controller receives the request. Express populates `req.session` with the user's session data.
3.  **Context Passing**: The controller calls the `calendarModule`, **passing the `req` object** along with other parameters.
    ```javascript
    // In calendar-controller.js
    availabilityData = await calendarModule.getAvailability({ ..., req: req });
    ```
4.  **Module (`calendar/index.cjs`)**: The module receives the `req` object and passes it down to the `graph-service`.
5.  **Graph Service (`graph/calendar-service.cjs`)**: The service needs an authenticated Graph client. It calls `msal-service.cjs` to get an access token, passing the `req` object.
6.  **Token Retrieval (`msal-service.cjs`)**: The MSAL service inspects `req.session` to identify the user, retrieves their specific MSAL token from storage, and returns it.
7.  **Authenticated API Call**: The graph service uses the retrieved token to make a successful, authenticated call to the Microsoft Graph API on behalf of the correct user.

Failure to pass the `req` object at any step breaks this chain, leading to authentication errors, as the system won't know which user's credentials to use.

## Authentication Modes

### Simple Bearer Token Mode

The adapter supports simple bearer token authentication for easy setup:

```javascript
// From mcp-adapter.cjs - Simple mode detection
const MCP_BEARER_TOKEN = process.env.MCP_BEARER_TOKEN;
const SIMPLE_MODE = !!MCP_BEARER_TOKEN;
```

When `MCP_BEARER_TOKEN` is present in environment variables, the adapter automatically uses this authentication mode and skips the more complex OAuth device flow.

### OAuth Device Flow Mode

For enterprise usage, the adapter supports the full OAuth 2.0 device flow:

```javascript
// OAuth 2.0 discovery and device authorization flow
async function discoverOAuthEndpoints() {
    // Fetches authorization_endpoint, token_endpoint, device_authorization_endpoint
    // Uses /.well-known/oauth-protected-resource standard endpoint
}
```

This mode activates automatically when no bearer token is provided, following MCP specification requirements for OAuth 2.0 Protected Resource Metadata.

## Environment Variables

### MCP Server Settings
```
# Core server configuration
MICROSOFT_CLIENT_ID      # Azure App Registration client ID
MICROSOFT_TENANT_ID      # Azure tenant ID

# LLM provider configuration (optional)
LLM_PROVIDER             # claude or openai
CLAUDE_API_KEY           # API key for Claude
OPENAI_API_KEY           # API key for OpenAI
```

### MCP Adapter Settings
```
# Remote server configuration
MCP_SERVER_URL          # URL of the MCP server (default: http://localhost:3000)

# Authentication mode
MCP_BEARER_TOKEN        # When provided, enables simple bearer token authentication

# Legacy settings (backward compatibility)
API_HOST                # Alternate host configuration
API_PORT                # Alternate port configuration
API_BASE_PATH           # Base path for API endpoints
```

## API Endpoints

### Authentication Endpoints
- `GET /.well-known/oauth-protected-resource` - OAuth 2.0 discovery endpoint
- `POST /api/auth/device/register` - Device registration 
- `POST /api/auth/device/token` - Token endpoint for device flow

### Microsoft 365 API Endpoints
- `GET /api/v1/mail` - Retrieve emails
- `POST /api/v1/mail` - Send emails
- `GET /api/v1/calendar` - List calendar events
- `POST /api/v1/calendar` - Create calendar events
- `GET /api/v1/files` - Browse files and folders
- `POST /api/v1/files` - Upload files
- `GET /api/v1/people` - Search contacts

All API endpoints follow consistent asynchronous patterns with proper error handling, as required by the MCP specification.

## Contributing

This project demonstrates **enterprise-grade MCP development**. Key patterns to follow:

1. **Tool Definitions**: Always define proper parameter schemas
2. **Async Patterns**: Use async/await with proper error handling
3. **Error Service**: Use centralized error creation and logging
4. **Data Normalization**: Consistent response formats
5. **Testing**: Real API integration tests

## License

MIT License - Build amazing things with Microsoft 365 and MCP!

---

**Built with  as a Windsurf Vibe Code project - demonstrating how to build production-ready MCP servers that bridge LLMs with enterprise APIs.**