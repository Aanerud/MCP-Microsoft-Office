[![MseeP.ai Security Assessment Badge](https://mseep.net/pr/aanerud-mcp-microsoft-office-badge.png)](https://mseep.ai/app/aanerud-mcp-microsoft-office)

# 🚀 MCP Microsoft Office - Enterprise-Grade Microsoft 365 Integration

**The most comprehensive, secure, and user-focused MCP server for Microsoft 365**

> Transform how you interact with Microsoft 365 through Claude and other LLMs with enterprise-grade security, comprehensive logging, and seamless multi-user support.

## ✨ Why This Project is Special

🔐 **User-Centric Security**: Every user gets their own isolated, encrypted data space  
📊 **Enterprise Logging**: 4-tier comprehensive logging system with full observability  
🛠️ **50+ Professional Tools**: Complete Microsoft 365 API coverage with validation  
⚡ **Zero-Config Setup**: Automatic project initialization - just `npm install` and go!  
🏢 **Multi-User Ready**: Session-based isolation with Microsoft authentication  
🔧 **Developer Friendly**: Extensive debugging, monitoring, and error handling  

---

## 🚀 Quick Start (Beginner-Friendly)

### 1. **One-Command Setup** ⚡
```bash
git clone https://github.com/Aanerud/MCP-Microsoft-Office.git
cd MCP-Microsoft-Office
npm install  # ✨ This does EVERYTHING automatically! ( i hope! )
```

**What happens automatically:**
- ✅ Creates secure database with user isolation
- ✅ Generates `.env` configuration file
- ✅ Sets up all required directories
- ✅ Initializes logging and monitoring systems
- ✅ Prepares multi-user session management

### 2. **Configure Azure App Registration** 🔑

You need to create an App Registration in Azure Entra ID (formerly Azure AD). Follow the detailed guide below.

📋 **Azure Portal:** [App Registrations](https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)

After setup, edit the auto-generated `.env` file:
```bash
MICROSOFT_CLIENT_ID=your_client_id_here
MICROSOFT_TENANT_ID=your_tenant_id_here
```

### 3. **Launch Your Server** 🎯
```bash
npm run dev:web  # Full development mode with comprehensive logging
```

🌐 **Access your server at:** `http://localhost:3000`

---

## 🔐 Azure App Registration Setup

This section provides detailed instructions for configuring your Azure Entra ID (Azure AD) App Registration.

### Step 1: Create App Registration

1. Go to [Azure Portal](https://portal.azure.com) → **Microsoft Entra ID** → **App registrations**
2. Click **+ New registration**
3. Configure:
   - **Name**: `MCP-Microsoft-Office` (or your preferred name)
   - **Supported account types**: Choose based on your needs:
     - *Single tenant*: Only your organization
     - *Multitenant*: Any Microsoft Entra ID tenant
   - **Redirect URI**: Leave blank for now (configure in Step 3)
4. Click **Register**
5. Note your **Application (client) ID** and **Directory (tenant) ID**

### Step 2: Configure API Permissions

Go to **API permissions** → **+ Add a permission** → **Microsoft Graph** → **Delegated permissions**

#### Required Permissions

| Permission | Description | Admin Consent Required |
|------------|-------------|----------------------|
| `User.Read` | Sign in and read user profile | No |
| `Mail.ReadWrite` | Read and write user mail | Yes* |
| `Mail.Send` | Send mail as user | Yes* |
| `Calendars.ReadWrite` | Full access to user calendars | Yes* |
| `Files.ReadWrite` | Full access to user files | Yes* |

*Admin consent may be required in enterprise tenants

#### Optional Permissions (for People API)

| Permission | Description |
|------------|-------------|
| `People.Read` | Read users' relevant people lists |
| `Contacts.Read` | Read user contacts |

### Step 3: Configure Authentication

Go to **Authentication** in the left sidebar:

#### Add Platform
1. Click **+ Add a platform**
2. Select **Web** (NOT "Single-page application")
3. Add Redirect URI:
   ```
   https://your-domain.com/api/auth/callback
   ```
   For local development:
   ```
   http://localhost:3000/api/auth/callback
   ```
4. Click **Configure**

#### Enable Implicit Grant (optional)
Under **Implicit grant and hybrid flows**:
- ✅ Access tokens
- ✅ ID tokens

#### Advanced Settings
Scroll down to **Advanced settings**:
- **Allow public client flows**: Set to **Yes**
  - This enables the Device Code Flow used by MCP adapters

### Step 4: Admin Consent

**Important for Enterprise Tenants:**

If you see "Needs admin approval" when logging in, the permissions require admin consent.

#### Option A: Grant Admin Consent (if you're an admin)
1. Go to **API permissions**
2. Click **Grant admin consent for [Your Organization]**
3. Confirm the prompt

#### Option B: Request Admin Consent (if you're not an admin)
1. Contact your IT administrator
2. Provide them the App Registration name/ID
3. Ask them to grant admin consent for the listed permissions

### Step 5: Environment Configuration

Add to your `.env` file:
```bash
# From App Registration Overview page
MICROSOFT_CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
MICROSOFT_TENANT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx

# Your server URL (must match redirect URI)
MICROSOFT_REDIRECT_URI=https://your-domain.com/api/auth/callback
```

### Common Issues

#### "AADSTS7000218: client_assertion or client_secret required"
**Cause**: App is configured as Confidential Client
**Fix**: Enable "Allow public client flows" in Authentication settings

#### "Needs administrator approval"
**Cause**: Permissions require admin consent in your tenant
**Fix**: Have an admin grant consent, or use a personal Microsoft account for testing

#### "Invalid redirect URI"
**Cause**: Redirect URI mismatch between app registration and server
**Fix**: Ensure the URI in Authentication matches your `MICROSOFT_REDIRECT_URI` exactly

#### "Platform type mismatch"
**Cause**: Using "Single-page application" instead of "Web"
**Fix**: Remove SPA platform, add "Web" platform with same redirect URI

---

## 🛠️ Complete Tool Arsenal 

### 📧 **Email Management** (9 Tools)
- `getMail` / `readMail` - Retrieve inbox messages with filtering
- `sendMail` - Compose and send emails with attachments
- `searchMail` - Powerful email search with KQL queries
- `flagMail` - Flag/unflag important emails
- `getEmailDetails` - View complete email content and metadata
- `markAsRead` / `markEmailRead` - Update read status
- `getMailAttachments` - Download email attachments
- `addMailAttachment` - Add files to emails
- `removeMailAttachment` - Remove email attachments

### 📅 **Calendar Operations** (13 Tools)
- `getCalendar` / `getEvents` - View upcoming events with filtering
- `createEvent` - Schedule meetings with attendees and rooms
- `updateEvent` - Modify existing calendar entries
- `cancelEvent` - Remove events from calendar
- `getAvailability` - Check free/busy times
- `acceptEvent` - Accept meeting invitations
- `tentativelyAcceptEvent` - Tentatively accept meetings
- `declineEvent` - Decline meeting invitations
- `findMeetingTimes` - Find optimal meeting slots
- `getRooms` - Find available meeting rooms
- `getCalendars` - List all user calendars
- `addAttachment` - Add files to calendar events
- `removeAttachment` - Remove event attachments

### 📁 **File Management** (11 Tools)
- `listFiles` - Browse OneDrive and SharePoint files
- `searchFiles` - Find files by name or content
- `downloadFile` - Retrieve file content
- `uploadFile` - Add new files to cloud storage
- `getFileMetadata` - View file properties and permissions
- `getFileContent` - Read document contents
- `setFileContent` / `updateFileContent` - Modify file contents
- `createSharingLink` - Generate secure sharing URLs

---

## 🔧 Advanced Configuration

### **Environment Variables**
```bash
# Microsoft 365 Configuration
MICROSOFT_CLIENT_ID=your_client_id
MICROSOFT_TENANT_ID=your_tenant_id

# Server Configuration
PORT=3000
NODE_ENV=development

# Security
MCP_ENCRYPTION_KEY=your_32_byte_encryption_key
MCP_TOKEN_SECRET=your_jwt_secret

# Database (Optional - defaults to SQLite)
DATABASE_TYPE=sqlite  # or 'mysql', 'postgresql'
DATABASE_URL=your_database_url

# Logging
LOG_LEVEL=info
LOG_RETENTION_DAYS=30
```

### **Database Support**
- ✅ **SQLite** (Default - Zero configuration)
- ✅ **MySQL** (Production ready)
- ✅ **PostgreSQL** (Enterprise grade)

### **Backup & Migration**
```bash
# Backup user data
npm run backup

# Restore from backup
npm run restore backup-file.sql

# Database migration
npm run migrate
```

---

## 💡 Real-World Usage Examples

### **Natural Language Queries** 🗣️
```text
# Email Management
"Show me unread emails from last week"
"Send a meeting recap to the project team"
"Find emails about the Q4 budget"

# Calendar Operations  
"What meetings do I have tomorrow?"
"Schedule a 1-on-1 with Sarah next Tuesday at 2pm"
"Find a time when John, Mary, and I are all free"

# File Management
"Find my PowerPoint presentations from last month"
"Share the project proposal with the team"
"Upload the latest budget spreadsheet"

# People & Contacts
"Find contacts in the marketing department"
"Get John Smith's contact information"
"Who are my most frequent email contacts?"
```

### **Advanced API Usage** 🔧
```javascript
// Direct API calls with full validation
POST /api/v1/mail/send
{
    "to": ["colleague@company.com"],
    "subject": "Project Update",
    "body": "Here's the latest update...",
    "attachments": [{
        "name": "report.pdf",
        "contentBytes": "base64_encoded_content"
    }]
}

// Calendar event creation with attendees
POST /api/v1/calendar/events
{
    "subject": "Team Standup",
    "start": "2024-01-15T09:00:00Z",
    "end": "2024-01-15T09:30:00Z",
    "attendees": ["team@company.com"],
    "location": "Conference Room A"
}

// File search with advanced filters
GET /api/v1/files?query=presentation&limit=10&type=powerpoint
```

---

## 🚨 Troubleshooting & Support

### **Common Issues & Solutions**

#### **Database Issues** 🗄️
```bash
# Reset database completely
npm run reset-db

# Check database health
curl http://localhost:3000/api/health

# View database logs
tail -f data/logs/database.log
```

#### **Authentication Problems** 🔐
```bash
# Check Microsoft 365 configuration
echo $MICROSOFT_CLIENT_ID
echo $MICROSOFT_TENANT_ID

# Test authentication endpoint
curl http://localhost:3000/api/auth/status

# Clear data
rm -rf data/*
```

#### **Permission Errors** 📁
```bash
# Fix directory permissions
chmod -R 755 data/
chown -R $USER:$USER data/

# Check disk space
df -h
```

### **Advanced Debugging** 🔍

#### **Enable Comprehensive Logging**
```bash
# Full debug mode
npm run dev:web

# Specific category logging
DEBUG=mail,calendar npm run dev:web

# View real-time logs
curl http://localhost:3000/api/logs?limit=100&level=debug
```

#### **Performance Monitoring**
```bash
# Check system metrics
curl http://localhost:3000/api/health

# Monitor API response times
curl -w "@curl-format.txt" http://localhost:3000/api/mail

# Database performance
sqlite3 data/mcp.sqlite ".timer on" "SELECT COUNT(*) FROM user_logs;"
```

---

## 🎯 Production Deployment

### **Environment Setup**
```bash
# Production environment variables
NODE_ENV=production
PORT=3000
DATABASE_TYPE=postgresql
DATABASE_URL=postgresql://user:pass@host:5432/mcpdb
MCP_ENCRYPTION_KEY=your_32_byte_production_key
LOG_LEVEL=info
LOG_RETENTION_DAYS=90
```

### **Security Hardening**
```bash
# Generate secure encryption key
openssl rand -hex 32

# Set proper file permissions
chmod 600 .env
chmod 700 data/

# Enable HTTPS (recommended)
HTTPS_ENABLED=true
SSL_CERT_PATH=/path/to/cert.pem
SSL_KEY_PATH=/path/to/key.pem
```

### **Monitoring & Alerts**
```bash
# Health check endpoint
GET /api/health

# User activity monitoring
GET /api/logs?scope=user&limit=1000

# System metrics
GET /api/metrics
```

---

## 🏆 What Makes This Project Outstanding

### **🔒 Enterprise-Grade Security**
- **Zero Trust Architecture**: Every request is authenticated and authorized
- **User Data Isolation**: Complete separation between users' data
- **Encryption at Rest**: All sensitive data encrypted in database
- **Session Security**: Secure session management with automatic cleanup
- **Audit Trail**: Complete logging of all user activities

### **📊 Comprehensive Observability**
- **4-Tier Logging**: From development debugging to user activity tracking
- **Real-Time Monitoring**: Live system health and performance metrics
- **Error Tracking**: Structured error handling with full context
- **Performance Analytics**: Response times, success rates, and usage patterns

### **🛠️ Developer Experience**
- **Zero Configuration**: Automatic setup with `npm install`
- **Extensive Validation**: Joi schemas for all API endpoints
- **Type Safety**: Comprehensive parameter validation and transformation
- **Error Handling**: Graceful error handling with detailed diagnostics
- **Development Tools**: Rich debugging and monitoring capabilities

### **🏢 Production Ready**
- **Multi-Database Support**: SQLite, MySQL, PostgreSQL
- **Horizontal Scaling**: Session-based architecture supports load balancing
- **Health Checks**: Comprehensive health monitoring endpoints
- **Backup & Recovery**: Built-in backup and migration tools
- **Security Hardening**: Production-ready security configurations

---

## 📚 API Documentation

### **Authentication Endpoints**
```bash
GET  /api/auth/status     # Check authentication status
POST /api/auth/login      # Initiate Microsoft 365 login
GET  /api/auth/callback   # OAuth callback handler
POST /api/auth/logout     # Logout and cleanup session
```

### **Mail API Endpoints**
```bash
GET    /api/v1/mail              # Get inbox messages
POST   /api/v1/mail/send         # Send email with attachments
GET    /api/v1/mail/search       # Search emails
PATCH  /api/v1/mail/:id/flag     # Flag/unflag email
GET    /api/v1/mail/:id          # Get email details
PATCH  /api/v1/mail/:id/read     # Mark as read/unread
```

### **Calendar API Endpoints**
```bash
GET    /api/v1/calendar          # Get calendar events
POST   /api/v1/calendar/events   # Create new event
PUT    /api/v1/calendar/events/:id # Update event
DELETE /api/v1/calendar/events/:id # Cancel event
GET    /api/v1/calendar/rooms    # Get available rooms
```

### **Files API Endpoints**
```bash
GET    /api/v1/files             # List files
GET    /api/v1/files/search      # Search files
POST   /api/v1/files/upload      # Upload file
GET    /api/v1/files/:id         # Get file metadata
GET    /api/v1/files/:id/content # Download file
```

### **People API Endpoints**
```bash
GET    /api/v1/people            # Get relevant people
GET    /api/v1/people/search     # Search people
GET    /api/v1/people/:id        # Get person details
```

### **System Endpoints**
```bash
GET    /api/health               # System health check
GET    /api/logs                 # Get system/user logs
POST   /api/v1/query             # Natural language query
```

#### **Comprehensive Audit Trail**
- **Authentication Events**: Login, logout, token refresh activities
- **Authorization Changes**: Permission grants and revocations
- **Data Access**: File access, email reads, calendar views
- **Administrative Actions**: Configuration changes and system updates

### 🛡️ **Security & Compliance Benefits**

#### **Enterprise Security Standards**
- **Zero Trust Architecture**: Every operation logged and verified
- **Audit Compliance**: Complete activity trails for compliance reporting
- **Incident Response**: Detailed logs for security incident investigation
- **User Accountability**: Clear attribution of all actions to authenticated users

#### **Multi-Tenant Security**
- **Data Isolation**: Complete separation between different user accounts
- **Session Security**: Secure session management with proper cleanup
- **Token Security**: JWT tokens with user binding and expiration
- **Access Logging**: All access attempts logged with context

### 🔍 **Log Categories & Structure**

#### **Supported Log Categories**
- `auth` - Authentication and authorization events
- `mail` - Email operations and activities
- `calendar` - Calendar events and scheduling
- `files` - File access and management
- `people` - Contact and directory operations
- `graph` - Microsoft Graph API interactions
- `storage` - Database and storage operations
- `request` - HTTP request/response logging
- `monitoring` - System monitoring and metrics

#### **Log Entry Structure**
```javascript
{
    "id": "log_entry_uuid",
    "timestamp": "2025-07-06T17:04:23.131Z",
    "level": "info",
    "category": "mail",
    "message": "Email sent successfully",
    "context": {
        "userId": "ms365:user@company.com",
        "operation": "sendMail",
        "duration": 1250,
        "recipientCount": 3
    },
    "sessionId": "session_uuid",
    "deviceId": "device_uuid"
}
```

### 🚀 **Production-Ready Logging**

This logging system is **production-tested** and provides:
- **High Performance**: Minimal overhead on API operations
- **Scalability**: Efficient storage and retrieval of large log volumes
- **Reliability**: Robust error handling and fallback mechanisms
- **Maintainability**: Clear separation of concerns and structured data

The logging system ensures complete visibility into system operations while maintaining the highest standards of user privacy and data security.

## Multi-User Architecture

### System Overview

The MCP server supports multiple concurrent users, each with their own authentication method and isolated data:

```
┌─────────────────────────────────────────────────────────────────────────────┐
│                         MCP Server (Remote or Local)                        │
│                         Supports Multiple Users                             │
├─────────────────────────────────────────────────────────────────────────────┤
│                                                                             │
│  User A (ms365:userA@company.com)    │  User B (ms365:userB@corp.com)       │
│  └─ Auth: OAuth Flow                 │  └─ Auth: Enterprise Token           │
│  └─ Tokens stored with userId        │  └─ Tokens stored with userId        │
│                                                                             │
│  MCP Clients for User A:             │  MCP Clients for User B:             │
│  ├─ Claude Desktop (deviceId: xxx)   │  ├─ Claude Desktop (deviceId: aaa)   │
│  └─ Other MCP Client (deviceId: yyy) │  └─ Other MCP Client (deviceId: bbb) │
│                                                                             │
└─────────────────────────────────────────────────────────────────────────────┘
```

### Data Flow Architecture

```
┌──────────────┐     MCP Bearer Token      ┌────────────┐    MS Graph Token    ┌───────────────┐
│  MCP Client  │ ◄─────────────────────► │ MCP Server │ ◄─────────────────► │ Microsoft 365 │
│(Claude, etc.)│   (Transport Security)   │            │   (API Access)       │  Graph API    │
└──────────────┘                          └────────────┘                      └───────────────┘
```

**Two distinct tokens serve different purposes:**

| Token | Purpose | Contains | Used For |
|-------|---------|----------|----------|
| **MCP Bearer Token** | Transport security | `userId` + `deviceId` | MCP Client ↔ MCP Server communication |
| **Microsoft Graph Token** | API access | Microsoft identity | MCP Server ↔ Microsoft 365 API calls |

### Key Identifiers

- **`userId`** (`ms365:email@domain.com`): Identifies the Microsoft 365 user. Used for storing and retrieving Microsoft Graph tokens.
- **`deviceId`** (`mcp-token-xxx-abc`): Identifies the specific MCP client instance. Used for tracking, auditing, and client management.

---

## Authentication Methods

The server supports **two authentication methods** for obtaining Microsoft Graph tokens:

### Method 1: OAuth Authorization Code Flow

Traditional web-based authentication where users log in through Microsoft's OAuth flow.

**How it works:**
1. User visits the MCP server web UI
2. Clicks "Login with Microsoft"
3. Redirected to Microsoft login page
4. After authentication, redirected back to `/api/auth/callback`
5. Server stores tokens and manages automatic refresh

**Best for:**
- Interactive web sessions
- Development and testing
- Users who can access a browser

### Method 2: Enterprise Token Flow (Manual JWT)

Direct injection of a pre-obtained Microsoft Graph access token.

**How it works:**
1. User obtains a Microsoft Graph token from an enterprise tool (Windows SSO, PowerShell, etc.)
2. Pastes the token in the "Enterprise Token" field in the web UI
3. Server validates and stores the token
4. Token is used directly for Graph API calls (no refresh - must paste new token when expired)

**Best for:**
- Enterprise environments with Windows SSO
- Automated scripts and tools
- Environments where OAuth redirect isn't practical

### Authentication Flow Summary

| Step | OAuth Flow | Enterprise Token Flow |
|------|------------|----------------------|
| 1 | Click "Login with Microsoft" | Paste token in "Enterprise Token" field |
| 2 | Complete Microsoft login | Click "Submit Token" |
| 3 | Server receives tokens via callback | Server validates and stores token |
| 4 | Generate MCP Bearer token | Generate MCP Bearer token |
| 5 | Configure MCP client | Configure MCP client |

Both methods result in the same MCP Bearer token format for client configuration.

---

## Connecting MCP Clients

### Step 1: Authenticate on the Web UI

1. Go to your MCP server URL (e.g., `https://your-server.com` or `http://localhost:3000`)
2. Choose authentication method:
   - **OAuth**: Click "Login with Microsoft" and complete the flow
   - **Enterprise Token**: Paste your Microsoft Graph token and click "Submit"

### Step 2: Generate MCP Bearer Token

After authentication, click "Generate MCP Token" to get your configuration.

### Step 3: Configure Claude Desktop

Copy the generated configuration to your Claude Desktop config file:

**macOS:** `~/Library/Application Support/Claude/claude_desktop_config.json`
**Windows:** `%APPDATA%\Claude\claude_desktop_config.json`

Example configuration:
```json
{
  "mcpServers": {
    "microsoft-365": {
      "command": "npx",
      "args": [
        "mcp-remote",
        "https://your-server.com/api/mcp/sse",
        "--header",
        "Authorization:${MCP_AUTH}"
      ],
      "env": {
        "MCP_AUTH": "Bearer eyJhbGc..."
      }
    }
  }
}
```

### Step 4: Restart Claude Desktop

Restart Claude Desktop to load the new configuration. You should see the Microsoft 365 tools available.

---

## User Isolation & Session Management

Each user's data is completely isolated through the userId-based architecture:

- **Token Storage**: All tokens stored with `userId` as the key prefix
- **Session Isolation**: Each user has their own session context
- **Audit Logging**: All operations logged with `userId` for accountability

```javascript
// Token storage keys use userId for isolation
`ms365:user@company.com:external-graph-token`  // Enterprise token
`ms365:user@company.com:token-source`          // Auth method tracking
`ms365:user@company.com:ms-access-token`       // OAuth token
```

---

## 🤝 Contributing & Support

### **Contributing Guidelines**
1. **Fork the repository** and create a feature branch
2. **Follow the logging patterns** - All new code must implement the 4-tier logging system
3. **Add comprehensive tests** for new functionality
4. **Update documentation** for any API changes
5. **Ensure security** - All user data must be properly isolated

### **Development Setup**
```bash
# Clone and setup development environment
git clone https://github.com/Aanerud/MCP-Microsoft-Office.git
cd MCP-Microsoft-Office
npm install

# Run in development mode with full logging
npm run dev:web

# Run tests
npm test

# Check code quality
npm run lint
```

### **Support & Community**
- 🐛 **Bug Reports**: [GitHub Issues](https://github.com/Aanerud/MCP-Microsoft-Office/issues)
- 💡 **Feature Requests**: [GitHub Discussions](https://github.com/Aanerud/MCP-Microsoft-Office/discussions)

---

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

## 🙏 Acknowledgments

- **Microsoft Graph API** for providing comprehensive Microsoft 365 integration
- **Model Context Protocol (MCP)** for the innovative LLM integration framework
- **Claude AI** for inspiring advanced AI-human collaboration
- **Open Source Community** for the amazing tools and libraries that make this possible

---

<div align="center">

**⭐ Star this repository if you find it useful!**

**🔗 Share with your team and help them work smarter with Microsoft 365!**

</div>
