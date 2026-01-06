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

### 2. **Add Your Microsoft 365 Credentials** 🔑
Edit the auto-generated `.env` file:
```bash
MICROSOFT_CLIENT_ID=your_client_id_here
MICROSOFT_TENANT_ID=your_tenant_id_here
```

📋 **Get these from:** [Azure App Registration Portal](https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)

### 3. **Launch Your Server** 🎯
```bash
npm run dev:web  # Full development mode with comprehensive logging
```

🌐 **Access your server at:** `http://localhost:3000`

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
