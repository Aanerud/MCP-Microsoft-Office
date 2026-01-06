# ShadowM365 - Understanding the Current MCP Microsoft Office Architecture

## Overview
This document provides a comprehensive understanding of the current MCP Microsoft Office project architecture, focusing on components that enable multi-user Microsoft 365 integration through MCP (Model Context Protocol).

## Authentication Architecture

### Multi-Modal Authentication System
The project supports two primary authentication flows:

#### 1. Web Authentication Flow (`src/auth/msal-service.cjs`)
- **MSAL Integration**: Uses Microsoft Authentication Library for JavaScript
- **OAuth 2.0/OpenID Connect**: Standard Microsoft identity platform integration
- **PKCE Support**: Proof Key for Code Exchange for enhanced security
- **Multi-User Sessions**: Each user maintains isolated authentication context

```javascript
// Key authentication scopes for Microsoft 365 access:
const SCOPES = [
    'User.Read',           // Basic profile access
    'Calendars.ReadWrite', // Full calendar management
    'Mail.ReadWrite',      // Email read/write operations
    'Mail.Send',           // Email sending capabilities
    'Files.ReadWrite'      // OneDrive/SharePoint file operations
];
```

#### 2. Device Authentication Flow (`src/auth/device-auth-controller.cjs`)
- **Device Registration**: `/auth/device/register` - Register new devices
- **Device Authorization**: `/auth/device/authorize` - Authorize device access
- **Token Polling**: `/auth/device/token` - Poll for authorization completion
- **Token Refresh**: `/auth/device/refresh` - Automatic token renewal
- **MCP Token Generation**: `/auth/generate-mcp-token` - Generate tokens for MCP clients

### Session Management (`src/core/session-service.cjs`)

#### Database-Backed Sessions
- **Encrypted Storage**: AES-256-CBC encryption for sensitive session data
- **Token Association**: Microsoft Graph tokens linked to session IDs
- **User Context Isolation**: Each session maintains separate user context
- **Automatic Cleanup**: Session expiration and cleanup mechanisms

#### Session Schema
```sql
CREATE TABLE user_sessions (
    session_id TEXT PRIMARY KEY,
    session_secret TEXT NOT NULL,
    expires_at INTEGER NOT NULL,
    user_agent TEXT,
    ip_address TEXT,
    microsoft_token TEXT,          -- Encrypted Microsoft Graph token
    microsoft_refresh_token TEXT,  -- Encrypted refresh token
    user_info TEXT                 -- User profile information
);
```

## API Architecture (`src/api/routes.cjs`)

### Versioned API Structure
- **Base Path**: `/v1/` for all versioned endpoints
- **Authentication Middleware**: `requireAuth` applied to all v1 routes
- **Request Logging**: Comprehensive request/response logging
- **Error Handling**: Structured error responses with monitoring

### Core API Endpoints

#### Mail Operations (`/v1/mail/`)
- `GET /` - Retrieve mail messages
- `POST /send` - Send new email
- `GET /search` - Search mail content
- `GET /attachments` - Get mail attachments
- `PATCH /:id/read` - Mark message as read
- `POST /flag` - Flag/unflag messages
- `POST /:id/attachments` - Add attachments
- `DELETE /:id/attachments/:attachmentId` - Remove attachments

#### Calendar Operations (`/v1/calendar/`)
- `GET /` - Get calendar events
- `POST /events` - Create new events
- `PUT /events/:id` - Update existing events
- `POST /availability` - Check availability
- `POST /events/:id/accept` - Accept meeting invitations
- `POST /events/:id/decline` - Decline meeting invitations
- `POST /findMeetingTimes` - Find optimal meeting times
- `GET /rooms` - Get available meeting rooms

#### File Operations (`/v1/files/`)
- `GET /` - List files and folders
- `POST /upload` - Upload new files
- `GET /search` - Search file content
- `GET /metadata` - Get file metadata
- `GET /content` - Download file content
- `POST /content` - Set file content
- `POST /share` - Create sharing links
- `GET /sharing` - Get sharing permissions

#### People Operations (`/v1/people/`)
- `GET /` - Get relevant people
- `GET /find` - Search directory
- `GET /:id` - Get person details

### MCP Integration

#### Tools Manifest (`/tools`)
- **Dynamic Tool Discovery**: Automatically generates available tools
- **MCP Compatibility**: Follows MCP specification for tool definitions
- **Real-time Updates**: Tools reflect current system capabilities

#### Request Flow
1. **Authentication Check**: `requireAuth` middleware validates session
2. **Request Logging**: Comprehensive request tracking
3. **Controller Processing**: Business logic execution
4. **Graph API Calls**: Microsoft Graph client handles API interactions
5. **Response Formation**: Structured JSON responses
6. **Activity Logging**: User activity tracking for audit

## Database Architecture

### Multi-Database Support
- **SQLite**: Default for development and small deployments
- **PostgreSQL**: Production-ready with advanced features
- **MySQL**: Alternative production option

### Key Tables for Shadow User System

#### User Sessions (Migration 003)
```sql
CREATE TABLE user_sessions (
    session_id TEXT PRIMARY KEY,
    session_secret TEXT NOT NULL,
    expires_at INTEGER NOT NULL,
    user_agent TEXT,
    ip_address TEXT,
    microsoft_token TEXT,
    microsoft_refresh_token TEXT,
    user_info TEXT
);
```

#### User Activity Logs (Migration 004)
```sql
CREATE TABLE user_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    user_id TEXT NOT NULL,
    level TEXT NOT NULL,
    message TEXT NOT NULL,
    category TEXT,
    context TEXT,
    timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
    device_id TEXT
);
```

#### Device Management (Migration 002)
```sql
CREATE TABLE devices (
    device_id TEXT PRIMARY KEY,
    device_secret TEXT NOT NULL,
    user_id TEXT,
    device_name TEXT,
    device_type TEXT,
    metadata TEXT
);
```

## Microsoft Graph Integration (`src/graph/`)

### Graph Client (`src/graph/graph-client.cjs`)
- **Automatic Retry Logic**: Handles 429 rate limiting with exponential backoff
- **Batch Operations**: Efficient bulk API calls
- **Error Recovery**: Comprehensive error handling and recovery
- **Performance Monitoring**: Built-in metrics and performance tracking

### Service Layer
- **Mail Service**: `src/graph/mail-service.cjs` - Email operations
- **Calendar Service**: `src/graph/calendar-service.cjs` - Calendar management
- **Files Service**: `src/graph/files-service.cjs` - File operations
- **People Service**: `src/graph/people-service.cjs` - Directory operations

## Monitoring and Logging

### Four-Pattern Logging Architecture
1. **Development Debug Logs**: Technical debugging information
2. **User Activity Logs**: Track user actions and behaviors
3. **Infrastructure Error Logging**: System-level error tracking
4. **User Error Tracking**: User-specific error attribution

### Error Service (`src/core/error-service.cjs`)
- **Structured Error Creation**: Consistent error formatting
- **Context Preservation**: Maintains error context and stack traces
- **User Attribution**: Links errors to specific users/sessions

## Security Features

### Token Security
- **AES-256-CBC Encryption**: All tokens encrypted at rest
- **Secure Session Management**: Session secrets and proper expiration
- **CORS Configuration**: Proper cross-origin request handling

### Request Security
- **Authentication Middleware**: All API endpoints protected
- **Rate Limiting Preparation**: Infrastructure for rate limiting (placeholders in place)
- **Input Validation**: Request validation and sanitization

## Scalability Considerations

### Horizontal Scaling Support
- **Database Abstraction**: Multi-database support for scaling
- **Session Isolation**: Each user operates independently
- **Async Operations**: Non-blocking operations throughout
- **Connection Pooling**: Efficient resource management

### Performance Optimization
- **Batch Operations**: Efficient bulk Microsoft Graph operations
- **Caching Layer**: Built-in caching mechanisms
- **Request Optimization**: Optimized request patterns

## Key Components for Shadow User Integration

### 1. Multi-User Session Management
- **Isolated User Contexts**: Each user maintains separate state
- **Concurrent Operations**: Multiple users can operate simultaneously
- **Persistent Sessions**: Database-backed session storage

### 2. Comprehensive API Coverage
- **Rich Microsoft 365 Operations**: All major M365 services covered
- **Realistic Activities**: Full range of user activities available
- **MCP Tool Integration**: Seamless integration with MCP protocol

### 3. Authentication Infrastructure
- **Programmatic Authentication**: Device flow supports automated auth
- **Token Management**: Automatic refresh and renewal
- **Multi-Credential Support**: Infrastructure supports multiple users

### 4. Monitoring and Analytics
- **Activity Tracking**: Comprehensive user activity logging
- **Performance Metrics**: Built-in performance monitoring
- **Error Tracking**: Detailed error attribution and tracking

This architecture provides an excellent foundation for implementing shadow users, as it already handles the most complex aspects: multi-user authentication, session isolation, comprehensive Microsoft 365 integration, and robust monitoring systems.