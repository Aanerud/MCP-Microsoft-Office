# Microsoft 365 MCP Gateway Product Roadmap

This document outlines the development roadmap for the Microsoft 365 MCP Gateway, informed by our practical experience implementing the system. It consolidates and refines the original phase documents with a more streamlined approach focused on expanding Microsoft Graph integration.

## Overview

The Microsoft 365 MCP Gateway connects Claude to Microsoft 365 services through a unified API and a Model Context Protocol (MCP) adapter. The system allows Claude to access emails, calendar events, files, and contacts with proper authentication and appropriate formatting.

## System Architecture

```
┌────────────────┐      ┌───────────────────┐     ┌─────────────────────┐
│                │      │                   │     │                     │
│  Claude LLM    │◄────►│  MCP Adapter      │◄───►│  Dev Server         │
│                │      │  (mcp-adapter.cjs)│     │  (src/main/dev-server.cjs) │
│                │      │                   │     │                     │
└────────────────┘      └───────────────────┘     └──────────┬──────────┘
                                                             │
                                                             ▼
┌────────────────────────────┐                     ┌─────────────────────┐
│                            │                     │                     │
│  Microsoft Graph Services  │◄────────────────────┤  Auth & Storage     │
│  (src/graph/*)             │                     │  Services           │
│                            │                     │                     │
└──────────┬─────────────────┘                     └─────────────────────┘
           │
           ▼
┌────────────────────────────┐
│                            │
│  Microsoft Graph API       │
│                            │
└────────────────────────────┘
```

## End-to-End Tool Implementation Process

When adding a new Microsoft Graph capability to the system, the following components must be updated:

1. **Graph Service** (`src/graph/`): Create or update service to interact with Microsoft Graph API
2. **Data Normalizers** (`src/graph/normalizers.cjs`): Add normalizers for new data types
3. **Module Implementation** (`src/modules/`): Implement module capabilities
4. **API Controller** (`src/api/controllers/`): Create API endpoint(s) for the capability
5. **Route Registration** (`src/api/routes.cjs`): Register new endpoints
6. **Tool Definition** (`src/core/tools-service.cjs`): Define tool parameters and descriptions
7. **MCP Adapter Support** (`mcp-adapter.cjs`): Add capability mapping and parameter handling

## Development Phases

### Phase 1: Foundation (Completed)
- ✅ Core infrastructure
- ✅ Authentication with Microsoft Graph
- ✅ Basic mail, calendar, and files access
- ✅ MCP adapter implementation

### Phase 2: Enhanced Functionality (Current)

#### People & Contacts
- ✅ People search and lookup
- ✅ Contact management
- ✅ User profile access

#### Improved Mail
- 🔄 Better attachment handling
- 🔄 Email drafts
- 🔄 Enhanced search capabilities
- 🔄 Email categorization

#### Advanced Calendar
- ✅ Meeting scheduling with availability check
- 🔄 Meeting response handling
- 🔄 Recurring meeting support
- 🔄 Calendar sharing and permissions

#### Files Expansion
- 🔄 File content extraction and preview
- 🔄 Search within documents
- 🔄 File metadata enrichment
- 🔄 File sharing controls

### Phase 3: Advanced Integration

#### Teams Integration
- Meeting chat access
- Teams message access
- Channel and team management
- Teams file integration

#### OneNote Integration
- Notebook access
- Page content retrieval
- Note creation and editing
- Media embedding

#### Office Document Integration
- Word document creation/editing
- Excel spreadsheet operations
- PowerPoint presentation handling
- Real-time collaboration support

#### Admin and Security Features
- Tenant information access
- Security alerts and compliance
- User management functions
- Conditional access policies

## Adding a New Tool End-to-End Example

When adding a new Microsoft Graph capability (e.g., Teams messages), follow these steps:

### 1. Implement Graph Service

Create or update a service in `src/graph/` that handles the specific Graph API calls:

```javascript
// src/graph/teams-service.cjs
async function getTeamsMessages(graphClient, channelId, options = {}) {
  try {
    const response = await graphClient
      .api(`/teams/${options.teamId}/channels/${channelId}/messages`)
      .top(options.limit || 50)
      .get();
    
    return {
      messages: response.value.map(message => normalizeTeamsMessage(message)),
      nextLink: response["@odata.nextLink"]
    };
  } catch (error) {
    throw new Error(`Failed to get Teams messages: ${error.message}`);
  }
}

module.exports = {
  getTeamsMessages,
  // other Teams-related methods
};
```

### 2. Create Data Normalizer

Add a normalizer function in `src/graph/normalizers.cjs`:

```javascript
function normalizeTeamsMessage(message) {
  if (!message || typeof message !== 'object') {
    throw new Error('Invalid Teams message for normalization');
  }
  
  return {
    id: message.id,
    type: 'teams_message',
    content: message.body.content,
    contentType: message.body.contentType,
    createdDateTime: message.createdDateTime,
    lastModifiedDateTime: message.lastModifiedDateTime,
    from: message.from ? {
      user: message.from.user ? {
        id: message.from.user.id,
        displayName: message.from.user.displayName,
        userIdentityType: message.from.user.userIdentityType
      } : undefined
    } : undefined,
    importance: message.importance || 'normal',
    reactions: Array.isArray(message.reactions) ? 
      message.reactions.map(reaction => ({
        reactionType: reaction.reactionType,
        count: reaction.count
      })) : [],
    attachments: Array.isArray(message.attachments) ?
      message.attachments.map(attachment => ({
        id: attachment.id,
        contentType: attachment.contentType,
        contentUrl: attachment.contentUrl,
        name: attachment.name,
        thumbnailUrl: attachment.thumbnailUrl
      })) : []
  };
}

// Add to module.exports
module.exports = {
  normalizeEmail,
  normalizeFile,
  normalizeEvent,
  normalizeUser,
  normalizePerson,
  normalizeTeamsMessage  // Add new normalizer
};
```

### 3. Implement Module Capability

Create a Teams module or update an existing one:

```javascript
// src/modules/teams/index.cjs
module.exports = {
  id: 'teams',
  name: 'Microsoft Teams',
  description: 'Access Microsoft Teams data',
  
  capabilities: ['getMessages', 'sendMessage', 'listChannels'],
  
  init(services) {
    this.teamsService = services.teamsService;
    this.graphClient = services.graphClient;
    return this;
  },
  
  async getMessages(params) {
    try {
      return await this.teamsService.getTeamsMessages(
        this.graphClient, 
        params.channelId,
        params
      );
    } catch (error) {
      throw new Error(`Teams module error: ${error.message}`);
    }
  },
  
  // Other capabilities implementation...
};
```

### 4. Create API Controller

Implement a controller in `src/api/controllers/`:

```javascript
// src/api/controllers/teams-controller.cjs
function createTeamsController(services) {
  const { moduleRegistry, logger } = services;
  
  return {
    async getMessages(req, res) {
      try {
        const teamsModule = moduleRegistry.getModule('teams');
        if (!teamsModule) {
          return res.status(404).json({ error: 'Teams module not available' });
        }
        
        const result = await teamsModule.getMessages({
          channelId: req.params.channelId,
          teamId: req.query.teamId,
          limit: req.query.limit ? parseInt(req.query.limit, 10) : 50
        });
        
        res.json(result);
      } catch (error) {
        logger.error('Teams controller error:', error);
        res.status(500).json({ error: error.message });
      }
    },
    
    // Other controller methods...
  };
}

module.exports = createTeamsController;
```

### 5. Register Routes

Update the routes in `src/api/routes.cjs`:

```javascript
// In registerRoutes function
const teamsController = require('./controllers/teams-controller.cjs')(services);

// Teams routes
router.get('/v1/teams/:teamId/channels', teamsController.listChannels);
router.get('/v1/teams/channels/:channelId/messages', teamsController.getMessages);
router.post('/v1/teams/channels/:channelId/messages', teamsController.sendMessage);
```

### 6. Define Tool in Tools Service

Add the tool definition in `src/core/tools-service.cjs`:

```javascript
// In generateToolDefinition function
case 'getMessages':
  toolDef.description = 'Get messages from a Microsoft Teams channel';
  toolDef.endpoint = '/api/v1/teams/channels/:channelId/messages';
  toolDef.parameters = {
    channelId: { type: 'string', description: 'ID of the Teams channel' },
    teamId: { type: 'string', description: 'ID of the team', optional: true },
    limit: { type: 'number', description: 'Maximum number of messages to return', optional: true }
  };
  break;
```

### 7. Update MCP Adapter

Add handling in `mcp-adapter.cjs`:

```javascript
// In executeModuleMethod function's switch statement
case 'teams.getMessages':
  apiPath = `/v1/teams/channels/${params.channelId}/messages`;
  apiMethod = 'GET';
  // Transform any parameters if needed
  break;

// In handleToolCall function's fallback switch case
case 'getTeamsMessages':
  // Get messages from a Teams channel
  return await executeModuleMethod('teams', 'getMessages', toolArgs);
```

## Prioritization Strategy

1. **Most requested/used features first**: Focus on capabilities with highest user value
2. **Complexity vs. Impact**: Prioritize lower complexity, high impact features
3. **Foundation before extensions**: Ensure core features are solid before adding more
4. **API stability**: Prioritize stable Graph API endpoints over beta features

## Future Considerations

- **Multi-tenant support**: Support for multiple Microsoft tenants
- **Advanced security features**: Conditional access, data loss prevention
- **Mobile application**: Native mobile app with push notifications
- **More LLM integrations**: Support for other LLM providers beyond Claude
- **Real-time collaboration**: Live co-editing of documents and sharing

## Regular Review Process

This roadmap should be reviewed and updated quarterly based on:
- User feedback and feature requests
- Changes to Microsoft Graph API and services
- Advances in LLM capabilities
- Technical debt and performance metrics