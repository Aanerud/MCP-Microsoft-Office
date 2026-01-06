# ShadowMCP - Minimal Changes for Shadow User System

## Overview
This document outlines the minimal changes needed to extend the current MCP Microsoft Office project to support a shadow user generator system. The goal is to create "timed agents" that represent personas and perform realistic Microsoft 365 activities automatically.

## Core Concept: Persona-Driven Timed Agents

### Shadow User Workflow
1. **Create Persona**: Generate a realistic company persona with backstory
2. **Assign Credentials**: Link Microsoft 365 credentials to the persona
3. **Deploy Agent**: Create an LLM-powered agent that acts on behalf of the persona
4. **Schedule Activities**: Agent performs realistic activities at timed intervals
5. **Monitor Behavior**: Track and analyze agent activities

## Required Extensions to Current MCP System

### 1. Multi-Auth Support (`src/auth/`)

#### Extend MSAL Service (`src/auth/msal-service.cjs`)
```javascript
// Add to existing MSAL service
class MSALService {
    // Existing methods...
    
    /**
     * Programmatic authentication for shadow users
     * Bypasses interactive login for automated personas
     */
    async authenticateShadowUser(credentials) {
        const { username, password, tenantId } = credentials;
        
        // Use device flow or client credentials flow for automated auth
        const authResult = await this.clientApp.acquireTokenSilent({
            scopes: this.SCOPES,
            account: { username, tenantId }
        });
        
        return authResult;
    }
    
    /**
     * Bulk authentication for multiple shadow users
     */
    async authenticateMultipleShadowUsers(credentialsList) {
        const authPromises = credentialsList.map(creds => 
            this.authenticateShadowUser(creds)
        );
        return Promise.allSettled(authPromises);
    }
}
```

#### New Shadow Auth Controller (`src/api/controllers/shadow-auth-controller.cjs`)
```javascript
/**
 * Handles authentication for shadow users
 */
class ShadowAuthController {
    async bulkAuthenticateUsers(req, res) {
        // Authenticate multiple shadow users at once
        // Store sessions for each persona
    }
    
    async refreshShadowTokens(req, res) {
        // Automatically refresh tokens for all active shadow users
    }
}
```

### 2. Persona Management System

#### New Database Tables
```sql
-- Shadow Personas (new table)
CREATE TABLE shadow_personas (
    persona_id TEXT PRIMARY KEY,
    name TEXT NOT NULL,
    email TEXT NOT NULL,
    department TEXT,
    role TEXT,
    backstory TEXT,
    personality_traits TEXT, -- JSON
    activity_patterns TEXT,  -- JSON
    credentials_encrypted TEXT,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
    is_active BOOLEAN DEFAULT TRUE
);

-- Shadow Activities Schedule (new table)  
CREATE TABLE shadow_activities (
    activity_id TEXT PRIMARY KEY,
    persona_id TEXT NOT NULL,
    activity_type TEXT NOT NULL, -- 'email', 'calendar', 'files'
    activity_config TEXT,        -- JSON configuration
    schedule_pattern TEXT,       -- Cron-like pattern
    next_execution DATETIME,
    last_execution DATETIME,
    is_active BOOLEAN DEFAULT TRUE,
    FOREIGN KEY (persona_id) REFERENCES shadow_personas(persona_id)
);

-- Shadow Activity Logs (new table)
CREATE TABLE shadow_activity_logs (
    log_id TEXT PRIMARY KEY,
    persona_id TEXT NOT NULL,
    activity_id TEXT,
    action_type TEXT NOT NULL,
    action_details TEXT,         -- JSON
    execution_time DATETIME DEFAULT CURRENT_TIMESTAMP,
    status TEXT DEFAULT 'completed', -- 'pending', 'completed', 'failed'
    error_message TEXT,
    FOREIGN KEY (persona_id) REFERENCES shadow_personas(persona_id)
);
```

#### Persona Controller (`src/api/controllers/persona-controller.cjs`)
```javascript
/**
 * Manages shadow personas and their configurations
 */
class PersonaController {
    async createPersona(req, res) {
        // Create new shadow persona with backstory
        // Generate realistic company profile
        // Store encrypted credentials
    }
    
    async listPersonas(req, res) {
        // List all active shadow personas
    }
    
    async updatePersona(req, res) {
        // Update persona details or activity patterns
    }
    
    async bulkCreatePersonas(req, res) {
        // Create multiple personas from CSV/JSON
        // Support same password, different usernames scenario
    }
}
```

### 3. Activity Simulation Engine

#### Shadow Activity Service (`src/core/shadow-activity-service.cjs`)
```javascript
/**
 * Orchestrates shadow user activities
 */
class ShadowActivityService {
    constructor(graphClient, llmService) {
        this.graphClient = graphClient;
        this.llmService = llmService;
        this.activeAgents = new Map(); // persona_id -> agent instance
    }
    
    /**
     * Create and deploy an agent for a persona
     */
    async deployPersonaAgent(persona) {
        const agent = new PersonaAgent(persona, this.graphClient, this.llmService);
        this.activeAgents.set(persona.persona_id, agent);
        
        // Start the agent's activity loop
        agent.startActivityLoop();
        
        return agent;
    }
    
    /**
     * Schedule activities for all active personas
     */
    async scheduleActivities() {
        // Check database for due activities
        // Execute activities through appropriate agents
    }
}
```

#### Persona Agent (`src/shadow/persona-agent.cjs`)
```javascript
/**
 * LLM-powered agent that acts on behalf of a shadow persona
 */
class PersonaAgent {
    constructor(persona, graphClient, llmService) {
        this.persona = persona;
        this.graphClient = graphClient;
        this.llmService = llmService;
        this.isActive = false;
    }
    
    async startActivityLoop() {
        this.isActive = true;
        this.scheduleNextActivity();
    }
    
    async performActivity(activityType, context = {}) {
        // Use LLM to generate realistic activity based on persona
        const prompt = this.buildActivityPrompt(activityType, context);
        const response = await this.llmService.generateResponse(prompt);
        
        // Execute the activity using MCP tools
        return this.executeActivity(response, activityType);
    }
    
    async executeActivity(llmResponse, activityType) {
        switch (activityType) {
            case 'email':
                return this.sendRealisticEmail(llmResponse);
            case 'calendar':
                return this.manageCalendar(llmResponse);
            case 'files':
                return this.manageFiles(llmResponse);
        }
    }
}
```

### 4. API Extensions (`src/api/routes.cjs`)

#### New Shadow Routes
```javascript
// Add to existing routes.cjs
function registerShadowRoutes(v1Router) {
    // Shadow Persona Management
    const shadowRouter = express.Router();
    shadowRouter.use(controllerLogger());
    
    // Persona CRUD operations
    shadowRouter.post('/personas', placeholderRateLimit, personaController.createPersona);
    shadowRouter.get('/personas', personaController.listPersonas);
    shadowRouter.put('/personas/:id', personaController.updatePersona);
    shadowRouter.delete('/personas/:id', personaController.deletePersona);
    shadowRouter.post('/personas/bulk', placeholderRateLimit, personaController.bulkCreatePersonas);
    
    // Activity management
    shadowRouter.post('/personas/:id/activities', placeholderRateLimit, activityController.scheduleActivity);
    shadowRouter.get('/personas/:id/activities', activityController.getActivities);
    shadowRouter.get('/personas/:id/logs', activityController.getActivityLogs);
    
    // Bulk operations
    shadowRouter.post('/bulk-authenticate', placeholderRateLimit, shadowAuthController.bulkAuthenticateUsers);
    shadowRouter.post('/start-all-agents', placeholderRateLimit, activityController.startAllAgents);
    shadowRouter.post('/stop-all-agents', placeholderRateLimit, activityController.stopAllAgents);
    
    v1Router.use('/shadow', shadowRouter);
}

// Integrate with existing route registration
function registerRoutes(router) {
    // ... existing routes ...
    
    registerShadowRoutes(v1);
    
    // ... rest of existing routes ...
}
```

### 5. Simple UI Extensions (`src/renderer/`)

#### Shadow Management UI (`src/renderer/shadow-manager.js`)
```javascript
/**
 * Simple UI for managing shadow personas
 */
class ShadowManager {
    constructor() {
        this.personas = [];
        this.initializeUI();
    }
    
    initializeUI() {
        // Add "Shadow Users" tab to existing UI
        // Simple + button to add new personas
        // List view of active personas with controls
    }
    
    async addPersona() {
        // Show modal/form for new persona creation
        // Support both individual and bulk creation
    }
    
    async startAgent(personaId) {
        // Start the agent for a specific persona
    }
    
    async viewActivity(personaId) {
        // Show activity logs for a persona
    }
}
```

## Minimal File Changes Required

### Files to Modify
1. **`src/auth/msal-service.cjs`** - Add programmatic auth methods
2. **`src/api/routes.cjs`** - Add shadow routes registration
3. **`src/core/database-migrations.cjs`** - Add shadow tables
4. **`src/renderer/index.html`** - Add shadow management tab

### Files to Create
1. **`src/shadow/persona-agent.cjs`** - LLM-powered persona agent
2. **`src/core/shadow-activity-service.cjs`** - Activity orchestration
3. **`src/api/controllers/persona-controller.cjs`** - Persona management
4. **`src/api/controllers/shadow-auth-controller.cjs`** - Shadow auth handling
5. **`src/renderer/shadow-manager.js`** - UI for shadow management

## Integration Strategy

### Phase 1: Core Infrastructure (Minimal Viable Product)
```javascript
// Essential components for basic shadow user functionality
- Extend MSAL service for programmatic auth
- Create persona database tables
- Basic persona CRUD API
- Simple agent with timer-based activities
```

### Phase 2: Intelligence Layer
```javascript
// Add LLM-powered behavior
- Integrate LLM service for realistic content generation
- Implement activity patterns based on persona traits
- Add sophisticated scheduling system
```

### Phase 3: Management Interface
```javascript
// User-friendly management
- Complete UI for persona management
- Bulk import/export functionality
- Activity monitoring dashboard
- Performance analytics
```

## Architecture Integration

### MCP Tool Integration
The shadow system leverages existing MCP tools:
- **Existing tools remain unchanged** - No breaking changes to current MCP interface
- **Shadow agents use same API** - Agents call existing `/v1/mail`, `/v1/calendar`, etc.
- **Authentication isolation** - Each shadow user has isolated auth context
- **Activity attribution** - All activities properly logged per persona

### Deployment Model
```
Current MCP Server (Port 3001)
├── Existing MCP functionality (unchanged)
├── Shadow management API (/v1/shadow/*)  
├── Extended auth for programmatic access
└── Background shadow agents

Shadow Control Interface
├── Simple web UI for persona management
├── + button for adding personas
├── Bulk import for CSV/list of users
└── Real-time activity monitoring
```

## Benefits of This Approach

### 1. **Minimal Disruption**
- Existing MCP functionality remains completely unchanged
- No breaking changes to current API
- Backward compatible with existing MCP clients

### 2. **Leverages Existing Infrastructure**
- Uses existing auth, session, and API systems
- Benefits from existing monitoring and logging
- Maintains same security and error handling patterns

### 3. **Scalable Architecture** 
- Can handle multiple concurrent shadow users
- Built on existing multi-user foundation
- Leverages existing database and caching systems

### 4. **Realistic Behavior**
- LLM-powered agents generate contextually appropriate content
- Persona-driven activities feel authentic
- Configurable activity patterns for different scenarios

This approach transforms the existing MCP server into a powerful shadow user platform while maintaining all existing functionality and requiring minimal code changes.

## Potential Architecture Diagram

```
┌─────────────────────────────────────────────────────────────────────────────────┐
│                          SHADOW M365 ECOSYSTEM                                  │
└─────────────────────────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────────────────────────┐
│                              SHADOW CONTROL CENTER                              │
│                                 (Port 3002)                                     │
├─────────────────────────────────────────────────────────────────────────────────┤
│  ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐                 │
│  │   Persona UI    │  │  Bulk Import    │  │   Monitoring    │                 │
│  │  ┌─────────────┐│  │  ┌─────────────┐│  │  ┌─────────────┐│                 │
│  │  │ + Add User  ││  │  │ CSV Upload  ││  │  │ Activity    ││                 │
│  │  │ Edit Users  ││  │  │ Same Pass   ││  │  │ Dashboard   ││                 │
│  │  │ Start/Stop  ││  │  │ Multi-User  ││  │  │ Performance ││                 │
│  │  └─────────────┘│  │  └─────────────┘│  │  └─────────────┘│                 │
│  └─────────────────┘  └─────────────────┘  └─────────────────┘                 │
└─────────────────────────────────────────────────────────────────────────────────┘
                              │
                              │ API Calls
                              ▼
┌─────────────────────────────────────────────────────────────────────────────────┐
│                          ENHANCED MCP SERVER                                    │
│                          (Existing Port 3001)                                  │
├─────────────────────────────────────────────────────────────────────────────────┤
│                                                                                 │
│  ┌─────────────────────────────── EXISTING MCP LAYER ──────────────────────────┐ │
│  │                                                                             │ │
│  │  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐       │ │
│  │  │    /tools   │  │ /v1/mail/*  │  │/v1/calendar │  │ /v1/files/* │       │ │
│  │  │   (MCP)     │  │   (Email)   │  │  (Events)   │  │  (OneDrive) │       │ │
│  │  └─────────────┘  └─────────────┘  └─────────────┘  └─────────────┘       │ │
│  │                                                                             │ │
│  │  ┌─────────────┐  ┌─────────────┐  Current Auth & Session Management       │ │
│  │  │/v1/people/* │  │  /v1/logs   │  ┌─────────────────────────────────────┐ │ │
│  │  │(Directory)  │  │ (Activity)  │  │ • Multi-user sessions               │ │ │
│  │  └─────────────┘  └─────────────┘  │ • Token management                  │ │ │
│  │                                    │ • Database-backed storage          │ │ │
│  └─────────────────────────────────── └─────────────────────────────────────┘ │ │
│                                                                                 │
│  ┌─────────────────────────────── NEW SHADOW EXTENSIONS ───────────────────────┐ │
│  │                                                                             │ │
│  │  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐       │ │
│  │  │/v1/shadow/  │  │   Shadow    │  │   Persona   │  │    Bulk     │       │ │
│  │  │ personas/*  │  │    Auth     │  │   Agent     │  │   Import    │       │ │
│  │  │             │  │  Service    │  │  Manager    │  │   Service   │       │ │
│  │  └─────────────┘  └─────────────┘  └─────────────┘  └─────────────┘       │ │
│  │                                                                             │ │
│  └─────────────────────────────────────────────────────────────────────────────┘ │
└─────────────────────────────────────────────────────────────────────────────────┘
                              │
                              │ Database Operations
                              ▼
┌─────────────────────────────────────────────────────────────────────────────────┐
│                            DATABASE LAYER                                       │
├─────────────────────────────────────────────────────────────────────────────────┤
│                                                                                 │
│  ┌─────────────────────────── EXISTING TABLES ──────────────────────────────┐   │
│  │  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐                       │   │
│  │  │user_sessions│  │  user_logs  │  │   devices   │                       │   │
│  │  │ (encrypted) │  │ (activity)  │  │ (auth flow) │                       │   │
│  │  └─────────────┘  └─────────────┘  └─────────────┘                       │   │
│  └─────────────────────────────────────────────────────────────────────────────┘   │
│                                                                                 │
│  ┌─────────────────────────── NEW SHADOW TABLES ────────────────────────────┐   │
│  │  ┌─────────────┐  ┌─────────────┐  ┌─────────────┐                       │   │
│  │  │shadow_      │  │shadow_      │  │shadow_      │                       │   │
│  │  │personas     │  │activities   │  │activity_logs│                       │   │
│  │  │(backstory)  │  │(schedule)   │  │(execution)  │                       │   │
│  │  └─────────────┘  └─────────────┘  └─────────────┘                       │   │
│  └─────────────────────────────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────────────────────────┘
                              │
                              │ Background Processing
                              ▼
┌─────────────────────────────────────────────────────────────────────────────────┐
│                           SHADOW AGENT RUNTIME                                  │
├─────────────────────────────────────────────────────────────────────────────────┤
│                                                                                 │
│  ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐                 │
│  │   Persona A     │  │   Persona B     │  │   Persona C     │                 │
│  │                 │  │                 │  │                 │                 │
│  │ ┌─────────┐     │  │ ┌─────────┐     │  │ ┌─────────┐     │                 │
│  │ │ LLM     │────────│─│ LLM     │────────│─│ LLM     │     │                 │
│  │ │ Agent   │     │  │ │ Agent   │     │  │ │ Agent   │     │                 │
│  │ └─────────┘     │  │ └─────────┘     │  │ └─────────┘     │                 │
│  │       │         │  │       │         │  │       │         │                 │
│  │ ┌─────v─────┐   │  │ ┌─────v─────┐   │  │ ┌─────v─────┐   │                 │
│  │ │Scheduler  │   │  │ │Scheduler  │   │  │ │Scheduler  │   │                 │
│  │ │(Timed)    │   │  │ │(Timed)    │   │  │ │(Timed)    │   │                 │
│  │ └───────────┘   │  │ └───────────┘   │  │ └───────────┘   │                 │
│  └─────────────────┘  └─────────────────┘  └─────────────────┘                 │
│           │                     │                     │                         │
│           └─────────────────────┼─────────────────────┘                         │
│                                 │                                               │
│                       ┌─────────v─────────┐                                     │
│                       │  Activity Queue   │                                     │
│                       │  - Send emails    │                                     │
│                       │  - Schedule meets │                                     │
│                       │  - Create files   │                                     │
│                       │  - Update status  │                                     │
│                       └───────────────────┘                                     │
└─────────────────────────────────────────────────────────────────────────────────┘
                              │
                              │ Microsoft Graph API Calls
                              ▼
┌─────────────────────────────────────────────────────────────────────────────────┐
│                          MICROSOFT 365 TENANT                                   │
├─────────────────────────────────────────────────────────────────────────────────┤
│                                                                                 │
│  ┌─────────────────┐  ┌─────────────────┐  ┌─────────────────┐                 │
│  │    Exchange     │  │   SharePoint    │  │      Teams      │                 │
│  │   (Email/Cal)   │  │     (Files)     │  │   (Messages)    │                 │
│  │                 │  │                 │  │                 │                 │
│  │ • Persona A     │  │ • Persona A     │  │ • Persona A     │                 │
│  │   mailbox       │  │   documents     │  │   presence      │                 │
│  │ • Persona B     │  │ • Persona B     │  │ • Persona B     │                 │
│  │   mailbox       │  │   documents     │  │   presence      │                 │
│  │ • Persona C     │  │ • Persona C     │  │ • Persona C     │                 │
│  │   mailbox       │  │   documents     │  │   presence      │                 │
│  └─────────────────┘  └─────────────────┘  └─────────────────┘                 │
└─────────────────────────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────────────────────────┐
│                              DATA FLOW EXAMPLE                                  │
├─────────────────────────────────────────────────────────────────────────────────┤
│                                                                                 │
│  1. Developer adds persona "John Smith" via + button                            │
│  2. System creates persona record with backstory & credentials                  │
│  3. Shadow Agent Service deploys LLM agent for "John Smith"                    │
│  4. Agent schedules realistic activities (9am email check, 2pm meeting)        │
│  5. At scheduled time, agent calls LLM: "Act as John, send status update"      │
│  6. LLM generates realistic email content based on John's persona              │
│  7. Agent calls /v1/mail/send with generated content                           │
│  8. MCP server authenticates as John and sends via Microsoft Graph             │
│  9. Activity logged to shadow_activity_logs for monitoring                     │
│  10. Process repeats for all personas throughout the day                       │
│                                                                                 │
└─────────────────────────────────────────────────────────────────────────────────┘
```

## Implementation Workflow

### Step 1: Extend Existing MCP Server
```bash
# Add shadow extensions to current project
src/shadow/persona-agent.cjs
src/api/controllers/persona-controller.cjs  
src/core/shadow-activity-service.cjs
```

### Step 2: Deploy Shadow Control Interface
```bash
# Create separate simple web app
npm create vue@latest shadow-control
# Simple interface connecting to MCP server API
# + button, list view, bulk import
```

### Step 3: Configure Personas
```javascript
// Example persona configuration
{
  "name": "John Smith",
  "email": "john.smith@company.com", 
  "role": "Marketing Manager",
  "department": "Marketing",
  "backstory": "5 years experience, focused on digital campaigns",
  "personality": "professional, detail-oriented, collaborative",
  "activityPatterns": {
    "emailFrequency": "high",
    "meetingPreference": "mornings", 
    "workingHours": "8-17"
  }
}
```

This architecture maintains complete compatibility with existing MCP functionality while adding powerful shadow user capabilities through minimal, well-isolated extensions.