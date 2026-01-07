/**
 * @fileoverview Initializes and provides all API dependencies (modules, services) for injection.
 * Ensures proper dependency injection for testability and modularity.
 */

const moduleRegistry = require('../modules/module-registry.cjs');
const MailModule = require('../modules/mail/index.cjs');
const CalendarModule = require('../modules/calendar/index.cjs');
const FilesModule = require('../modules/files/index.cjs');
const PeopleModule = require('../modules/people/index.cjs');
const SearchModule = require('../modules/search/index.cjs');
const TeamsModule = require('../modules/teams/index.cjs');
const TodoModule = require('../modules/todo/index.cjs');
const ContactsModule = require('../modules/contacts/index.cjs');
const GroupsModule = require('../modules/groups/index.cjs');
const cacheService = require('../core/cache-service.cjs');
const eventService = require('../core/event-service.cjs');
const graphClientFactory = require('../graph/graph-client.cjs');
const calendarService = require('../graph/calendar-service.cjs');
const mailService = require('../graph/mail-service.cjs');
const filesService = require('../graph/files-service.cjs');
const peopleService = require('../graph/people-service.cjs');
const searchService = require('../graph/search-service.cjs');
const teamsService = require('../graph/teams-service.cjs');
const todoService = require('../graph/todo-service.cjs');
const contactsService = require('../graph/contacts-service.cjs');
const groupsService = require('../graph/groups-service.cjs');

// Import error and monitoring services
const ErrorService = require('../core/error-service.cjs');
const MonitoringService = require('../core/monitoring-service.cjs');

// Initialize modules with their dependencies
const mailModule = MailModule.init({ graphService: mailService, cacheService, eventService, errorService: ErrorService, monitoringService: MonitoringService });
const calendarModule = CalendarModule.init({ graphService: calendarService, cacheService, eventService, errorService: ErrorService, monitoringService: MonitoringService });
const filesModule = FilesModule.init({ graphService: filesService, cacheService, eventService, errorService: ErrorService, monitoringService: MonitoringService });
const peopleModule = PeopleModule.init({ graphService: peopleService, cacheService, eventService, errorService: ErrorService, monitoringService: MonitoringService });
const searchModule = SearchModule.init({ searchService, cacheService, eventService, errorService: ErrorService, monitoringService: MonitoringService });
const teamsModule = TeamsModule.init({ teamsService, cacheService, eventService, errorService: ErrorService, monitoringService: MonitoringService });
const todoModule = TodoModule.init({ todoService, cacheService, eventService, errorService: ErrorService, monitoringService: MonitoringService });
const contactsModule = ContactsModule.init({ contactsService, cacheService, eventService, errorService: ErrorService, monitoringService: MonitoringService });
const groupsModule = GroupsModule.init({ groupsService, cacheService, eventService, errorService: ErrorService, monitoringService: MonitoringService });

// Register modules
moduleRegistry.registerModule(mailModule);
moduleRegistry.registerModule(calendarModule);
moduleRegistry.registerModule(filesModule);
moduleRegistry.registerModule(peopleModule);
moduleRegistry.registerModule(searchModule);
moduleRegistry.registerModule(teamsModule);
moduleRegistry.registerModule(todoModule);
moduleRegistry.registerModule(contactsModule);
moduleRegistry.registerModule(groupsModule);

// Simple mock NLU agent for development/testing
const nluAgent = {
  processQuery: async ({ query, sessionId, userId }) => {
    const startTime = Date.now();
    
    try {
      // Pattern 1: Development Debug Logs
      if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Processing NLU query', {
          query: query?.substring(0, 100), // Truncate for privacy
          sessionId,
          timestamp: new Date().toISOString(),
          userId
        }, 'api');
      }
      
      // Simple intent detection based on keywords
      let intent = 'unknown';
      let entities = {};
      
      if (query.toLowerCase().includes('email') || 
          query.toLowerCase().includes('mail')) {
        intent = 'readMail';
        entities = { count: 5 };
      } else if (query.toLowerCase().includes('calendar') || 
               query.toLowerCase().includes('event') || 
               query.toLowerCase().includes('meeting')) {
        intent = 'readCalendar';
        entities = { count: 3 };
      } else if (query.toLowerCase().includes('file') || 
               query.toLowerCase().includes('document')) {
        intent = 'listFiles';
        entities = {};
      } else if (query.toLowerCase().includes('people') || 
               query.toLowerCase().includes('person') || 
               query.toLowerCase().includes('contact') || 
               query.toLowerCase().includes('find') || 
               query.toLowerCase().includes('who')) {
        // Extract name if present (simple extraction)
        const nameMatch = query.match(/\b(find|who is|about|contact)\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?)\b/i);
        const name = nameMatch ? nameMatch[2] : '';
        
        intent = 'findPeople';
        entities = { 
          criteria: { 
            name: name,
            query: name || query,
            limit: 5
          } 
        };
      }
      
      const result = {
        intent,
        entities,
        confidence: 0.85,
        query
      };
      
      // Pattern 2: User Activity Logs
      if (userId) {
        MonitoringService.info('NLU query processed successfully', {
          intent,
          confidence: 0.85,
          duration: Date.now() - startTime,
          timestamp: new Date().toISOString()
        }, 'api', null, userId);
      } else if (sessionId) {
        MonitoringService.info('NLU query processed with session', {
          sessionId,
          intent,
          confidence: 0.85,
          duration: Date.now() - startTime,
          timestamp: new Date().toISOString()
        }, 'api');
      }
      
      return result;
      
    } catch (error) {
      // Pattern 3: Infrastructure Error Logging
      const mcpError = ErrorService.createError(
        'api',
        'Failed to process NLU query',
        'error',
        {
          operation: 'processQuery',
          error: error.message,
          stack: error.stack,
          query: query?.substring(0, 50), // Truncated for privacy
          userId,
          sessionId,
          timestamp: new Date().toISOString()
        }
      );
      MonitoringService.logError(mcpError);
      
      // Pattern 4: User Error Tracking
      if (userId) {
        MonitoringService.error('NLU query processing failed', {
          error: error.message,
          operation: 'processQuery',
          timestamp: new Date().toISOString()
        }, 'api', null, userId);
      } else if (sessionId) {
        MonitoringService.error('NLU query processing failed', {
          sessionId,
          error: error.message,
          operation: 'processQuery',
          timestamp: new Date().toISOString()
        }, 'api');
      }
      
      // Return fallback result
      return {
        intent: 'unknown',
        entities: {},
        confidence: 0.0,
        query,
        error: true
      };
    }
  }
};

// Simple mock context service for development/testing
const contextService = {
  context: {
    sessionId: 'mock-session',
    timestamp: new Date().toISOString()
  },
  updateContext: async (updates, sessionId, userId) => {
    const startTime = Date.now();
    
    try {
      // Pattern 1: Development Debug Logs
      if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Updating context', {
          updateKeys: Object.keys(updates || {}),
          sessionId,
          timestamp: new Date().toISOString(),
          userId
        }, 'api');
      }
      
      // Validate updates parameter
      if (!updates || typeof updates !== 'object') {
        throw new Error('Invalid updates parameter: must be an object');
      }
      
      // Update context
      Object.assign(contextService.context, updates);
      contextService.context.timestamp = new Date().toISOString();
      
      // Pattern 2: User Activity Logs
      if (userId) {
        MonitoringService.info('Context updated successfully', {
          updateKeys: Object.keys(updates),
          duration: Date.now() - startTime,
          timestamp: new Date().toISOString()
        }, 'api', null, userId);
      } else if (sessionId) {
        MonitoringService.info('Context updated with session', {
          sessionId,
          updateKeys: Object.keys(updates),
          duration: Date.now() - startTime,
          timestamp: new Date().toISOString()
        }, 'api');
      }
      
      return contextService.context;
      
    } catch (error) {
      // Pattern 3: Infrastructure Error Logging
      const mcpError = ErrorService.createError(
        'api',
        'Failed to update context',
        'error',
        {
          operation: 'updateContext',
          error: error.message,
          stack: error.stack,
          updateKeys: updates ? Object.keys(updates) : [],
          userId,
          sessionId,
          timestamp: new Date().toISOString()
        }
      );
      MonitoringService.logError(mcpError);
      
      // Pattern 4: User Error Tracking
      if (userId) {
        MonitoringService.error('Context update failed', {
          error: error.message,
          operation: 'updateContext',
          timestamp: new Date().toISOString()
        }, 'api', null, userId);
      } else if (sessionId) {
        MonitoringService.error('Context update failed', {
          sessionId,
          error: error.message,
          operation: 'updateContext',
          timestamp: new Date().toISOString()
        }, 'api');
      }
      
      throw error;
    }
  },
  getCurrentContext: async (sessionId, userId) => {
    const startTime = Date.now();
    
    try {
      // Pattern 1: Development Debug Logs
      if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Retrieving current context', {
          sessionId,
          timestamp: new Date().toISOString(),
          userId
        }, 'api');
      }
      
      const context = { ...contextService.context };
      
      // Pattern 2: User Activity Logs
      if (userId) {
        MonitoringService.info('Context retrieved successfully', {
          contextKeys: Object.keys(context),
          duration: Date.now() - startTime,
          timestamp: new Date().toISOString()
        }, 'api', null, userId);
      } else if (sessionId) {
        MonitoringService.info('Context retrieved with session', {
          sessionId,
          contextKeys: Object.keys(context),
          duration: Date.now() - startTime,
          timestamp: new Date().toISOString()
        }, 'api');
      }
      
      return context;
      
    } catch (error) {
      // Pattern 3: Infrastructure Error Logging
      const mcpError = ErrorService.createError(
        'api',
        'Failed to retrieve context',
        'error',
        {
          operation: 'getCurrentContext',
          error: error.message,
          stack: error.stack,
          userId,
          sessionId,
          timestamp: new Date().toISOString()
        }
      );
      MonitoringService.logError(mcpError);
      
      // Pattern 4: User Error Tracking
      if (userId) {
        MonitoringService.error('Context retrieval failed', {
          error: error.message,
          operation: 'getCurrentContext',
          timestamp: new Date().toISOString()
        }, 'api', null, userId);
      } else if (sessionId) {
        MonitoringService.error('Context retrieval failed', {
          sessionId,
          error: error.message,
          operation: 'getCurrentContext',
          timestamp: new Date().toISOString()
        }, 'api');
      }
      
      throw error;
    }
  }
};

// Tools service initialization
const createToolsService = require('../core/tools-service.cjs');

// Initialize tools service with module registry
const toolsService = createToolsService({ moduleRegistry });

module.exports = {
  mailModule,
  calendarModule,
  filesModule,
  peopleModule,
  searchModule,
  teamsModule,
  todoModule,
  contactsModule,
  groupsModule,
  moduleRegistry,
  cacheService,
  eventService,
  graphClientFactory,
  calendarService,
  mailService,
  filesService,
  peopleService,
  searchService,
  teamsService,
  todoService,
  contactsService,
  groupsService,
  toolsService,
  nluAgent,
  contextService,
  errorService: ErrorService,
  monitoringService: MonitoringService
};
