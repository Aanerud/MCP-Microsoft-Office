/**
 * @fileoverview init-modules - Initializes all discovered MCP modules with dependency injection.
 * Follows async, modular, and testable design. Aligned with phase1_architecture.md and MCP rules.
 * Implements proper dependency injection to avoid circular dependencies.
 */

const moduleRegistry = require('./module-registry.cjs');
const MonitoringService = require('../core/monitoring-service.cjs');
const ErrorService = require('../core/error-service.cjs');
const StorageService = require('../core/storage-service.cjs');
const { databaseFactory } = require('../core/database-factory.cjs');
const graphService = require('../graph/graph-service.cjs');
const todoService = require('../graph/todo-service.cjs');
const contactsService = require('../graph/contacts-service.cjs');
const groupsService = require('../graph/groups-service.cjs');

/**
 * Initializes all registered modules with provided dependencies/services.
 * Calls each module's init(services) and replaces the module in the registry with the initialized instance.
 * Handles errors during initialization and logs them appropriately.
 * Tracks performance metrics for the initialization process.
 * @param {object} services - Dependency/service registry to inject
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<Array<object>>} Array of successfully initialized modules
 * @throws {Object} Will not throw errors from individual module initialization failures
 */
async function initializeModules(services = {}, userId, sessionId) {
    const startTime = Date.now();
    
    // Pattern 1: Development Debug Logs
    if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Starting module initialization process', {
            sessionId: sessionId || 'unknown',
            userId: userId ? userId.substring(0, 20) + '...' : 'unknown',
            serviceCount: Object.keys(services).length,
            timestamp: new Date().toISOString()
        }, 'modules');
    }
    
    // Initialize database factory first
    try {
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Initializing database factory', {
                sessionId: sessionId || 'unknown',
                userId: userId ? userId.substring(0, 20) + '...' : 'unknown',
                timestamp: new Date().toISOString()
            }, 'modules');
        }
        
        const config = {
            DB_TYPE: 'sqlite',
            DB_PATH: './data/mcp.sqlite'
        };
        await databaseFactory.init(config);
        
        // Pattern 2: User Activity Logs
        if (userId) {
            MonitoringService.info('Database factory initialized successfully', {
                databaseType: config.DB_TYPE,
                timestamp: new Date().toISOString()
            }, 'modules', null, userId);
        } else if (sessionId) {
            MonitoringService.info('Database factory initialized with session', {
                sessionId: sessionId,
                databaseType: config.DB_TYPE,
                timestamp: new Date().toISOString()
            }, 'modules');
        }
    } catch (error) {
        // Pattern 3: Infrastructure Error Logging
        const mcpError = ErrorService.createError(
            'modules',
            `Failed to initialize database factory: ${error.message}`,
            'critical',
            {
                error: error.message,
                stack: error.stack,
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(mcpError);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('Database factory initialization failed', {
                error: error.message,
                timestamp: new Date().toISOString()
            }, 'modules', null, userId);
        } else if (sessionId) {
            MonitoringService.error('Database factory initialization failed', {
                sessionId: sessionId,
                error: error.message,
                timestamp: new Date().toISOString()
            }, 'modules');
        }
        
        throw mcpError;
    }
    
    // Initialize storage service after database factory
    try {
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Initializing storage service', {
                sessionId: sessionId || 'unknown',
                userId: userId ? userId.substring(0, 20) + '...' : 'unknown',
                timestamp: new Date().toISOString()
            }, 'modules');
        }
        
        await StorageService.init();
        
        // Pattern 2: User Activity Logs
        if (userId) {
            MonitoringService.info('Storage service initialized successfully', {
                timestamp: new Date().toISOString()
            }, 'modules', null, userId);
        } else if (sessionId) {
            MonitoringService.info('Storage service initialized with session', {
                sessionId: sessionId,
                timestamp: new Date().toISOString()
            }, 'modules');
        }
    } catch (error) {
        // Pattern 3: Infrastructure Error Logging
        const mcpError = ErrorService.createError(
            'modules',
            `Failed to initialize storage service: ${error.message}`,
            'critical',
            {
                error: error.message,
                stack: error.stack,
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(mcpError);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('Storage service initialization failed', {
                error: error.message,
                timestamp: new Date().toISOString()
            }, 'modules', null, userId);
        } else if (sessionId) {
            MonitoringService.error('Storage service initialization failed', {
                sessionId: sessionId,
                error: error.message,
                timestamp: new Date().toISOString()
            }, 'modules');
        }
        
        throw mcpError;
    }
    
    // Set up dependency injection between core services to avoid circular references
    // This is critical to prevent infinite error loops
    if (ErrorService && MonitoringService) {
        // Set the logging service in the error service
        ErrorService.setLoggingService(MonitoringService);
    }
    
    // Ensure all core services are available in the services object
    const enrichedServices = {
        ...services,
        errorService: ErrorService,
        monitoringService: MonitoringService,
        storageService: StorageService,
        databaseFactory: databaseFactory,
        graphService: graphService,
        todoService: todoService,
        contactsService: contactsService,
        groupsService: groupsService
    };
    
    // Log initialization start
    if (MonitoringService) {
        // Use helper function to redact sensitive information
        const safeServicesInfo = redactSensitiveServiceInfo(enrichedServices, userId, sessionId);
        
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Starting module initialization with services', {
                sessionId: sessionId || 'unknown',
                userId: userId ? userId.substring(0, 20) + '...' : 'unknown',
                moduleCount: moduleRegistry.getAllModules().length,
                serviceKeys: safeServicesInfo,
                timestamp: new Date().toISOString()
            }, 'modules');
        }
        
        // Pattern 2: User Activity Logs
        if (userId) {
            MonitoringService.info('Starting module initialization process', {
                moduleCount: moduleRegistry.getAllModules().length,
                servicesProvided: safeServicesInfo
            }, 'module', null, userId);
        } else if (sessionId) {
            MonitoringService.info('Starting module initialization with session', {
                sessionId: sessionId,
                moduleCount: moduleRegistry.getAllModules().length,
                servicesProvided: safeServicesInfo
            }, 'modules');
        }
    } else {
        // Use helper function to redact sensitive information for fallback logging
        const safeServicesInfo = redactSensitiveServiceInfo(enrichedServices, userId, sessionId);
        
        console.info('[MCP MODULE] Starting module initialization process', JSON.stringify({ 
            moduleCount: moduleRegistry.getAllModules().length,
            servicesProvided: safeServicesInfo
        }));
    }
    
    const modules = moduleRegistry.getAllModules();
    const initialized = [];
    for (const mod of modules) {
        if (mod && typeof mod.init === 'function') {
            try {
                // Pattern 1: Development Debug Logs
                if (process.env.NODE_ENV === 'development') {
                    MonitoringService.debug('Initializing individual module', {
                        sessionId: sessionId || 'unknown',
                        userId: userId ? userId.substring(0, 20) + '...' : 'unknown',
                        moduleId: mod.id || 'unknown',
                        moduleName: mod.name || 'unknown',
                        timestamp: new Date().toISOString()
                    }, 'modules');
                }
                
                // Call the module's init function with the services
                const initializedModule = await mod.init(enrichedServices);

                // Replace the module in the registry with the initialized instance
                // Use updateModule to avoid "already registered" errors
                moduleRegistry.updateModule(initializedModule);
                
                // Add to the list of successfully initialized modules
                initialized.push(initializedModule);
                
                // Pattern 2: User Activity Logs
                if (userId) {
                    MonitoringService.info('Module initialized successfully', {
                        moduleId: mod.id,
                        moduleName: mod.name,
                        timestamp: new Date().toISOString()
                    }, 'modules', null, userId);
                } else if (sessionId) {
                    MonitoringService.info('Module initialized with session', {
                        sessionId: sessionId,
                        moduleId: mod.id,
                        moduleName: mod.name,
                        timestamp: new Date().toISOString()
                    }, 'modules');
                }
            } catch (error) {
                // Pattern 3: Infrastructure Error Logging
                const errorTraceId = `module-init-error-${mod.id}-${Date.now()}`;
                const mcpError = ErrorService.createError(
                    'modules',
                    `Failed to initialize module: ${mod.id}`,
                    'error',
                    { 
                        moduleId: mod.id, 
                        moduleName: mod.name,
                        error: error.message,
                        stack: error.stack,
                        timestamp: new Date().toISOString()
                    },
                    errorTraceId
                );
                MonitoringService.logError(mcpError);
                
                // Pattern 4: User Error Tracking
                if (userId) {
                    MonitoringService.error('Module initialization failed', {
                        moduleId: mod.id,
                        moduleName: mod.name,
                        error: error.message,
                        timestamp: new Date().toISOString()
                    }, 'modules', null, userId);
                } else if (sessionId) {
                    MonitoringService.error('Module initialization failed', {
                        sessionId: sessionId,
                        moduleId: mod.id,
                        moduleName: mod.name,
                        error: error.message,
                        timestamp: new Date().toISOString()
                    }, 'modules');
                }
            }
        } else {
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                MonitoringService.debug('Module missing init function', {
                    sessionId: sessionId || 'unknown',
                    userId: userId ? userId.substring(0, 20) + '...' : 'unknown',
                    moduleId: mod.id || 'unknown',
                    moduleName: mod.name || 'unknown',
                    timestamp: new Date().toISOString()
                }, 'modules');
            }
            
            // Pattern 4: User Error Tracking (missing init function is a user-visible issue)
            if (userId) {
                MonitoringService.error('Module missing init function', {
                    moduleId: mod.id || 'unknown',
                    moduleName: mod.name || 'unknown',
                    timestamp: new Date().toISOString()
                }, 'modules', null, userId);
            } else if (sessionId) {
                MonitoringService.error('Module missing init function', {
                    sessionId: sessionId,
                    moduleId: mod.id || 'unknown',
                    moduleName: mod.name || 'unknown',
                    timestamp: new Date().toISOString()
                }, 'modules');
            }
        }
    }
    
    const elapsedTime = Date.now() - startTime;
    
    // Log completion information
    if (MonitoringService) {
        const completionData = { 
            totalModules: modules.length,
            initializedCount: initialized.length,
            moduleIds: initialized.map(m => m.id),
            elapsedTimeMs: elapsedTime,
            timestamp: new Date().toISOString()
        };
        
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Module initialization completed with details', {
                sessionId: sessionId || 'unknown',
                userId: userId ? userId.substring(0, 20) + '...' : 'unknown',
                ...completionData
            }, 'modules');
        }
        
        // Pattern 2: User Activity Logs
        if (userId) {
            MonitoringService.info('Module initialization completed', completionData, 'modules', null, userId);
        } else if (sessionId) {
            MonitoringService.info('Module initialization completed with session', {
                sessionId: sessionId,
                ...completionData
            }, 'modules');
        }
        
        // Emit event for real-time UI updates if logEmitter is available
        try {
            if (MonitoringService.logEmitter && typeof MonitoringService.logEmitter.emit === 'function') {
                MonitoringService.logEmitter.emit('moduleInitialized', {
                    event: 'moduleInitialized',
                    data: completionData,
                    timestamp: new Date().toISOString()
                });
            }
        } catch (error) {
            // Silently fail if emitter is not available or fails
            // This is non-critical functionality
        }
    } else {
        console.info('[MCP MODULE] Module initialization completed', JSON.stringify({
            totalModules: modules.length,
            initializedCount: initialized.length,
            moduleIds: initialized.map(m => m.id),
            elapsedTimeMs: elapsedTime
        }));
    }
    
    return initialized;
}

/**
 * Helper function to redact sensitive information from service keys
 * @param {Object} services - The services object to redact
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Array<string>} Array of service keys with sensitive ones redacted
 */
function redactSensitiveServiceInfo(services, userId, sessionId) {
    const startTime = Date.now();
    
    // Pattern 1: Development Debug Logs
    if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Redacting sensitive service information', {
            sessionId: sessionId || 'unknown',
            userId: userId ? userId.substring(0, 20) + '...' : 'unknown',
            serviceCount: Object.keys(services).length,
            timestamp: new Date().toISOString()
        }, 'modules');
    }
    
    try {
        const redactedKeys = Object.keys(services).map(key => {
            // Redact any keys that might contain sensitive information
            if (key.toLowerCase().includes('token') || 
                key.toLowerCase().includes('secret') || 
                key.toLowerCase().includes('password') || 
                key.toLowerCase().includes('auth')) {
                return `${key}: [REDACTED]`;
            }
            return key;
        });
        
        const elapsedTime = Date.now() - startTime;
        
        // Pattern 2: User Activity Logs
        if (userId) {
            MonitoringService.info('Service information redacted successfully', {
                serviceCount: Object.keys(services).length,
                redactedCount: redactedKeys.filter(k => k.includes('[REDACTED]')).length,
                elapsedTimeMs: elapsedTime,
                timestamp: new Date().toISOString()
            }, 'modules', null, userId);
        } else if (sessionId) {
            MonitoringService.info('Service information redacted with session', {
                sessionId: sessionId,
                serviceCount: Object.keys(services).length,
                redactedCount: redactedKeys.filter(k => k.includes('[REDACTED]')).length,
                elapsedTimeMs: elapsedTime,
                timestamp: new Date().toISOString()
            }, 'modules');
        }
        
        return redactedKeys;
        
    } catch (error) {
        // Pattern 3: Infrastructure Error Logging
        const mcpError = ErrorService.createError(
            'modules',
            `Failed to redact sensitive service information: ${error.message}`,
            'error',
            {
                error: error.message,
                stack: error.stack,
                serviceCount: Object.keys(services).length,
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(mcpError);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('Service information redaction failed', {
                error: error.message,
                timestamp: new Date().toISOString()
            }, 'modules', null, userId);
        } else if (sessionId) {
            MonitoringService.error('Service information redaction failed', {
                sessionId: sessionId,
                error: error.message,
                timestamp: new Date().toISOString()
            }, 'modules');
        }
        
        // Return empty array as fallback
        return [];
    }
}

module.exports = { 
    initializeModules,
    // Export for testing purposes
    redactSensitiveServiceInfo
};
