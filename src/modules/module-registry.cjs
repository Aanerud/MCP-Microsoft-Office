/**
 * @fileoverview ModuleRegistry - Manages discovery and lifecycle of MCP modules.
 * Follows MCP modularity, explicit interface, and dependency injection rules.
 * Implements standardized logging and error handling via MonitoringService and ErrorService.
 */

const MonitoringService = require('../core/monitoring-service.cjs');
const ErrorService = require('../core/error-service.cjs');

class ModuleRegistry {
    // TODO: [constructor] Accept logger via DI; track registration timestamps. (LOW) - Implemented
    constructor(options = {}, userId, sessionId) {
        const startTime = Date.now();
        
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('ModuleRegistry constructor called', {
                sessionId: sessionId ? sessionId.substring(0, 8) + '...' : 'none',
                userId: userId ? userId.substring(0, 20) + '...' : 'none',
                timestamp: new Date().toISOString()
            }, 'module');
        }
        
        /** @type {Map<string, object>} */
        this.modules = new Map();
        /** @type {Map<string, Set<string>>} */
        this.capabilityMap = new Map(); // capability -> Set of module ids
        /** @type {Map<string, Date>} */
        this.registrationTimestamps = new Map(); // module.id -> registration time
        
        // Pattern 2: User Activity Logs
        if (userId) {
            MonitoringService.info('ModuleRegistry initialized successfully', {
                executionTime: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'module', null, userId);
        } else if (sessionId) {
            MonitoringService.info('ModuleRegistry initialized with session', {
                sessionId: sessionId.substring(0, 8) + '...',
                executionTime: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'module');
        }
    }

    /**
     * Registers a module with the registry and its capabilities.
     * @param {object} module - Must have id, name, capabilities, init, handleIntent
     * @param {string} userId - User ID for logging context
     * @param {string} sessionId - Session ID for logging context
     */
    // TODO: [registerModule] Validate module interface contract (MEDIUM) - Implemented (basic validation)
    registerModule(module, userId, sessionId) {
        const startTime = Date.now();
        
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Module registration started', {
                moduleId: module?.id || 'unknown',
                moduleName: module?.name || 'unknown',
                sessionId: sessionId ? sessionId.substring(0, 8) + '...' : 'none',
                userId: userId ? userId.substring(0, 20) + '...' : 'none',
                timestamp: new Date().toISOString()
            }, 'module');
        }
        
        // Validate module object
        if (!module || typeof module !== 'object' || !module.id) {
            // Pattern 3: Infrastructure Error Logging
            const error = ErrorService.createError(
                'module',
                'Invalid module object or missing ID',
                'error',
                { 
                    moduleId: module?.id || 'unknown',
                    moduleType: typeof module,
                    timestamp: new Date().toISOString()
                }
            );
            
            MonitoringService.logError(error);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                MonitoringService.error('Module registration failed - invalid module', {
                    moduleId: module?.id || 'unknown',
                    error: 'Invalid module object or missing ID',
                    timestamp: new Date().toISOString()
                }, 'module', null, userId);
            } else if (sessionId) {
                MonitoringService.error('Module registration failed - invalid module', {
                    sessionId: sessionId.substring(0, 8) + '...',
                    moduleId: module?.id || 'unknown',
                    error: 'Invalid module object or missing ID',
                    timestamp: new Date().toISOString()
                }, 'module');
            }
            
            throw error;
        }
        
        // Check for duplicate module ID
        if (this.modules.has(module.id)) {
            // Pattern 3: Infrastructure Error Logging
            const error = ErrorService.createError(
                'module',
                `Module with ID '${module.id}' already registered`,
                'error',
                { 
                    moduleId: module.id,
                    moduleName: module.name,
                    timestamp: new Date().toISOString()
                }
            );
            
            MonitoringService.logError(error);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                MonitoringService.error('Module registration failed - duplicate ID', {
                    moduleId: module.id,
                    error: 'Module already registered',
                    timestamp: new Date().toISOString()
                }, 'module', null, userId);
            } else if (sessionId) {
                MonitoringService.error('Module registration failed - duplicate ID', {
                    sessionId: sessionId.substring(0, 8) + '...',
                    moduleId: module.id,
                    error: 'Module already registered',
                    timestamp: new Date().toISOString()
                }, 'module');
            }
            
            throw error;
        }

        // Validate interface contract
        const requiredProps = {
            id: 'string',
            name: 'string',
            capabilities: 'array',
            init: 'function',
            handleIntent: 'function'
        };
        const optionalProps = {
            priority: 'number' // Added for future sorting
        };

        // Check required properties
        for (const prop in requiredProps) {
            if (typeof module[prop] !== requiredProps[prop] && !(prop === 'capabilities' && Array.isArray(module[prop]))) {
                const error = ErrorService?.createError(
                    'module',
                    `Module '${module.id}' missing or invalid required property '${prop}'`,
                    'error',
                    { 
                        moduleId: module.id,
                        moduleName: module.name,
                        property: prop,
                        expectedType: requiredProps[prop],
                        actualType: Array.isArray(module[prop]) ? 'array' : typeof module[prop],
                        timestamp: new Date().toISOString()
                    }
                ) || {
                    category: 'module',
                    message: `Module '${module.id}' missing or invalid required property '${prop}'`,
                    severity: 'error',
                    context: { 
                        moduleId: module.id, 
                        property: prop,
                        expectedType: requiredProps[prop]
                    }
                };
                
                MonitoringService.logError(error);
                
                throw error;
            }
        }
        
        // Check optional properties
        for (const prop in optionalProps) {
            if (module.hasOwnProperty(prop) && typeof module[prop] !== optionalProps[prop]) {
                MonitoringService.warn(`Module has optional property with incorrect type`, { 
                    moduleId: module.id,
                    moduleName: module.name,
                    property: prop,
                    expectedType: optionalProps[prop],
                    actualType: typeof module[prop],
                    timestamp: new Date().toISOString()
                }, 'module');
            }
        }

        // Store a normalized module object
        const moduleToStore = { ...module, priority: module.priority || 0 }; // Default priority 0

        // Register the module
        this.modules.set(module.id, moduleToStore);
        this.registrationTimestamps.set(module.id, new Date());
        
        // Register capabilities
        if (Array.isArray(module.capabilities)) {
            for (const cap of module.capabilities) {
                if (!this.capabilityMap.has(cap)) this.capabilityMap.set(cap, new Set());
                this.capabilityMap.get(cap).add(module.id);
            }
        }
        
        // Pattern 2: User Activity Logs
        if (userId) {
            MonitoringService.info('Module registered successfully', {
                moduleId: module.id,
                moduleName: module.name,
                capabilities: Array.isArray(module.capabilities) ? module.capabilities.length : 0,
                executionTime: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'module', null, userId);
        } else if (sessionId) {
            MonitoringService.info('Module registered with session', {
                sessionId: sessionId.substring(0, 8) + '...',
                moduleId: module.id,
                moduleName: module.name,
                capabilities: Array.isArray(module.capabilities) ? module.capabilities.length : 0,
                executionTime: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'module');
        }
        
        // Track performance metric
        const elapsedTime = Date.now() - startTime;
        MonitoringService.trackMetric('module_registration_time', elapsedTime, {
            moduleId: module.id,
            timestamp: new Date().toISOString()
        });
        
        // Emit event for real-time UI updates if logEmitter is available
        try {
            if (MonitoringService.logEmitter && typeof MonitoringService.logEmitter.emit === 'function') {
                MonitoringService.logEmitter.emit('moduleRegistered', {
                    event: 'moduleRegistered',
                    data: {
                        moduleId: module.id,
                        moduleName: module.name,
                        capabilities: Array.isArray(module.capabilities) ? module.capabilities : []
                    },
                    timestamp: new Date().toISOString()
                });
            }
        } catch (emitError) {
            // Silently fail if emitter is not available or fails
            // This is non-critical functionality
        }
    }

    /**
     * Updates an existing module in the registry (replaces with new version).
     * If the module doesn't exist, registers it as new.
     * @param {object} module - Must have id, name, capabilities, init, handleIntent
     * @param {string} userId - User ID for logging context
     * @param {string} sessionId - Session ID for logging context
     */
    updateModule(module, userId, sessionId) {
        if (!module || !module.id) {
            return;
        }

        // If module doesn't exist, just register it
        if (!this.modules.has(module.id)) {
            this.registerModule(module, userId, sessionId);
            return;
        }

        // Update the module in place
        const moduleToStore = { ...module, priority: module.priority || 0 };
        this.modules.set(module.id, moduleToStore);
        this.registrationTimestamps.set(module.id, new Date());

        // Update capabilities
        if (Array.isArray(module.capabilities)) {
            for (const cap of module.capabilities) {
                if (!this.capabilityMap.has(cap)) this.capabilityMap.set(cap, new Set());
                this.capabilityMap.get(cap).add(module.id);
            }
        }

        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Module updated in registry', {
                moduleId: module.id,
                moduleName: module.name,
                timestamp: new Date().toISOString()
            }, 'module');
        }
    }

    /**
     * Returns a registered module by id.
     * @param {string} id
     * @param {object} [options] - Optional settings.
     * @param {boolean} [options.strict=false] - Throw an error if the module is not found.
     * @param {string} userId - User ID for logging context
     * @param {string} sessionId - Session ID for logging context
     * @returns {object|undefined}
     */
    // TODO: [getModule] Throw if not found when strict flag true (LOW) - Implemented
    getModule(id, options = {}, userId, sessionId) {
        const startTime = Date.now();
        const { strict = false } = options;
        
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Getting module by ID', {
                moduleId: id ? id.substring(0, 20) + '...' : 'none',
                strict,
                sessionId: sessionId ? sessionId.substring(0, 8) + '...' : 'none',
                userId: userId ? userId.substring(0, 20) + '...' : 'none',
                timestamp: new Date().toISOString()
            }, 'module');
        }
        
        const module = this.modules.get(id);

        if (!module && strict) {
            // Pattern 3: Infrastructure Error Logging
            const error = ErrorService.createError(
                'module',
                `Module with ID '${id}' not found (strict mode enabled)`,
                'error',
                { 
                    moduleId: id,
                    strict,
                    timestamp: new Date().toISOString()
                }
            );
            
            MonitoringService.logError(error);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                MonitoringService.error('Module lookup failed - not found', {
                    moduleId: id,
                    error: 'Module not found in strict mode',
                    timestamp: new Date().toISOString()
                }, 'module', null, userId);
            } else if (sessionId) {
                MonitoringService.error('Module lookup failed - not found', {
                    sessionId: sessionId.substring(0, 8) + '...',
                    moduleId: id,
                    error: 'Module not found in strict mode',
                    timestamp: new Date().toISOString()
                }, 'module');
            }
            
            throw error;
        }
        
        // Pattern 2: User Activity Logs for successful operations
        if (module) {
            if (userId) {
                MonitoringService.info('Module retrieved successfully', {
                    moduleId: id,
                    moduleName: module.name,
                    executionTime: Date.now() - startTime,
                    timestamp: new Date().toISOString()
                }, 'module', null, userId);
            } else if (sessionId) {
                MonitoringService.info('Module retrieved with session', {
                    sessionId: sessionId.substring(0, 8) + '...',
                    moduleId: id,
                    moduleName: module.name,
                    executionTime: Date.now() - startTime,
                    timestamp: new Date().toISOString()
                }, 'module');
            }
        }
        
        // Track performance metric
        const elapsedTime = Date.now() - startTime;
        MonitoringService.trackMetric('module_lookup_time', elapsedTime, {
            moduleId: id,
            found: !!module,
            timestamp: new Date().toISOString()
        });
        
        return module; // Returns the module or undefined if not found and strict is false
    }

    /**
     * Returns all registered modules.
     * @param {string} userId - User ID for logging context
     * @param {string} sessionId - Session ID for logging context
     * @returns {Array<object>}
     */
    getAllModules(userId, sessionId) {
        const startTime = Date.now();
        
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Getting all modules', {
                sessionId: sessionId ? sessionId.substring(0, 8) + '...' : 'none',
                userId: userId ? userId.substring(0, 20) + '...' : 'none',
                timestamp: new Date().toISOString()
            }, 'module');
        }
        
        const modules = Array.from(this.modules.values());
        
        // Pattern 2: User Activity Logs
        if (userId) {
            MonitoringService.info('All modules retrieved successfully', {
                count: modules.length,
                moduleIds: modules.map(m => m.id),
                executionTime: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'module', null, userId);
        } else if (sessionId) {
            MonitoringService.info('All modules retrieved with session', {
                sessionId: sessionId.substring(0, 8) + '...',
                count: modules.length,
                moduleIds: modules.map(m => m.id),
                executionTime: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'module');
        }
        
        return modules;
    }

    /**
     * Finds all modules that support a given capability/intent.
     * @param {string} capability
     * @param {string} userId - User ID for logging context
     * @param {string} sessionId - Session ID for logging context
     * @returns {Array<object>} Modules supporting the capability
     */
    // TODO: [findModulesForIntent] Return by priority when multiple modules share capability (LOW) - Implemented
    findModulesForIntent(capability, userId, sessionId) {
        const startTime = Date.now();
        
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Finding modules for intent', {
                intent: capability ? capability.substring(0, 50) + '...' : 'none',
                sessionId: sessionId ? sessionId.substring(0, 8) + '...' : 'none',
                userId: userId ? userId.substring(0, 20) + '...' : 'none',
                timestamp: new Date().toISOString()
            }, 'module');
        }
        
        const ids = this.capabilityMap.get(capability);
        if (!ids || ids.size === 0) {
            // Pattern 2: User Activity Logs for no results
            if (userId) {
                MonitoringService.info('No modules found for intent', {
                    intent: capability,
                    executionTime: Date.now() - startTime,
                    timestamp: new Date().toISOString()
                }, 'module', null, userId);
            } else if (sessionId) {
                MonitoringService.info('No modules found for intent with session', {
                    sessionId: sessionId.substring(0, 8) + '...',
                    intent: capability,
                    executionTime: Date.now() - startTime,
                    timestamp: new Date().toISOString()
                }, 'module');
            }
            
            return [];
        }

        // Retrieve modules and filter out any potential nulls if an ID somehow exists without a module
        const modules = Array.from(ids)
            .map(id => this.modules.get(id))
            .filter(Boolean);

        // Sort by priority (descending). Modules without priority have default 0.
        modules.sort((a, b) => (b.priority || 0) - (a.priority || 0));

        // Pattern 2: User Activity Logs for successful results
        if (userId) {
            MonitoringService.info('Found modules for intent', {
                intent: capability,
                count: modules.length,
                moduleIds: modules.map(m => m.id),
                executionTime: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'module', null, userId);
        } else if (sessionId) {
            MonitoringService.info('Found modules for intent with session', {
                sessionId: sessionId.substring(0, 8) + '...',
                intent: capability,
                count: modules.length,
                moduleIds: modules.map(m => m.id),
                executionTime: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'module');
        }
        
        // Track performance metric
        const elapsedTime = Date.now() - startTime;
        MonitoringService.trackMetric('intent_lookup_time', elapsedTime, {
            intent: capability,
            found: true,
            count: modules.length,
            timestamp: new Date().toISOString()
        });
        
        // Emit event for real-time UI updates if logEmitter is available
        try {
            if (MonitoringService.logEmitter && typeof MonitoringService.logEmitter.emit === 'function') {
                MonitoringService.logEmitter.emit('modulesFoundForIntent', {
                    event: 'modulesFoundForIntent',
                    data: {
                        intent: capability,
                        count: modules.length,
                        moduleIds: modules.map(m => m.id)
                    },
                    timestamp: new Date().toISOString()
                });
            }
        } catch (emitError) {
            // Silently fail if emitter is not available or fails
            // This is non-critical functionality
        }
        
        return modules;
    }

    /**
     * Returns all capabilities registered in the system.
     * @param {string} userId - User ID for logging context
     * @param {string} sessionId - Session ID for logging context
     * @returns {Array<string>}
     */
    // TODO: [listCapabilities] Alphabetically sort output (LOW) - Implemented
    listCapabilities(userId, sessionId) {
        const startTime = Date.now();
        
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Listing all capabilities', {
                sessionId: sessionId ? sessionId.substring(0, 8) + '...' : 'none',
                userId: userId ? userId.substring(0, 20) + '...' : 'none',
                timestamp: new Date().toISOString()
            }, 'module');
        }
        
        // Get keys and sort them alphabetically
        const capabilities = Array.from(this.capabilityMap.keys()).sort();
        
        // Pattern 2: User Activity Logs
        if (userId) {
            MonitoringService.info('Capabilities listed successfully', {
                count: capabilities.length,
                capabilities,
                executionTime: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'module', null, userId);
        } else if (sessionId) {
            MonitoringService.info('Capabilities listed with session', {
                sessionId: sessionId.substring(0, 8) + '...',
                count: capabilities.length,
                capabilities,
                executionTime: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'module');
        }
        
        return capabilities;
    }

    // TODO: [getRegistrationInfo] New method to return module registration timestamp. (LOW) - Implemented
    /**
     * Gets registration information for a specific module.
     * @param {string} moduleId - The ID of the module.
     * @param {string} userId - User ID for logging context
     * @param {string} sessionId - Session ID for logging context
     * @returns {object|null} An object containing registration info (like timestamp) or null if not found.
     */
    getRegistrationInfo(moduleId, userId, sessionId) {
        const startTime = Date.now();
        
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Getting module registration info', {
                moduleId: moduleId ? moduleId.substring(0, 20) + '...' : 'none',
                sessionId: sessionId ? sessionId.substring(0, 8) + '...' : 'none',
                userId: userId ? userId.substring(0, 20) + '...' : 'none',
                timestamp: new Date().toISOString()
            }, 'module');
        }
        
        const registrationTime = this.registrationTimestamps.get(moduleId);
        if (!registrationTime) {
            // Pattern 4: User Error Tracking for not found
            if (userId) {
                MonitoringService.error('Module registration info not found', {
                    moduleId,
                    error: 'No registration timestamp found',
                    timestamp: new Date().toISOString()
                }, 'module', null, userId);
            } else if (sessionId) {
                MonitoringService.error('Module registration info not found', {
                    sessionId: sessionId.substring(0, 8) + '...',
                    moduleId,
                    error: 'No registration timestamp found',
                    timestamp: new Date().toISOString()
                }, 'module');
            }
            
            return null;
        }
        
        const registrationInfo = {
            registrationTime: registrationTime
            // Add other info like state here in the future
        };
        
        // Pattern 2: User Activity Logs for successful retrieval
        if (userId) {
            MonitoringService.info('Module registration info retrieved successfully', {
                moduleId,
                registeredAt: registrationTime,
                executionTime: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'module', null, userId);
        } else if (sessionId) {
            MonitoringService.info('Module registration info retrieved with session', {
                sessionId: sessionId.substring(0, 8) + '...',
                moduleId,
                registeredAt: registrationTime,
                executionTime: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'module');
        }
        
        return registrationInfo;
    }
}

module.exports = new ModuleRegistry();
