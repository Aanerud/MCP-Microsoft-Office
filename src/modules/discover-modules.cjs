/**
 * @fileoverview discover-modules - Dynamically discovers and loads MCP modules from the modules directory.
 * Registers each found module with the ModuleRegistry.
 * Follows async, modular, and testable design.
 */

const fs = require('fs').promises;
const path = require('path');
const moduleRegistry = require('./module-registry.cjs');
const MonitoringService = require('../core/monitoring-service.cjs');
const ErrorService = require('../core/error-service.cjs');

/**
 * Discovers and registers all modules in the given directory.
 * Only files ending with .cjs and exporting required module interface are loaded.
 * Files must export an object with id, name, capabilities array, init function, and handleIntent function.
 * @param {string} modulesDir - Absolute path to modules directory
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<Array<object>>} Array of registered modules
 * @throws {Object} Standardized error object if directory cannot be read
 */
async function discoverModules(modulesDir, userId, sessionId) {
    const startTime = Date.now();
    
    // Pattern 1: Development Debug Logs
    if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Starting module discovery process', {
            modulesDir,
            userId: userId ? userId.substring(0, 20) + '...' : 'anonymous',
            sessionId: sessionId ? sessionId.substring(0, 8) + '...' : 'none',
            timestamp: new Date().toISOString()
        }, 'module');
    }
    
    let files;
    try {
        files = await fs.readdir(modulesDir);
        
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Successfully read modules directory', {
                modulesDir,
                fileCount: files.length,
                files: files.filter(f => f.endsWith('.cjs')),
                userId: userId ? userId.substring(0, 20) + '...' : 'anonymous',
                sessionId: sessionId ? sessionId.substring(0, 8) + '...' : 'none',
                timestamp: new Date().toISOString()
            }, 'module');
        }
        
    } catch (error) {
        // Pattern 3: Infrastructure Error Logging
        const mcpError = ErrorService.createError(
            'module',
            'Failed to read modules directory',
            'error',
            { 
                modulesDir, 
                errorCode: error.code || 'UNKNOWN',
                errorMessage: error.message,
                errorStack: error.stack,
                userId: userId ? userId.substring(0, 20) + '...' : 'anonymous',
                sessionId: sessionId ? sessionId.substring(0, 8) + '...' : 'none',
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(mcpError);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('Module discovery failed - directory read error', {
                modulesDir: modulesDir.substring(0, 50) + '...',
                error: error.message,
                timestamp: new Date().toISOString()
            }, 'module', null, userId);
        } else if (sessionId) {
            MonitoringService.error('Module discovery failed - directory read error', {
                sessionId: sessionId.substring(0, 8) + '...',
                modulesDir: modulesDir.substring(0, 50) + '...',
                error: error.message,
                timestamp: new Date().toISOString()
            }, 'module');
        }
        
        throw mcpError;
    }
    
    const registered = [];
    const skipped = [];
    const failed = [];
    
    for (const file of files) {
        if (!file.endsWith('.cjs') || file === 'module-registry.cjs' || file === 'discover-modules.cjs') {
            skipped.push(file);
            continue;
        }
        
        const modPath = path.join(modulesDir, file);
        let mod;
        
        try {
            mod = require(modPath);
            
            // Must have id, name, capabilities, init, handleIntent
            if (mod && mod.id && mod.name && Array.isArray(mod.capabilities) &&
                typeof mod.init === 'function' && typeof mod.handleIntent === 'function') {

                // Use updateModule to avoid "already registered" errors
                moduleRegistry.updateModule(mod);
                registered.push(mod);
                
                // Pattern 1: Development Debug Logs
                if (process.env.NODE_ENV === 'development') {
                    MonitoringService.debug('Module registered successfully', {
                        moduleId: mod.id,
                        moduleName: mod.name,
                        file,
                        capabilities: mod.capabilities,
                        userId: userId ? userId.substring(0, 20) + '...' : 'anonymous',
                        sessionId: sessionId ? sessionId.substring(0, 8) + '...' : 'none',
                        timestamp: new Date().toISOString()
                    }, 'module');
                }
                
            } else {
                failed.push({ file, reason: 'invalid_interface', missing: getMissingProperties(mod, userId, sessionId) });
                
                // Pattern 3: Infrastructure Error Logging
                const mcpError = ErrorService.createError(
                    'module',
                    `Invalid module interface: ${file}`,
                    'warning',
                    { 
                        file, 
                        path: modPath, 
                        missingProperties: getMissingProperties(mod, userId, sessionId),
                        userId: userId ? userId.substring(0, 20) + '...' : 'anonymous',
                        sessionId: sessionId ? sessionId.substring(0, 8) + '...' : 'none',
                        timestamp: new Date().toISOString()
                    }
                );
                MonitoringService.logError(mcpError);
                
                // Pattern 4: User Error Tracking
                if (userId) {
                    MonitoringService.error('Module validation failed - invalid interface', {
                        file,
                        missingProperties: getMissingProperties(mod, userId, sessionId),
                        timestamp: new Date().toISOString()
                    }, 'module', null, userId);
                } else if (sessionId) {
                    MonitoringService.error('Module validation failed - invalid interface', {
                        sessionId: sessionId.substring(0, 8) + '...',
                        file,
                        missingProperties: getMissingProperties(mod, userId, sessionId),
                        timestamp: new Date().toISOString()
                    }, 'module');
                }
            }
        } catch (error) {
            failed.push({ file, reason: 'load_error', error: error.message });
            
            // Pattern 3: Infrastructure Error Logging
            const mcpError = ErrorService.createError(
                'module',
                `Failed to load module: ${file}`,
                'error',
                { 
                    file, 
                    path: modPath, 
                    errorMessage: error.message,
                    errorStack: error.stack,
                    userId: userId ? userId.substring(0, 20) + '...' : 'anonymous',
                    sessionId: sessionId ? sessionId.substring(0, 8) + '...' : 'none',
                    timestamp: new Date().toISOString()
                }
            );
            MonitoringService.logError(mcpError);
            
            // Pattern 4: User Error Tracking
            if (userId) {
                MonitoringService.error('Module loading failed', {
                    file,
                    error: error.message,
                    timestamp: new Date().toISOString()
                }, 'module', null, userId);
            } else if (sessionId) {
                MonitoringService.error('Module loading failed', {
                    sessionId: sessionId.substring(0, 8) + '...',
                    file,
                    error: error.message,
                    timestamp: new Date().toISOString()
                }, 'module');
            }
        }
    }
    
    const elapsedTime = Date.now() - startTime;
    
    // Pattern 2: User Activity Logs
    if (userId) {
        MonitoringService.info('Module discovery completed successfully', {
            totalModules: registered.length,
            moduleIds: registered.map(m => m.id),
            skippedFiles: skipped.length,
            failedModules: failed.length,
            elapsedTimeMs: elapsedTime,
            timestamp: new Date().toISOString()
        }, 'module', null, userId);
    } else if (sessionId) {
        MonitoringService.info('Module discovery completed successfully', {
            sessionId: sessionId.substring(0, 8) + '...',
            totalModules: registered.length,
            moduleIds: registered.map(m => m.id),
            skippedFiles: skipped.length,
            failedModules: failed.length,
            elapsedTimeMs: elapsedTime,
            timestamp: new Date().toISOString()
        }, 'module');
    }
    
    // Pattern 1: Development Debug Logs
    if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Module discovery process completed', {
            summary: {
                registered: registered.length,
                skipped: skipped.length,
                failed: failed.length,
                totalProcessed: files.length
            },
            registeredModules: registered.map(m => ({ id: m.id, name: m.name, capabilities: m.capabilities.length })),
            failedModules: failed,
            elapsedTimeMs: elapsedTime,
            userId: userId ? userId.substring(0, 20) + '...' : 'anonymous',
            sessionId: sessionId ? sessionId.substring(0, 8) + '...' : 'none',
            timestamp: new Date().toISOString()
        }, 'module');
    }
    
    return registered;
}

/**
 * Helper function to identify which required properties are missing from a module.
 * @param {Object} mod - The module object to check
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Array<string>} List of missing required properties
 */
function getMissingProperties(mod, userId, sessionId) {
    const startTime = Date.now();
    const requiredProps = ['id', 'name', 'capabilities', 'init', 'handleIntent'];
    const missing = [];
    
    // Pattern 1: Development Debug Logs
    if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Validating module properties', {
            moduleProvided: !!mod,
            moduleType: typeof mod,
            requiredProps,
            userId: userId ? userId.substring(0, 20) + '...' : 'anonymous',
            sessionId: sessionId ? sessionId.substring(0, 8) + '...' : 'none',
            timestamp: new Date().toISOString()
        }, 'module');
    }
    
    try {
        if (!mod) {
            // Pattern 1: Development Debug Logs
            if (process.env.NODE_ENV === 'development') {
                MonitoringService.debug('Module is null or undefined - all properties missing', {
                    missingProperties: requiredProps,
                    userId: userId ? userId.substring(0, 20) + '...' : 'anonymous',
                    sessionId: sessionId ? sessionId.substring(0, 8) + '...' : 'none',
                    timestamp: new Date().toISOString()
                }, 'module');
            }
            return requiredProps;
        }
        
        for (const prop of requiredProps) {
            if (prop === 'capabilities') {
                if (!Array.isArray(mod[prop])) missing.push(prop);
            } else if (prop === 'init' || prop === 'handleIntent') {
                if (typeof mod[prop] !== 'function') missing.push(prop);
            } else if (!mod[prop]) {
                missing.push(prop);
            }
        }
        
        const elapsedTime = Date.now() - startTime;
        
        // Pattern 2: User Activity Logs (only if there are missing properties)
        if (missing.length > 0) {
            if (userId) {
                MonitoringService.info('Module validation completed - missing properties found', {
                    missingProperties: missing,
                    totalMissing: missing.length,
                    elapsedTimeMs: elapsedTime,
                    timestamp: new Date().toISOString()
                }, 'module', null, userId);
            } else if (sessionId) {
                MonitoringService.info('Module validation completed - missing properties found', {
                    sessionId: sessionId.substring(0, 8) + '...',
                    missingProperties: missing,
                    totalMissing: missing.length,
                    elapsedTimeMs: elapsedTime,
                    timestamp: new Date().toISOString()
                }, 'module');
            }
        }
        
        // Pattern 1: Development Debug Logs
        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Module property validation completed', {
                missingProperties: missing,
                validProperties: requiredProps.filter(prop => !missing.includes(prop)),
                validationResult: missing.length === 0 ? 'valid' : 'invalid',
                elapsedTimeMs: elapsedTime,
                userId: userId ? userId.substring(0, 20) + '...' : 'anonymous',
                sessionId: sessionId ? sessionId.substring(0, 8) + '...' : 'none',
                timestamp: new Date().toISOString()
            }, 'module');
        }
        
        return missing;
        
    } catch (error) {
        // Pattern 3: Infrastructure Error Logging
        const mcpError = ErrorService.createError(
            'module',
            'Failed to validate module properties',
            'error',
            {
                errorMessage: error.message,
                errorStack: error.stack,
                moduleProvided: !!mod,
                userId: userId ? userId.substring(0, 20) + '...' : 'anonymous',
                sessionId: sessionId ? sessionId.substring(0, 8) + '...' : 'none',
                timestamp: new Date().toISOString()
            }
        );
        MonitoringService.logError(mcpError);
        
        // Pattern 4: User Error Tracking
        if (userId) {
            MonitoringService.error('Module property validation failed', {
                error: error.message,
                timestamp: new Date().toISOString()
            }, 'module', null, userId);
        } else if (sessionId) {
            MonitoringService.error('Module property validation failed', {
                sessionId: sessionId.substring(0, 8) + '...',
                error: error.message,
                timestamp: new Date().toISOString()
            }, 'module');
        }
        
        // Return all properties as missing if validation fails
        return requiredProps;
    }
}

module.exports = { 
    discoverModules,
    // Export for testing purposes
    getMissingProperties
};
