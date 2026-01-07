/**
 * @fileoverview MCP To-Do Module - Handles Microsoft To-Do intents and actions for MCP.
 * Exposes: id, name, capabilities, init, handleIntent. Aligned with MCP module system.
 */

const MonitoringService = require('../../core/monitoring-service.cjs');
const ErrorService = require('../../core/error-service.cjs');

const TODO_CAPABILITIES = [
  'listTaskLists',
  'getTaskList',
  'createTaskList',
  'updateTaskList',
  'deleteTaskList',
  'listTasks',
  'getTask',
  'createTask',
  'updateTask',
  'deleteTask',
  'completeTask'
];

// Log module initialization
MonitoringService.info('To-Do Module initialized', {
  serviceName: 'todo-module',
  capabilities: TODO_CAPABILITIES.length,
  timestamp: new Date().toISOString()
}, 'todo');

const TodoModule = {
  /**
   * Initializes the To-Do module with dependencies.
   * @param {object} services - { todoService, cacheService, errorService, monitoringService }
   * @returns {object} Initialized module
   */
  init(services = {}) {
    const { todoService, cacheService, errorService = ErrorService, monitoringService = MonitoringService } = services;

    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Initializing To-Do Module', {
        hasTodoService: !!todoService,
        hasCacheService: !!cacheService,
        timestamp: new Date().toISOString()
      }, 'todo');
    }

    if (!todoService) {
      const mcpError = ErrorService.createError(
        'todo',
        "TodoModule init failed: Required service 'todoService' is missing",
        'error',
        { missingService: 'todoService', timestamp: new Date().toISOString() }
      );
      MonitoringService.logError(mcpError);
      throw mcpError;
    }

    this.services = { todoService, cacheService, errorService, monitoringService };
    return this;
  },

  // ═══════════════════════════════════════════════════════════════
  // TASK LIST OPERATIONS
  // ═══════════════════════════════════════════════════════════════

  /**
   * List all task lists
   */
  async listTaskLists(options = {}, req, userId, sessionId) {
    const { todoService } = this.services || {};
    const contextUserId = userId || req?.user?.userId;
    const contextSessionId = sessionId || req?.session?.id;

    if (!todoService) {
      throw ErrorService.createError('todo', 'TodoService not available', 'error', {});
    }

    return await todoService.listTaskLists(options, req, contextUserId, contextSessionId);
  },

  /**
   * Get a specific task list
   */
  async getTaskList(listId, req, userId, sessionId) {
    const { todoService } = this.services || {};
    const contextUserId = userId || req?.user?.userId;
    const contextSessionId = sessionId || req?.session?.id;

    if (!todoService) {
      throw ErrorService.createError('todo', 'TodoService not available', 'error', {});
    }

    return await todoService.getTaskList(listId, req, contextUserId, contextSessionId);
  },

  /**
   * Create a new task list
   */
  async createTaskList(listData, req, userId, sessionId) {
    const { todoService } = this.services || {};
    const contextUserId = userId || req?.user?.userId;
    const contextSessionId = sessionId || req?.session?.id;

    if (!todoService) {
      throw ErrorService.createError('todo', 'TodoService not available', 'error', {});
    }

    return await todoService.createTaskList(listData, req, contextUserId, contextSessionId);
  },

  /**
   * Update a task list
   */
  async updateTaskList(listId, updates, req, userId, sessionId) {
    const { todoService } = this.services || {};
    const contextUserId = userId || req?.user?.userId;
    const contextSessionId = sessionId || req?.session?.id;

    if (!todoService) {
      throw ErrorService.createError('todo', 'TodoService not available', 'error', {});
    }

    return await todoService.updateTaskList(listId, updates, req, contextUserId, contextSessionId);
  },

  /**
   * Delete a task list
   */
  async deleteTaskList(listId, req, userId, sessionId) {
    const { todoService } = this.services || {};
    const contextUserId = userId || req?.user?.userId;
    const contextSessionId = sessionId || req?.session?.id;

    if (!todoService) {
      throw ErrorService.createError('todo', 'TodoService not available', 'error', {});
    }

    return await todoService.deleteTaskList(listId, req, contextUserId, contextSessionId);
  },

  // ═══════════════════════════════════════════════════════════════
  // TASK OPERATIONS
  // ═══════════════════════════════════════════════════════════════

  /**
   * List tasks in a task list
   */
  async listTasks(listId, options = {}, req, userId, sessionId) {
    const { todoService } = this.services || {};
    const contextUserId = userId || req?.user?.userId;
    const contextSessionId = sessionId || req?.session?.id;

    if (!todoService) {
      throw ErrorService.createError('todo', 'TodoService not available', 'error', {});
    }

    return await todoService.listTasks(listId, options, req, contextUserId, contextSessionId);
  },

  /**
   * Get a specific task
   */
  async getTask(listId, taskId, req, userId, sessionId) {
    const { todoService } = this.services || {};
    const contextUserId = userId || req?.user?.userId;
    const contextSessionId = sessionId || req?.session?.id;

    if (!todoService) {
      throw ErrorService.createError('todo', 'TodoService not available', 'error', {});
    }

    return await todoService.getTask(listId, taskId, req, contextUserId, contextSessionId);
  },

  /**
   * Create a new task
   */
  async createTask(listId, taskData, req, userId, sessionId) {
    const { todoService } = this.services || {};
    const contextUserId = userId || req?.user?.userId;
    const contextSessionId = sessionId || req?.session?.id;

    if (!todoService) {
      throw ErrorService.createError('todo', 'TodoService not available', 'error', {});
    }

    return await todoService.createTask(listId, taskData, req, contextUserId, contextSessionId);
  },

  /**
   * Update a task
   */
  async updateTask(listId, taskId, updates, req, userId, sessionId) {
    const { todoService } = this.services || {};
    const contextUserId = userId || req?.user?.userId;
    const contextSessionId = sessionId || req?.session?.id;

    if (!todoService) {
      throw ErrorService.createError('todo', 'TodoService not available', 'error', {});
    }

    return await todoService.updateTask(listId, taskId, updates, req, contextUserId, contextSessionId);
  },

  /**
   * Delete a task
   */
  async deleteTask(listId, taskId, req, userId, sessionId) {
    const { todoService } = this.services || {};
    const contextUserId = userId || req?.user?.userId;
    const contextSessionId = sessionId || req?.session?.id;

    if (!todoService) {
      throw ErrorService.createError('todo', 'TodoService not available', 'error', {});
    }

    return await todoService.deleteTask(listId, taskId, req, contextUserId, contextSessionId);
  },

  /**
   * Mark a task as complete
   */
  async completeTask(listId, taskId, req, userId, sessionId) {
    const { todoService } = this.services || {};
    const contextUserId = userId || req?.user?.userId;
    const contextSessionId = sessionId || req?.session?.id;

    if (!todoService) {
      throw ErrorService.createError('todo', 'TodoService not available', 'error', {});
    }

    return await todoService.completeTask(listId, taskId, req, contextUserId, contextSessionId);
  },

  /**
   * Handles To-Do related intents routed to this module.
   * @param {string} intent
   * @param {object} entities
   * @param {object} context
   * @returns {Promise<object>} Normalized response
   */
  async handleIntent(intent, entities = {}, context = {}) {
    const startTime = Date.now();
    const contextUserId = context?.req?.user?.userId;
    const contextSessionId = context?.req?.session?.id;

    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Handling To-Do intent', {
        sessionId: contextSessionId,
        intent,
        entities: Object.keys(entities),
        timestamp: new Date().toISOString()
      }, 'todo');
    }

    try {
      let result;

      switch (intent) {
        case 'listTaskLists': {
          const lists = await this.listTaskLists(entities.options || {}, context.req, contextUserId, contextSessionId);
          result = { type: 'taskListCollection', items: lists };
          break;
        }

        case 'getTaskList': {
          if (!entities.listId) {
            throw ErrorService.createError('todo', 'listId is required', 'warn', { intent });
          }
          const list = await this.getTaskList(entities.listId, context.req, contextUserId, contextSessionId);
          result = { type: 'taskList', data: list };
          break;
        }

        case 'createTaskList': {
          if (!entities.displayName) {
            throw ErrorService.createError('todo', 'displayName is required', 'warn', { intent });
          }
          const list = await this.createTaskList({ displayName: entities.displayName }, context.req, contextUserId, contextSessionId);
          result = { type: 'taskList', data: list, created: true };
          break;
        }

        case 'updateTaskList': {
          if (!entities.listId) {
            throw ErrorService.createError('todo', 'listId is required', 'warn', { intent });
          }
          const list = await this.updateTaskList(entities.listId, entities.updates || {}, context.req, contextUserId, contextSessionId);
          result = { type: 'taskList', data: list, updated: true };
          break;
        }

        case 'deleteTaskList': {
          if (!entities.listId) {
            throw ErrorService.createError('todo', 'listId is required', 'warn', { intent });
          }
          await this.deleteTaskList(entities.listId, context.req, contextUserId, contextSessionId);
          result = { type: 'taskList', deleted: true, listId: entities.listId };
          break;
        }

        case 'listTasks': {
          if (!entities.listId) {
            throw ErrorService.createError('todo', 'listId is required', 'warn', { intent });
          }
          const tasks = await this.listTasks(entities.listId, entities.options || {}, context.req, contextUserId, contextSessionId);
          result = { type: 'taskCollection', items: tasks, listId: entities.listId };
          break;
        }

        case 'getTask': {
          if (!entities.listId || !entities.taskId) {
            throw ErrorService.createError('todo', 'listId and taskId are required', 'warn', { intent });
          }
          const task = await this.getTask(entities.listId, entities.taskId, context.req, contextUserId, contextSessionId);
          result = { type: 'task', data: task };
          break;
        }

        case 'createTask': {
          if (!entities.listId || !entities.title) {
            throw ErrorService.createError('todo', 'listId and title are required', 'warn', { intent });
          }
          const taskData = {
            title: entities.title,
            body: entities.body,
            importance: entities.importance,
            dueDateTime: entities.dueDateTime,
            reminderDateTime: entities.reminderDateTime,
            isReminderOn: entities.isReminderOn,
            categories: entities.categories
          };
          const task = await this.createTask(entities.listId, taskData, context.req, contextUserId, contextSessionId);
          result = { type: 'task', data: task, created: true };
          break;
        }

        case 'updateTask': {
          if (!entities.listId || !entities.taskId) {
            throw ErrorService.createError('todo', 'listId and taskId are required', 'warn', { intent });
          }
          const task = await this.updateTask(entities.listId, entities.taskId, entities.updates || {}, context.req, contextUserId, contextSessionId);
          result = { type: 'task', data: task, updated: true };
          break;
        }

        case 'deleteTask': {
          if (!entities.listId || !entities.taskId) {
            throw ErrorService.createError('todo', 'listId and taskId are required', 'warn', { intent });
          }
          await this.deleteTask(entities.listId, entities.taskId, context.req, contextUserId, contextSessionId);
          result = { type: 'task', deleted: true, taskId: entities.taskId };
          break;
        }

        case 'completeTask': {
          if (!entities.listId || !entities.taskId) {
            throw ErrorService.createError('todo', 'listId and taskId are required', 'warn', { intent });
          }
          const task = await this.completeTask(entities.listId, entities.taskId, context.req, contextUserId, contextSessionId);
          result = { type: 'task', data: task, completed: true };
          break;
        }

        default:
          throw ErrorService.createError('todo', `TodoModule cannot handle intent: ${intent}`, 'warn', { intent });
      }

      const elapsedTime = Date.now() - startTime;

      if (contextUserId) {
        MonitoringService.info('Successfully handled To-Do intent', {
          intent,
          elapsedTime,
          timestamp: new Date().toISOString()
        }, 'todo', null, contextUserId);
      }

      return result;
    } catch (error) {
      const elapsedTime = Date.now() - startTime;

      if (error.category && error.severity) {
        throw error;
      }

      const mcpError = ErrorService.createError(
        'todo',
        `Error handling To-Do intent '${intent}': ${error.message}`,
        'error',
        { intent, originalError: error.stack, elapsedTime, timestamp: new Date().toISOString() }
      );
      MonitoringService.logError(mcpError);
      throw mcpError;
    }
  },

  id: 'todo',
  name: 'Microsoft To-Do',
  capabilities: TODO_CAPABILITIES
};

module.exports = TodoModule;
