/**
 * @fileoverview To-Do Controller - Handles API requests for Microsoft To-Do API.
 * Follows MCP modular, testable, and consistent API contract rules.
 */

const Joi = require('joi');
const ErrorService = require('../../core/error-service.cjs');
const MonitoringService = require('../../core/monitoring-service.cjs');

/**
 * Joi validation schemas for todo endpoints
 */
const schemas = {
  listTaskLists: Joi.object({
    limit: Joi.number().integer().min(1).max(100).optional()
  }),

  getTaskList: Joi.object({
    listId: Joi.string().required()
  }),

  createTaskList: Joi.object({
    displayName: Joi.string().required().max(255)
  }),

  updateTaskList: Joi.object({
    listId: Joi.string().required(),
    displayName: Joi.string().optional().max(255)
  }),

  deleteTaskList: Joi.object({
    listId: Joi.string().required()
  }),

  listTasks: Joi.object({
    listId: Joi.string().required(),
    limit: Joi.number().integer().min(1).max(100).optional(),
    filter: Joi.string().optional(),
    orderby: Joi.string().optional()
  }),

  getTask: Joi.object({
    listId: Joi.string().required(),
    taskId: Joi.string().required()
  }),

  createTask: Joi.object({
    listId: Joi.string().required(),
    title: Joi.string().required().max(255),
    body: Joi.alternatives().try(
      Joi.string(),
      Joi.object({
        content: Joi.string().required(),
        contentType: Joi.string().valid('text', 'html').optional()
      })
    ).optional(),
    importance: Joi.string().valid('low', 'normal', 'high').optional(),
    dueDateTime: Joi.alternatives().try(
      Joi.string().isoDate(),
      Joi.object({
        dateTime: Joi.string().required(),
        timeZone: Joi.string().optional()
      })
    ).optional(),
    reminderDateTime: Joi.alternatives().try(
      Joi.string().isoDate(),
      Joi.object({
        dateTime: Joi.string().required(),
        timeZone: Joi.string().optional()
      })
    ).optional(),
    isReminderOn: Joi.boolean().optional(),
    categories: Joi.array().items(Joi.string()).optional()
  }),

  updateTask: Joi.object({
    listId: Joi.string().required(),
    taskId: Joi.string().required(),
    title: Joi.string().max(255).optional(),
    body: Joi.alternatives().try(
      Joi.string(),
      Joi.object({
        content: Joi.string().required(),
        contentType: Joi.string().valid('text', 'html').optional()
      })
    ).optional(),
    importance: Joi.string().valid('low', 'normal', 'high').optional(),
    status: Joi.string().valid('notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred').optional(),
    dueDateTime: Joi.alternatives().try(
      Joi.string().isoDate(),
      Joi.object({
        dateTime: Joi.string().required(),
        timeZone: Joi.string().optional()
      })
    ).optional(),
    reminderDateTime: Joi.alternatives().try(
      Joi.string().isoDate(),
      Joi.object({
        dateTime: Joi.string().required(),
        timeZone: Joi.string().optional()
      })
    ).optional(),
    isReminderOn: Joi.boolean().optional(),
    categories: Joi.array().items(Joi.string()).optional()
  }),

  deleteTask: Joi.object({
    listId: Joi.string().required(),
    taskId: Joi.string().required()
  }),

  completeTask: Joi.object({
    listId: Joi.string().required(),
    taskId: Joi.string().required()
  })
};

/**
 * Creates a todo controller with injected dependencies.
 * @param {object} deps - Controller dependencies
 * @param {object} deps.todoModule - Initialized todo module
 * @returns {object} Controller methods
 */
function createTodoController({ todoModule }) {
  if (!todoModule) {
    throw new Error('To-Do module is required for TodoController');
  }

  return {
    // ═══════════════════════════════════════════════════════════════
    // TASK LIST ENDPOINTS
    // ═══════════════════════════════════════════════════════════════

    /**
     * List all task lists for the current user.
     */
    async listTaskLists(req, res) {
      const { userId = null } = req.user || {};
      const sessionId = req.session?.id;
      const startTime = Date.now();

      try {
        const { error, value } = schemas.listTaskLists.validate(req.query);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const options = { top: value.limit || 50 };
        const lists = await todoModule.listTaskLists(options, req, userId, sessionId);

        MonitoringService.trackMetric('todo.listTaskLists.duration', Date.now() - startTime, {
          count: lists.length,
          success: true,
          userId
        });

        res.json({ success: true, data: lists, count: lists.length });
      } catch (error) {
        MonitoringService.logError(ErrorService.createError('todo', 'Failed to list task lists', 'error', {
          error: error.message,
          userId,
          timestamp: new Date().toISOString()
        }));

        res.status(500).json({
          error: 'TODO_OPERATION_FAILED',
          error_description: 'Failed to list task lists'
        });
      }
    },

    /**
     * Get a specific task list by ID.
     */
    async getTaskList(req, res) {
      const { userId = null } = req.user || {};
      const sessionId = req.session?.id;
      const startTime = Date.now();

      try {
        const { error, value } = schemas.getTaskList.validate(req.params);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const list = await todoModule.getTaskList(value.listId, req, userId, sessionId);

        MonitoringService.trackMetric('todo.getTaskList.duration', Date.now() - startTime, {
          success: true,
          userId
        });

        res.json({ success: true, data: list });
      } catch (error) {
        MonitoringService.logError(ErrorService.createError('todo', 'Failed to get task list', 'error', {
          listId: req.params.listId,
          error: error.message,
          userId,
          timestamp: new Date().toISOString()
        }));

        res.status(500).json({
          error: 'TODO_OPERATION_FAILED',
          error_description: 'Failed to get task list'
        });
      }
    },

    /**
     * Create a new task list.
     */
    async createTaskList(req, res) {
      const { userId = null } = req.user || {};
      const sessionId = req.session?.id;
      const startTime = Date.now();

      try {
        const { error, value } = schemas.createTaskList.validate(req.body);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const list = await todoModule.createTaskList({ displayName: value.displayName }, req, userId, sessionId);

        MonitoringService.trackMetric('todo.createTaskList.duration', Date.now() - startTime, {
          success: true,
          userId
        });

        res.status(201).json({ success: true, data: list, created: true });
      } catch (error) {
        MonitoringService.logError(ErrorService.createError('todo', 'Failed to create task list', 'error', {
          displayName: req.body.displayName,
          error: error.message,
          userId,
          timestamp: new Date().toISOString()
        }));

        res.status(500).json({
          error: 'TODO_OPERATION_FAILED',
          error_description: 'Failed to create task list'
        });
      }
    },

    /**
     * Update a task list.
     */
    async updateTaskList(req, res) {
      const { userId = null } = req.user || {};
      const sessionId = req.session?.id;
      const startTime = Date.now();

      try {
        const combined = { ...req.params, ...req.body };
        const { error, value } = schemas.updateTaskList.validate(combined);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const updates = {};
        if (value.displayName) updates.displayName = value.displayName;

        const list = await todoModule.updateTaskList(value.listId, updates, req, userId, sessionId);

        MonitoringService.trackMetric('todo.updateTaskList.duration', Date.now() - startTime, {
          success: true,
          userId
        });

        res.json({ success: true, data: list, updated: true });
      } catch (error) {
        MonitoringService.logError(ErrorService.createError('todo', 'Failed to update task list', 'error', {
          listId: req.params.listId,
          error: error.message,
          userId,
          timestamp: new Date().toISOString()
        }));

        res.status(500).json({
          error: 'TODO_OPERATION_FAILED',
          error_description: 'Failed to update task list'
        });
      }
    },

    /**
     * Delete a task list.
     */
    async deleteTaskList(req, res) {
      const { userId = null } = req.user || {};
      const sessionId = req.session?.id;
      const startTime = Date.now();

      try {
        const { error, value } = schemas.deleteTaskList.validate(req.params);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        await todoModule.deleteTaskList(value.listId, req, userId, sessionId);

        MonitoringService.trackMetric('todo.deleteTaskList.duration', Date.now() - startTime, {
          success: true,
          userId
        });

        res.json({ success: true, deleted: true, listId: value.listId });
      } catch (error) {
        MonitoringService.logError(ErrorService.createError('todo', 'Failed to delete task list', 'error', {
          listId: req.params.listId,
          error: error.message,
          userId,
          timestamp: new Date().toISOString()
        }));

        res.status(500).json({
          error: 'TODO_OPERATION_FAILED',
          error_description: 'Failed to delete task list'
        });
      }
    },

    // ═══════════════════════════════════════════════════════════════
    // TASK ENDPOINTS
    // ═══════════════════════════════════════════════════════════════

    /**
     * List tasks in a task list.
     */
    async listTasks(req, res) {
      const { userId = null } = req.user || {};
      const sessionId = req.session?.id;
      const startTime = Date.now();

      try {
        const combined = { ...req.params, ...req.query };
        const { error, value } = schemas.listTasks.validate(combined);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const options = {
          top: value.limit || 50,
          filter: value.filter,
          orderby: value.orderby
        };

        const tasks = await todoModule.listTasks(value.listId, options, req, userId, sessionId);

        MonitoringService.trackMetric('todo.listTasks.duration', Date.now() - startTime, {
          count: tasks.length,
          success: true,
          userId
        });

        res.json({ success: true, data: tasks, count: tasks.length });
      } catch (error) {
        MonitoringService.logError(ErrorService.createError('todo', 'Failed to list tasks', 'error', {
          listId: req.params.listId,
          error: error.message,
          userId,
          timestamp: new Date().toISOString()
        }));

        res.status(500).json({
          error: 'TODO_OPERATION_FAILED',
          error_description: 'Failed to list tasks'
        });
      }
    },

    /**
     * Get a specific task.
     */
    async getTask(req, res) {
      const { userId = null } = req.user || {};
      const sessionId = req.session?.id;
      const startTime = Date.now();

      try {
        const { error, value } = schemas.getTask.validate(req.params);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const task = await todoModule.getTask(value.listId, value.taskId, req, userId, sessionId);

        MonitoringService.trackMetric('todo.getTask.duration', Date.now() - startTime, {
          success: true,
          userId
        });

        res.json({ success: true, data: task });
      } catch (error) {
        MonitoringService.logError(ErrorService.createError('todo', 'Failed to get task', 'error', {
          listId: req.params.listId,
          taskId: req.params.taskId,
          error: error.message,
          userId,
          timestamp: new Date().toISOString()
        }));

        res.status(500).json({
          error: 'TODO_OPERATION_FAILED',
          error_description: 'Failed to get task'
        });
      }
    },

    /**
     * Create a new task.
     */
    async createTask(req, res) {
      const { userId = null } = req.user || {};
      const sessionId = req.session?.id;
      const startTime = Date.now();

      try {
        const combined = { listId: req.params.listId, ...req.body };
        const { error, value } = schemas.createTask.validate(combined);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const taskData = {
          title: value.title,
          body: value.body,
          importance: value.importance,
          dueDateTime: value.dueDateTime,
          reminderDateTime: value.reminderDateTime,
          isReminderOn: value.isReminderOn,
          categories: value.categories
        };

        const task = await todoModule.createTask(value.listId, taskData, req, userId, sessionId);

        MonitoringService.trackMetric('todo.createTask.duration', Date.now() - startTime, {
          success: true,
          userId
        });

        res.status(201).json({ success: true, data: task, created: true });
      } catch (error) {
        MonitoringService.logError(ErrorService.createError('todo', 'Failed to create task', 'error', {
          listId: req.params.listId,
          title: req.body.title,
          error: error.message,
          userId,
          timestamp: new Date().toISOString()
        }));

        res.status(500).json({
          error: 'TODO_OPERATION_FAILED',
          error_description: 'Failed to create task'
        });
      }
    },

    /**
     * Update a task.
     */
    async updateTask(req, res) {
      const { userId = null } = req.user || {};
      const sessionId = req.session?.id;
      const startTime = Date.now();

      try {
        const combined = { ...req.params, ...req.body };
        const { error, value } = schemas.updateTask.validate(combined);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const updates = {};
        if (value.title) updates.title = value.title;
        if (value.body) updates.body = value.body;
        if (value.importance) updates.importance = value.importance;
        if (value.status) updates.status = value.status;
        if (value.dueDateTime) updates.dueDateTime = value.dueDateTime;
        if (value.reminderDateTime) updates.reminderDateTime = value.reminderDateTime;
        if (typeof value.isReminderOn === 'boolean') updates.isReminderOn = value.isReminderOn;
        if (value.categories) updates.categories = value.categories;

        const task = await todoModule.updateTask(value.listId, value.taskId, updates, req, userId, sessionId);

        MonitoringService.trackMetric('todo.updateTask.duration', Date.now() - startTime, {
          success: true,
          userId
        });

        res.json({ success: true, data: task, updated: true });
      } catch (error) {
        MonitoringService.logError(ErrorService.createError('todo', 'Failed to update task', 'error', {
          listId: req.params.listId,
          taskId: req.params.taskId,
          error: error.message,
          userId,
          timestamp: new Date().toISOString()
        }));

        res.status(500).json({
          error: 'TODO_OPERATION_FAILED',
          error_description: 'Failed to update task'
        });
      }
    },

    /**
     * Delete a task.
     */
    async deleteTask(req, res) {
      const { userId = null } = req.user || {};
      const sessionId = req.session?.id;
      const startTime = Date.now();

      try {
        const { error, value } = schemas.deleteTask.validate(req.params);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        await todoModule.deleteTask(value.listId, value.taskId, req, userId, sessionId);

        MonitoringService.trackMetric('todo.deleteTask.duration', Date.now() - startTime, {
          success: true,
          userId
        });

        res.json({ success: true, deleted: true, taskId: value.taskId });
      } catch (error) {
        MonitoringService.logError(ErrorService.createError('todo', 'Failed to delete task', 'error', {
          listId: req.params.listId,
          taskId: req.params.taskId,
          error: error.message,
          userId,
          timestamp: new Date().toISOString()
        }));

        res.status(500).json({
          error: 'TODO_OPERATION_FAILED',
          error_description: 'Failed to delete task'
        });
      }
    },

    /**
     * Mark a task as complete.
     */
    async completeTask(req, res) {
      const { userId = null } = req.user || {};
      const sessionId = req.session?.id;
      const startTime = Date.now();

      try {
        const { error, value } = schemas.completeTask.validate(req.params);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const task = await todoModule.completeTask(value.listId, value.taskId, req, userId, sessionId);

        MonitoringService.trackMetric('todo.completeTask.duration', Date.now() - startTime, {
          success: true,
          userId
        });

        res.json({ success: true, data: task, completed: true });
      } catch (error) {
        MonitoringService.logError(ErrorService.createError('todo', 'Failed to complete task', 'error', {
          listId: req.params.listId,
          taskId: req.params.taskId,
          error: error.message,
          userId,
          timestamp: new Date().toISOString()
        }));

        res.status(500).json({
          error: 'TODO_OPERATION_FAILED',
          error_description: 'Failed to complete task'
        });
      }
    }
  };
}

module.exports = createTodoController;
