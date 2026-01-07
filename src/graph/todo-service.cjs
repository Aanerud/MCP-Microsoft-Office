/**
 * @fileoverview TodoService - Microsoft Graph To-Do API operations.
 * All methods are async, modular, and use GraphClient for requests.
 * Follows project error handling, validation, and normalization rules.
 */

const graphClientFactory = require('./graph-client.cjs');
const MonitoringService = require('../core/monitoring-service.cjs');
const ErrorService = require('../core/error-service.cjs');

/**
 * Normalizes a task list object from Graph API response.
 * @param {object} list - Raw task list object from Graph API
 * @returns {object} Normalized task list object
 */
function normalizeTaskList(list) {
  return {
    id: list.id,
    displayName: list.displayName,
    isOwner: list.isOwner,
    isShared: list.isShared,
    wellknownListName: list.wellknownListName || 'none'
  };
}

/**
 * Normalizes a task object from Graph API response.
 * @param {object} task - Raw task object from Graph API
 * @returns {object} Normalized task object
 */
function normalizeTask(task) {
  return {
    id: task.id,
    title: task.title,
    body: task.body ? {
      content: task.body.content,
      contentType: task.body.contentType
    } : null,
    importance: task.importance || 'normal',
    status: task.status || 'notStarted',
    isReminderOn: task.isReminderOn || false,
    createdDateTime: task.createdDateTime,
    lastModifiedDateTime: task.lastModifiedDateTime,
    completedDateTime: task.completedDateTime ? task.completedDateTime.dateTime : null,
    dueDateTime: task.dueDateTime ? {
      dateTime: task.dueDateTime.dateTime,
      timeZone: task.dueDateTime.timeZone
    } : null,
    reminderDateTime: task.reminderDateTime ? {
      dateTime: task.reminderDateTime.dateTime,
      timeZone: task.reminderDateTime.timeZone
    } : null,
    categories: task.categories || [],
    linkedResources: task.linkedResources || []
  };
}

// ═══════════════════════════════════════════════════════════════
// TASK LIST OPERATIONS
// ═══════════════════════════════════════════════════════════════

/**
 * Gets all task lists for the current user.
 * @param {object} options - Query options
 * @param {number} [options.top=50] - Number of lists to retrieve
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<Array<object>>} Normalized task list objects
 */
async function listTaskLists(options = {}, req, userId, sessionId) {
  const startTime = new Date();

  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;

  try {
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Listing task lists', {
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString(),
        options: { top: options.top || 50 }
      }, 'todo');
    }

    const client = await graphClientFactory.createClient(req, resolvedUserId, resolvedSessionId);
    const top = options.top || 50;

    const url = `/me/todo/lists?$top=${top}`;
    const res = await client.api(url, resolvedUserId, resolvedSessionId).get();

    const endTime = new Date();
    const duration = endTime - startTime;

    const lists = (res.value || []).map(normalizeTaskList);

    if (resolvedUserId) {
      MonitoringService.info('Retrieved task lists successfully', {
        count: lists.length,
        duration: `${duration}ms`,
        timestamp: new Date().toISOString()
      }, 'todo', null, resolvedUserId);
    }

    return lists;
  } catch (error) {
    const endTime = new Date();
    const duration = endTime - startTime;

    const mcpError = ErrorService.createError(
      'todo',
      'Failed to list task lists',
      'error',
      {
        endpoint: '/me/todo/lists',
        error: error.message,
        stack: error.stack,
        duration: `${duration}ms`,
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString()
      }
    );
    MonitoringService.logError(mcpError);

    throw error;
  }
}

/**
 * Gets a specific task list by ID.
 * @param {string} listId - Task list ID
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<object>} Normalized task list object
 */
async function getTaskList(listId, req, userId, sessionId) {
  const startTime = new Date();

  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;

  try {
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Getting task list', {
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString(),
        listId
      }, 'todo');
    }

    const client = await graphClientFactory.createClient(req, resolvedUserId, resolvedSessionId);

    const url = `/me/todo/lists/${listId}`;
    const res = await client.api(url, resolvedUserId, resolvedSessionId).get();

    const endTime = new Date();
    const duration = endTime - startTime;

    if (resolvedUserId) {
      MonitoringService.info('Retrieved task list successfully', {
        listId,
        duration: `${duration}ms`,
        timestamp: new Date().toISOString()
      }, 'todo', null, resolvedUserId);
    }

    return normalizeTaskList(res);
  } catch (error) {
    const endTime = new Date();
    const duration = endTime - startTime;

    const mcpError = ErrorService.createError(
      'todo',
      'Failed to get task list',
      'error',
      {
        endpoint: `/me/todo/lists/${listId}`,
        listId,
        error: error.message,
        stack: error.stack,
        duration: `${duration}ms`,
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString()
      }
    );
    MonitoringService.logError(mcpError);

    throw error;
  }
}

/**
 * Creates a new task list.
 * @param {object} listData - Task list data
 * @param {string} listData.displayName - Display name for the list
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<object>} Created task list object
 */
async function createTaskList(listData, req, userId, sessionId) {
  const startTime = new Date();

  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;

  try {
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Creating task list', {
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString(),
        displayName: listData.displayName
      }, 'todo');
    }

    const client = await graphClientFactory.createClient(req, resolvedUserId, resolvedSessionId);

    const url = '/me/todo/lists';
    const res = await client.api(url, resolvedUserId, resolvedSessionId).post({
      displayName: listData.displayName
    });

    const endTime = new Date();
    const duration = endTime - startTime;

    if (resolvedUserId) {
      MonitoringService.info('Created task list successfully', {
        listId: res.id,
        displayName: listData.displayName,
        duration: `${duration}ms`,
        timestamp: new Date().toISOString()
      }, 'todo', null, resolvedUserId);
    }

    return normalizeTaskList(res);
  } catch (error) {
    const endTime = new Date();
    const duration = endTime - startTime;

    const mcpError = ErrorService.createError(
      'todo',
      'Failed to create task list',
      'error',
      {
        endpoint: '/me/todo/lists',
        displayName: listData.displayName,
        error: error.message,
        stack: error.stack,
        duration: `${duration}ms`,
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString()
      }
    );
    MonitoringService.logError(mcpError);

    throw error;
  }
}

/**
 * Updates a task list.
 * @param {string} listId - Task list ID
 * @param {object} updates - Update data
 * @param {string} [updates.displayName] - New display name
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<object>} Updated task list object
 */
async function updateTaskList(listId, updates, req, userId, sessionId) {
  const startTime = new Date();

  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;

  try {
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Updating task list', {
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString(),
        listId,
        updates: Object.keys(updates)
      }, 'todo');
    }

    const client = await graphClientFactory.createClient(req, resolvedUserId, resolvedSessionId);

    const url = `/me/todo/lists/${listId}`;
    const res = await client.api(url, resolvedUserId, resolvedSessionId).patch(updates);

    const endTime = new Date();
    const duration = endTime - startTime;

    if (resolvedUserId) {
      MonitoringService.info('Updated task list successfully', {
        listId,
        duration: `${duration}ms`,
        timestamp: new Date().toISOString()
      }, 'todo', null, resolvedUserId);
    }

    return normalizeTaskList(res);
  } catch (error) {
    const endTime = new Date();
    const duration = endTime - startTime;

    const mcpError = ErrorService.createError(
      'todo',
      'Failed to update task list',
      'error',
      {
        endpoint: `/me/todo/lists/${listId}`,
        listId,
        error: error.message,
        stack: error.stack,
        duration: `${duration}ms`,
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString()
      }
    );
    MonitoringService.logError(mcpError);

    throw error;
  }
}

/**
 * Deletes a task list.
 * @param {string} listId - Task list ID
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<boolean>} True if deleted successfully
 */
async function deleteTaskList(listId, req, userId, sessionId) {
  const startTime = new Date();

  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;

  try {
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Deleting task list', {
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString(),
        listId
      }, 'todo');
    }

    const client = await graphClientFactory.createClient(req, resolvedUserId, resolvedSessionId);

    const url = `/me/todo/lists/${listId}`;
    await client.api(url, resolvedUserId, resolvedSessionId).delete();

    const endTime = new Date();
    const duration = endTime - startTime;

    if (resolvedUserId) {
      MonitoringService.info('Deleted task list successfully', {
        listId,
        duration: `${duration}ms`,
        timestamp: new Date().toISOString()
      }, 'todo', null, resolvedUserId);
    }

    return true;
  } catch (error) {
    const endTime = new Date();
    const duration = endTime - startTime;

    const mcpError = ErrorService.createError(
      'todo',
      'Failed to delete task list',
      'error',
      {
        endpoint: `/me/todo/lists/${listId}`,
        listId,
        error: error.message,
        stack: error.stack,
        duration: `${duration}ms`,
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString()
      }
    );
    MonitoringService.logError(mcpError);

    throw error;
  }
}

// ═══════════════════════════════════════════════════════════════
// TASK OPERATIONS
// ═══════════════════════════════════════════════════════════════

/**
 * Gets all tasks in a task list.
 * @param {string} listId - Task list ID
 * @param {object} options - Query options
 * @param {number} [options.top=50] - Number of tasks to retrieve
 * @param {string} [options.filter] - OData filter (e.g., "status eq 'notStarted'")
 * @param {string} [options.orderby] - Order by field
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<Array<object>>} Normalized task objects
 */
async function listTasks(listId, options = {}, req, userId, sessionId) {
  const startTime = new Date();

  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;

  try {
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Listing tasks', {
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString(),
        listId,
        options: { top: options.top || 50, filter: options.filter }
      }, 'todo');
    }

    const client = await graphClientFactory.createClient(req, resolvedUserId, resolvedSessionId);
    const top = options.top || 50;

    let queryParams = [`$top=${top}`];
    if (options.filter) {
      queryParams.push(`$filter=${encodeURIComponent(options.filter)}`);
    }
    if (options.orderby) {
      queryParams.push(`$orderby=${encodeURIComponent(options.orderby)}`);
    }

    const url = `/me/todo/lists/${listId}/tasks?${queryParams.join('&')}`;
    const res = await client.api(url, resolvedUserId, resolvedSessionId).get();

    const endTime = new Date();
    const duration = endTime - startTime;

    const tasks = (res.value || []).map(normalizeTask);

    if (resolvedUserId) {
      MonitoringService.info('Retrieved tasks successfully', {
        listId,
        count: tasks.length,
        duration: `${duration}ms`,
        timestamp: new Date().toISOString()
      }, 'todo', null, resolvedUserId);
    }

    return tasks;
  } catch (error) {
    const endTime = new Date();
    const duration = endTime - startTime;

    const mcpError = ErrorService.createError(
      'todo',
      'Failed to list tasks',
      'error',
      {
        endpoint: `/me/todo/lists/${listId}/tasks`,
        listId,
        error: error.message,
        stack: error.stack,
        duration: `${duration}ms`,
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString()
      }
    );
    MonitoringService.logError(mcpError);

    throw error;
  }
}

/**
 * Gets a specific task by ID.
 * @param {string} listId - Task list ID
 * @param {string} taskId - Task ID
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<object>} Normalized task object
 */
async function getTask(listId, taskId, req, userId, sessionId) {
  const startTime = new Date();

  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;

  try {
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Getting task', {
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString(),
        listId,
        taskId
      }, 'todo');
    }

    const client = await graphClientFactory.createClient(req, resolvedUserId, resolvedSessionId);

    const url = `/me/todo/lists/${listId}/tasks/${taskId}`;
    const res = await client.api(url, resolvedUserId, resolvedSessionId).get();

    const endTime = new Date();
    const duration = endTime - startTime;

    if (resolvedUserId) {
      MonitoringService.info('Retrieved task successfully', {
        listId,
        taskId,
        duration: `${duration}ms`,
        timestamp: new Date().toISOString()
      }, 'todo', null, resolvedUserId);
    }

    return normalizeTask(res);
  } catch (error) {
    const endTime = new Date();
    const duration = endTime - startTime;

    const mcpError = ErrorService.createError(
      'todo',
      'Failed to get task',
      'error',
      {
        endpoint: `/me/todo/lists/${listId}/tasks/${taskId}`,
        listId,
        taskId,
        error: error.message,
        stack: error.stack,
        duration: `${duration}ms`,
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString()
      }
    );
    MonitoringService.logError(mcpError);

    throw error;
  }
}

/**
 * Creates a new task.
 * @param {string} listId - Task list ID
 * @param {object} taskData - Task data
 * @param {string} taskData.title - Task title (required)
 * @param {object} [taskData.body] - Task body/notes
 * @param {string} [taskData.importance] - 'low', 'normal', or 'high'
 * @param {object} [taskData.dueDateTime] - Due date/time
 * @param {object} [taskData.reminderDateTime] - Reminder date/time
 * @param {boolean} [taskData.isReminderOn] - Enable reminder
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<object>} Created task object
 */
async function createTask(listId, taskData, req, userId, sessionId) {
  const startTime = new Date();

  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;

  try {
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Creating task', {
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString(),
        listId,
        title: taskData.title
      }, 'todo');
    }

    const client = await graphClientFactory.createClient(req, resolvedUserId, resolvedSessionId);

    // Build the task object
    const task = {
      title: taskData.title
    };

    if (taskData.body) {
      task.body = {
        content: taskData.body.content || taskData.body,
        contentType: taskData.body.contentType || 'text'
      };
    }

    if (taskData.importance) {
      task.importance = taskData.importance;
    }

    if (taskData.dueDateTime) {
      task.dueDateTime = {
        dateTime: taskData.dueDateTime.dateTime || taskData.dueDateTime,
        timeZone: taskData.dueDateTime.timeZone || 'UTC'
      };
    }

    if (taskData.reminderDateTime) {
      task.reminderDateTime = {
        dateTime: taskData.reminderDateTime.dateTime || taskData.reminderDateTime,
        timeZone: taskData.reminderDateTime.timeZone || 'UTC'
      };
      task.isReminderOn = true;
    }

    if (typeof taskData.isReminderOn === 'boolean') {
      task.isReminderOn = taskData.isReminderOn;
    }

    if (taskData.categories) {
      task.categories = taskData.categories;
    }

    const url = `/me/todo/lists/${listId}/tasks`;
    const res = await client.api(url, resolvedUserId, resolvedSessionId).post(task);

    const endTime = new Date();
    const duration = endTime - startTime;

    if (resolvedUserId) {
      MonitoringService.info('Created task successfully', {
        listId,
        taskId: res.id,
        title: taskData.title,
        duration: `${duration}ms`,
        timestamp: new Date().toISOString()
      }, 'todo', null, resolvedUserId);
    }

    return normalizeTask(res);
  } catch (error) {
    const endTime = new Date();
    const duration = endTime - startTime;

    const mcpError = ErrorService.createError(
      'todo',
      'Failed to create task',
      'error',
      {
        endpoint: `/me/todo/lists/${listId}/tasks`,
        listId,
        title: taskData.title,
        error: error.message,
        stack: error.stack,
        duration: `${duration}ms`,
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString()
      }
    );
    MonitoringService.logError(mcpError);

    throw error;
  }
}

/**
 * Updates a task.
 * @param {string} listId - Task list ID
 * @param {string} taskId - Task ID
 * @param {object} updates - Update data
 * @param {string} [updates.title] - Task title
 * @param {object} [updates.body] - Task body/notes
 * @param {string} [updates.importance] - 'low', 'normal', or 'high'
 * @param {string} [updates.status] - 'notStarted', 'inProgress', 'completed', 'waitingOnOthers', 'deferred'
 * @param {object} [updates.dueDateTime] - Due date/time
 * @param {object} [updates.reminderDateTime] - Reminder date/time
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<object>} Updated task object
 */
async function updateTask(listId, taskId, updates, req, userId, sessionId) {
  const startTime = new Date();

  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;

  try {
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Updating task', {
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString(),
        listId,
        taskId,
        updates: Object.keys(updates)
      }, 'todo');
    }

    const client = await graphClientFactory.createClient(req, resolvedUserId, resolvedSessionId);

    // Build update object with proper formatting
    const taskUpdates = {};

    if (updates.title) {
      taskUpdates.title = updates.title;
    }

    if (updates.body) {
      taskUpdates.body = {
        content: updates.body.content || updates.body,
        contentType: updates.body.contentType || 'text'
      };
    }

    if (updates.importance) {
      taskUpdates.importance = updates.importance;
    }

    if (updates.status) {
      taskUpdates.status = updates.status;
    }

    if (updates.dueDateTime) {
      taskUpdates.dueDateTime = {
        dateTime: updates.dueDateTime.dateTime || updates.dueDateTime,
        timeZone: updates.dueDateTime.timeZone || 'UTC'
      };
    }

    if (updates.reminderDateTime) {
      taskUpdates.reminderDateTime = {
        dateTime: updates.reminderDateTime.dateTime || updates.reminderDateTime,
        timeZone: updates.reminderDateTime.timeZone || 'UTC'
      };
      taskUpdates.isReminderOn = true;
    }

    if (typeof updates.isReminderOn === 'boolean') {
      taskUpdates.isReminderOn = updates.isReminderOn;
    }

    if (updates.categories) {
      taskUpdates.categories = updates.categories;
    }

    const url = `/me/todo/lists/${listId}/tasks/${taskId}`;
    const res = await client.api(url, resolvedUserId, resolvedSessionId).patch(taskUpdates);

    const endTime = new Date();
    const duration = endTime - startTime;

    if (resolvedUserId) {
      MonitoringService.info('Updated task successfully', {
        listId,
        taskId,
        duration: `${duration}ms`,
        timestamp: new Date().toISOString()
      }, 'todo', null, resolvedUserId);
    }

    return normalizeTask(res);
  } catch (error) {
    const endTime = new Date();
    const duration = endTime - startTime;

    const mcpError = ErrorService.createError(
      'todo',
      'Failed to update task',
      'error',
      {
        endpoint: `/me/todo/lists/${listId}/tasks/${taskId}`,
        listId,
        taskId,
        error: error.message,
        stack: error.stack,
        duration: `${duration}ms`,
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString()
      }
    );
    MonitoringService.logError(mcpError);

    throw error;
  }
}

/**
 * Deletes a task.
 * @param {string} listId - Task list ID
 * @param {string} taskId - Task ID
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<boolean>} True if deleted successfully
 */
async function deleteTask(listId, taskId, req, userId, sessionId) {
  const startTime = new Date();

  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;

  try {
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Deleting task', {
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString(),
        listId,
        taskId
      }, 'todo');
    }

    const client = await graphClientFactory.createClient(req, resolvedUserId, resolvedSessionId);

    const url = `/me/todo/lists/${listId}/tasks/${taskId}`;
    await client.api(url, resolvedUserId, resolvedSessionId).delete();

    const endTime = new Date();
    const duration = endTime - startTime;

    if (resolvedUserId) {
      MonitoringService.info('Deleted task successfully', {
        listId,
        taskId,
        duration: `${duration}ms`,
        timestamp: new Date().toISOString()
      }, 'todo', null, resolvedUserId);
    }

    return true;
  } catch (error) {
    const endTime = new Date();
    const duration = endTime - startTime;

    const mcpError = ErrorService.createError(
      'todo',
      'Failed to delete task',
      'error',
      {
        endpoint: `/me/todo/lists/${listId}/tasks/${taskId}`,
        listId,
        taskId,
        error: error.message,
        stack: error.stack,
        duration: `${duration}ms`,
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString()
      }
    );
    MonitoringService.logError(mcpError);

    throw error;
  }
}

/**
 * Marks a task as complete.
 * @param {string} listId - Task list ID
 * @param {string} taskId - Task ID
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<object>} Updated task object
 */
async function completeTask(listId, taskId, req, userId, sessionId) {
  return updateTask(listId, taskId, { status: 'completed' }, req, userId, sessionId);
}

module.exports = {
  // Task list operations
  listTaskLists,
  getTaskList,
  createTaskList,
  updateTaskList,
  deleteTaskList,
  // Task operations
  listTasks,
  getTask,
  createTask,
  updateTask,
  deleteTask,
  completeTask
};
