/**
 * @fileoverview Groups Controller - Handles API requests for Microsoft Groups API.
 * Follows MCP modular, testable, and consistent API contract rules.
 */

const Joi = require('joi');
const ErrorService = require('../../core/error-service.cjs');
const MonitoringService = require('../../core/monitoring-service.cjs');

/**
 * Joi validation schemas for groups endpoints
 */
const schemas = {
  listGroups: Joi.object({
    limit: Joi.number().integer().min(1).max(100).optional(),
    filter: Joi.string().optional(),
    search: Joi.string().optional()
  }),

  getGroup: Joi.object({
    groupId: Joi.string().required()
  }),

  listGroupMembers: Joi.object({
    groupId: Joi.string().required(),
    limit: Joi.number().integer().min(1).max(100).optional()
  }),

  listMyGroups: Joi.object({
    limit: Joi.number().integer().min(1).max(100).optional()
  })
};

/**
 * Creates a groups controller with injected dependencies.
 * @param {object} deps - Controller dependencies
 * @param {object} deps.groupsModule - Initialized groups module
 * @returns {object} Controller methods
 */
function createGroupsController({ groupsModule }) {
  if (!groupsModule) {
    throw new Error('Groups module is required for GroupsController');
  }

  return {
    /**
     * List all groups.
     */
    async listGroups(req, res) {
      const { userId = null } = req.user || {};
      const sessionId = req.session?.id;
      const startTime = Date.now();

      try {
        const { error, value } = schemas.listGroups.validate(req.query);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const options = {
          top: value.limit || 50,
          filter: value.filter,
          search: value.search
        };

        const groups = await groupsModule.listGroups(options, req, userId, sessionId);

        MonitoringService.trackMetric('groups.listGroups.duration', Date.now() - startTime, {
          count: groups.length,
          success: true,
          userId
        });

        res.json({ success: true, data: groups, count: groups.length });
      } catch (error) {
        MonitoringService.logError(ErrorService.createError('groups', 'Failed to list groups', 'error', {
          error: error.message,
          userId,
          timestamp: new Date().toISOString()
        }));

        res.status(500).json({
          error: 'GROUPS_OPERATION_FAILED',
          error_description: 'Failed to list groups'
        });
      }
    },

    /**
     * Get a specific group by ID.
     */
    async getGroup(req, res) {
      const { userId = null } = req.user || {};
      const sessionId = req.session?.id;
      const startTime = Date.now();

      try {
        const { error, value } = schemas.getGroup.validate(req.params);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const group = await groupsModule.getGroup(value.groupId, req, userId, sessionId);

        MonitoringService.trackMetric('groups.getGroup.duration', Date.now() - startTime, {
          success: true,
          userId
        });

        res.json({ success: true, data: group });
      } catch (error) {
        MonitoringService.logError(ErrorService.createError('groups', 'Failed to get group', 'error', {
          groupId: req.params.groupId,
          error: error.message,
          userId,
          timestamp: new Date().toISOString()
        }));

        res.status(500).json({
          error: 'GROUPS_OPERATION_FAILED',
          error_description: 'Failed to get group'
        });
      }
    },

    /**
     * List members of a group.
     */
    async listGroupMembers(req, res) {
      const { userId = null } = req.user || {};
      const sessionId = req.session?.id;
      const startTime = Date.now();

      try {
        const combined = { ...req.params, ...req.query };
        const { error, value } = schemas.listGroupMembers.validate(combined);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const options = { top: value.limit || 100 };
        const members = await groupsModule.listGroupMembers(value.groupId, options, req, userId, sessionId);

        MonitoringService.trackMetric('groups.listGroupMembers.duration', Date.now() - startTime, {
          count: members.length,
          success: true,
          userId
        });

        res.json({ success: true, data: members, count: members.length, groupId: value.groupId });
      } catch (error) {
        MonitoringService.logError(ErrorService.createError('groups', 'Failed to list group members', 'error', {
          groupId: req.params.groupId,
          error: error.message,
          userId,
          timestamp: new Date().toISOString()
        }));

        res.status(500).json({
          error: 'GROUPS_OPERATION_FAILED',
          error_description: 'Failed to list group members'
        });
      }
    },

    /**
     * List groups the current user is a member of.
     */
    async listMyGroups(req, res) {
      const { userId = null } = req.user || {};
      const sessionId = req.session?.id;
      const startTime = Date.now();

      try {
        const { error, value } = schemas.listMyGroups.validate(req.query);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const options = { top: value.limit || 100 };
        const groups = await groupsModule.listMyGroups(options, req, userId, sessionId);

        MonitoringService.trackMetric('groups.listMyGroups.duration', Date.now() - startTime, {
          count: groups.length,
          success: true,
          userId
        });

        res.json({ success: true, data: groups, count: groups.length });
      } catch (error) {
        MonitoringService.logError(ErrorService.createError('groups', 'Failed to list my groups', 'error', {
          error: error.message,
          userId,
          timestamp: new Date().toISOString()
        }));

        res.status(500).json({
          error: 'GROUPS_OPERATION_FAILED',
          error_description: 'Failed to list my groups'
        });
      }
    }
  };
}

module.exports = createGroupsController;
