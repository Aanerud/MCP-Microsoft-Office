/**
 * @fileoverview MCP Groups Module - Handles Microsoft Groups intents and actions for MCP.
 * Exposes: id, name, capabilities, init, handleIntent. Aligned with MCP module system.
 */

const MonitoringService = require('../../core/monitoring-service.cjs');
const ErrorService = require('../../core/error-service.cjs');

const GROUPS_CAPABILITIES = [
  'listGroups',
  'getGroup',
  'listGroupMembers',
  'listMyGroups'
];

// Log module initialization
MonitoringService.info('Groups Module initialized', {
  serviceName: 'groups-module',
  capabilities: GROUPS_CAPABILITIES.length,
  timestamp: new Date().toISOString()
}, 'groups');

const GroupsModule = {
  /**
   * Initializes the Groups module with dependencies.
   * @param {object} services - { groupsService, cacheService, errorService, monitoringService }
   * @returns {object} Initialized module
   */
  init(services = {}) {
    const { groupsService, cacheService, errorService = ErrorService, monitoringService = MonitoringService } = services;

    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Initializing Groups Module', {
        hasGroupsService: !!groupsService,
        hasCacheService: !!cacheService,
        timestamp: new Date().toISOString()
      }, 'groups');
    }

    if (!groupsService) {
      const mcpError = ErrorService.createError(
        'groups',
        "GroupsModule init failed: Required service 'groupsService' is missing",
        'error',
        { missingService: 'groupsService', timestamp: new Date().toISOString() }
      );
      MonitoringService.logError(mcpError);
      throw mcpError;
    }

    this.services = { groupsService, cacheService, errorService, monitoringService };
    return this;
  },

  /**
   * List all groups
   */
  async listGroups(options = {}, req, userId, sessionId) {
    const { groupsService } = this.services || {};
    const contextUserId = userId || req?.user?.userId;
    const contextSessionId = sessionId || req?.session?.id;

    if (!groupsService) {
      throw ErrorService.createError('groups', 'GroupsService not available', 'error', {});
    }

    return await groupsService.listGroups(options, req, contextUserId, contextSessionId);
  },

  /**
   * Get a specific group
   */
  async getGroup(groupId, req, userId, sessionId) {
    const { groupsService } = this.services || {};
    const contextUserId = userId || req?.user?.userId;
    const contextSessionId = sessionId || req?.session?.id;

    if (!groupsService) {
      throw ErrorService.createError('groups', 'GroupsService not available', 'error', {});
    }

    return await groupsService.getGroup(groupId, req, contextUserId, contextSessionId);
  },

  /**
   * List members of a group
   */
  async listGroupMembers(groupId, options = {}, req, userId, sessionId) {
    const { groupsService } = this.services || {};
    const contextUserId = userId || req?.user?.userId;
    const contextSessionId = sessionId || req?.session?.id;

    if (!groupsService) {
      throw ErrorService.createError('groups', 'GroupsService not available', 'error', {});
    }

    return await groupsService.listGroupMembers(groupId, options, req, contextUserId, contextSessionId);
  },

  /**
   * List groups the current user is a member of
   */
  async listMyGroups(options = {}, req, userId, sessionId) {
    const { groupsService } = this.services || {};
    const contextUserId = userId || req?.user?.userId;
    const contextSessionId = sessionId || req?.session?.id;

    if (!groupsService) {
      throw ErrorService.createError('groups', 'GroupsService not available', 'error', {});
    }

    return await groupsService.listMyGroups(options, req, contextUserId, contextSessionId);
  },

  /**
   * Handles Groups related intents routed to this module.
   */
  async handleIntent(intent, entities = {}, context = {}) {
    const startTime = Date.now();
    const contextUserId = context?.req?.user?.userId;
    const contextSessionId = context?.req?.session?.id;

    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Handling Groups intent', {
        sessionId: contextSessionId,
        intent,
        entities: Object.keys(entities),
        timestamp: new Date().toISOString()
      }, 'groups');
    }

    try {
      let result;

      switch (intent) {
        case 'listGroups': {
          const groups = await this.listGroups(entities.options || {}, context.req, contextUserId, contextSessionId);
          result = { type: 'groupCollection', items: groups };
          break;
        }

        case 'getGroup': {
          if (!entities.groupId) {
            throw ErrorService.createError('groups', 'groupId is required', 'warn', { intent });
          }
          const group = await this.getGroup(entities.groupId, context.req, contextUserId, contextSessionId);
          result = { type: 'group', data: group };
          break;
        }

        case 'listGroupMembers': {
          if (!entities.groupId) {
            throw ErrorService.createError('groups', 'groupId is required', 'warn', { intent });
          }
          const members = await this.listGroupMembers(entities.groupId, entities.options || {}, context.req, contextUserId, contextSessionId);
          result = { type: 'memberCollection', items: members, groupId: entities.groupId };
          break;
        }

        case 'listMyGroups': {
          const groups = await this.listMyGroups(entities.options || {}, context.req, contextUserId, contextSessionId);
          result = { type: 'groupCollection', items: groups };
          break;
        }

        default:
          throw ErrorService.createError('groups', `GroupsModule cannot handle intent: ${intent}`, 'warn', { intent });
      }

      const elapsedTime = Date.now() - startTime;

      if (contextUserId) {
        MonitoringService.info('Successfully handled Groups intent', {
          intent,
          elapsedTime,
          timestamp: new Date().toISOString()
        }, 'groups', null, contextUserId);
      }

      return result;
    } catch (error) {
      const elapsedTime = Date.now() - startTime;

      if (error.category && error.severity) {
        throw error;
      }

      const mcpError = ErrorService.createError(
        'groups',
        `Error handling Groups intent '${intent}': ${error.message}`,
        'error',
        { intent, originalError: error.stack, elapsedTime, timestamp: new Date().toISOString() }
      );
      MonitoringService.logError(mcpError);
      throw mcpError;
    }
  },

  id: 'groups',
  name: 'Microsoft Groups',
  capabilities: GROUPS_CAPABILITIES
};

module.exports = GroupsModule;
