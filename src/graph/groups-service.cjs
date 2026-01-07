/**
 * @fileoverview GroupsService - Microsoft Graph Groups API operations.
 * All methods are async, modular, and use GraphClient for requests.
 * Follows project error handling, validation, and normalization rules.
 */

const graphClientFactory = require('./graph-client.cjs');
const MonitoringService = require('../core/monitoring-service.cjs');
const ErrorService = require('../core/error-service.cjs');

/**
 * Normalizes a group object from Graph API response.
 * @param {object} group - Raw group object from Graph API
 * @returns {object} Normalized group object
 */
function normalizeGroup(group) {
  return {
    id: group.id,
    displayName: group.displayName,
    description: group.description,
    mail: group.mail,
    mailEnabled: group.mailEnabled,
    mailNickname: group.mailNickname,
    securityEnabled: group.securityEnabled,
    groupTypes: group.groupTypes || [],
    visibility: group.visibility,
    createdDateTime: group.createdDateTime,
    renewedDateTime: group.renewedDateTime,
    resourceProvisioningOptions: group.resourceProvisioningOptions || [],
    onPremisesSyncEnabled: group.onPremisesSyncEnabled,
    membershipRule: group.membershipRule
  };
}

/**
 * Normalizes a member object from Graph API response.
 * @param {object} member - Raw member object from Graph API
 * @returns {object} Normalized member object
 */
function normalizeMember(member) {
  return {
    id: member.id,
    displayName: member.displayName,
    mail: member.mail,
    userPrincipalName: member.userPrincipalName,
    jobTitle: member.jobTitle,
    department: member.department,
    officeLocation: member.officeLocation,
    '@odata.type': member['@odata.type']
  };
}

/**
 * Gets all groups (requires appropriate permissions).
 * @param {object} options - Query options
 * @param {number} [options.top=50] - Number of groups to retrieve
 * @param {string} [options.filter] - OData filter
 * @param {string} [options.search] - Search query
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<Array<object>>} Normalized group objects
 */
async function listGroups(options = {}, req, userId, sessionId) {
  const startTime = new Date();

  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;

  try {
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Listing groups', {
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString(),
        options: { top: options.top || 50 }
      }, 'groups');
    }

    const client = await graphClientFactory.createClient(req, resolvedUserId, resolvedSessionId);
    const top = options.top || 50;

    let queryParams = [`$top=${top}`];
    queryParams.push('$select=id,displayName,description,mail,mailEnabled,mailNickname,securityEnabled,groupTypes,visibility,createdDateTime');

    if (options.filter) {
      queryParams.push(`$filter=${encodeURIComponent(options.filter)}`);
    }
    if (options.search) {
      queryParams.push(`$search="${encodeURIComponent(options.search)}"`);
    }

    const url = `/groups?${queryParams.join('&')}`;
    const res = await client.api(url, resolvedUserId, resolvedSessionId).get();

    const endTime = new Date();
    const duration = endTime - startTime;

    const groups = (res.value || []).map(normalizeGroup);

    if (resolvedUserId) {
      MonitoringService.info('Retrieved groups successfully', {
        count: groups.length,
        duration: `${duration}ms`,
        timestamp: new Date().toISOString()
      }, 'groups', null, resolvedUserId);
    }

    return groups;
  } catch (error) {
    const endTime = new Date();
    const duration = endTime - startTime;

    const mcpError = ErrorService.createError(
      'groups',
      'Failed to list groups',
      'error',
      {
        endpoint: '/groups',
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
 * Gets a specific group by ID.
 * @param {string} groupId - Group ID
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<object>} Normalized group object
 */
async function getGroup(groupId, req, userId, sessionId) {
  const startTime = new Date();

  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;

  try {
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Getting group', {
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString(),
        groupId
      }, 'groups');
    }

    const client = await graphClientFactory.createClient(req, resolvedUserId, resolvedSessionId);

    const url = `/groups/${groupId}`;
    const res = await client.api(url, resolvedUserId, resolvedSessionId).get();

    const endTime = new Date();
    const duration = endTime - startTime;

    if (resolvedUserId) {
      MonitoringService.info('Retrieved group successfully', {
        groupId,
        duration: `${duration}ms`,
        timestamp: new Date().toISOString()
      }, 'groups', null, resolvedUserId);
    }

    return normalizeGroup(res);
  } catch (error) {
    const endTime = new Date();
    const duration = endTime - startTime;

    const mcpError = ErrorService.createError(
      'groups',
      'Failed to get group',
      'error',
      {
        endpoint: `/groups/${groupId}`,
        groupId,
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
 * Gets members of a specific group.
 * @param {string} groupId - Group ID
 * @param {object} options - Query options
 * @param {number} [options.top=100] - Number of members to retrieve
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<Array<object>>} Normalized member objects
 */
async function listGroupMembers(groupId, options = {}, req, userId, sessionId) {
  const startTime = new Date();

  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;

  try {
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Listing group members', {
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString(),
        groupId,
        options: { top: options.top || 100 }
      }, 'groups');
    }

    const client = await graphClientFactory.createClient(req, resolvedUserId, resolvedSessionId);
    const top = options.top || 100;

    const url = `/groups/${groupId}/members?$top=${top}&$select=id,displayName,mail,userPrincipalName,jobTitle,department,officeLocation`;
    const res = await client.api(url, resolvedUserId, resolvedSessionId).get();

    const endTime = new Date();
    const duration = endTime - startTime;

    const members = (res.value || []).map(normalizeMember);

    if (resolvedUserId) {
      MonitoringService.info('Retrieved group members successfully', {
        groupId,
        count: members.length,
        duration: `${duration}ms`,
        timestamp: new Date().toISOString()
      }, 'groups', null, resolvedUserId);
    }

    return members;
  } catch (error) {
    const endTime = new Date();
    const duration = endTime - startTime;

    const mcpError = ErrorService.createError(
      'groups',
      'Failed to list group members',
      'error',
      {
        endpoint: `/groups/${groupId}/members`,
        groupId,
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
 * Gets groups the current user is a member of.
 * @param {object} options - Query options
 * @param {number} [options.top=100] - Number of groups to retrieve
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<Array<object>>} Normalized group objects
 */
async function listMyGroups(options = {}, req, userId, sessionId) {
  const startTime = new Date();

  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;

  try {
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Listing my groups', {
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString(),
        options: { top: options.top || 100 }
      }, 'groups');
    }

    const client = await graphClientFactory.createClient(req, resolvedUserId, resolvedSessionId);
    const top = options.top || 100;

    // Get group memberships - filter to only groups (not directory roles, etc.)
    const url = `/me/memberOf/microsoft.graph.group?$top=${top}&$select=id,displayName,description,mail,mailEnabled,mailNickname,securityEnabled,groupTypes,visibility,createdDateTime`;
    const res = await client.api(url, resolvedUserId, resolvedSessionId).get();

    const endTime = new Date();
    const duration = endTime - startTime;

    const groups = (res.value || []).map(normalizeGroup);

    if (resolvedUserId) {
      MonitoringService.info('Retrieved my groups successfully', {
        count: groups.length,
        duration: `${duration}ms`,
        timestamp: new Date().toISOString()
      }, 'groups', null, resolvedUserId);
    }

    return groups;
  } catch (error) {
    const endTime = new Date();
    const duration = endTime - startTime;

    const mcpError = ErrorService.createError(
      'groups',
      'Failed to list my groups',
      'error',
      {
        endpoint: '/me/memberOf',
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

module.exports = {
  listGroups,
  getGroup,
  listGroupMembers,
  listMyGroups
};
