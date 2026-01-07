/**
 * @fileoverview TeamsService - Microsoft Graph Teams API operations.
 * Provides chat, channel, and online meeting functionality.
 * All methods are async, modular, and use GraphClient for requests.
 * Follows project error handling, validation, and normalization rules.
 */

const graphClientFactory = require('./graph-client.cjs');
const ErrorService = require('../core/error-service.cjs');
const MonitoringService = require('../core/monitoring-service.cjs');

// Log service initialization
MonitoringService.info('Graph Teams Service initialized', {
    serviceName: 'graph-teams-service',
    capabilities: ['chats', 'channels', 'messages', 'onlineMeetings'],
    timestamp: new Date().toISOString()
}, 'graph');

// ============================================================================
// NORMALIZERS
// ============================================================================

/**
 * Normalizes a Graph chat object to MCP schema.
 * @param {object} graphChat - Raw chat object from Graph API
 * @returns {object} Normalized chat object
 */
function normalizeChat(graphChat) {
    return {
        id: graphChat.id,
        type: 'chat',
        chatType: graphChat.chatType, // oneOnOne, group, meeting
        topic: graphChat.topic || null,
        createdDateTime: graphChat.createdDateTime,
        lastUpdatedDateTime: graphChat.lastUpdatedDateTime,
        webUrl: graphChat.webUrl,
        members: (graphChat.members || []).map(m => ({
            id: m.id,
            displayName: m.displayName,
            email: m.email
        }))
    };
}

/**
 * Normalizes a Graph Teams message object to MCP schema.
 * @param {object} graphMessage - Raw message object from Graph API
 * @returns {object} Normalized message object
 */
function normalizeTeamsMessage(graphMessage) {
    return {
        id: graphMessage.id,
        type: 'teamsMessage',
        messageType: graphMessage.messageType,
        createdDateTime: graphMessage.createdDateTime,
        lastModifiedDateTime: graphMessage.lastModifiedDateTime,
        subject: graphMessage.subject || null,
        body: {
            contentType: graphMessage.body?.contentType,
            content: graphMessage.body?.content
        },
        from: graphMessage.from?.user ? {
            id: graphMessage.from.user.id,
            displayName: graphMessage.from.user.displayName,
            email: graphMessage.from.user.email
        } : null,
        importance: graphMessage.importance,
        webUrl: graphMessage.webUrl,
        attachments: (graphMessage.attachments || []).map(a => ({
            id: a.id,
            contentType: a.contentType,
            name: a.name,
            contentUrl: a.contentUrl
        })),
        mentions: (graphMessage.mentions || []).map(m => ({
            id: m.id,
            mentionText: m.mentionText,
            mentioned: m.mentioned?.user?.displayName
        })),
        reactions: (graphMessage.reactions || []).map(r => ({
            reactionType: r.reactionType,
            user: r.user?.user?.displayName
        }))
    };
}

/**
 * Normalizes a Graph team object to MCP schema.
 * @param {object} graphTeam - Raw team object from Graph API
 * @returns {object} Normalized team object
 */
function normalizeTeam(graphTeam) {
    return {
        id: graphTeam.id,
        type: 'team',
        displayName: graphTeam.displayName,
        description: graphTeam.description,
        visibility: graphTeam.visibility,
        webUrl: graphTeam.webUrl,
        createdDateTime: graphTeam.createdDateTime
    };
}

/**
 * Normalizes a Graph channel object to MCP schema.
 * @param {object} graphChannel - Raw channel object from Graph API
 * @returns {object} Normalized channel object
 */
function normalizeChannel(graphChannel) {
    return {
        id: graphChannel.id,
        type: 'channel',
        displayName: graphChannel.displayName,
        description: graphChannel.description,
        membershipType: graphChannel.membershipType, // standard, private, shared
        webUrl: graphChannel.webUrl,
        email: graphChannel.email
    };
}

/**
 * Normalizes a Graph online meeting object to MCP schema.
 * @param {object} graphMeeting - Raw online meeting object from Graph API
 * @returns {object} Normalized meeting object
 */
function normalizeOnlineMeeting(graphMeeting) {
    return {
        id: graphMeeting.id,
        type: 'onlineMeeting',
        subject: graphMeeting.subject,
        startDateTime: graphMeeting.startDateTime,
        endDateTime: graphMeeting.endDateTime,
        joinUrl: graphMeeting.joinUrl || graphMeeting.joinWebUrl,
        joinInformation: graphMeeting.joinInformation?.content,
        videoTeleconferenceId: graphMeeting.videoTeleconferenceId,
        participants: {
            organizer: graphMeeting.participants?.organizer?.upn,
            attendees: (graphMeeting.participants?.attendees || []).map(a => a.upn)
        },
        lobbyBypassSettings: graphMeeting.lobbyBypassSettings?.scope,
        chatInfo: graphMeeting.chatInfo ? {
            threadId: graphMeeting.chatInfo.threadId,
            messageId: graphMeeting.chatInfo.messageId
        } : null
    };
}

// ============================================================================
// CHAT OPERATIONS
// ============================================================================

/**
 * Retrieves user's chats.
 * @param {object} options - Query options (limit, filter)
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<Array<object>>} Array of normalized chat objects
 */
async function getChats(options = {}, req, userId, sessionId) {
    const startTime = Date.now();
    const contextUserId = userId || req?.user?.userId;
    const contextSessionId = sessionId || req?.session?.id;

    if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Teams getChats operation started', {
            method: 'getChats',
            optionKeys: Object.keys(options),
            sessionId: contextSessionId,
            timestamp: new Date().toISOString()
        }, 'teams');
    }

    try {
        const client = await graphClientFactory.createClient(req, contextUserId, contextSessionId);
        const top = options.limit || options.top || 20;

        let endpoint = `/me/chats?$top=${top}&$expand=members`;
        if (options.filter) {
            endpoint += `&$filter=${encodeURIComponent(options.filter)}`;
        }

        const res = await client.api(endpoint, contextUserId, contextSessionId).get();
        const chats = (res.value || []).map(normalizeChat);

        const executionTime = Date.now() - startTime;

        if (contextUserId) {
            MonitoringService.info('Retrieved chats successfully', {
                chatCount: chats.length,
                requestedTop: top,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'teams', null, contextUserId);
        }

        MonitoringService.trackMetric('teams_get_chats_time', executionTime, {
            chatCount: chats.length
        });

        return chats;
    } catch (error) {
        const executionTime = Date.now() - startTime;

        const mcpError = ErrorService.createError(
            'teams',
            `Failed to get chats: ${error.message}`,
            'error',
            {
                service: 'graph-teams-service',
                method: 'getChats',
                executionTimeMs: executionTime,
                error: error.message,
                statusCode: error.statusCode || error.code,
                timestamp: new Date().toISOString()
            }
        );

        MonitoringService.logError(mcpError);

        if (contextUserId) {
            MonitoringService.error('Failed to retrieve chats', {
                error: error.message,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'teams', null, contextUserId);
        }

        throw mcpError;
    }
}

/**
 * Gets messages from a specific chat.
 * @param {string} chatId - Chat ID
 * @param {object} options - Query options (limit)
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<Array<object>>} Array of normalized message objects
 */
async function getChatMessages(chatId, options = {}, req, userId, sessionId) {
    const startTime = Date.now();
    const contextUserId = userId || req?.user?.userId;
    const contextSessionId = sessionId || req?.session?.id;

    if (!chatId) {
        const error = ErrorService.createError(
            'teams',
            'Chat ID is required',
            'warning',
            { method: 'getChatMessages' }
        );
        MonitoringService.logError(error);
        throw error;
    }

    if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Teams getChatMessages operation started', {
            method: 'getChatMessages',
            chatId: chatId.substring(0, 20) + '...',
            sessionId: contextSessionId,
            timestamp: new Date().toISOString()
        }, 'teams');
    }

    try {
        const client = await graphClientFactory.createClient(req, contextUserId, contextSessionId);
        const top = options.limit || options.top || 50;

        const res = await client.api(`/chats/${chatId}/messages?$top=${top}`, contextUserId, contextSessionId).get();
        const messages = (res.value || []).map(normalizeTeamsMessage);

        const executionTime = Date.now() - startTime;

        if (contextUserId) {
            MonitoringService.info('Retrieved chat messages successfully', {
                chatId: chatId.substring(0, 20) + '...',
                messageCount: messages.length,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'teams', null, contextUserId);
        }

        return messages;
    } catch (error) {
        const executionTime = Date.now() - startTime;

        const mcpError = ErrorService.createError(
            'teams',
            `Failed to get chat messages: ${error.message}`,
            'error',
            {
                service: 'graph-teams-service',
                method: 'getChatMessages',
                chatId: chatId.substring(0, 20) + '...',
                executionTimeMs: executionTime,
                error: error.message,
                statusCode: error.statusCode || error.code,
                timestamp: new Date().toISOString()
            }
        );

        MonitoringService.logError(mcpError);
        throw mcpError;
    }
}

/**
 * Sends a message to a chat.
 * @param {string} chatId - Chat ID
 * @param {object} messageData - Message content { content, contentType }
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<object>} Created message object
 */
async function sendChatMessage(chatId, messageData, req, userId, sessionId) {
    const startTime = Date.now();
    const contextUserId = userId || req?.user?.userId;
    const contextSessionId = sessionId || req?.session?.id;

    if (!chatId) {
        const error = ErrorService.createError(
            'teams',
            'Chat ID is required',
            'warning',
            { method: 'sendChatMessage' }
        );
        MonitoringService.logError(error);
        throw error;
    }

    if (!messageData?.content) {
        const error = ErrorService.createError(
            'teams',
            'Message content is required',
            'warning',
            { method: 'sendChatMessage' }
        );
        MonitoringService.logError(error);
        throw error;
    }

    if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Teams sendChatMessage operation started', {
            method: 'sendChatMessage',
            chatId: chatId.substring(0, 20) + '...',
            contentLength: messageData.content.length,
            sessionId: contextSessionId,
            timestamp: new Date().toISOString()
        }, 'teams');
    }

    try {
        const client = await graphClientFactory.createClient(req, contextUserId, contextSessionId);

        const body = {
            body: {
                contentType: messageData.contentType || 'text',
                content: messageData.content
            }
        };

        const res = await client.api(`/chats/${chatId}/messages`, contextUserId, contextSessionId).post(body);
        const message = normalizeTeamsMessage(res);

        const executionTime = Date.now() - startTime;

        if (contextUserId) {
            MonitoringService.info('Sent chat message successfully', {
                chatId: chatId.substring(0, 20) + '...',
                messageId: message.id,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'teams', null, contextUserId);
        }

        return message;
    } catch (error) {
        const executionTime = Date.now() - startTime;

        const mcpError = ErrorService.createError(
            'teams',
            `Failed to send chat message: ${error.message}`,
            'error',
            {
                service: 'graph-teams-service',
                method: 'sendChatMessage',
                chatId: chatId.substring(0, 20) + '...',
                executionTimeMs: executionTime,
                error: error.message,
                statusCode: error.statusCode || error.code,
                timestamp: new Date().toISOString()
            }
        );

        MonitoringService.logError(mcpError);
        throw mcpError;
    }
}

// ============================================================================
// TEAM & CHANNEL OPERATIONS
// ============================================================================

/**
 * Retrieves user's joined teams.
 * @param {object} options - Query options (limit)
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<Array<object>>} Array of normalized team objects
 */
async function getJoinedTeams(options = {}, req, userId, sessionId) {
    const startTime = Date.now();
    const contextUserId = userId || req?.user?.userId;
    const contextSessionId = sessionId || req?.session?.id;

    if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Teams getJoinedTeams operation started', {
            method: 'getJoinedTeams',
            sessionId: contextSessionId,
            timestamp: new Date().toISOString()
        }, 'teams');
    }

    try {
        const client = await graphClientFactory.createClient(req, contextUserId, contextSessionId);
        // Note: /me/joinedTeams does not support $top query parameter
        const res = await client.api('/me/joinedTeams', contextUserId, contextSessionId).get();
        const teams = (res.value || []).map(normalizeTeam);

        const executionTime = Date.now() - startTime;

        if (contextUserId) {
            MonitoringService.info('Retrieved joined teams successfully', {
                teamCount: teams.length,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'teams', null, contextUserId);
        }

        return teams;
    } catch (error) {
        const executionTime = Date.now() - startTime;

        const mcpError = ErrorService.createError(
            'teams',
            `Failed to get joined teams: ${error.message}`,
            'error',
            {
                service: 'graph-teams-service',
                method: 'getJoinedTeams',
                executionTimeMs: executionTime,
                error: error.message,
                statusCode: error.statusCode || error.code,
                timestamp: new Date().toISOString()
            }
        );

        MonitoringService.logError(mcpError);
        throw mcpError;
    }
}

/**
 * Gets channels for a team.
 * @param {string} teamId - Team ID
 * @param {object} options - Query options
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<Array<object>>} Array of normalized channel objects
 */
async function getTeamChannels(teamId, options = {}, req, userId, sessionId) {
    const startTime = Date.now();
    const contextUserId = userId || req?.user?.userId;
    const contextSessionId = sessionId || req?.session?.id;

    if (!teamId) {
        const error = ErrorService.createError(
            'teams',
            'Team ID is required',
            'warning',
            { method: 'getTeamChannels' }
        );
        MonitoringService.logError(error);
        throw error;
    }

    if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Teams getTeamChannels operation started', {
            method: 'getTeamChannels',
            teamId: teamId.substring(0, 20) + '...',
            sessionId: contextSessionId,
            timestamp: new Date().toISOString()
        }, 'teams');
    }

    try {
        const client = await graphClientFactory.createClient(req, contextUserId, contextSessionId);

        const res = await client.api(`/teams/${teamId}/channels`, contextUserId, contextSessionId).get();
        const channels = (res.value || []).map(normalizeChannel);

        const executionTime = Date.now() - startTime;

        if (contextUserId) {
            MonitoringService.info('Retrieved team channels successfully', {
                teamId: teamId.substring(0, 20) + '...',
                channelCount: channels.length,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'teams', null, contextUserId);
        }

        return channels;
    } catch (error) {
        const executionTime = Date.now() - startTime;

        const mcpError = ErrorService.createError(
            'teams',
            `Failed to get team channels: ${error.message}`,
            'error',
            {
                service: 'graph-teams-service',
                method: 'getTeamChannels',
                teamId: teamId.substring(0, 20) + '...',
                executionTimeMs: executionTime,
                error: error.message,
                statusCode: error.statusCode || error.code,
                timestamp: new Date().toISOString()
            }
        );

        MonitoringService.logError(mcpError);
        throw mcpError;
    }
}

/**
 * Gets messages from a channel.
 * @param {string} teamId - Team ID
 * @param {string} channelId - Channel ID
 * @param {object} options - Query options (limit)
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<Array<object>>} Array of normalized message objects
 */
async function getChannelMessages(teamId, channelId, options = {}, req, userId, sessionId) {
    const startTime = Date.now();
    const contextUserId = userId || req?.user?.userId;
    const contextSessionId = sessionId || req?.session?.id;

    if (!teamId || !channelId) {
        const error = ErrorService.createError(
            'teams',
            'Team ID and Channel ID are required',
            'warning',
            { method: 'getChannelMessages' }
        );
        MonitoringService.logError(error);
        throw error;
    }

    if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Teams getChannelMessages operation started', {
            method: 'getChannelMessages',
            teamId: teamId.substring(0, 20) + '...',
            channelId: channelId.substring(0, 20) + '...',
            sessionId: contextSessionId,
            timestamp: new Date().toISOString()
        }, 'teams');
    }

    try {
        const client = await graphClientFactory.createClient(req, contextUserId, contextSessionId);
        const top = options.limit || options.top || 50;

        const res = await client.api(`/teams/${teamId}/channels/${channelId}/messages?$top=${top}`, contextUserId, contextSessionId).get();
        const messages = (res.value || []).map(normalizeTeamsMessage);

        const executionTime = Date.now() - startTime;

        if (contextUserId) {
            MonitoringService.info('Retrieved channel messages successfully', {
                teamId: teamId.substring(0, 20) + '...',
                channelId: channelId.substring(0, 20) + '...',
                messageCount: messages.length,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'teams', null, contextUserId);
        }

        return messages;
    } catch (error) {
        const executionTime = Date.now() - startTime;

        const mcpError = ErrorService.createError(
            'teams',
            `Failed to get channel messages: ${error.message}`,
            'error',
            {
                service: 'graph-teams-service',
                method: 'getChannelMessages',
                teamId: teamId.substring(0, 20) + '...',
                channelId: channelId.substring(0, 20) + '...',
                executionTimeMs: executionTime,
                error: error.message,
                statusCode: error.statusCode || error.code,
                timestamp: new Date().toISOString()
            }
        );

        MonitoringService.logError(mcpError);
        throw mcpError;
    }
}

/**
 * Sends a message to a channel.
 * @param {string} teamId - Team ID
 * @param {string} channelId - Channel ID
 * @param {object} messageData - Message content { content, contentType, subject }
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<object>} Created message object
 */
async function sendChannelMessage(teamId, channelId, messageData, req, userId, sessionId) {
    const startTime = Date.now();
    const contextUserId = userId || req?.user?.userId;
    const contextSessionId = sessionId || req?.session?.id;

    if (!teamId || !channelId) {
        const error = ErrorService.createError(
            'teams',
            'Team ID and Channel ID are required',
            'warning',
            { method: 'sendChannelMessage' }
        );
        MonitoringService.logError(error);
        throw error;
    }

    if (!messageData?.content) {
        const error = ErrorService.createError(
            'teams',
            'Message content is required',
            'warning',
            { method: 'sendChannelMessage' }
        );
        MonitoringService.logError(error);
        throw error;
    }

    if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Teams sendChannelMessage operation started', {
            method: 'sendChannelMessage',
            teamId: teamId.substring(0, 20) + '...',
            channelId: channelId.substring(0, 20) + '...',
            contentLength: messageData.content.length,
            sessionId: contextSessionId,
            timestamp: new Date().toISOString()
        }, 'teams');
    }

    try {
        const client = await graphClientFactory.createClient(req, contextUserId, contextSessionId);

        const body = {
            body: {
                contentType: messageData.contentType || 'text',
                content: messageData.content
            }
        };

        if (messageData.subject) {
            body.subject = messageData.subject;
        }

        const res = await client.api(`/teams/${teamId}/channels/${channelId}/messages`, contextUserId, contextSessionId).post(body);
        const message = normalizeTeamsMessage(res);

        const executionTime = Date.now() - startTime;

        if (contextUserId) {
            MonitoringService.info('Sent channel message successfully', {
                teamId: teamId.substring(0, 20) + '...',
                channelId: channelId.substring(0, 20) + '...',
                messageId: message.id,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'teams', null, contextUserId);
        }

        return message;
    } catch (error) {
        const executionTime = Date.now() - startTime;

        const mcpError = ErrorService.createError(
            'teams',
            `Failed to send channel message: ${error.message}`,
            'error',
            {
                service: 'graph-teams-service',
                method: 'sendChannelMessage',
                teamId: teamId.substring(0, 20) + '...',
                channelId: channelId.substring(0, 20) + '...',
                executionTimeMs: executionTime,
                error: error.message,
                statusCode: error.statusCode || error.code,
                timestamp: new Date().toISOString()
            }
        );

        MonitoringService.logError(mcpError);
        throw mcpError;
    }
}

/**
 * Replies to a message in a channel.
 * @param {string} teamId - Team ID
 * @param {string} channelId - Channel ID
 * @param {string} messageId - Parent message ID
 * @param {object} replyData - Reply content { content, contentType }
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<object>} Created reply object
 */
async function replyToMessage(teamId, channelId, messageId, replyData, req, userId, sessionId) {
    const startTime = Date.now();
    const contextUserId = userId || req?.user?.userId;
    const contextSessionId = sessionId || req?.session?.id;

    if (!teamId || !channelId || !messageId) {
        const error = ErrorService.createError(
            'teams',
            'Team ID, Channel ID, and Message ID are required',
            'warning',
            { method: 'replyToMessage' }
        );
        MonitoringService.logError(error);
        throw error;
    }

    if (!replyData?.content) {
        const error = ErrorService.createError(
            'teams',
            'Reply content is required',
            'warning',
            { method: 'replyToMessage' }
        );
        MonitoringService.logError(error);
        throw error;
    }

    if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Teams replyToMessage operation started', {
            method: 'replyToMessage',
            teamId: teamId.substring(0, 20) + '...',
            channelId: channelId.substring(0, 20) + '...',
            messageId: messageId.substring(0, 20) + '...',
            contentLength: replyData.content.length,
            sessionId: contextSessionId,
            timestamp: new Date().toISOString()
        }, 'teams');
    }

    try {
        const client = await graphClientFactory.createClient(req, contextUserId, contextSessionId);

        const body = {
            body: {
                contentType: replyData.contentType || 'text',
                content: replyData.content
            }
        };

        const res = await client.api(`/teams/${teamId}/channels/${channelId}/messages/${messageId}/replies`, contextUserId, contextSessionId).post(body);
        const reply = normalizeTeamsMessage(res);

        const executionTime = Date.now() - startTime;

        if (contextUserId) {
            MonitoringService.info('Replied to message successfully', {
                teamId: teamId.substring(0, 20) + '...',
                channelId: channelId.substring(0, 20) + '...',
                parentMessageId: messageId.substring(0, 20) + '...',
                replyId: reply.id,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'teams', null, contextUserId);
        }

        return reply;
    } catch (error) {
        const executionTime = Date.now() - startTime;

        const mcpError = ErrorService.createError(
            'teams',
            `Failed to reply to message: ${error.message}`,
            'error',
            {
                service: 'graph-teams-service',
                method: 'replyToMessage',
                teamId: teamId.substring(0, 20) + '...',
                channelId: channelId.substring(0, 20) + '...',
                messageId: messageId.substring(0, 20) + '...',
                executionTimeMs: executionTime,
                error: error.message,
                statusCode: error.statusCode || error.code,
                timestamp: new Date().toISOString()
            }
        );

        MonitoringService.logError(mcpError);
        throw mcpError;
    }
}

// ============================================================================
// ONLINE MEETING OPERATIONS
// ============================================================================

/**
 * Creates a new online meeting.
 * @param {object} meetingData - Meeting details { subject, startDateTime, endDateTime, participants, lobbyBypassSettings }
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<object>} Created meeting object with join URL
 */
async function createOnlineMeeting(meetingData, req, userId, sessionId) {
    const startTime = Date.now();
    const contextUserId = userId || req?.user?.userId;
    const contextSessionId = sessionId || req?.session?.id;

    if (!meetingData?.subject || !meetingData?.startDateTime || !meetingData?.endDateTime) {
        const error = ErrorService.createError(
            'teams',
            'Subject, start time, and end time are required for online meeting',
            'warning',
            { method: 'createOnlineMeeting' }
        );
        MonitoringService.logError(error);
        throw error;
    }

    if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Teams createOnlineMeeting operation started', {
            method: 'createOnlineMeeting',
            subject: meetingData.subject.substring(0, 50),
            startDateTime: meetingData.startDateTime,
            sessionId: contextSessionId,
            timestamp: new Date().toISOString()
        }, 'teams');
    }

    try {
        const client = await graphClientFactory.createClient(req, contextUserId, contextSessionId);

        const body = {
            subject: meetingData.subject,
            startDateTime: meetingData.startDateTime,
            endDateTime: meetingData.endDateTime
        };

        if (meetingData.participants && Array.isArray(meetingData.participants)) {
            body.participants = {
                attendees: meetingData.participants.map(email => ({
                    upn: email,
                    role: 'attendee'
                }))
            };
        }

        if (meetingData.lobbyBypassSettings) {
            body.lobbyBypassSettings = {
                scope: meetingData.lobbyBypassSettings
            };
        }

        const res = await client.api('/me/onlineMeetings', contextUserId, contextSessionId).post(body);
        const meeting = normalizeOnlineMeeting(res);

        const executionTime = Date.now() - startTime;

        if (contextUserId) {
            MonitoringService.info('Created online meeting successfully', {
                meetingId: meeting.id,
                subject: meetingData.subject.substring(0, 50),
                joinUrl: meeting.joinUrl ? 'present' : 'missing',
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'teams', null, contextUserId);
        }

        return meeting;
    } catch (error) {
        const executionTime = Date.now() - startTime;

        const mcpError = ErrorService.createError(
            'teams',
            `Failed to create online meeting: ${error.message}`,
            'error',
            {
                service: 'graph-teams-service',
                method: 'createOnlineMeeting',
                subject: meetingData.subject?.substring(0, 50),
                executionTimeMs: executionTime,
                error: error.message,
                statusCode: error.statusCode || error.code,
                timestamp: new Date().toISOString()
            }
        );

        MonitoringService.logError(mcpError);
        throw mcpError;
    }
}

/**
 * Gets details of an online meeting.
 * @param {string} meetingId - Meeting ID
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<object>} Meeting details
 */
async function getOnlineMeeting(meetingId, req, userId, sessionId) {
    const startTime = Date.now();
    const contextUserId = userId || req?.user?.userId;
    const contextSessionId = sessionId || req?.session?.id;

    if (!meetingId) {
        const error = ErrorService.createError(
            'teams',
            'Meeting ID is required',
            'warning',
            { method: 'getOnlineMeeting' }
        );
        MonitoringService.logError(error);
        throw error;
    }

    if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Teams getOnlineMeeting operation started', {
            method: 'getOnlineMeeting',
            meetingId: meetingId.substring(0, 20) + '...',
            sessionId: contextSessionId,
            timestamp: new Date().toISOString()
        }, 'teams');
    }

    try {
        const client = await graphClientFactory.createClient(req, contextUserId, contextSessionId);

        const res = await client.api(`/me/onlineMeetings/${meetingId}`, contextUserId, contextSessionId).get();
        const meeting = normalizeOnlineMeeting(res);

        const executionTime = Date.now() - startTime;

        if (contextUserId) {
            MonitoringService.info('Retrieved online meeting successfully', {
                meetingId: meeting.id,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'teams', null, contextUserId);
        }

        return meeting;
    } catch (error) {
        const executionTime = Date.now() - startTime;

        const mcpError = ErrorService.createError(
            'teams',
            `Failed to get online meeting: ${error.message}`,
            'error',
            {
                service: 'graph-teams-service',
                method: 'getOnlineMeeting',
                meetingId: meetingId.substring(0, 20) + '...',
                executionTimeMs: executionTime,
                error: error.message,
                statusCode: error.statusCode || error.code,
                timestamp: new Date().toISOString()
            }
        );

        MonitoringService.logError(mcpError);
        throw mcpError;
    }
}

/**
 * Finds an online meeting by join URL.
 * @param {string} joinUrl - Meeting join URL
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<object>} Meeting details
 */
async function getMeetingByJoinUrl(joinUrl, req, userId, sessionId) {
    const startTime = Date.now();
    const contextUserId = userId || req?.user?.userId;
    const contextSessionId = sessionId || req?.session?.id;

    if (!joinUrl) {
        const error = ErrorService.createError(
            'teams',
            'Join URL is required',
            'warning',
            { method: 'getMeetingByJoinUrl' }
        );
        MonitoringService.logError(error);
        throw error;
    }

    if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Teams getMeetingByJoinUrl operation started', {
            method: 'getMeetingByJoinUrl',
            joinUrl: joinUrl.substring(0, 50) + '...',
            sessionId: contextSessionId,
            timestamp: new Date().toISOString()
        }, 'teams');
    }

    try {
        const client = await graphClientFactory.createClient(req, contextUserId, contextSessionId);

        const filter = `JoinWebUrl eq '${joinUrl}'`;
        const res = await client.api(`/me/onlineMeetings?$filter=${encodeURIComponent(filter)}`, contextUserId, contextSessionId).get();

        if (!res.value || res.value.length === 0) {
            const error = ErrorService.createError(
                'teams',
                'Meeting not found for the given join URL',
                'warning',
                { method: 'getMeetingByJoinUrl', joinUrl: joinUrl.substring(0, 50) }
            );
            MonitoringService.logError(error);
            throw error;
        }

        const meeting = normalizeOnlineMeeting(res.value[0]);

        const executionTime = Date.now() - startTime;

        if (contextUserId) {
            MonitoringService.info('Found meeting by join URL successfully', {
                meetingId: meeting.id,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'teams', null, contextUserId);
        }

        return meeting;
    } catch (error) {
        const executionTime = Date.now() - startTime;

        // If it's already our error, rethrow
        if (error.module === 'teams') {
            throw error;
        }

        const mcpError = ErrorService.createError(
            'teams',
            `Failed to find meeting by join URL: ${error.message}`,
            'error',
            {
                service: 'graph-teams-service',
                method: 'getMeetingByJoinUrl',
                executionTimeMs: executionTime,
                error: error.message,
                statusCode: error.statusCode || error.code,
                timestamp: new Date().toISOString()
            }
        );

        MonitoringService.logError(mcpError);
        throw mcpError;
    }
}

/**
 * Lists user's online meetings.
 * @param {object} options - Query options (limit)
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<Array<object>>} Array of meeting objects
 */
async function listOnlineMeetings(options = {}, req, userId, sessionId) {
    const startTime = Date.now();
    const contextUserId = userId || req?.user?.userId;
    const contextSessionId = sessionId || req?.session?.id;

    if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Teams listOnlineMeetings operation started', {
            method: 'listOnlineMeetings',
            sessionId: contextSessionId,
            timestamp: new Date().toISOString()
        }, 'teams');
    }

    try {
        const client = await graphClientFactory.createClient(req, contextUserId, contextSessionId);
        // Note: /me/onlineMeetings does not support $top query parameter
        const res = await client.api('/me/onlineMeetings', contextUserId, contextSessionId).get();
        const meetings = (res.value || []).map(normalizeOnlineMeeting);

        const executionTime = Date.now() - startTime;

        if (contextUserId) {
            MonitoringService.info('Listed online meetings successfully', {
                meetingCount: meetings.length,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'teams', null, contextUserId);
        }

        return meetings;
    } catch (error) {
        const executionTime = Date.now() - startTime;

        const mcpError = ErrorService.createError(
            'teams',
            `Failed to list online meetings: ${error.message}`,
            'error',
            {
                service: 'graph-teams-service',
                method: 'listOnlineMeetings',
                executionTimeMs: executionTime,
                error: error.message,
                statusCode: error.statusCode || error.code,
                timestamp: new Date().toISOString()
            }
        );

        MonitoringService.logError(mcpError);
        throw mcpError;
    }
}

// ============================================================================
// TRANSCRIPT OPERATIONS
// ============================================================================

/**
 * Normalizes a transcript object from Graph API response.
 * @param {object} transcript - Raw transcript from Graph API
 * @returns {object} Normalized transcript object
 */
function normalizeTranscript(transcript) {
    return {
        type: 'meetingTranscript',
        id: transcript.id,
        meetingId: transcript.meetingId,
        createdDateTime: transcript.createdDateTime,
        meetingOrganizerId: transcript.meetingOrganizer?.user?.id,
        meetingOrganizerName: transcript.meetingOrganizer?.user?.displayName,
        contentUrl: transcript.transcriptContentUrl
    };
}

/**
 * Parses VTT content into structured transcript entries.
 * @param {string} vttContent - Raw VTT content
 * @returns {Array<object>} Parsed transcript entries
 */
function parseVttContent(vttContent) {
    const entries = [];
    const lines = vttContent.split('\n');
    let currentEntry = null;

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();

        // Skip WEBVTT header and empty lines
        if (line === 'WEBVTT' || line === '' || line.startsWith('NOTE')) {
            continue;
        }

        // Timestamp line (e.g., "00:00:05.000 --> 00:00:10.000")
        if (line.includes('-->')) {
            const [start, end] = line.split('-->').map(t => t.trim());
            currentEntry = { startTime: start, endTime: end, speaker: null, text: '' };
            continue;
        }

        // Speaker and text line (e.g., "<v Speaker Name>Text content")
        if (currentEntry) {
            const speakerMatch = line.match(/^<v\s+([^>]+)>(.*)$/);
            if (speakerMatch) {
                currentEntry.speaker = speakerMatch[1].trim();
                currentEntry.text = speakerMatch[2].trim();
            } else if (line && !line.match(/^\d+$/)) {
                // Continuation of text or plain text without speaker tag
                currentEntry.text += (currentEntry.text ? ' ' : '') + line;
            }

            // Check if next line is empty or a new timestamp (entry complete)
            const nextLine = lines[i + 1]?.trim() || '';
            if (nextLine === '' || nextLine.includes('-->') || nextLine.match(/^\d+$/)) {
                if (currentEntry.text) {
                    entries.push({ ...currentEntry });
                }
                currentEntry = null;
            }
        }
    }

    return entries;
}

/**
 * Gets all transcripts for an online meeting.
 * Requires OnlineMeetingTranscript.Read.All permission.
 * @param {string} meetingId - Online meeting ID
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<Array<object>>} List of transcripts
 */
async function getMeetingTranscripts(meetingId, req, userId, sessionId) {
    const startTime = Date.now();
    const contextUserId = userId || req?.user?.userId;
    const contextSessionId = sessionId || req?.session?.id;

    if (!meetingId) {
        const error = ErrorService.createError(
            'teams',
            'Meeting ID is required',
            'warning',
            { method: 'getMeetingTranscripts' }
        );
        MonitoringService.logError(error);
        throw error;
    }

    if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Teams getMeetingTranscripts operation started', {
            method: 'getMeetingTranscripts',
            meetingId: meetingId.substring(0, 20) + '...',
            sessionId: contextSessionId,
            timestamp: new Date().toISOString()
        }, 'teams');
    }

    try {
        const client = await graphClientFactory.createClient(req, contextUserId, contextSessionId);

        const res = await client.api(`/me/onlineMeetings/${meetingId}/transcripts`, contextUserId, contextSessionId).get();
        const transcripts = (res.value || []).map(normalizeTranscript);

        const executionTime = Date.now() - startTime;

        if (contextUserId) {
            MonitoringService.info('Retrieved meeting transcripts successfully', {
                meetingId: meetingId.substring(0, 20) + '...',
                transcriptCount: transcripts.length,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'teams', null, contextUserId);
        }

        return transcripts;
    } catch (error) {
        const executionTime = Date.now() - startTime;

        const mcpError = ErrorService.createError(
            'teams',
            `Failed to get meeting transcripts: ${error.message}`,
            'error',
            {
                service: 'graph-teams-service',
                method: 'getMeetingTranscripts',
                meetingId: meetingId.substring(0, 20) + '...',
                executionTimeMs: executionTime,
                error: error.message,
                statusCode: error.statusCode || error.code,
                timestamp: new Date().toISOString()
            }
        );

        MonitoringService.logError(mcpError);
        throw mcpError;
    }
}

/**
 * Gets the content of a specific meeting transcript.
 * Returns parsed VTT content with speaker attribution.
 * Requires OnlineMeetingTranscript.Read.All permission.
 * @param {string} meetingId - Online meeting ID
 * @param {string} transcriptId - Transcript ID
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<object>} Transcript content with parsed entries
 */
async function getMeetingTranscriptContent(meetingId, transcriptId, req, userId, sessionId) {
    const startTime = Date.now();
    const contextUserId = userId || req?.user?.userId;
    const contextSessionId = sessionId || req?.session?.id;

    if (!meetingId) {
        const error = ErrorService.createError(
            'teams',
            'Meeting ID is required',
            'warning',
            { method: 'getMeetingTranscriptContent' }
        );
        MonitoringService.logError(error);
        throw error;
    }

    if (!transcriptId) {
        const error = ErrorService.createError(
            'teams',
            'Transcript ID is required',
            'warning',
            { method: 'getMeetingTranscriptContent' }
        );
        MonitoringService.logError(error);
        throw error;
    }

    if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Teams getMeetingTranscriptContent operation started', {
            method: 'getMeetingTranscriptContent',
            meetingId: meetingId.substring(0, 20) + '...',
            transcriptId: transcriptId.substring(0, 20) + '...',
            sessionId: contextSessionId,
            timestamp: new Date().toISOString()
        }, 'teams');
    }

    try {
        const client = await graphClientFactory.createClient(req, contextUserId, contextSessionId);

        // Request VTT format for transcript content
        const vttContent = await client
            .api(`/me/onlineMeetings/${meetingId}/transcripts/${transcriptId}/content`, contextUserId, contextSessionId)
            .query({ '$format': 'text/vtt' })
            .get();

        // Parse VTT content into structured format
        const entries = parseVttContent(vttContent);

        const executionTime = Date.now() - startTime;

        if (contextUserId) {
            MonitoringService.info('Retrieved meeting transcript content successfully', {
                meetingId: meetingId.substring(0, 20) + '...',
                transcriptId: transcriptId.substring(0, 20) + '...',
                entryCount: entries.length,
                executionTimeMs: executionTime,
                timestamp: new Date().toISOString()
            }, 'teams', null, contextUserId);
        }

        return {
            type: 'meetingTranscriptContent',
            meetingId,
            transcriptId,
            entries,
            rawVtt: vttContent,
            entryCount: entries.length
        };
    } catch (error) {
        const executionTime = Date.now() - startTime;

        const mcpError = ErrorService.createError(
            'teams',
            `Failed to get meeting transcript content: ${error.message}`,
            'error',
            {
                service: 'graph-teams-service',
                method: 'getMeetingTranscriptContent',
                meetingId: meetingId.substring(0, 20) + '...',
                transcriptId: transcriptId.substring(0, 20) + '...',
                executionTimeMs: executionTime,
                error: error.message,
                statusCode: error.statusCode || error.code,
                timestamp: new Date().toISOString()
            }
        );

        MonitoringService.logError(mcpError);
        throw mcpError;
    }
}

// ============================================================================
// MODULE EXPORTS
// ============================================================================

module.exports = {
    // Chat operations
    getChats,
    getChatMessages,
    sendChatMessage,

    // Team & channel operations
    getJoinedTeams,
    getTeamChannels,
    getChannelMessages,
    sendChannelMessage,
    replyToMessage,

    // Online meeting operations
    createOnlineMeeting,
    getOnlineMeeting,
    getMeetingByJoinUrl,
    listOnlineMeetings,

    // Transcript operations
    getMeetingTranscripts,
    getMeetingTranscriptContent,

    // Normalizers (exported for use in other modules)
    normalizeChat,
    normalizeTeamsMessage,
    normalizeTeam,
    normalizeChannel,
    normalizeOnlineMeeting,
    normalizeTranscript
};
