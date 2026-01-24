/**
 * @fileoverview MCP Teams Module - Microsoft Teams integration.
 * Provides chat messaging, channel operations, and online meeting management.
 * Uses Microsoft Graph Teams API with support for chats, channels, and meetings.
 */

const ErrorService = require('../../core/error-service.cjs');
const MonitoringService = require('../../core/monitoring-service.cjs');

const TEAMS_CAPABILITIES = [
    // Chat operations
    'listChats',
    'createChat',
    'getChatMessages',
    'sendChatMessage',
    // Team & channel operations
    'listJoinedTeams',
    'listTeamChannels',
    'getChannelMessages',
    'sendChannelMessage',
    'replyToMessage',
    // Channel management operations
    'createTeamChannel',
    'addChannelMember',
    // Channel file operations
    'listChannelFiles',
    'uploadFileToChannel',
    'readChannelFile',
    // Meeting operations
    'createOnlineMeeting',
    'getOnlineMeeting',
    'getMeetingByJoinUrl',
    'listOnlineMeetings',
    // Transcript operations
    'getMeetingTranscripts',
    'getMeetingTranscriptContent'
];

// Log module initialization
MonitoringService.info('Teams Module initialized', {
    serviceName: 'teams-module',
    capabilities: TEAMS_CAPABILITIES.length,
    timestamp: new Date().toISOString()
}, 'teams');

const TeamsModule = {
    /**
     * Module ID
     */
    id: 'teams',

    /**
     * Module name
     */
    name: 'Microsoft Teams',

    /**
     * Module capabilities
     */
    capabilities: TEAMS_CAPABILITIES,

    /**
     * Service dependencies (injected during init)
     */
    services: null,

    /**
     * Helper method to redact sensitive data from objects before logging
     * @param {object} data - The data object to redact
     * @returns {object} Redacted copy of the data
     * @private
     */
    redactSensitiveData(data) {
        if (!data || typeof data !== 'object') {
            return data;
        }

        const result = Array.isArray(data) ? [...data] : { ...data };

        const sensitiveFields = [
            'body', 'content', 'email', 'emailAddress', 'address',
            'from', 'to', 'subject', 'message', 'joinUrl'
        ];

        for (const key in result) {
            if (Object.prototype.hasOwnProperty.call(result, key)) {
                if (sensitiveFields.includes(key.toLowerCase())) {
                    if (typeof result[key] === 'string') {
                        result[key] = result[key].substring(0, 20) + '...';
                    } else if (Array.isArray(result[key])) {
                        result[key] = `[${result[key].length} items]`;
                    } else if (typeof result[key] === 'object' && result[key] !== null) {
                        result[key] = '{...}';
                    }
                } else if (typeof result[key] === 'object' && result[key] !== null) {
                    result[key] = this.redactSensitiveData(result[key]);
                }
            }
        }

        return result;
    },

    /**
     * Initialize the teams module with dependencies
     * @param {object} services - Service dependencies (teamsService)
     * @param {string} userId - User ID for context
     * @param {string} sessionId - Session ID for context
     * @returns {TeamsModule} This module for chaining
     */
    init(services, userId, sessionId) {
        this.services = services;

        MonitoringService.info('Teams Module services initialized', {
            hasTeamsService: !!services?.teamsService,
            timestamp: new Date().toISOString()
        }, 'teams', null, userId);

        return this;
    },

    // ========================================================================
    // CHAT OPERATIONS
    // ========================================================================

    /**
     * List user's chats
     * @param {object} options - Query options { limit, filter }
     * @param {object} req - Express request
     * @param {string} userId - User ID
     * @param {string} sessionId - Session ID
     * @returns {Promise<Array<object>>} Array of chats
     */
    async listChats(options = {}, req, userId, sessionId) {
        const startTime = Date.now();
        const { teamsService } = this.services || {};

        if (process.env.NODE_ENV === 'development') {
            MonitoringService.debug('Teams module: listChats started', {
                limit: options.limit,
                timestamp: new Date().toISOString()
            }, 'teams');
        }

        if (!teamsService || typeof teamsService.getChats !== 'function') {
            const mcpError = ErrorService.createError(
                'teams',
                'TeamsService not available',
                'error',
                { method: 'listChats', moduleId: 'teams' }
            );
            MonitoringService.logError(mcpError);
            throw mcpError;
        }

        try {
            const chats = await teamsService.getChats(options, req, userId, sessionId);
            const executionTime = Date.now() - startTime;

            if (userId) {
                MonitoringService.info('Listed chats via module', {
                    chatCount: chats.length,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'teams', null, userId);
            }

            return chats;
        } catch (error) {
            MonitoringService.error('Teams module: listChats failed', {
                error: error.message,
                executionTimeMs: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'teams', null, userId);
            throw error;
        }
    },

    /**
     * Create a new chat (1:1 or group)
     * @param {object} chatData - Chat details { members: [{ email }], chatType, topic }
     * @param {object} req - Express request
     * @param {string} userId - User ID
     * @param {string} sessionId - Session ID
     * @returns {Promise<object>} Created chat
     */
    async createChat(chatData, req, userId, sessionId) {
        const startTime = Date.now();
        const { teamsService } = this.services || {};

        if (!teamsService || typeof teamsService.createChat !== 'function') {
            const mcpError = ErrorService.createError(
                'teams',
                'TeamsService not available',
                'error',
                { method: 'createChat', moduleId: 'teams' }
            );
            MonitoringService.logError(mcpError);
            throw mcpError;
        }

        try {
            const chat = await teamsService.createChat(chatData, req, userId, sessionId);
            const executionTime = Date.now() - startTime;

            if (userId) {
                MonitoringService.info('Created chat via module', {
                    chatId: chat.id,
                    chatType: chatData.chatType || 'oneOnOne',
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'teams', null, userId);
            }

            return chat;
        } catch (error) {
            MonitoringService.error('Teams module: createChat failed', {
                error: error.message,
                executionTimeMs: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'teams', null, userId);
            throw error;
        }
    },

    /**
     * Get messages from a chat
     * @param {string} chatId - Chat ID
     * @param {object} options - Query options { limit }
     * @param {object} req - Express request
     * @param {string} userId - User ID
     * @param {string} sessionId - Session ID
     * @returns {Promise<Array<object>>} Array of messages
     */
    async getChatMessages(chatId, options = {}, req, userId, sessionId) {
        const startTime = Date.now();
        const { teamsService } = this.services || {};

        if (!teamsService || typeof teamsService.getChatMessages !== 'function') {
            const mcpError = ErrorService.createError(
                'teams',
                'TeamsService not available',
                'error',
                { method: 'getChatMessages', moduleId: 'teams' }
            );
            MonitoringService.logError(mcpError);
            throw mcpError;
        }

        try {
            const messages = await teamsService.getChatMessages(chatId, options, req, userId, sessionId);
            const executionTime = Date.now() - startTime;

            if (userId) {
                MonitoringService.info('Retrieved chat messages via module', {
                    chatId: chatId.substring(0, 20) + '...',
                    messageCount: messages.length,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'teams', null, userId);
            }

            return messages;
        } catch (error) {
            MonitoringService.error('Teams module: getChatMessages failed', {
                error: error.message,
                executionTimeMs: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'teams', null, userId);
            throw error;
        }
    },

    /**
     * Send a message to a chat
     * @param {string} chatId - Chat ID
     * @param {object} messageData - Message content { content, contentType }
     * @param {object} req - Express request
     * @param {string} userId - User ID
     * @param {string} sessionId - Session ID
     * @returns {Promise<object>} Created message
     */
    async sendChatMessage(chatId, messageData, req, userId, sessionId) {
        const startTime = Date.now();
        const { teamsService } = this.services || {};

        if (!teamsService || typeof teamsService.sendChatMessage !== 'function') {
            const mcpError = ErrorService.createError(
                'teams',
                'TeamsService not available',
                'error',
                { method: 'sendChatMessage', moduleId: 'teams' }
            );
            MonitoringService.logError(mcpError);
            throw mcpError;
        }

        try {
            const message = await teamsService.sendChatMessage(chatId, messageData, req, userId, sessionId);
            const executionTime = Date.now() - startTime;

            if (userId) {
                MonitoringService.info('Sent chat message via module', {
                    chatId: chatId.substring(0, 20) + '...',
                    messageId: message.id,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'teams', null, userId);
            }

            return message;
        } catch (error) {
            MonitoringService.error('Teams module: sendChatMessage failed', {
                error: error.message,
                executionTimeMs: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'teams', null, userId);
            throw error;
        }
    },

    // ========================================================================
    // TEAM & CHANNEL OPERATIONS
    // ========================================================================

    /**
     * List user's joined teams
     * @param {object} options - Query options { limit }
     * @param {object} req - Express request
     * @param {string} userId - User ID
     * @param {string} sessionId - Session ID
     * @returns {Promise<Array<object>>} Array of teams
     */
    async listJoinedTeams(options = {}, req, userId, sessionId) {
        const startTime = Date.now();
        const { teamsService } = this.services || {};

        if (!teamsService || typeof teamsService.getJoinedTeams !== 'function') {
            const mcpError = ErrorService.createError(
                'teams',
                'TeamsService not available',
                'error',
                { method: 'listJoinedTeams', moduleId: 'teams' }
            );
            MonitoringService.logError(mcpError);
            throw mcpError;
        }

        try {
            const teams = await teamsService.getJoinedTeams(options, req, userId, sessionId);
            const executionTime = Date.now() - startTime;

            if (userId) {
                MonitoringService.info('Listed joined teams via module', {
                    teamCount: teams.length,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'teams', null, userId);
            }

            return teams;
        } catch (error) {
            MonitoringService.error('Teams module: listJoinedTeams failed', {
                error: error.message,
                executionTimeMs: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'teams', null, userId);
            throw error;
        }
    },

    /**
     * List channels in a team
     * @param {string} teamId - Team ID
     * @param {object} options - Query options
     * @param {object} req - Express request
     * @param {string} userId - User ID
     * @param {string} sessionId - Session ID
     * @returns {Promise<Array<object>>} Array of channels
     */
    async listTeamChannels(teamId, options = {}, req, userId, sessionId) {
        const startTime = Date.now();
        const { teamsService } = this.services || {};

        if (!teamsService || typeof teamsService.getTeamChannels !== 'function') {
            const mcpError = ErrorService.createError(
                'teams',
                'TeamsService not available',
                'error',
                { method: 'listTeamChannels', moduleId: 'teams' }
            );
            MonitoringService.logError(mcpError);
            throw mcpError;
        }

        try {
            const channels = await teamsService.getTeamChannels(teamId, options, req, userId, sessionId);
            const executionTime = Date.now() - startTime;

            if (userId) {
                MonitoringService.info('Listed team channels via module', {
                    teamId: teamId.substring(0, 20) + '...',
                    channelCount: channels.length,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'teams', null, userId);
            }

            return channels;
        } catch (error) {
            MonitoringService.error('Teams module: listTeamChannels failed', {
                error: error.message,
                executionTimeMs: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'teams', null, userId);
            throw error;
        }
    },

    /**
     * Get messages from a channel
     * @param {string} teamId - Team ID
     * @param {string} channelId - Channel ID
     * @param {object} options - Query options { limit }
     * @param {object} req - Express request
     * @param {string} userId - User ID
     * @param {string} sessionId - Session ID
     * @returns {Promise<Array<object>>} Array of messages
     */
    async getChannelMessages(teamId, channelId, options = {}, req, userId, sessionId) {
        const startTime = Date.now();
        const { teamsService } = this.services || {};

        if (!teamsService || typeof teamsService.getChannelMessages !== 'function') {
            const mcpError = ErrorService.createError(
                'teams',
                'TeamsService not available',
                'error',
                { method: 'getChannelMessages', moduleId: 'teams' }
            );
            MonitoringService.logError(mcpError);
            throw mcpError;
        }

        try {
            const messages = await teamsService.getChannelMessages(teamId, channelId, options, req, userId, sessionId);
            const executionTime = Date.now() - startTime;

            if (userId) {
                MonitoringService.info('Retrieved channel messages via module', {
                    teamId: teamId.substring(0, 20) + '...',
                    channelId: channelId.substring(0, 20) + '...',
                    messageCount: messages.length,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'teams', null, userId);
            }

            return messages;
        } catch (error) {
            MonitoringService.error('Teams module: getChannelMessages failed', {
                error: error.message,
                executionTimeMs: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'teams', null, userId);
            throw error;
        }
    },

    /**
     * Send a message to a channel
     * @param {string} teamId - Team ID
     * @param {string} channelId - Channel ID
     * @param {object} messageData - Message content { content, contentType, subject }
     * @param {object} req - Express request
     * @param {string} userId - User ID
     * @param {string} sessionId - Session ID
     * @returns {Promise<object>} Created message
     */
    async sendChannelMessage(teamId, channelId, messageData, req, userId, sessionId) {
        const startTime = Date.now();
        const { teamsService } = this.services || {};

        if (!teamsService || typeof teamsService.sendChannelMessage !== 'function') {
            const mcpError = ErrorService.createError(
                'teams',
                'TeamsService not available',
                'error',
                { method: 'sendChannelMessage', moduleId: 'teams' }
            );
            MonitoringService.logError(mcpError);
            throw mcpError;
        }

        try {
            const message = await teamsService.sendChannelMessage(teamId, channelId, messageData, req, userId, sessionId);
            const executionTime = Date.now() - startTime;

            if (userId) {
                MonitoringService.info('Sent channel message via module', {
                    teamId: teamId.substring(0, 20) + '...',
                    channelId: channelId.substring(0, 20) + '...',
                    messageId: message.id,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'teams', null, userId);
            }

            return message;
        } catch (error) {
            MonitoringService.error('Teams module: sendChannelMessage failed', {
                error: error.message,
                executionTimeMs: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'teams', null, userId);
            throw error;
        }
    },

    /**
     * Reply to a message in a channel
     * @param {string} teamId - Team ID
     * @param {string} channelId - Channel ID
     * @param {string} messageId - Parent message ID
     * @param {object} replyData - Reply content { content, contentType }
     * @param {object} req - Express request
     * @param {string} userId - User ID
     * @param {string} sessionId - Session ID
     * @returns {Promise<object>} Created reply
     */
    async replyToMessage(teamId, channelId, messageId, replyData, req, userId, sessionId) {
        const startTime = Date.now();
        const { teamsService } = this.services || {};

        if (!teamsService || typeof teamsService.replyToMessage !== 'function') {
            const mcpError = ErrorService.createError(
                'teams',
                'TeamsService not available',
                'error',
                { method: 'replyToMessage', moduleId: 'teams' }
            );
            MonitoringService.logError(mcpError);
            throw mcpError;
        }

        try {
            const reply = await teamsService.replyToMessage(teamId, channelId, messageId, replyData, req, userId, sessionId);
            const executionTime = Date.now() - startTime;

            if (userId) {
                MonitoringService.info('Replied to message via module', {
                    teamId: teamId.substring(0, 20) + '...',
                    channelId: channelId.substring(0, 20) + '...',
                    parentMessageId: messageId.substring(0, 20) + '...',
                    replyId: reply.id,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'teams', null, userId);
            }

            return reply;
        } catch (error) {
            MonitoringService.error('Teams module: replyToMessage failed', {
                error: error.message,
                executionTimeMs: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'teams', null, userId);
            throw error;
        }
    },

    // ========================================================================
    // CHANNEL MANAGEMENT OPERATIONS
    // ========================================================================

    /**
     * Create a new channel in a team
     * @param {string} teamId - Team ID
     * @param {object} channelData - Channel details { displayName, description, membershipType }
     * @param {object} req - Express request
     * @param {string} userId - User ID
     * @param {string} sessionId - Session ID
     * @returns {Promise<object>} Created channel
     */
    async createTeamChannel(teamId, channelData, req, userId, sessionId) {
        const startTime = Date.now();
        const { teamsService } = this.services || {};

        if (!teamsService || typeof teamsService.createChannel !== 'function') {
            const mcpError = ErrorService.createError(
                'teams',
                'TeamsService not available',
                'error',
                { method: 'createTeamChannel', moduleId: 'teams' }
            );
            MonitoringService.logError(mcpError);
            throw mcpError;
        }

        try {
            const channel = await teamsService.createChannel(teamId, channelData, req, userId, sessionId);
            const executionTime = Date.now() - startTime;

            if (userId) {
                MonitoringService.info('Created team channel via module', {
                    teamId: teamId.substring(0, 20) + '...',
                    channelId: channel.id,
                    displayName: channelData.displayName,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'teams', null, userId);
            }

            return channel;
        } catch (error) {
            MonitoringService.error('Teams module: createTeamChannel failed', {
                error: error.message,
                executionTimeMs: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'teams', null, userId);
            throw error;
        }
    },

    /**
     * Add a member to a channel (for private channels)
     * @param {string} teamId - Team ID
     * @param {string} channelId - Channel ID
     * @param {object} memberData - Member details { userEmail, roles }
     * @param {object} req - Express request
     * @param {string} userId - User ID
     * @param {string} sessionId - Session ID
     * @returns {Promise<object>} Added member
     */
    async addChannelMember(teamId, channelId, memberData, req, userId, sessionId) {
        const startTime = Date.now();
        const { teamsService } = this.services || {};

        if (!teamsService || typeof teamsService.addChannelMember !== 'function') {
            const mcpError = ErrorService.createError(
                'teams',
                'TeamsService not available',
                'error',
                { method: 'addChannelMember', moduleId: 'teams' }
            );
            MonitoringService.logError(mcpError);
            throw mcpError;
        }

        try {
            const member = await teamsService.addChannelMember(teamId, channelId, memberData, req, userId, sessionId);
            const executionTime = Date.now() - startTime;

            if (userId) {
                MonitoringService.info('Added channel member via module', {
                    teamId: teamId.substring(0, 20) + '...',
                    channelId: channelId.substring(0, 20) + '...',
                    memberId: member.id,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'teams', null, userId);
            }

            return member;
        } catch (error) {
            MonitoringService.error('Teams module: addChannelMember failed', {
                error: error.message,
                executionTimeMs: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'teams', null, userId);
            throw error;
        }
    },

    // ========================================================================
    // CHANNEL FILE OPERATIONS
    // ========================================================================

    /**
     * List files in a channel's Files tab
     * @param {string} teamId - Team ID
     * @param {string} channelId - Channel ID
     * @param {object} options - Query options { limit }
     * @param {object} req - Express request
     * @param {string} userId - User ID
     * @param {string} sessionId - Session ID
     * @returns {Promise<Array<object>>} Array of files
     */
    async listChannelFiles(teamId, channelId, options = {}, req, userId, sessionId) {
        const startTime = Date.now();
        const { teamsService } = this.services || {};

        if (!teamsService || typeof teamsService.listChannelFiles !== 'function') {
            const mcpError = ErrorService.createError(
                'teams',
                'TeamsService not available',
                'error',
                { method: 'listChannelFiles', moduleId: 'teams' }
            );
            MonitoringService.logError(mcpError);
            throw mcpError;
        }

        try {
            const files = await teamsService.listChannelFiles(teamId, channelId, options, req, userId, sessionId);
            const executionTime = Date.now() - startTime;

            if (userId) {
                MonitoringService.info('Listed channel files via module', {
                    teamId: teamId.substring(0, 20) + '...',
                    channelId: channelId.substring(0, 20) + '...',
                    fileCount: files.length,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'teams', null, userId);
            }

            return files;
        } catch (error) {
            MonitoringService.error('Teams module: listChannelFiles failed', {
                error: error.message,
                executionTimeMs: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'teams', null, userId);
            throw error;
        }
    },

    /**
     * Upload a file to a channel's Files tab
     * @param {string} teamId - Team ID
     * @param {string} channelId - Channel ID
     * @param {object} fileData - File details { fileName, content, contentType, isBase64 }
     * @param {object} req - Express request
     * @param {string} userId - User ID
     * @param {string} sessionId - Session ID
     * @returns {Promise<object>} Uploaded file
     */
    async uploadFileToChannel(teamId, channelId, fileData, req, userId, sessionId) {
        const startTime = Date.now();
        const { teamsService } = this.services || {};

        if (!teamsService || typeof teamsService.uploadFileToChannel !== 'function') {
            const mcpError = ErrorService.createError(
                'teams',
                'TeamsService not available',
                'error',
                { method: 'uploadFileToChannel', moduleId: 'teams' }
            );
            MonitoringService.logError(mcpError);
            throw mcpError;
        }

        try {
            const file = await teamsService.uploadFileToChannel(teamId, channelId, fileData, req, userId, sessionId);
            const executionTime = Date.now() - startTime;

            if (userId) {
                MonitoringService.info('Uploaded file to channel via module', {
                    teamId: teamId.substring(0, 20) + '...',
                    channelId: channelId.substring(0, 20) + '...',
                    fileId: file.id,
                    fileName: fileData.fileName,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'teams', null, userId);
            }

            return file;
        } catch (error) {
            MonitoringService.error('Teams module: uploadFileToChannel failed', {
                error: error.message,
                executionTimeMs: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'teams', null, userId);
            throw error;
        }
    },

    /**
     * Read content of a file from a channel's Files tab
     * @param {string} teamId - Team ID
     * @param {string} channelId - Channel ID
     * @param {string} fileName - File name to read
     * @param {object} req - Express request
     * @param {string} userId - User ID
     * @param {string} sessionId - Session ID
     * @returns {Promise<object>} File content
     */
    async readChannelFile(teamId, channelId, fileName, req, userId, sessionId) {
        const startTime = Date.now();
        const { teamsService } = this.services || {};

        if (!teamsService || typeof teamsService.readChannelFile !== 'function') {
            const mcpError = ErrorService.createError(
                'teams',
                'TeamsService not available',
                'error',
                { method: 'readChannelFile', moduleId: 'teams' }
            );
            MonitoringService.logError(mcpError);
            throw mcpError;
        }

        try {
            const content = await teamsService.readChannelFile(teamId, channelId, fileName, req, userId, sessionId);
            const executionTime = Date.now() - startTime;

            if (userId) {
                MonitoringService.info('Read channel file via module', {
                    teamId: teamId.substring(0, 20) + '...',
                    channelId: channelId.substring(0, 20) + '...',
                    fileName: fileName,
                    size: content.size,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'teams', null, userId);
            }

            return content;
        } catch (error) {
            MonitoringService.error('Teams module: readChannelFile failed', {
                error: error.message,
                executionTimeMs: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'teams', null, userId);
            throw error;
        }
    },

    // ========================================================================
    // MEETING OPERATIONS
    // ========================================================================

    /**
     * Create an online meeting
     * @param {object} meetingData - Meeting details { subject, startDateTime, endDateTime, participants, lobbyBypassSettings }
     * @param {object} req - Express request
     * @param {string} userId - User ID
     * @param {string} sessionId - Session ID
     * @returns {Promise<object>} Created meeting with join URL
     */
    async createOnlineMeeting(meetingData, req, userId, sessionId) {
        const startTime = Date.now();
        const { teamsService } = this.services || {};

        if (!teamsService || typeof teamsService.createOnlineMeeting !== 'function') {
            const mcpError = ErrorService.createError(
                'teams',
                'TeamsService not available',
                'error',
                { method: 'createOnlineMeeting', moduleId: 'teams' }
            );
            MonitoringService.logError(mcpError);
            throw mcpError;
        }

        try {
            const meeting = await teamsService.createOnlineMeeting(meetingData, req, userId, sessionId);
            const executionTime = Date.now() - startTime;

            if (userId) {
                MonitoringService.info('Created online meeting via module', {
                    meetingId: meeting.id,
                    subject: meetingData.subject?.substring(0, 30),
                    hasJoinUrl: !!meeting.joinUrl,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'teams', null, userId);
            }

            return meeting;
        } catch (error) {
            MonitoringService.error('Teams module: createOnlineMeeting failed', {
                error: error.message,
                executionTimeMs: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'teams', null, userId);
            throw error;
        }
    },

    /**
     * Get online meeting details
     * @param {string} meetingId - Meeting ID
     * @param {object} req - Express request
     * @param {string} userId - User ID
     * @param {string} sessionId - Session ID
     * @returns {Promise<object>} Meeting details
     */
    async getOnlineMeeting(meetingId, req, userId, sessionId) {
        const startTime = Date.now();
        const { teamsService } = this.services || {};

        if (!teamsService || typeof teamsService.getOnlineMeeting !== 'function') {
            const mcpError = ErrorService.createError(
                'teams',
                'TeamsService not available',
                'error',
                { method: 'getOnlineMeeting', moduleId: 'teams' }
            );
            MonitoringService.logError(mcpError);
            throw mcpError;
        }

        try {
            const meeting = await teamsService.getOnlineMeeting(meetingId, req, userId, sessionId);
            const executionTime = Date.now() - startTime;

            if (userId) {
                MonitoringService.info('Retrieved online meeting via module', {
                    meetingId: meeting.id,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'teams', null, userId);
            }

            return meeting;
        } catch (error) {
            MonitoringService.error('Teams module: getOnlineMeeting failed', {
                error: error.message,
                executionTimeMs: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'teams', null, userId);
            throw error;
        }
    },

    /**
     * Find meeting by join URL
     * @param {string} joinUrl - Meeting join URL
     * @param {object} req - Express request
     * @param {string} userId - User ID
     * @param {string} sessionId - Session ID
     * @returns {Promise<object>} Meeting details
     */
    async getMeetingByJoinUrl(joinUrl, req, userId, sessionId) {
        const startTime = Date.now();
        const { teamsService } = this.services || {};

        if (!teamsService || typeof teamsService.getMeetingByJoinUrl !== 'function') {
            const mcpError = ErrorService.createError(
                'teams',
                'TeamsService not available',
                'error',
                { method: 'getMeetingByJoinUrl', moduleId: 'teams' }
            );
            MonitoringService.logError(mcpError);
            throw mcpError;
        }

        try {
            const meeting = await teamsService.getMeetingByJoinUrl(joinUrl, req, userId, sessionId);
            const executionTime = Date.now() - startTime;

            if (userId) {
                MonitoringService.info('Found meeting by join URL via module', {
                    meetingId: meeting.id,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'teams', null, userId);
            }

            return meeting;
        } catch (error) {
            MonitoringService.error('Teams module: getMeetingByJoinUrl failed', {
                error: error.message,
                executionTimeMs: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'teams', null, userId);
            throw error;
        }
    },

    /**
     * List online meetings
     * @param {object} options - Query options { limit }
     * @param {object} req - Express request
     * @param {string} userId - User ID
     * @param {string} sessionId - Session ID
     * @returns {Promise<Array<object>>} Array of meetings
     */
    async listOnlineMeetings(options = {}, req, userId, sessionId) {
        const startTime = Date.now();
        const { teamsService } = this.services || {};

        if (!teamsService || typeof teamsService.listOnlineMeetings !== 'function') {
            const mcpError = ErrorService.createError(
                'teams',
                'TeamsService not available',
                'error',
                { method: 'listOnlineMeetings', moduleId: 'teams' }
            );
            MonitoringService.logError(mcpError);
            throw mcpError;
        }

        try {
            const meetings = await teamsService.listOnlineMeetings(options, req, userId, sessionId);
            const executionTime = Date.now() - startTime;

            if (userId) {
                MonitoringService.info('Listed online meetings via module', {
                    meetingCount: meetings.length,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'teams', null, userId);
            }

            return meetings;
        } catch (error) {
            MonitoringService.error('Teams module: listOnlineMeetings failed', {
                error: error.message,
                executionTimeMs: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'teams', null, userId);
            throw error;
        }
    },

    /**
     * Get transcripts for an online meeting
     * @param {string} meetingId - Meeting ID
     * @param {object} req - Express request
     * @param {string} userId - User ID
     * @param {string} sessionId - Session ID
     * @returns {Promise<Array<object>>} Array of transcripts
     */
    async getMeetingTranscripts(meetingId, req, userId, sessionId) {
        const startTime = Date.now();
        const { teamsService } = this.services || {};

        if (!teamsService || typeof teamsService.getMeetingTranscripts !== 'function') {
            const mcpError = ErrorService.createError(
                'teams',
                'TeamsService not available',
                'error',
                { method: 'getMeetingTranscripts', moduleId: 'teams' }
            );
            MonitoringService.logError(mcpError);
            throw mcpError;
        }

        try {
            const transcripts = await teamsService.getMeetingTranscripts(meetingId, req, userId, sessionId);
            const executionTime = Date.now() - startTime;

            if (userId) {
                MonitoringService.info('Retrieved meeting transcripts via module', {
                    meetingId: meetingId.substring(0, 20) + '...',
                    transcriptCount: transcripts.length,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'teams', null, userId);
            }

            return transcripts;
        } catch (error) {
            MonitoringService.error('Teams module: getMeetingTranscripts failed', {
                error: error.message,
                executionTimeMs: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'teams', null, userId);
            throw error;
        }
    },

    /**
     * Get transcript content for a specific transcript
     * @param {string} meetingId - Meeting ID
     * @param {string} transcriptId - Transcript ID
     * @param {object} req - Express request
     * @param {string} userId - User ID
     * @param {string} sessionId - Session ID
     * @returns {Promise<object>} Transcript content with parsed entries
     */
    async getMeetingTranscriptContent(meetingId, transcriptId, req, userId, sessionId) {
        const startTime = Date.now();
        const { teamsService } = this.services || {};

        if (!teamsService || typeof teamsService.getMeetingTranscriptContent !== 'function') {
            const mcpError = ErrorService.createError(
                'teams',
                'TeamsService not available',
                'error',
                { method: 'getMeetingTranscriptContent', moduleId: 'teams' }
            );
            MonitoringService.logError(mcpError);
            throw mcpError;
        }

        try {
            const content = await teamsService.getMeetingTranscriptContent(meetingId, transcriptId, req, userId, sessionId);
            const executionTime = Date.now() - startTime;

            if (userId) {
                MonitoringService.info('Retrieved meeting transcript content via module', {
                    meetingId: meetingId.substring(0, 20) + '...',
                    transcriptId: transcriptId.substring(0, 20) + '...',
                    entryCount: content.entryCount,
                    executionTimeMs: executionTime,
                    timestamp: new Date().toISOString()
                }, 'teams', null, userId);
            }

            return content;
        } catch (error) {
            MonitoringService.error('Teams module: getMeetingTranscriptContent failed', {
                error: error.message,
                executionTimeMs: Date.now() - startTime,
                timestamp: new Date().toISOString()
            }, 'teams', null, userId);
            throw error;
        }
    },

    // ========================================================================
    // INTENT HANDLER
    // ========================================================================

    /**
     * Handle Teams intents from MCP
     * @param {string} intent - The intent to handle
     * @param {object} params - Intent parameters
     * @param {object} context - Execution context { req, userId, sessionId }
     * @returns {Promise<object>} Intent result
     */
    async handleIntent(intent, params = {}, context = {}) {
        const { req, userId, sessionId } = context;

        MonitoringService.debug('Teams module handling intent', {
            intent,
            hasParams: Object.keys(params).length > 0,
            timestamp: new Date().toISOString()
        }, 'teams');

        switch (intent) {
            // Chat intents
            case 'listChats':
                return this.listChats(params, req, userId, sessionId);

            case 'createChat':
                return this.createChat(params, req, userId, sessionId);

            case 'getChatMessages':
                return this.getChatMessages(params.chatId, params, req, userId, sessionId);

            case 'sendChatMessage':
                return this.sendChatMessage(params.chatId, params, req, userId, sessionId);

            // Team & channel intents
            case 'listJoinedTeams':
                return this.listJoinedTeams(params, req, userId, sessionId);

            case 'listTeamChannels':
                return this.listTeamChannels(params.teamId, params, req, userId, sessionId);

            case 'getChannelMessages':
                return this.getChannelMessages(params.teamId, params.channelId, params, req, userId, sessionId);

            case 'sendChannelMessage':
                return this.sendChannelMessage(params.teamId, params.channelId, params, req, userId, sessionId);

            case 'replyToMessage':
                return this.replyToMessage(params.teamId, params.channelId, params.messageId, params, req, userId, sessionId);

            // Channel management intents
            case 'createTeamChannel':
                return this.createTeamChannel(params.teamId, params, req, userId, sessionId);

            case 'addChannelMember':
                return this.addChannelMember(params.teamId, params.channelId, params, req, userId, sessionId);

            // Channel file intents
            case 'listChannelFiles':
                return this.listChannelFiles(params.teamId, params.channelId, params, req, userId, sessionId);

            case 'uploadFileToChannel':
                return this.uploadFileToChannel(params.teamId, params.channelId, params, req, userId, sessionId);

            case 'readChannelFile':
                return this.readChannelFile(params.teamId, params.channelId, params.fileName, req, userId, sessionId);

            // Meeting intents
            case 'createOnlineMeeting':
                return this.createOnlineMeeting(params, req, userId, sessionId);

            case 'getOnlineMeeting':
                return this.getOnlineMeeting(params.meetingId, req, userId, sessionId);

            case 'getMeetingByJoinUrl':
                return this.getMeetingByJoinUrl(params.joinUrl, req, userId, sessionId);

            case 'listOnlineMeetings':
                return this.listOnlineMeetings(params, req, userId, sessionId);

            // Transcript intents
            case 'getMeetingTranscripts':
                return this.getMeetingTranscripts(params.meetingId, req, userId, sessionId);

            case 'getMeetingTranscriptContent':
                return this.getMeetingTranscriptContent(params.meetingId, params.transcriptId, req, userId, sessionId);

            default:
                const error = ErrorService.createError(
                    'teams',
                    `Unknown teams intent: ${intent}`,
                    'warning',
                    { intent, availableIntents: TEAMS_CAPABILITIES }
                );
                MonitoringService.logError(error);
                throw error;
        }
    }
};

module.exports = TeamsModule;
