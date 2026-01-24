/**
 * @fileoverview Teams Controller - Handles Microsoft Teams API requests.
 * Provides chat, channel, and online meeting operations.
 * Follows MCP modular, testable, and consistent API contract rules.
 */

const Joi = require('joi');
const ErrorService = require('../../core/error-service.cjs');
const MonitoringService = require('../../core/monitoring-service.cjs');
const { validateAndLog } = require('../middleware/validation-utils.cjs');

/**
 * Joi validation schemas for teams endpoints
 */
const schemas = {
    // Chat schemas
    listChats: Joi.object({
        limit: Joi.number().integer().min(1).max(50).optional().default(20),
        filter: Joi.string().max(500).optional()
    }),

    getChatMessages: Joi.object({
        limit: Joi.number().integer().min(1).max(100).optional().default(50)
    }),

    sendChatMessage: Joi.object({
        content: Joi.string().min(1).max(10000).required(),
        contentType: Joi.string().valid('text', 'html').optional().default('text')
    }),

    createChat: Joi.object({
        members: Joi.array().items(Joi.object({
            email: Joi.string().email().required(),
            roles: Joi.array().items(Joi.string().valid('owner')).optional()
        })).min(1).required(),
        chatType: Joi.string().valid('oneOnOne', 'group').default('oneOnOne'),
        topic: Joi.string().max(250).optional()
    }),

    // Team & channel schemas
    listTeams: Joi.object({
        limit: Joi.number().integer().min(1).max(100).optional().default(100)
    }),

    getChannelMessages: Joi.object({
        limit: Joi.number().integer().min(1).max(100).optional().default(50)
    }),

    sendChannelMessage: Joi.object({
        content: Joi.string().min(1).max(10000).required(),
        contentType: Joi.string().valid('text', 'html').optional().default('text'),
        subject: Joi.string().max(200).optional()
    }),

    replyToMessage: Joi.object({
        content: Joi.string().min(1).max(10000).required(),
        contentType: Joi.string().valid('text', 'html').optional().default('text')
    }),

    // Channel management schemas
    createTeamChannel: Joi.object({
        displayName: Joi.string().min(1).max(50).required(),
        description: Joi.string().max(1024).optional(),
        membershipType: Joi.string().valid('standard', 'private', 'shared').default('standard')
    }),

    addChannelMember: Joi.object({
        userEmail: Joi.string().email().required(),
        roles: Joi.array().items(Joi.string().valid('owner')).default([])
    }),

    // Channel files schemas
    listChannelFiles: Joi.object({
        limit: Joi.number().integer().min(1).max(100).default(50)
    }),

    uploadFileToChannel: Joi.object({
        fileName: Joi.string().min(1).max(256).required(),
        content: Joi.string().required(),
        contentType: Joi.string().max(100).optional(),
        isBase64: Joi.boolean().default(false)
    }),

    readChannelFile: Joi.object({
        // fileName comes from URL params, not body
    }),

    // Meeting schemas
    createOnlineMeeting: Joi.object({
        subject: Joi.string().min(1).max(200).required(),
        startDateTime: Joi.string().isoDate().required(),
        endDateTime: Joi.string().isoDate().required(),
        participants: Joi.array().items(Joi.string().email()).optional(),
        lobbyBypassSettings: Joi.string()
            .valid('everyone', 'organization', 'organizationAndFederated', 'organizer')
            .optional()
            .default('organization')
    }),

    listOnlineMeetings: Joi.object({
        limit: Joi.number().integer().min(1).max(50).optional().default(20)
    }),

    getMeetingByJoinUrl: Joi.object({
        joinUrl: Joi.string().uri().required()
    })
};

/**
 * Creates a teams controller with injected dependencies.
 * @param {object} deps - Controller dependencies
 * @param {object} deps.teamsModule - Initialized teams module
 * @returns {object} Controller methods
 */
function createTeamsController({ teamsModule }) {
    if (!teamsModule) {
        throw new Error('Teams module is required for TeamsController');
    }

    return {
        // ====================================================================
        // CHAT ENDPOINTS
        // ====================================================================

        /**
         * List user's chats
         * GET /api/v1/teams/chats
         */
        async listChats(req, res) {
            const { userId = null, deviceId = null } = req.user || {};
            const sessionId = req.session?.id;
            const startTime = Date.now();

            try {
                if (process.env.NODE_ENV === 'development') {
                    MonitoringService.debug('Processing listChats request', {
                        path: req.path,
                        sessionId,
                        userId,
                        timestamp: new Date().toISOString()
                    }, 'teams');
                }

                const { error: validationError, value: validatedData } = validateAndLog(
                    req,
                    schemas.listChats,
                    'listChats',
                    { userId, deviceId }
                );

                if (validationError) {
                    return res.status(400).json({
                        error: 'INVALID_REQUEST',
                        error_description: validationError.details[0].message
                    });
                }

                const chats = await teamsModule.listChats(validatedData, req, userId, sessionId);

                if (userId) {
                    MonitoringService.info('Listed chats successfully', {
                        chatCount: chats.length,
                        duration: Date.now() - startTime,
                        timestamp: new Date().toISOString()
                    }, 'teams', null, userId);
                }

                res.json({ chats, count: chats.length });
            } catch (error) {
                handleControllerError(res, error, 'listChats', userId, sessionId, startTime);
            }
        },

        /**
         * Create a new chat (1:1 or group)
         * POST /api/v1/teams/chats
         */
        async createChat(req, res) {
            const { userId = null, deviceId = null } = req.user || {};
            const sessionId = req.session?.id;
            const startTime = Date.now();

            try {
                const { error: validationError, value: validatedData } = validateAndLog(
                    req,
                    schemas.createChat,
                    'createChat',
                    { userId, deviceId }
                );

                if (validationError) {
                    return res.status(400).json({
                        error: 'INVALID_REQUEST',
                        error_description: validationError.details[0].message
                    });
                }

                const chat = await teamsModule.createChat(validatedData, req, userId, sessionId);

                if (userId) {
                    MonitoringService.info('Created chat successfully', {
                        chatId: chat.id,
                        chatType: validatedData.chatType,
                        memberCount: validatedData.members.length,
                        duration: Date.now() - startTime,
                        timestamp: new Date().toISOString()
                    }, 'teams', null, userId);
                }

                res.status(201).json({ chat, success: true });
            } catch (error) {
                handleControllerError(res, error, 'createChat', userId, sessionId, startTime);
            }
        },

        /**
         * Get messages from a chat
         * GET /api/v1/teams/chats/:chatId/messages
         */
        async getChatMessages(req, res) {
            const { userId = null, deviceId = null } = req.user || {};
            const sessionId = req.session?.id;
            const startTime = Date.now();
            const { chatId } = req.params;

            try {
                if (!chatId) {
                    return res.status(400).json({
                        error: 'INVALID_REQUEST',
                        error_description: 'Chat ID is required'
                    });
                }

                const { error: validationError, value: validatedData } = validateAndLog(
                    req,
                    schemas.getChatMessages,
                    'getChatMessages',
                    { userId, deviceId }
                );

                if (validationError) {
                    return res.status(400).json({
                        error: 'INVALID_REQUEST',
                        error_description: validationError.details[0].message
                    });
                }

                const messages = await teamsModule.getChatMessages(chatId, validatedData, req, userId, sessionId);

                if (userId) {
                    MonitoringService.info('Retrieved chat messages successfully', {
                        chatId: chatId.substring(0, 20) + '...',
                        messageCount: messages.length,
                        duration: Date.now() - startTime,
                        timestamp: new Date().toISOString()
                    }, 'teams', null, userId);
                }

                res.json({ messages, count: messages.length });
            } catch (error) {
                handleControllerError(res, error, 'getChatMessages', userId, sessionId, startTime);
            }
        },

        /**
         * Send a message to a chat
         * POST /api/v1/teams/chats/:chatId/messages
         */
        async sendChatMessage(req, res) {
            const { userId = null, deviceId = null } = req.user || {};
            const sessionId = req.session?.id;
            const startTime = Date.now();
            const { chatId } = req.params;

            try {
                if (!chatId) {
                    return res.status(400).json({
                        error: 'INVALID_REQUEST',
                        error_description: 'Chat ID is required'
                    });
                }

                const { error: validationError, value: validatedData } = validateAndLog(
                    req,
                    schemas.sendChatMessage,
                    'sendChatMessage',
                    { userId, deviceId }
                );

                if (validationError) {
                    return res.status(400).json({
                        error: 'INVALID_REQUEST',
                        error_description: validationError.details[0].message
                    });
                }

                const message = await teamsModule.sendChatMessage(chatId, validatedData, req, userId, sessionId);

                if (userId) {
                    MonitoringService.info('Sent chat message successfully', {
                        chatId: chatId.substring(0, 20) + '...',
                        messageId: message.id,
                        duration: Date.now() - startTime,
                        timestamp: new Date().toISOString()
                    }, 'teams', null, userId);
                }

                res.status(201).json({ message, success: true });
            } catch (error) {
                handleControllerError(res, error, 'sendChatMessage', userId, sessionId, startTime);
            }
        },

        // ====================================================================
        // TEAM & CHANNEL ENDPOINTS
        // ====================================================================

        /**
         * List user's joined teams
         * GET /api/v1/teams
         */
        async listJoinedTeams(req, res) {
            const { userId = null, deviceId = null } = req.user || {};
            const sessionId = req.session?.id;
            const startTime = Date.now();

            try {
                const { error: validationError, value: validatedData } = validateAndLog(
                    req,
                    schemas.listTeams,
                    'listJoinedTeams',
                    { userId, deviceId }
                );

                if (validationError) {
                    return res.status(400).json({
                        error: 'INVALID_REQUEST',
                        error_description: validationError.details[0].message
                    });
                }

                const teams = await teamsModule.listJoinedTeams(validatedData, req, userId, sessionId);

                if (userId) {
                    MonitoringService.info('Listed joined teams successfully', {
                        teamCount: teams.length,
                        duration: Date.now() - startTime,
                        timestamp: new Date().toISOString()
                    }, 'teams', null, userId);
                }

                res.json({ teams, count: teams.length });
            } catch (error) {
                handleControllerError(res, error, 'listJoinedTeams', userId, sessionId, startTime);
            }
        },

        /**
         * List channels in a team
         * GET /api/v1/teams/:teamId/channels
         */
        async listTeamChannels(req, res) {
            const { userId = null, deviceId = null } = req.user || {};
            const sessionId = req.session?.id;
            const startTime = Date.now();
            const { teamId } = req.params;

            try {
                if (!teamId) {
                    return res.status(400).json({
                        error: 'INVALID_REQUEST',
                        error_description: 'Team ID is required'
                    });
                }

                const channels = await teamsModule.listTeamChannels(teamId, {}, req, userId, sessionId);

                if (userId) {
                    MonitoringService.info('Listed team channels successfully', {
                        teamId: teamId.substring(0, 20) + '...',
                        channelCount: channels.length,
                        duration: Date.now() - startTime,
                        timestamp: new Date().toISOString()
                    }, 'teams', null, userId);
                }

                res.json({ channels, count: channels.length });
            } catch (error) {
                handleControllerError(res, error, 'listTeamChannels', userId, sessionId, startTime);
            }
        },

        /**
         * Get messages from a channel
         * GET /api/v1/teams/:teamId/channels/:channelId/messages
         */
        async getChannelMessages(req, res) {
            const { userId = null, deviceId = null } = req.user || {};
            const sessionId = req.session?.id;
            const startTime = Date.now();
            const { teamId, channelId } = req.params;

            try {
                if (!teamId || !channelId) {
                    return res.status(400).json({
                        error: 'INVALID_REQUEST',
                        error_description: 'Team ID and Channel ID are required'
                    });
                }

                const { error: validationError, value: validatedData } = validateAndLog(
                    req,
                    schemas.getChannelMessages,
                    'getChannelMessages',
                    { userId, deviceId }
                );

                if (validationError) {
                    return res.status(400).json({
                        error: 'INVALID_REQUEST',
                        error_description: validationError.details[0].message
                    });
                }

                const messages = await teamsModule.getChannelMessages(teamId, channelId, validatedData, req, userId, sessionId);

                if (userId) {
                    MonitoringService.info('Retrieved channel messages successfully', {
                        teamId: teamId.substring(0, 20) + '...',
                        channelId: channelId.substring(0, 20) + '...',
                        messageCount: messages.length,
                        duration: Date.now() - startTime,
                        timestamp: new Date().toISOString()
                    }, 'teams', null, userId);
                }

                res.json({ messages, count: messages.length });
            } catch (error) {
                handleControllerError(res, error, 'getChannelMessages', userId, sessionId, startTime);
            }
        },

        /**
         * Send a message to a channel
         * POST /api/v1/teams/:teamId/channels/:channelId/messages
         */
        async sendChannelMessage(req, res) {
            const { userId = null, deviceId = null } = req.user || {};
            const sessionId = req.session?.id;
            const startTime = Date.now();
            const { teamId, channelId } = req.params;

            try {
                if (!teamId || !channelId) {
                    return res.status(400).json({
                        error: 'INVALID_REQUEST',
                        error_description: 'Team ID and Channel ID are required'
                    });
                }

                const { error: validationError, value: validatedData } = validateAndLog(
                    req,
                    schemas.sendChannelMessage,
                    'sendChannelMessage',
                    { userId, deviceId }
                );

                if (validationError) {
                    return res.status(400).json({
                        error: 'INVALID_REQUEST',
                        error_description: validationError.details[0].message
                    });
                }

                const message = await teamsModule.sendChannelMessage(teamId, channelId, validatedData, req, userId, sessionId);

                if (userId) {
                    MonitoringService.info('Sent channel message successfully', {
                        teamId: teamId.substring(0, 20) + '...',
                        channelId: channelId.substring(0, 20) + '...',
                        messageId: message.id,
                        duration: Date.now() - startTime,
                        timestamp: new Date().toISOString()
                    }, 'teams', null, userId);
                }

                res.status(201).json({ message, success: true });
            } catch (error) {
                handleControllerError(res, error, 'sendChannelMessage', userId, sessionId, startTime);
            }
        },

        /**
         * Reply to a message in a channel
         * POST /api/v1/teams/:teamId/channels/:channelId/messages/:messageId/replies
         */
        async replyToMessage(req, res) {
            const { userId = null, deviceId = null } = req.user || {};
            const sessionId = req.session?.id;
            const startTime = Date.now();
            const { teamId, channelId, messageId } = req.params;

            try {
                if (!teamId || !channelId || !messageId) {
                    return res.status(400).json({
                        error: 'INVALID_REQUEST',
                        error_description: 'Team ID, Channel ID, and Message ID are required'
                    });
                }

                const { error: validationError, value: validatedData } = validateAndLog(
                    req,
                    schemas.replyToMessage,
                    'replyToMessage',
                    { userId, deviceId }
                );

                if (validationError) {
                    return res.status(400).json({
                        error: 'INVALID_REQUEST',
                        error_description: validationError.details[0].message
                    });
                }

                const reply = await teamsModule.replyToMessage(teamId, channelId, messageId, validatedData, req, userId, sessionId);

                if (userId) {
                    MonitoringService.info('Replied to message successfully', {
                        teamId: teamId.substring(0, 20) + '...',
                        channelId: channelId.substring(0, 20) + '...',
                        parentMessageId: messageId.substring(0, 20) + '...',
                        replyId: reply.id,
                        duration: Date.now() - startTime,
                        timestamp: new Date().toISOString()
                    }, 'teams', null, userId);
                }

                res.status(201).json({ reply, success: true });
            } catch (error) {
                handleControllerError(res, error, 'replyToMessage', userId, sessionId, startTime);
            }
        },

        // ====================================================================
        // CHANNEL MANAGEMENT ENDPOINTS
        // ====================================================================

        /**
         * Create a new channel in a team
         * POST /api/v1/teams/:teamId/channels
         */
        async createTeamChannel(req, res) {
            const { userId = null, deviceId = null } = req.user || {};
            const sessionId = req.session?.id;
            const startTime = Date.now();
            const { teamId } = req.params;

            try {
                if (!teamId) {
                    return res.status(400).json({
                        error: 'INVALID_REQUEST',
                        error_description: 'Team ID is required'
                    });
                }

                const { error: validationError, value: validatedData } = validateAndLog(
                    req,
                    schemas.createTeamChannel,
                    'createTeamChannel',
                    { userId, deviceId }
                );

                if (validationError) {
                    return res.status(400).json({
                        error: 'INVALID_REQUEST',
                        error_description: validationError.details[0].message
                    });
                }

                const channel = await teamsModule.createTeamChannel(teamId, validatedData, req, userId, sessionId);

                if (userId) {
                    MonitoringService.info('Created team channel successfully', {
                        teamId: teamId.substring(0, 20) + '...',
                        channelId: channel.id,
                        displayName: validatedData.displayName,
                        duration: Date.now() - startTime,
                        timestamp: new Date().toISOString()
                    }, 'teams', null, userId);
                }

                res.status(201).json({ channel, success: true });
            } catch (error) {
                handleControllerError(res, error, 'createTeamChannel', userId, sessionId, startTime);
            }
        },

        /**
         * Add a member to a channel
         * POST /api/v1/teams/:teamId/channels/:channelId/members
         */
        async addChannelMember(req, res) {
            const { userId = null, deviceId = null } = req.user || {};
            const sessionId = req.session?.id;
            const startTime = Date.now();
            const { teamId, channelId } = req.params;

            try {
                if (!teamId || !channelId) {
                    return res.status(400).json({
                        error: 'INVALID_REQUEST',
                        error_description: 'Team ID and Channel ID are required'
                    });
                }

                const { error: validationError, value: validatedData } = validateAndLog(
                    req,
                    schemas.addChannelMember,
                    'addChannelMember',
                    { userId, deviceId }
                );

                if (validationError) {
                    return res.status(400).json({
                        error: 'INVALID_REQUEST',
                        error_description: validationError.details[0].message
                    });
                }

                const member = await teamsModule.addChannelMember(teamId, channelId, validatedData, req, userId, sessionId);

                if (userId) {
                    MonitoringService.info('Added channel member successfully', {
                        teamId: teamId.substring(0, 20) + '...',
                        channelId: channelId.substring(0, 20) + '...',
                        memberId: member.id,
                        duration: Date.now() - startTime,
                        timestamp: new Date().toISOString()
                    }, 'teams', null, userId);
                }

                res.status(201).json({ member, success: true });
            } catch (error) {
                handleControllerError(res, error, 'addChannelMember', userId, sessionId, startTime);
            }
        },

        // ====================================================================
        // CHANNEL FILES ENDPOINTS
        // ====================================================================

        /**
         * List files in a channel
         * GET /api/v1/teams/:teamId/channels/:channelId/files
         */
        async listChannelFiles(req, res) {
            const { userId = null, deviceId = null } = req.user || {};
            const sessionId = req.session?.id;
            const startTime = Date.now();
            const { teamId, channelId } = req.params;

            try {
                if (!teamId || !channelId) {
                    return res.status(400).json({
                        error: 'INVALID_REQUEST',
                        error_description: 'Team ID and Channel ID are required'
                    });
                }

                const { error: validationError, value: validatedData } = validateAndLog(
                    req,
                    schemas.listChannelFiles,
                    'listChannelFiles',
                    { userId, deviceId }
                );

                if (validationError) {
                    return res.status(400).json({
                        error: 'INVALID_REQUEST',
                        error_description: validationError.details[0].message
                    });
                }

                const files = await teamsModule.listChannelFiles(teamId, channelId, validatedData, req, userId, sessionId);

                if (userId) {
                    MonitoringService.info('Listed channel files successfully', {
                        teamId: teamId.substring(0, 20) + '...',
                        channelId: channelId.substring(0, 20) + '...',
                        fileCount: files.length,
                        duration: Date.now() - startTime,
                        timestamp: new Date().toISOString()
                    }, 'teams', null, userId);
                }

                res.json({ files, count: files.length });
            } catch (error) {
                handleControllerError(res, error, 'listChannelFiles', userId, sessionId, startTime);
            }
        },

        /**
         * Upload a file to a channel
         * POST /api/v1/teams/:teamId/channels/:channelId/files
         */
        async uploadFileToChannel(req, res) {
            const { userId = null, deviceId = null } = req.user || {};
            const sessionId = req.session?.id;
            const startTime = Date.now();
            const { teamId, channelId } = req.params;

            try {
                if (!teamId || !channelId) {
                    return res.status(400).json({
                        error: 'INVALID_REQUEST',
                        error_description: 'Team ID and Channel ID are required'
                    });
                }

                const { error: validationError, value: validatedData } = validateAndLog(
                    req,
                    schemas.uploadFileToChannel,
                    'uploadFileToChannel',
                    { userId, deviceId }
                );

                if (validationError) {
                    return res.status(400).json({
                        error: 'INVALID_REQUEST',
                        error_description: validationError.details[0].message
                    });
                }

                const file = await teamsModule.uploadFileToChannel(teamId, channelId, validatedData, req, userId, sessionId);

                if (userId) {
                    MonitoringService.info('Uploaded file to channel successfully', {
                        teamId: teamId.substring(0, 20) + '...',
                        channelId: channelId.substring(0, 20) + '...',
                        fileId: file.id,
                        fileName: validatedData.fileName,
                        duration: Date.now() - startTime,
                        timestamp: new Date().toISOString()
                    }, 'teams', null, userId);
                }

                res.status(201).json({ file, success: true });
            } catch (error) {
                handleControllerError(res, error, 'uploadFileToChannel', userId, sessionId, startTime);
            }
        },

        /**
         * Read content of a file from a channel
         * GET /api/v1/teams/:teamId/channels/:channelId/files/:fileName
         */
        async readChannelFile(req, res) {
            const { userId = null } = req.user || {};
            const sessionId = req.session?.id;
            const startTime = Date.now();
            const { teamId, channelId, fileName } = req.params;

            try {
                if (!teamId || !channelId) {
                    return res.status(400).json({
                        error: 'INVALID_REQUEST',
                        error_description: 'Team ID and Channel ID are required'
                    });
                }

                if (!fileName) {
                    return res.status(400).json({
                        error: 'INVALID_REQUEST',
                        error_description: 'File name is required'
                    });
                }

                // Decode the filename (it comes URL-encoded from the route)
                const decodedFileName = decodeURIComponent(fileName);

                const content = await teamsModule.readChannelFile(teamId, channelId, decodedFileName, req, userId, sessionId);

                if (userId) {
                    MonitoringService.info('Read channel file successfully', {
                        teamId: teamId.substring(0, 20) + '...',
                        channelId: channelId.substring(0, 20) + '...',
                        fileName: decodedFileName,
                        size: content.size,
                        duration: Date.now() - startTime,
                        timestamp: new Date().toISOString()
                    }, 'teams', null, userId);
                }

                res.json({ file: content });
            } catch (error) {
                handleControllerError(res, error, 'readChannelFile', userId, sessionId, startTime);
            }
        },

        // ====================================================================
        // MEETING ENDPOINTS
        // ====================================================================

        /**
         * List online meetings
         * GET /api/v1/teams/meetings
         */
        async listOnlineMeetings(req, res) {
            const { userId = null, deviceId = null } = req.user || {};
            const sessionId = req.session?.id;
            const startTime = Date.now();

            try {
                const { error: validationError, value: validatedData } = validateAndLog(
                    req,
                    schemas.listOnlineMeetings,
                    'listOnlineMeetings',
                    { userId, deviceId }
                );

                if (validationError) {
                    return res.status(400).json({
                        error: 'INVALID_REQUEST',
                        error_description: validationError.details[0].message
                    });
                }

                const meetings = await teamsModule.listOnlineMeetings(validatedData, req, userId, sessionId);

                if (userId) {
                    MonitoringService.info('Listed online meetings successfully', {
                        meetingCount: meetings.length,
                        duration: Date.now() - startTime,
                        timestamp: new Date().toISOString()
                    }, 'teams', null, userId);
                }

                res.json({ meetings, count: meetings.length });
            } catch (error) {
                handleControllerError(res, error, 'listOnlineMeetings', userId, sessionId, startTime);
            }
        },

        /**
         * Create an online meeting
         * POST /api/v1/teams/meetings
         */
        async createOnlineMeeting(req, res) {
            const { userId = null, deviceId = null } = req.user || {};
            const sessionId = req.session?.id;
            const startTime = Date.now();

            try {
                const { error: validationError, value: validatedData } = validateAndLog(
                    req,
                    schemas.createOnlineMeeting,
                    'createOnlineMeeting',
                    { userId, deviceId }
                );

                if (validationError) {
                    return res.status(400).json({
                        error: 'INVALID_REQUEST',
                        error_description: validationError.details[0].message
                    });
                }

                const meeting = await teamsModule.createOnlineMeeting(validatedData, req, userId, sessionId);

                if (userId) {
                    MonitoringService.info('Created online meeting successfully', {
                        meetingId: meeting.id,
                        subject: validatedData.subject?.substring(0, 30),
                        hasJoinUrl: !!meeting.joinUrl,
                        duration: Date.now() - startTime,
                        timestamp: new Date().toISOString()
                    }, 'teams', null, userId);
                }

                res.status(201).json({ meeting, success: true });
            } catch (error) {
                handleControllerError(res, error, 'createOnlineMeeting', userId, sessionId, startTime);
            }
        },

        /**
         * Get meeting by join URL
         * GET /api/v1/teams/meetings/findByJoinUrl
         */
        async getMeetingByJoinUrl(req, res) {
            const { userId = null, deviceId = null } = req.user || {};
            const sessionId = req.session?.id;
            const startTime = Date.now();

            try {
                const { error: validationError, value: validatedData } = validateAndLog(
                    req,
                    schemas.getMeetingByJoinUrl,
                    'getMeetingByJoinUrl',
                    { userId, deviceId }
                );

                if (validationError) {
                    return res.status(400).json({
                        error: 'INVALID_REQUEST',
                        error_description: validationError.details[0].message
                    });
                }

                const meeting = await teamsModule.getMeetingByJoinUrl(validatedData.joinUrl, req, userId, sessionId);

                if (userId) {
                    MonitoringService.info('Found meeting by join URL successfully', {
                        meetingId: meeting.id,
                        duration: Date.now() - startTime,
                        timestamp: new Date().toISOString()
                    }, 'teams', null, userId);
                }

                res.json({ meeting });
            } catch (error) {
                handleControllerError(res, error, 'getMeetingByJoinUrl', userId, sessionId, startTime);
            }
        },

        /**
         * Get online meeting by ID
         * GET /api/v1/teams/meetings/:meetingId
         */
        async getOnlineMeeting(req, res) {
            const { userId = null, deviceId = null } = req.user || {};
            const sessionId = req.session?.id;
            const startTime = Date.now();
            const { meetingId } = req.params;

            try {
                if (!meetingId) {
                    return res.status(400).json({
                        error: 'INVALID_REQUEST',
                        error_description: 'Meeting ID is required'
                    });
                }

                const meeting = await teamsModule.getOnlineMeeting(meetingId, req, userId, sessionId);

                if (userId) {
                    MonitoringService.info('Retrieved online meeting successfully', {
                        meetingId: meeting.id,
                        duration: Date.now() - startTime,
                        timestamp: new Date().toISOString()
                    }, 'teams', null, userId);
                }

                res.json({ meeting });
            } catch (error) {
                handleControllerError(res, error, 'getOnlineMeeting', userId, sessionId, startTime);
            }
        },

        /**
         * Get transcripts for an online meeting
         * GET /api/v1/teams/meetings/:meetingId/transcripts
         */
        async getMeetingTranscripts(req, res) {
            const { userId = null, deviceId = null } = req.user || {};
            const sessionId = req.session?.id;
            const startTime = Date.now();
            const { meetingId } = req.params;

            try {
                if (!meetingId) {
                    return res.status(400).json({
                        error: 'INVALID_REQUEST',
                        error_description: 'Meeting ID is required'
                    });
                }

                const transcripts = await teamsModule.getMeetingTranscripts(meetingId, req, userId, sessionId);

                if (userId) {
                    MonitoringService.info('Retrieved meeting transcripts successfully', {
                        meetingId: meetingId.substring(0, 20) + '...',
                        transcriptCount: transcripts.length,
                        duration: Date.now() - startTime,
                        timestamp: new Date().toISOString()
                    }, 'teams', null, userId);
                }

                res.json({ transcripts });
            } catch (error) {
                handleControllerError(res, error, 'getMeetingTranscripts', userId, sessionId, startTime);
            }
        },

        /**
         * Get transcript content for a specific transcript
         * GET /api/v1/teams/meetings/:meetingId/transcripts/:transcriptId
         */
        async getMeetingTranscriptContent(req, res) {
            const { userId = null, deviceId = null } = req.user || {};
            const sessionId = req.session?.id;
            const startTime = Date.now();
            const { meetingId, transcriptId } = req.params;

            try {
                if (!meetingId) {
                    return res.status(400).json({
                        error: 'INVALID_REQUEST',
                        error_description: 'Meeting ID is required'
                    });
                }

                if (!transcriptId) {
                    return res.status(400).json({
                        error: 'INVALID_REQUEST',
                        error_description: 'Transcript ID is required'
                    });
                }

                const content = await teamsModule.getMeetingTranscriptContent(meetingId, transcriptId, req, userId, sessionId);

                if (userId) {
                    MonitoringService.info('Retrieved meeting transcript content successfully', {
                        meetingId: meetingId.substring(0, 20) + '...',
                        transcriptId: transcriptId.substring(0, 20) + '...',
                        entryCount: content.entryCount,
                        duration: Date.now() - startTime,
                        timestamp: new Date().toISOString()
                    }, 'teams', null, userId);
                }

                res.json({ transcript: content });
            } catch (error) {
                handleControllerError(res, error, 'getMeetingTranscriptContent', userId, sessionId, startTime);
            }
        }
    };
}

/**
 * Handle controller errors consistently
 */
function handleControllerError(res, error, operation, userId, sessionId, startTime) {
    const duration = Date.now() - startTime;

    // Log infrastructure error
    const mcpError = ErrorService.createError(
        'teams',
        `Teams ${operation} failed`,
        'error',
        {
            endpoint: `/api/v1/teams/${operation}`,
            error: error.message,
            stack: error.stack,
            operation,
            userId,
            timestamp: new Date().toISOString()
        }
    );
    MonitoringService.logError(mcpError);

    // Log user error tracking
    if (userId) {
        MonitoringService.error(`Teams ${operation} failed`, {
            error: error.message,
            operation,
            duration,
            timestamp: new Date().toISOString()
        }, 'teams', null, userId);
    } else if (sessionId) {
        MonitoringService.error(`Teams ${operation} failed`, {
            sessionId,
            error: error.message,
            operation,
            duration,
            timestamp: new Date().toISOString()
        }, 'teams');
    }

    // Track error metrics
    MonitoringService.trackMetric(`teams.${operation}.error`, 1, {
        errorMessage: error.message,
        duration,
        success: false,
        userId
    });

    // Return appropriate status code
    const statusCode = error.statusCode || error.code === 'ENOTFOUND' ? 503 : 500;
    res.status(statusCode).json({
        error: 'TEAMS_OPERATION_FAILED',
        error_description: `Failed to ${operation.replace(/([A-Z])/g, ' $1').toLowerCase()}`
    });
}

module.exports = createTeamsController;
