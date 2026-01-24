/**
 * @fileoverview MailService - Microsoft Graph Mail API operations.
 * All methods are async, modular, and use GraphClient for requests.
 * Follows project error handling, validation, and normalization rules.
 */

const graphClientFactory = require('./graph-client.cjs');
const ErrorService = require('../core/error-service.cjs');
const MonitoringService = require('../core/monitoring-service.cjs');

// PERF-3: Graph API query limit constants
const GRAPH_API_MAX_LIMIT = 1000; // Microsoft Graph API maximum for mail messages

/**
 * Validates and clamps a query limit to Graph API constraints
 * @param {number} limit - Requested limit
 * @param {number} defaultLimit - Default if not specified (default: 10)
 * @returns {number} - Validated limit
 */
function validateQueryLimit(limit, defaultLimit = 10) {
    const requested = limit || defaultLimit;
    const validated = Math.min(Math.max(1, requested), GRAPH_API_MAX_LIMIT);
    if (requested > GRAPH_API_MAX_LIMIT) {
        MonitoringService.warn('Query limit exceeded Graph API max', {
            requested,
            applied: validated,
            max: GRAPH_API_MAX_LIMIT,
            timestamp: new Date().toISOString()
        }, 'mail');
    }
    return validated;
}

// Log service initialization
MonitoringService.info('Graph Mail Service initialized', {
    serviceName: 'graph-mail-service',
    timestamp: new Date().toISOString()
}, 'graph');

/**
 * Normalizes a Graph email object to MCP schema.
 */
function normalizeEmail(graphEmail) {
  return {
    id: graphEmail.id,
    subject: graphEmail.subject,
    from: {
      name: graphEmail.from?.emailAddress?.name,
      email: graphEmail.from?.emailAddress?.address
    },
    received: graphEmail.receivedDateTime,
    preview: graphEmail.bodyPreview?.substring(0, 150),
    isRead: graphEmail.isRead,
    importance: graphEmail.importance,
    hasAttachments: graphEmail.hasAttachments
  };
}

/**
 * Retrieves inbox emails.
 * @param {object} options
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<Array<object>>}
 */
async function getInbox(options = {}, req, userId, sessionId) {
  const startTime = Date.now();
  
  // Extract user context from request if not provided
  const contextUserId = userId || req?.user?.userId;
  const contextSessionId = sessionId || req?.session?.id;
  
  // Pattern 1: Development Debug Logs
  if (process.env.NODE_ENV === 'development') {
    MonitoringService.debug('Mail getInbox operation started', {
      method: 'getInbox',
      optionKeys: Object.keys(options),
      sessionId: contextSessionId,
      userAgent: req?.get('User-Agent'),
      timestamp: new Date().toISOString()
    }, 'mail');
  }
  
  try {
    const client = await graphClientFactory.createClient(req, contextUserId, contextSessionId);
    const top = validateQueryLimit(options.top || options.limit);
    const res = await client.api(`/me/mailFolders/inbox/messages?$top=${top}`, contextUserId, contextSessionId).get();
    const emails = (res.value || []).map(normalizeEmail);

    const executionTime = Date.now() - startTime;

    // Pattern 2: User Activity Logs
    if (contextUserId) {
      MonitoringService.info('Retrieved inbox emails successfully', {
        emailCount: emails.length,
        requestedTop: top,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'mail', null, contextUserId);
    } else if (contextSessionId) {
      MonitoringService.info('Retrieved inbox emails with session', {
        sessionId: contextSessionId,
        emailCount: emails.length,
        requestedTop: top,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'mail');
    }
    
    return emails;
  } catch (error) {
    const executionTime = Date.now() - startTime;
    
    // Pattern 3: Infrastructure Error Logging
    const mcpError = ErrorService.createError(
      'mail',
      `Failed to get inbox emails: ${error.message}`,
      'error',
      {
        service: 'graph-mail-service',
        method: 'getInbox',
        requestedTop: options.top || options.limit || 10,
        executionTimeMs: executionTime,
        error: error.message,
        stack: error.stack,
        timestamp: new Date().toISOString()
      }
    );
    
    MonitoringService.logError(mcpError);
    
    // Pattern 4: User Error Tracking
    if (contextUserId) {
      MonitoringService.error('Failed to retrieve inbox emails', {
        error: error.message,
        requestedTop: options.top || options.limit || 10,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'mail', null, contextUserId);
    } else if (contextSessionId) {
      MonitoringService.error('Failed to retrieve inbox emails', {
        sessionId: contextSessionId,
        error: error.message,
        requestedTop: options.top || options.limit || 10,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'mail');
    }
    
    throw mcpError;
  }
}

/**
 * Searches emails by query string using Microsoft Graph KQL syntax.
 * @param {string} query - KQL search query (e.g., "from:user@domain.com subject:meeting")
 * @param {object} options
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<Array<object>>}
 */
async function searchEmails(query, options = {}, req, userId, sessionId) {
  const startTime = Date.now();
  
  // Extract user context from request if not provided
  const contextUserId = userId || req?.user?.userId;
  const contextSessionId = sessionId || req?.session?.id;
  
  // Pattern 1: Development Debug Logs
  if (process.env.NODE_ENV === 'development') {
    MonitoringService.debug('Mail searchEmails operation started', {
      method: 'searchEmails',
      query: query ? query.substring(0, 50) + '...' : null, // Truncate for privacy
      queryLength: query ? query.length : 0,
      optionKeys: Object.keys(options),
      sessionId: contextSessionId,
      userAgent: req?.get('User-Agent'),
      timestamp: new Date().toISOString()
    }, 'mail');
  }
  
  try {
    if (!query || typeof query !== 'string') {
      const mcpError = ErrorService.createError(
        'mail',
        'Search query must be a non-empty string',
        'warning',
        {
          service: 'graph-mail-service',
          method: 'searchEmails',
          queryType: typeof query,
          timestamp: new Date().toISOString()
        }
      );
      MonitoringService.logError(mcpError);
      throw mcpError;
    }
    
    const client = await graphClientFactory.createClient(req, contextUserId, contextSessionId);
    const top = options.top || options.limit || 10;
    
    // Pass query through - let the LLM handle formatting
    const cleanQuery = query.trim();

    const searchUrl = `/me/messages?$search="${encodeURIComponent(cleanQuery)}"&$top=${top}`;
    const res = await client.api(searchUrl, contextUserId, contextSessionId).get();
    const emails = (res.value || []).map(normalizeEmail);
    
    const executionTime = Date.now() - startTime;
    
    // Pattern 2: User Activity Logs
    if (contextUserId) {
      MonitoringService.info('Searched emails successfully', {
        queryLength: query.length,
        resultCount: emails.length,
        requestedTop: top,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'mail', null, contextUserId);
    } else if (contextSessionId) {
      MonitoringService.info('Searched emails with session', {
        sessionId: contextSessionId,
        queryLength: query.length,
        resultCount: emails.length,
        requestedTop: top,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'mail');
    }
    
    return emails;
  } catch (error) {
    const executionTime = Date.now() - startTime;
    
    // If it's already an MCP error, just track and rethrow
    if (error.category) {
      throw error;
    }
    
    // Pattern 3: Infrastructure Error Logging
    const mcpError = ErrorService.createError(
      'mail',
      `Failed to search emails: ${error.message}`,
      'error',
      {
        service: 'graph-mail-service',
        method: 'searchEmails',
        query: query ? query.substring(0, 50) + '...' : null,
        requestedTop: options.top || options.limit || 10,
        executionTimeMs: executionTime,
        error: error.message,
        stack: error.stack,
        timestamp: new Date().toISOString()
      }
    );
    
    MonitoringService.logError(mcpError);
    
    // Pattern 4: User Error Tracking
    if (contextUserId) {
      MonitoringService.error('Failed to search emails', {
        error: error.message,
        queryLength: query ? query.length : 0,
        requestedTop: options.top || options.limit || 10,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'mail', null, contextUserId);
    } else if (contextSessionId) {
      MonitoringService.error('Failed to search emails', {
        sessionId: contextSessionId,
        error: error.message,
        queryLength: query ? query.length : 0,
        requestedTop: options.top || options.limit || 10,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'mail');
    }
    
    throw mcpError;
  }
}

/**
 * Sends an email.
 * @param {object} emailData
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<boolean>}
 */
async function sendEmail(emailData, req, userId, sessionId) {
  const startTime = Date.now();
  
  // Extract user context from request if not provided
  const contextUserId = userId || req?.user?.userId;
  const contextSessionId = sessionId || req?.session?.id;
  
  // Pattern 1: Development Debug Logs
  if (process.env.NODE_ENV === 'development') {
    MonitoringService.debug('Mail sendEmail operation started', {
      method: 'sendEmail',
      hasAttachments: emailData.attachments && Array.isArray(emailData.attachments),
      attachmentCount: emailData.attachments ? emailData.attachments.length : 0,
      sessionId: contextSessionId,
      userAgent: req?.get('User-Agent'),
      timestamp: new Date().toISOString()
    }, 'mail');
  }
  
  try {
    const client = await graphClientFactory.createClient(req);
    const { to, subject, body, cc, bcc, contentType, attachments } = emailData;
    
    // Handle recipients in various formats (string or array)
    function formatRecipients(recipients) {
      if (!recipients) return [];
      
      // Convert string to array if needed
      const recipientArray = Array.isArray(recipients) ? recipients : [recipients];
      
      // Format each recipient
      return recipientArray.map(recipient => ({
        emailAddress: { address: recipient }
      }));
    }
    
    const message = {
      subject,
      body: {
        contentType: contentType || 'Text',
        content: body
      },
      toRecipients: formatRecipients(to),
      ccRecipients: formatRecipients(cc),
      bccRecipients: formatRecipients(bcc)
    };
    
    // Add attachments if provided
    if (attachments && Array.isArray(attachments) && attachments.length > 0) {
      if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Processing email attachments', {
          attachmentCount: attachments.length,
          timestamp: new Date().toISOString()
        }, 'graph');
      }
      
      // Process attachments asynchronously
      const processedAttachments = await Promise.all(attachments.map(async attachment => {
        try {
          // Check if attachment is a file ID (string) or an attachment object
          if (typeof attachment === 'string') {
            if (process.env.NODE_ENV === 'development') {
              MonitoringService.debug('Processing file attachment by ID', {
                fileId: attachment,
                timestamp: new Date().toISOString()
              }, 'graph');
            }
            // This is a file ID from the files service
            // We need to get the file content from the files service
            try {
              // Import the files service
              const filesService = require('./files-service.cjs');
              
              // Get the file metadata and content
              const fileMetadata = await filesService.getFileMetadata(attachment, req);
              const fileContent = await filesService.getFileContent(attachment, req);
              
              if (!fileMetadata || !fileContent) {
                MonitoringService.warn('Could not retrieve file for email attachment', {
                  fileId: attachment,
                  hasMetadata: !!fileMetadata,
                  hasContent: !!fileContent,
                  timestamp: new Date().toISOString()
                }, 'graph');
                return null;
              }
              
              if (process.env.NODE_ENV === 'development') {
                MonitoringService.debug('Retrieved file for email attachment', {
                  fileName: fileMetadata.name,
                  fileSize: fileContent.length,
                  timestamp: new Date().toISOString()
                }, 'graph');
              }
              
              // Convert file content to base64
              const contentBytes = Buffer.from(fileContent).toString('base64');
              
              return {
                '@odata.type': '#microsoft.graph.fileAttachment',
                name: fileMetadata.name,
                contentType: fileMetadata.contentType || 'application/octet-stream',
                contentBytes: contentBytes,
                isInline: false
              };
            } catch (fileError) {
              const mcpError = ErrorService.createError(
                'mail',
                `Error retrieving file for email attachment: ${fileError.message}`,
                'warning',
                {
                  service: 'graph-mail-service',
                  method: 'sendEmail',
                  fileId: attachment,
                  stack: fileError.stack,
                  timestamp: new Date().toISOString()
                }
              );
              MonitoringService.logError(mcpError);
              return null;
            }
          }
          
          // Handle attachment object
          // Check if we have contentBytes or need to convert from content
          let contentBytes = attachment.contentBytes;
          
          // If we have content but not contentBytes, convert content to base64
          if (!contentBytes && attachment.content) {
            if (process.env.NODE_ENV === 'development') {
              MonitoringService.debug('Converting content to contentBytes for attachment', {
                attachmentName: attachment.name,
                timestamp: new Date().toISOString()
              }, 'graph');
            }
            contentBytes = Buffer.from(attachment.content).toString('base64');
          }
          
          // Ensure contentBytes is properly formatted - must be a valid base64 string
          if (contentBytes && typeof contentBytes === 'string') {
            // Make sure it's properly padded base64
            const paddingNeeded = contentBytes.length % 4;
            if (paddingNeeded > 0) {
              contentBytes += '='.repeat(4 - paddingNeeded);
            }
          }
          
          // Ensure we have all required fields for a valid attachment
          if (!contentBytes || !attachment.name || !attachment.contentType) {
            MonitoringService.warn('Invalid attachment missing required fields', {
              hasName: !!attachment.name,
              hasContentType: !!attachment.contentType,
              hasContentBytes: !!contentBytes,
              timestamp: new Date().toISOString()
            }, 'graph');
            return null; // Skip invalid attachments
          }
          
          return {
            '@odata.type': '#microsoft.graph.fileAttachment',
            name: attachment.name,
            contentType: attachment.contentType,
            contentBytes: contentBytes,
            isInline: attachment.isInline || false
          };
        } catch (attachmentError) {
          const mcpError = ErrorService.createError(
            'mail',
            `Error processing email attachment: ${attachmentError.message}`,
            'warning',
            {
              service: 'graph-mail-service',
              method: 'sendEmail',
              attachmentName: attachment.name || 'unknown',
              stack: attachmentError.stack,
              timestamp: new Date().toISOString()
            }
          );
          MonitoringService.logError(mcpError);
          return null;
        }
      }));
      
      // Filter out null entries from invalid attachments
      message.attachments = processedAttachments.filter(Boolean);
      
      MonitoringService.trackMetric('graph_mail_attachments_processed', Date.now() - startTime, {
        originalCount: attachments.length,
        validCount: message.attachments.length,
        timestamp: new Date().toISOString()
      });
      
      if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Email attachments processed', {
          originalCount: attachments.length,
          validCount: message.attachments.length,
          timestamp: new Date().toISOString()
        }, 'graph');
      }
    }

    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Sending email via Graph API', {
        hasSubject: !!message.subject,
        toCount: message.toRecipients ? message.toRecipients.length : 0,
        ccCount: message.ccRecipients ? message.ccRecipients.length : 0,
        bccCount: message.bccRecipients ? message.bccRecipients.length : 0,
        attachmentCount: message.attachments ? message.attachments.length : 0,
        contentType: message.body.contentType,
        timestamp: new Date().toISOString()
      }, 'graph');
    }
    
    // Explicitly set saveToSentItems to true to ensure the email is saved with attachments
    // Also explicitly set the hasAttachments flag if we have attachments
    if (message.attachments && message.attachments.length > 0) {
      // Microsoft Graph API uses 'hasAttachments' (not 'isHasAttachments')
      message.hasAttachments = true;
    }
    
    const requestBody = {
      message,
      saveToSentItems: true
    };
    
    await client.api('/me/sendMail').post(requestBody);
    
    const executionTime = Date.now() - startTime;
    
    // Pattern 2: User Activity Logs
    if (contextUserId) {
      MonitoringService.info('Email sent successfully', {
        recipientCount: message.toRecipients ? message.toRecipients.length : 0,
        hasAttachments: !!(message.attachments && message.attachments.length > 0),
        attachmentCount: message.attachments ? message.attachments.length : 0,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'mail', null, contextUserId);
    } else if (contextSessionId) {
      MonitoringService.info('Email sent with session', {
        sessionId: contextSessionId,
        recipientCount: message.toRecipients ? message.toRecipients.length : 0,
        hasAttachments: !!(message.attachments && message.attachments.length > 0),
        attachmentCount: message.attachments ? message.attachments.length : 0,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'mail');
    }
    
    MonitoringService.trackMetric('graph_mail_send_email_success', executionTime, {
      service: 'graph-mail-service',
      method: 'sendEmail',
      toCount: message.toRecipients ? message.toRecipients.length : 0,
      ccCount: message.ccRecipients ? message.ccRecipients.length : 0,
      bccCount: message.bccRecipients ? message.bccRecipients.length : 0,
      attachmentCount: message.attachments ? message.attachments.length : 0,
      contentType: message.body.contentType,
      timestamp: new Date().toISOString()
    });
    
    return true;
  } catch (error) {
    const executionTime = Date.now() - startTime;
    
    // Pattern 3: Infrastructure Error Logging
    const mcpError = ErrorService.createError(
      'mail',
      `Failed to send email: ${error.message}`,
      'error',
      {
        service: 'graph-mail-service',
        method: 'sendEmail',
        hasSubject: emailData.subject && emailData.subject.length > 0,
        hasTo: emailData.to && emailData.to.length > 0,
        attachmentCount: emailData.attachments ? emailData.attachments.length : 0,
        graphMessage: error.message,
        stack: error.stack,
        timestamp: new Date().toISOString()
      }
    );
    
    MonitoringService.logError(mcpError);
    
    // Pattern 4: User Error Tracking
    if (contextUserId) {
      MonitoringService.error('User experienced error sending email', {
        errorMessage: 'Failed to send email',
        hasAttachments: !!(emailData.attachments && emailData.attachments.length > 0),
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'mail', null, contextUserId);
    } else if (contextSessionId) {
      MonitoringService.error('Session experienced error sending email', {
        sessionId: contextSessionId,
        errorMessage: 'Failed to send email',
        hasAttachments: !!(emailData.attachments && emailData.attachments.length > 0),
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'mail');
    }
    
    MonitoringService.trackMetric('graph_mail_send_email_failure', executionTime, {
      service: 'graph-mail-service',
      method: 'sendEmail',
      errorType: error.code || 'unknown',
      timestamp: new Date().toISOString()
    });
    
    throw mcpError;
  }
}

/**
 * Flags/unflag an email.
 * @param {string} id - Email ID
 * @param {boolean} flag - Flag state
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<boolean>}
 */
async function flagEmail(id, flag = true, req, userId, sessionId) {
  const startTime = Date.now();
  
  // Extract user context from request if not provided
  const contextUserId = userId || req?.user?.userId;
  const contextSessionId = sessionId || req?.session?.id;
  
  // Pattern 1: Development Debug Logs
  if (process.env.NODE_ENV === 'development') {
    MonitoringService.debug('Mail flagEmail operation started', {
      method: 'flagEmail',
      emailId: id ? id.substring(0, 20) + '...' : null,
      flagState: flag,
      sessionId: contextSessionId,
      userAgent: req?.get('User-Agent'),
      timestamp: new Date().toISOString()
    }, 'mail');
  }
  
  try {
    if (!id || typeof id !== 'string') {
      const mcpError = ErrorService.createError(
        'mail',
        'Email ID must be a non-empty string',
        'warning',
        {
          service: 'graph-mail-service',
          method: 'flagEmail',
          idType: typeof id,
          timestamp: new Date().toISOString()
        }
      );
      MonitoringService.logError(mcpError);
      throw mcpError;
    }
    
    const client = await graphClientFactory.createClient(req, contextUserId, contextSessionId);
    await client.api(`/me/messages/${id}`, contextUserId, contextSessionId).patch({
      flag: { flagStatus: flag ? 'flagged' : 'notFlagged' }
    });
    
    const executionTime = Date.now() - startTime;
    
    // Pattern 2: User Activity Logs
    if (contextUserId) {
      MonitoringService.info('Email flagged successfully', {
        emailId: id ? id.substring(0, 20) + '...' : null,
        flagState: flag,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'mail', null, contextUserId);
    } else if (contextSessionId) {
      MonitoringService.info('Email flagged with session', {
        sessionId: contextSessionId,
        emailId: id ? id.substring(0, 20) + '...' : null,
        flagState: flag,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'mail');
    }
    
    return true;
  } catch (error) {
    const executionTime = Date.now() - startTime;
    
    // If it's already an MCP error, just rethrow
    if (error.category) {
      throw error;
    }
    
    // Pattern 3: Infrastructure Error Logging
    const mcpError = ErrorService.createError(
      'mail',
      `Failed to flag email: ${error.message}`,
      'error',
      {
        service: 'graph-mail-service',
        method: 'flagEmail',
        emailId: id ? id.substring(0, 20) + '...' : null,
        flagState: flag,
        executionTimeMs: executionTime,
        error: error.message,
        stack: error.stack,
        timestamp: new Date().toISOString()
      }
    );
    
    MonitoringService.logError(mcpError);
    
    // Pattern 4: User Error Tracking
    if (contextUserId) {
      MonitoringService.error('Failed to flag email', {
        error: error.message,
        emailId: id ? id.substring(0, 20) + '...' : null,
        flagState: flag,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'mail', null, contextUserId);
    } else if (contextSessionId) {
      MonitoringService.error('Failed to flag email', {
        sessionId: contextSessionId,
        error: error.message,
        emailId: id ? id.substring(0, 20) + '...' : null,
        flagState: flag,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'mail');
    }
    
    throw mcpError;
  }
}

/**
 * Retrieves attachments for an email.
 * @param {string} id
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<Array<object>>}
 */
async function getAttachments(id, req, userId, sessionId) {
  const startTime = Date.now();
  
  // Extract user context from request if not provided
  const contextUserId = userId || req?.user?.userId;
  const contextSessionId = sessionId || req?.session?.id;
  
  // Pattern 1: Development Debug Logs
  if (process.env.NODE_ENV === 'development') {
    MonitoringService.debug('Mail getAttachments operation started', {
      method: 'getAttachments',
      emailId: id ? id.substring(0, 20) + '...' : null,
      sessionId: contextSessionId,
      userAgent: req?.get('User-Agent'),
      timestamp: new Date().toISOString()
    }, 'mail');
  }
  
  try {
    if (!id || typeof id !== 'string') {
      const mcpError = ErrorService.createError(
        'mail',
        'Email ID must be a non-empty string',
        'warning',
        {
          service: 'graph-mail-service',
          method: 'getAttachments',
          idType: typeof id,
          timestamp: new Date().toISOString()
        }
      );
      MonitoringService.logError(mcpError);
      throw mcpError;
    }
    
    const client = await graphClientFactory.createClient(req, contextUserId, contextSessionId);
    
    // First check if the email exists and has attachments
    try {
      const emailDetails = await client.api(`/me/messages/${id}`).get();
      
      if (process.env.NODE_ENV === 'development') {
        MonitoringService.debug('Email metadata retrieved for attachments check', {
          emailId: id,
          hasAttachments: emailDetails.hasAttachments,
          timestamp: new Date().toISOString()
        }, 'graph');
      }
      
      // If the email doesn't have attachments according to metadata, return empty array early
      if (!emailDetails.hasAttachments) {
        const executionTime = Date.now() - startTime;
        MonitoringService.trackMetric('graph_mail_get_attachments_no_attachments', executionTime, {
          service: 'graph-mail-service',
          method: 'getAttachments',
          timestamp: new Date().toISOString()
        });
        return [];
      }
    } catch (metadataError) {
      MonitoringService.warn('Error checking email metadata for attachments', {
        emailId: id,
        error: metadataError.message,
        timestamp: new Date().toISOString()
      }, 'graph');
      // Continue anyway to try getting attachments directly
    }
    
    // Use $select to ensure we get all attachment properties
    const res = await client.api(`/me/messages/${id}/attachments`).get();
    
    const attachments = res.value || [];
    
    if (process.env.NODE_ENV === 'development' && attachments.length > 0) {
      // Log attachment details for debugging
      attachments.forEach(attachment => {
        MonitoringService.debug('Email attachment found', {
          emailId: id,
          attachmentName: attachment.name,
          contentType: attachment.contentType,
          size: attachment.size || 'unknown',
          timestamp: new Date().toISOString()
        }, 'graph');
      });
    }
    
    const normalizedAttachments = attachments.map(attachment => ({
      id: attachment.id,
      name: attachment.name,
      contentType: attachment.contentType,
      size: attachment.size || 0,
      isInline: attachment.isInline || false,
      lastModifiedDateTime: attachment.lastModifiedDateTime || new Date().toISOString()
    }));
    
    const executionTime = Date.now() - startTime;
    
    // Pattern 2: User Activity Logs
    if (contextUserId) {
      MonitoringService.info('Email attachments retrieved successfully', {
        attachmentCount: normalizedAttachments.length,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'mail', null, contextUserId);
    } else if (contextSessionId) {
      MonitoringService.info('Email attachments retrieved with session', {
        sessionId: contextSessionId,
        attachmentCount: normalizedAttachments.length,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'mail');
    }
    
    MonitoringService.trackMetric('graph_mail_get_attachments_success', executionTime, {
      service: 'graph-mail-service',
      method: 'getAttachments',
      attachmentCount: normalizedAttachments.length,
      timestamp: new Date().toISOString()
    });
    
    return normalizedAttachments;
  } catch (error) {
    const executionTime = Date.now() - startTime;
    
    // If it's already an MCP error, just track metrics and rethrow
    if (error.category) {
      MonitoringService.trackMetric('graph_mail_get_attachments_failure', executionTime, {
        service: 'graph-mail-service',
        method: 'getAttachments',
        errorType: error.code || 'validation_error',
        timestamp: new Date().toISOString()
      });
      throw error;
    }
    
    // Pattern 3: Infrastructure Error Logging
    const mcpError = ErrorService.createError(
      'mail',
      `Failed to get email attachments: ${error.message}`,
      'error',
      {
        service: 'graph-mail-service',
        method: 'getAttachments',
        emailId: id ? id.substring(0, 20) + '...' : null,
        graphMessage: error.message,
        stack: error.stack,
        timestamp: new Date().toISOString()
      }
    );
    
    MonitoringService.logError(mcpError);
    
    // Pattern 4: User Error Tracking
    if (contextUserId) {
      MonitoringService.error('User experienced error retrieving email attachments', {
        errorMessage: 'Failed to retrieve email attachments',
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'mail', null, contextUserId);
    } else if (contextSessionId) {
      MonitoringService.error('Session experienced error retrieving email attachments', {
        sessionId: contextSessionId,
        errorMessage: 'Failed to retrieve email attachments',
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'mail');
    }
    
    MonitoringService.trackMetric('graph_mail_get_attachments_failure', executionTime, {
      service: 'graph-mail-service',
      method: 'getAttachments',
      errorType: error.code || 'unknown',
      timestamp: new Date().toISOString()
    });
    
    throw mcpError;
  }
}

/**
 * Retrieves raw inbox data (no normalization).
 * @param {object} options
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<Array<object>>}
 */
async function getInboxRaw(options = {}, req, userId, sessionId) {
  const startTime = Date.now();
  
  // Extract user context from request if not provided
  const contextUserId = userId || req?.user?.userId;
  const contextSessionId = sessionId || req?.session?.id;
  
  // Pattern 1: Development Debug Logs
  if (process.env.NODE_ENV === 'development') {
    MonitoringService.debug('Mail getInboxRaw operation started', {
      method: 'getInboxRaw',
      optionKeys: Object.keys(options),
      sessionId: contextSessionId,
      userAgent: req?.get('User-Agent'),
      timestamp: new Date().toISOString()
    }, 'mail');
  }
  
  try {
    const client = await graphClientFactory.createClient(req, contextUserId, contextSessionId);
    const top = options.top || options.limit || 10;
    const res = await client.api(`/me/mailFolders/inbox/messages?$top=${top}`, contextUserId, contextSessionId).get();
    const emails = res.value || [];
    
    const executionTime = Date.now() - startTime;
    
    // Pattern 2: User Activity Logs
    if (contextUserId) {
      MonitoringService.info('Retrieved raw inbox emails successfully', {
        emailCount: emails.length,
        requestedTop: top,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'mail', null, contextUserId);
    } else if (contextSessionId) {
      MonitoringService.info('Retrieved raw inbox emails with session', {
        sessionId: contextSessionId,
        emailCount: emails.length,
        requestedTop: top,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'mail');
    }
    
    return emails;
  } catch (error) {
    const executionTime = Date.now() - startTime;
    
    // Pattern 3: Infrastructure Error Logging
    const mcpError = ErrorService.createError(
      'mail',
      `Failed to get raw inbox emails: ${error.message}`,
      'error',
      {
        service: 'graph-mail-service',
        method: 'getInboxRaw',
        requestedTop: options.top || options.limit || 10,
        executionTimeMs: executionTime,
        error: error.message,
        stack: error.stack,
        timestamp: new Date().toISOString()
      }
    );
    
    MonitoringService.logError(mcpError);
    
    // Pattern 4: User Error Tracking
    if (contextUserId) {
      MonitoringService.error('Failed to retrieve raw inbox emails', {
        error: error.message,
        requestedTop: options.top || options.limit || 10,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'mail', null, contextUserId);
    } else if (contextSessionId) {
      MonitoringService.error('Failed to retrieve raw inbox emails', {
        sessionId: contextSessionId,
        error: error.message,
        requestedTop: options.top || options.limit || 10,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'mail');
    }
    
    throw mcpError;
  }
}

/**
 * Retrieves detailed information for a specific email by ID.
 * @param {string} id - Email ID
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<object>}
 */
async function getEmailDetails(id, req, userId, sessionId) {
  const startTime = Date.now();
  
  // Extract user context from request if not provided
  const contextUserId = userId || req?.user?.userId;
  const contextSessionId = sessionId || req?.session?.id;
  
  // Pattern 1: Development Debug Logs
  if (process.env.NODE_ENV === 'development') {
    MonitoringService.debug('Mail getEmailDetails operation started', {
      method: 'getEmailDetails',
      emailId: id ? id.substring(0, 20) + '...' : null,
      sessionId: contextSessionId,
      userAgent: req?.get('User-Agent'),
      timestamp: new Date().toISOString()
    }, 'mail');
  }
  
  try {
    if (!id || typeof id !== 'string') {
      const mcpError = ErrorService.createError(
        'mail',
        'Email ID must be a non-empty string',
        'warning',
        {
          service: 'graph-mail-service',
          method: 'getEmailDetails',
          idType: typeof id,
          timestamp: new Date().toISOString()
        }
      );
      MonitoringService.logError(mcpError);
      throw mcpError;
    }
    
    const client = await graphClientFactory.createClient(req, contextUserId, contextSessionId);
    const message = await client.api(`/me/messages/${id}`, contextUserId, contextSessionId).get();
    
    if (!message) {
      const mcpError = ErrorService.createError(
        'mail',
        `No message found with ID: ${id}`,
        'warning',
        {
          service: 'graph-mail-service',
          method: 'getEmailDetails',
          emailId: id ? id.substring(0, 20) + '...' : null,
          timestamp: new Date().toISOString()
        }
      );
      MonitoringService.logError(mcpError);
      throw mcpError;
    }
    
    const emailDetails = {
      id: message.id,
      subject: message.subject,
      from: {
        name: message.from?.emailAddress?.name,
        email: message.from?.emailAddress?.address
      },
      to: message.toRecipients?.map(r => ({
        name: r.emailAddress?.name,
        email: r.emailAddress?.address
      })) || [],
      cc: message.ccRecipients?.map(r => ({
        name: r.emailAddress?.name,
        email: r.emailAddress?.address
      })) || [],
      bcc: message.bccRecipients?.map(r => ({
        name: r.emailAddress?.name,
        email: r.emailAddress?.address
      })) || [],
      body: message.body?.content,
      contentType: message.body?.contentType,
      received: message.receivedDateTime,
      sent: message.sentDateTime,
      isRead: message.isRead,
      importance: message.importance,
      hasAttachments: message.hasAttachments,
      categories: message.categories || []
    };
    
    const executionTime = Date.now() - startTime;
    
    // Pattern 2: User Activity Logs
    if (contextUserId) {
      MonitoringService.info('Retrieved email details successfully', {
        emailId: id ? id.substring(0, 20) + '...' : null,
        hasAttachments: emailDetails.hasAttachments,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'mail', null, contextUserId);
    } else if (contextSessionId) {
      MonitoringService.info('Retrieved email details with session', {
        sessionId: contextSessionId,
        emailId: id ? id.substring(0, 20) + '...' : null,
        hasAttachments: emailDetails.hasAttachments,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'mail');
    }
    
    return emailDetails;
  } catch (error) {
    const executionTime = Date.now() - startTime;
    
    // If it's already an MCP error, just rethrow
    if (error.category) {
      throw error;
    }
    
    // Pattern 3: Infrastructure Error Logging
    const mcpError = ErrorService.createError(
      'mail',
      `Failed to get email details: ${error.message}`,
      'error',
      {
        service: 'graph-mail-service',
        method: 'getEmailDetails',
        emailId: id ? id.substring(0, 20) + '...' : null,
        executionTimeMs: executionTime,
        error: error.message,
        stack: error.stack,
        timestamp: new Date().toISOString()
      }
    );
    
    MonitoringService.logError(mcpError);
    
    // Pattern 4: User Error Tracking
    if (contextUserId) {
      MonitoringService.error('Failed to retrieve email details', {
        error: error.message,
        emailId: id ? id.substring(0, 20) + '...' : null,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'mail', null, contextUserId);
    } else if (contextSessionId) {
      MonitoringService.error('Failed to retrieve email details', {
        sessionId: contextSessionId,
        error: error.message,
        emailId: id ? id.substring(0, 20) + '...' : null,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'mail');
    }
    
    throw mcpError;
  }
}

/**
 * Marks an email as read.
 * @param {string} id - Email ID
 * @param {boolean} isRead - Read status to set
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<boolean>}
 */
async function markAsRead(id, isRead = true, req, userId, sessionId) {
  const startTime = Date.now();
  
  // Extract user context from request if not provided
  const contextUserId = userId || req?.user?.userId;
  const contextSessionId = sessionId || req?.session?.id;
  
  // Pattern 1: Development Debug Logs
  if (process.env.NODE_ENV === 'development') {
    MonitoringService.debug('Mail markAsRead operation started', {
      method: 'markAsRead',
      emailId: id ? id.substring(0, 20) + '...' : null,
      isRead,
      sessionId: contextSessionId,
      userAgent: req?.get('User-Agent'),
      timestamp: new Date().toISOString()
    }, 'mail');
  }
  
  try {
    if (!id || typeof id !== 'string') {
      const mcpError = ErrorService.createError(
        'mail',
        'Email ID must be a non-empty string',
        'warning',
        {
          service: 'graph-mail-service',
          method: 'markAsRead',
          idType: typeof id,
          timestamp: new Date().toISOString()
        }
      );
      MonitoringService.logError(mcpError);
      throw mcpError;
    }
    
    const client = await graphClientFactory.createClient(req, contextUserId, contextSessionId);
    await client.api(`/me/messages/${id}`, contextUserId, contextSessionId).patch({
      isRead: isRead
    });
    
    const executionTime = Date.now() - startTime;
    
    // Pattern 2: User Activity Logs
    if (contextUserId) {
      MonitoringService.info('Email marked as read successfully', {
        emailId: id ? id.substring(0, 20) + '...' : null,
        isRead,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'mail', null, contextUserId);
    } else if (contextSessionId) {
      MonitoringService.info('Email marked as read with session', {
        sessionId: contextSessionId,
        emailId: id ? id.substring(0, 20) + '...' : null,
        isRead,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'mail');
    }
    
    return true;
  } catch (error) {
    const executionTime = Date.now() - startTime;
    
    // If it's already an MCP error, just rethrow
    if (error.category) {
      throw error;
    }
    
    // Pattern 3: Infrastructure Error Logging
    const mcpError = ErrorService.createError(
      'mail',
      `Failed to mark email as read: ${error.message}`,
      'error',
      {
        service: 'graph-mail-service',
        method: 'markAsRead',
        emailId: id ? id.substring(0, 20) + '...' : null,
        isRead,
        executionTimeMs: executionTime,
        error: error.message,
        stack: error.stack,
        timestamp: new Date().toISOString()
      }
    );
    
    MonitoringService.logError(mcpError);
    
    // Pattern 4: User Error Tracking
    if (contextUserId) {
      MonitoringService.error('Failed to mark email as read', {
        error: error.message,
        emailId: id ? id.substring(0, 20) + '...' : null,
        isRead,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'mail', null, contextUserId);
    } else if (contextSessionId) {
      MonitoringService.error('Failed to mark email as read', {
        sessionId: contextSessionId,
        error: error.message,
        emailId: id ? id.substring(0, 20) + '...' : null,
        isRead,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'mail');
    }
    
    throw mcpError;
  }
}

/**
 * Add an attachment to an existing email message.
 * @param {string} messageId - ID of the email message
 * @param {object} attachment - Attachment data
 * @param {string} attachment.name - Name of the attachment
 * @param {string} attachment.contentType - MIME type of the attachment
 * @param {string} attachment.contentBytes - Base64 encoded content
 * @param {boolean} [attachment.isInline=false] - Whether the attachment is inline
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<object>} Created attachment object
 */
async function addMailAttachment(messageId, attachment, req, userId, sessionId) {
  const startTime = Date.now();
  
  // Extract user context from request if not provided
  const contextUserId = userId || req?.user?.userId;
  const contextSessionId = sessionId || req?.session?.id;
  
  // Pattern 1: Development Debug Logs
  if (process.env.NODE_ENV === 'development') {
    MonitoringService.debug('Mail addMailAttachment operation started', {
      method: 'addMailAttachment',
      messageId: messageId ? messageId.substring(0, 20) + '...' : null,
      attachmentName: attachment?.name,
      sessionId: contextSessionId,
      userAgent: req?.get('User-Agent'),
      timestamp: new Date().toISOString()
    }, 'mail');
  }
  
  try {
    if (!messageId || typeof messageId !== 'string') {
      const mcpError = ErrorService.createError(
        'mail',
        'Message ID must be a non-empty string',
        'warning',
        {
          service: 'graph-mail-service',
          method: 'addMailAttachment',
          messageIdType: typeof messageId,
          timestamp: new Date().toISOString()
        }
      );
      MonitoringService.logError(mcpError);
      throw mcpError;
    }
    
    if (!attachment || !attachment.name || !attachment.contentBytes) {
      const mcpError = ErrorService.createError(
        'mail',
        'Attachment must have name and contentBytes',
        'warning',
        {
          service: 'graph-mail-service',
          method: 'addMailAttachment',
          hasName: !!attachment?.name,
          hasContentBytes: !!attachment?.contentBytes,
          timestamp: new Date().toISOString()
        }
      );
      MonitoringService.logError(mcpError);
      throw mcpError;
    }
    
    const client = await graphClientFactory.createClient(req);
    
    // Prepare the attachment object for Microsoft Graph API
    const attachmentData = {
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: attachment.name,
      contentBytes: attachment.contentBytes,
      contentType: attachment.contentType || 'application/octet-stream',
      isInline: attachment.isInline || false
    };
    
    MonitoringService.debug('Adding attachment to email', {
      messageId: messageId,
      attachmentName: attachment.name,
      contentType: attachmentData.contentType,
      isInline: attachmentData.isInline,
      timestamp: new Date().toISOString()
    }, 'graph-mail-service');
    
    // Add the attachment to the message
    const result = await client.api(`/me/messages/${messageId}/attachments`).post(attachmentData);
    
    const executionTime = Date.now() - startTime;
    MonitoringService.trackMetric('graph_mail_add_attachment_success', executionTime, {
      service: 'graph-mail-service',
      method: 'addMailAttachment',
      attachmentName: attachment.name,
      timestamp: new Date().toISOString()
    });
    
    MonitoringService.info('Successfully added attachment to email', {
      messageId: messageId,
      attachmentId: result.id,
      attachmentName: result.name,
      executionTime: executionTime,
      timestamp: new Date().toISOString()
    }, 'graph-mail-service');
    
    return result;
    
  } catch (error) {
    const executionTime = Date.now() - startTime;
    
    const mcpError = ErrorService.createError(
      ErrorService.CATEGORIES.GRAPH,
      `Failed to add attachment to email: ${error.message}`,
      ErrorService.SEVERITIES.ERROR,
      {
        service: 'graph-mail-service',
        method: 'addMailAttachment',
        messageId: messageId,
        attachmentName: attachment?.name,
        stack: error.stack,
        timestamp: new Date().toISOString()
      }
    );
    
    MonitoringService.logError(mcpError);
    MonitoringService.trackMetric('graph_mail_add_attachment_failure', executionTime, {
      service: 'graph-mail-service',
      method: 'addMailAttachment',
      errorType: error.code || 'unknown',
      timestamp: new Date().toISOString()
    });
    
    throw mcpError;
  }
}

/**
 * Remove an attachment from an existing email message.
 * @param {string} messageId - ID of the email message
 * @param {string} attachmentId - ID of the attachment to remove
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<object>} Success status
 */
async function removeMailAttachment(messageId, attachmentId, req, userId, sessionId) {
  const startTime = Date.now();
  
  // Extract user context from request if not provided
  const contextUserId = userId || req?.user?.userId;
  const contextSessionId = sessionId || req?.session?.id;
  
  // Pattern 1: Development Debug Logs
  if (process.env.NODE_ENV === 'development') {
    MonitoringService.debug('Mail removeMailAttachment operation started', {
      method: 'removeMailAttachment',
      messageId: messageId ? messageId.substring(0, 20) + '...' : null,
      attachmentId: attachmentId ? attachmentId.substring(0, 20) + '...' : null,
      sessionId: contextSessionId,
      userAgent: req?.get('User-Agent'),
      timestamp: new Date().toISOString()
    }, 'mail');
  }
  
  try {
    if (!messageId || typeof messageId !== 'string') {
      const mcpError = ErrorService.createError(
        'mail',
        'Message ID must be a non-empty string',
        'warning',
        {
          service: 'graph-mail-service',
          method: 'removeMailAttachment',
          messageIdType: typeof messageId,
          timestamp: new Date().toISOString()
        }
      );
      MonitoringService.logError(mcpError);
      throw mcpError;
    }
    
    if (!attachmentId || typeof attachmentId !== 'string') {
      const mcpError = ErrorService.createError(
        'mail',
        'Attachment ID must be a non-empty string',
        'warning',
        {
          service: 'graph-mail-service',
          method: 'removeMailAttachment',
          attachmentIdType: typeof attachmentId,
          timestamp: new Date().toISOString()
        }
      );
      MonitoringService.logError(mcpError);
      throw mcpError;
    }
    
    const client = await graphClientFactory.createClient(req);
    
    MonitoringService.info('Attempting to remove attachment from email', {
      messageId: messageId,
      attachmentId: attachmentId,
      encodedAttachmentId: encodeURIComponent(attachmentId),
      apiPath: `/me/messages/${messageId}/attachments/${encodeURIComponent(attachmentId)}`,
      timestamp: new Date().toISOString()
    }, 'graph-mail-service');
    
    // URL encode the attachment ID to handle special characters
    const encodedAttachmentId = encodeURIComponent(attachmentId);
    
    MonitoringService.debug('Attempting to remove attachment', {
      messageId: messageId,
      attachmentId: attachmentId,
      encodedAttachmentId: encodedAttachmentId,
      apiPath: `/me/messages/${messageId}/attachments/${encodedAttachmentId}`,
      timestamp: new Date().toISOString()
    }, 'graph-mail-service');
    
    // Remove the attachment from the message
    try {
      const deleteResponse = await client.api(`/me/messages/${messageId}/attachments/${encodedAttachmentId}`).delete();
      MonitoringService.info('Graph API delete response received', {
        messageId: messageId,
        attachmentId: attachmentId,
        response: deleteResponse,
        timestamp: new Date().toISOString()
      }, 'graph-mail-service');
    } catch (graphError) {
      MonitoringService.error('Graph API delete request failed', {
        messageId: messageId,
        attachmentId: attachmentId,
        encodedAttachmentId: encodedAttachmentId,
        apiPath: `/me/messages/${messageId}/attachments/${encodedAttachmentId}`,
        error: graphError.message,
        statusCode: graphError.statusCode || graphError.code,
        errorDetails: graphError.body || graphError.response || graphError,
        timestamp: new Date().toISOString()
      }, 'graph-mail-service');
      throw graphError;
    }
    
    const executionTime = Date.now() - startTime;
    MonitoringService.trackMetric('graph_mail_remove_attachment_success', executionTime, {
      service: 'graph-mail-service',
      method: 'removeMailAttachment',
      timestamp: new Date().toISOString()
    });
    
    MonitoringService.info('Successfully removed attachment from email', {
      messageId: messageId,
      attachmentId: attachmentId,
      executionTime: executionTime,
      timestamp: new Date().toISOString()
    }, 'graph-mail-service');
    
    return { success: true, messageId, attachmentId };
    
  } catch (error) {
    const executionTime = Date.now() - startTime;
    
    const mcpError = ErrorService.createError(
      ErrorService.CATEGORIES.GRAPH,
      `Failed to remove attachment from email: ${error.message}`,
      ErrorService.SEVERITIES.ERROR,
      {
        service: 'graph-mail-service',
        method: 'removeMailAttachment',
        messageId: messageId,
        attachmentId: attachmentId,
        encodedAttachmentId: encodeURIComponent(attachmentId),
        apiPath: `/me/messages/${messageId}/attachments/${encodeURIComponent(attachmentId)}`,
        graphError: error.code || 'unknown',
        graphMessage: error.message,
        stack: error.stack,
        timestamp: new Date().toISOString()
      }
    );
    
    MonitoringService.logError(mcpError);
    MonitoringService.trackMetric('graph_mail_remove_attachment_failure', executionTime, {
      service: 'graph-mail-service',
      method: 'removeMailAttachment',
      errorType: error.code || 'unknown',
      timestamp: new Date().toISOString()
    });
    
    throw mcpError;
  }
}

/**
 * Replies to an email message.
 * @param {string} messageId - The ID of the message to reply to
 * @param {object} replyData - Reply data
 * @param {string} replyData.body - The reply body content
 * @param {string} [replyData.contentType] - Content type ('Text' or 'HTML'), defaults to 'Text'
 * @param {boolean} [replyData.replyAll] - If true, reply to all recipients
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<object>} - Result of the reply operation
 */
async function replyToEmail(messageId, replyData, req, userId, sessionId) {
  const startTime = Date.now();

  // Extract user context from request if not provided
  const contextUserId = userId || req?.user?.userId;
  const contextSessionId = sessionId || req?.session?.id;

  // Pattern 1: Development Debug Logs
  if (process.env.NODE_ENV === 'development') {
    MonitoringService.debug('Mail replyToEmail operation started', {
      method: 'replyToEmail',
      messageId,
      replyAll: replyData?.replyAll || false,
      sessionId: contextSessionId,
      timestamp: new Date().toISOString()
    }, 'mail');
  }

  try {
    // Validate required fields
    if (!messageId) {
      const validationError = ErrorService.createError(
        ErrorService.CATEGORIES.VALIDATION,
        'Message ID is required for reply',
        ErrorService.SEVERITIES.WARNING,
        { messageId },
        null,
        contextUserId
      );
      throw validationError;
    }

    if (!replyData?.body) {
      const validationError = ErrorService.createError(
        ErrorService.CATEGORIES.VALIDATION,
        'Reply body is required',
        ErrorService.SEVERITIES.WARNING,
        { messageId },
        null,
        contextUserId
      );
      throw validationError;
    }

    const client = await graphClientFactory.createClient(req, contextUserId, contextSessionId);

    // Determine endpoint based on replyAll flag
    const endpoint = replyData.replyAll
      ? `/me/messages/${messageId}/replyAll`
      : `/me/messages/${messageId}/reply`;

    // Build the reply payload
    const payload = {
      message: {
        body: {
          contentType: replyData.contentType || 'Text',
          content: replyData.body
        }
      }
    };

    // Make the API call
    await client.api(endpoint, contextUserId, contextSessionId).post(payload);

    const executionTime = Date.now() - startTime;

    // Pattern 2: User Activity Logs
    if (contextUserId) {
      MonitoringService.info('Email reply sent successfully', {
        messageId,
        replyAll: replyData.replyAll || false,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'mail', null, contextUserId);
    }

    return {
      success: true,
      messageId,
      replyAll: replyData.replyAll || false
    };

  } catch (error) {
    const executionTime = Date.now() - startTime;

    // Pattern 4: User Error Tracking
    const mcpError = ErrorService.createError(
      ErrorService.CATEGORIES.API,
      `Failed to reply to email: ${error.message}`,
      ErrorService.SEVERITIES.ERROR,
      {
        messageId,
        originalError: error.message,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      },
      error,
      contextUserId
    );

    throw mcpError;
  }
}

module.exports = {
  getInbox,
  searchEmails,
  sendEmail,
  replyToEmail,
  flagEmail,
  getAttachments,
  getInboxRaw,
  getEmailDetails,
  markAsRead,
  addMailAttachment,
  removeMailAttachment
};
