/**
 * @fileoverview MCP Contacts Module - Handles Microsoft Contacts intents and actions for MCP.
 * Exposes: id, name, capabilities, init, handleIntent. Aligned with MCP module system.
 */

const MonitoringService = require('../../core/monitoring-service.cjs');
const ErrorService = require('../../core/error-service.cjs');

const CONTACTS_CAPABILITIES = [
  'listContacts',
  'getContact',
  'createContact',
  'updateContact',
  'deleteContact',
  'searchContacts'
];

// Log module initialization
MonitoringService.info('Contacts Module initialized', {
  serviceName: 'contacts-module',
  capabilities: CONTACTS_CAPABILITIES.length,
  timestamp: new Date().toISOString()
}, 'contacts');

const ContactsModule = {
  /**
   * Initializes the Contacts module with dependencies.
   * @param {object} services - { contactsService, cacheService, errorService, monitoringService }
   * @returns {object} Initialized module
   */
  init(services = {}) {
    const { contactsService, cacheService, errorService = ErrorService, monitoringService = MonitoringService } = services;

    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Initializing Contacts Module', {
        hasContactsService: !!contactsService,
        hasCacheService: !!cacheService,
        timestamp: new Date().toISOString()
      }, 'contacts');
    }

    if (!contactsService) {
      const mcpError = ErrorService.createError(
        'contacts',
        "ContactsModule init failed: Required service 'contactsService' is missing",
        'error',
        { missingService: 'contactsService', timestamp: new Date().toISOString() }
      );
      MonitoringService.logError(mcpError);
      throw mcpError;
    }

    this.services = { contactsService, cacheService, errorService, monitoringService };
    return this;
  },

  /**
   * List all contacts
   */
  async listContacts(options = {}, req, userId, sessionId) {
    const { contactsService } = this.services || {};
    const contextUserId = userId || req?.user?.userId;
    const contextSessionId = sessionId || req?.session?.id;

    if (!contactsService) {
      throw ErrorService.createError('contacts', 'ContactsService not available', 'error', {});
    }

    return await contactsService.listContacts(options, req, contextUserId, contextSessionId);
  },

  /**
   * Get a specific contact
   */
  async getContact(contactId, req, userId, sessionId) {
    const { contactsService } = this.services || {};
    const contextUserId = userId || req?.user?.userId;
    const contextSessionId = sessionId || req?.session?.id;

    if (!contactsService) {
      throw ErrorService.createError('contacts', 'ContactsService not available', 'error', {});
    }

    return await contactsService.getContact(contactId, req, contextUserId, contextSessionId);
  },

  /**
   * Create a new contact
   */
  async createContact(contactData, req, userId, sessionId) {
    const { contactsService } = this.services || {};
    const contextUserId = userId || req?.user?.userId;
    const contextSessionId = sessionId || req?.session?.id;

    if (!contactsService) {
      throw ErrorService.createError('contacts', 'ContactsService not available', 'error', {});
    }

    return await contactsService.createContact(contactData, req, contextUserId, contextSessionId);
  },

  /**
   * Update a contact
   */
  async updateContact(contactId, updates, req, userId, sessionId) {
    const { contactsService } = this.services || {};
    const contextUserId = userId || req?.user?.userId;
    const contextSessionId = sessionId || req?.session?.id;

    if (!contactsService) {
      throw ErrorService.createError('contacts', 'ContactsService not available', 'error', {});
    }

    return await contactsService.updateContact(contactId, updates, req, contextUserId, contextSessionId);
  },

  /**
   * Delete a contact
   */
  async deleteContact(contactId, req, userId, sessionId) {
    const { contactsService } = this.services || {};
    const contextUserId = userId || req?.user?.userId;
    const contextSessionId = sessionId || req?.session?.id;

    if (!contactsService) {
      throw ErrorService.createError('contacts', 'ContactsService not available', 'error', {});
    }

    return await contactsService.deleteContact(contactId, req, contextUserId, contextSessionId);
  },

  /**
   * Search contacts
   */
  async searchContacts(query, options = {}, req, userId, sessionId) {
    const { contactsService } = this.services || {};
    const contextUserId = userId || req?.user?.userId;
    const contextSessionId = sessionId || req?.session?.id;

    if (!contactsService) {
      throw ErrorService.createError('contacts', 'ContactsService not available', 'error', {});
    }

    return await contactsService.searchContacts(query, options, req, contextUserId, contextSessionId);
  },

  /**
   * Handles Contacts related intents routed to this module.
   */
  async handleIntent(intent, entities = {}, context = {}) {
    const startTime = Date.now();
    const contextUserId = context?.req?.user?.userId;
    const contextSessionId = context?.req?.session?.id;

    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Handling Contacts intent', {
        sessionId: contextSessionId,
        intent,
        entities: Object.keys(entities),
        timestamp: new Date().toISOString()
      }, 'contacts');
    }

    try {
      let result;

      switch (intent) {
        case 'listContacts': {
          const contacts = await this.listContacts(entities.options || {}, context.req, contextUserId, contextSessionId);
          result = { type: 'contactCollection', items: contacts };
          break;
        }

        case 'getContact': {
          if (!entities.contactId) {
            throw ErrorService.createError('contacts', 'contactId is required', 'warn', { intent });
          }
          const contact = await this.getContact(entities.contactId, context.req, contextUserId, contextSessionId);
          result = { type: 'contact', data: contact };
          break;
        }

        case 'createContact': {
          const contactData = {
            givenName: entities.givenName,
            surname: entities.surname,
            displayName: entities.displayName,
            emailAddresses: entities.emailAddresses,
            businessPhones: entities.businessPhones,
            mobilePhone: entities.mobilePhone,
            jobTitle: entities.jobTitle,
            companyName: entities.companyName,
            department: entities.department,
            officeLocation: entities.officeLocation,
            businessAddress: entities.businessAddress,
            homeAddress: entities.homeAddress,
            birthday: entities.birthday,
            personalNotes: entities.personalNotes,
            categories: entities.categories
          };
          const contact = await this.createContact(contactData, context.req, contextUserId, contextSessionId);
          result = { type: 'contact', data: contact, created: true };
          break;
        }

        case 'updateContact': {
          if (!entities.contactId) {
            throw ErrorService.createError('contacts', 'contactId is required', 'warn', { intent });
          }
          const contact = await this.updateContact(entities.contactId, entities.updates || {}, context.req, contextUserId, contextSessionId);
          result = { type: 'contact', data: contact, updated: true };
          break;
        }

        case 'deleteContact': {
          if (!entities.contactId) {
            throw ErrorService.createError('contacts', 'contactId is required', 'warn', { intent });
          }
          await this.deleteContact(entities.contactId, context.req, contextUserId, contextSessionId);
          result = { type: 'contact', deleted: true, contactId: entities.contactId };
          break;
        }

        case 'searchContacts': {
          if (!entities.query) {
            throw ErrorService.createError('contacts', 'query is required', 'warn', { intent });
          }
          const contacts = await this.searchContacts(entities.query, entities.options || {}, context.req, contextUserId, contextSessionId);
          result = { type: 'contactCollection', items: contacts, query: entities.query };
          break;
        }

        default:
          throw ErrorService.createError('contacts', `ContactsModule cannot handle intent: ${intent}`, 'warn', { intent });
      }

      const elapsedTime = Date.now() - startTime;

      if (contextUserId) {
        MonitoringService.info('Successfully handled Contacts intent', {
          intent,
          elapsedTime,
          timestamp: new Date().toISOString()
        }, 'contacts', null, contextUserId);
      }

      return result;
    } catch (error) {
      const elapsedTime = Date.now() - startTime;

      if (error.category && error.severity) {
        throw error;
      }

      const mcpError = ErrorService.createError(
        'contacts',
        `Error handling Contacts intent '${intent}': ${error.message}`,
        'error',
        { intent, originalError: error.stack, elapsedTime, timestamp: new Date().toISOString() }
      );
      MonitoringService.logError(mcpError);
      throw mcpError;
    }
  },

  id: 'contacts',
  name: 'Microsoft Contacts',
  capabilities: CONTACTS_CAPABILITIES
};

module.exports = ContactsModule;
