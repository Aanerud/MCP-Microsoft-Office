/**
 * @fileoverview Contacts Controller - Handles API requests for Microsoft Contacts API.
 * Follows MCP modular, testable, and consistent API contract rules.
 */

const Joi = require('joi');
const ErrorService = require('../../core/error-service.cjs');
const MonitoringService = require('../../core/monitoring-service.cjs');

/**
 * Joi validation schemas for contacts endpoints
 */
const schemas = {
  listContacts: Joi.object({
    limit: Joi.number().integer().min(1).max(100).optional(),
    orderby: Joi.string().optional(),
    filter: Joi.string().optional()
  }),

  getContact: Joi.object({
    contactId: Joi.string().required()
  }),

  createContact: Joi.object({
    givenName: Joi.string().optional(),
    surname: Joi.string().optional(),
    displayName: Joi.string().optional(),
    emailAddresses: Joi.array().items(Joi.object({
      address: Joi.string().email().required(),
      name: Joi.string().optional()
    })).optional(),
    businessPhones: Joi.array().items(Joi.string()).optional(),
    homePhones: Joi.array().items(Joi.string()).optional(),
    mobilePhone: Joi.string().optional(),
    jobTitle: Joi.string().optional(),
    companyName: Joi.string().optional(),
    department: Joi.string().optional(),
    officeLocation: Joi.string().optional(),
    businessAddress: Joi.object({
      street: Joi.string().optional(),
      city: Joi.string().optional(),
      state: Joi.string().optional(),
      postalCode: Joi.string().optional(),
      countryOrRegion: Joi.string().optional()
    }).optional(),
    homeAddress: Joi.object({
      street: Joi.string().optional(),
      city: Joi.string().optional(),
      state: Joi.string().optional(),
      postalCode: Joi.string().optional(),
      countryOrRegion: Joi.string().optional()
    }).optional(),
    birthday: Joi.string().isoDate().optional(),
    personalNotes: Joi.string().optional(),
    categories: Joi.array().items(Joi.string()).optional()
  }),

  updateContact: Joi.object({
    contactId: Joi.string().required(),
    givenName: Joi.string().optional(),
    surname: Joi.string().optional(),
    displayName: Joi.string().optional(),
    emailAddresses: Joi.array().items(Joi.object({
      address: Joi.string().email().required(),
      name: Joi.string().optional()
    })).optional(),
    businessPhones: Joi.array().items(Joi.string()).optional(),
    homePhones: Joi.array().items(Joi.string()).optional(),
    mobilePhone: Joi.string().optional(),
    jobTitle: Joi.string().optional(),
    companyName: Joi.string().optional(),
    department: Joi.string().optional(),
    officeLocation: Joi.string().optional(),
    businessAddress: Joi.object({
      street: Joi.string().optional(),
      city: Joi.string().optional(),
      state: Joi.string().optional(),
      postalCode: Joi.string().optional(),
      countryOrRegion: Joi.string().optional()
    }).optional(),
    homeAddress: Joi.object({
      street: Joi.string().optional(),
      city: Joi.string().optional(),
      state: Joi.string().optional(),
      postalCode: Joi.string().optional(),
      countryOrRegion: Joi.string().optional()
    }).optional(),
    birthday: Joi.string().isoDate().optional(),
    personalNotes: Joi.string().optional(),
    categories: Joi.array().items(Joi.string()).optional()
  }),

  deleteContact: Joi.object({
    contactId: Joi.string().required()
  }),

  searchContacts: Joi.object({
    query: Joi.string().required(),
    limit: Joi.number().integer().min(1).max(100).optional()
  })
};

/**
 * Creates a contacts controller with injected dependencies.
 * @param {object} deps - Controller dependencies
 * @param {object} deps.contactsModule - Initialized contacts module
 * @returns {object} Controller methods
 */
function createContactsController({ contactsModule }) {
  if (!contactsModule) {
    throw new Error('Contacts module is required for ContactsController');
  }

  return {
    /**
     * List all contacts for the current user.
     */
    async listContacts(req, res) {
      const { userId = null } = req.user || {};
      const sessionId = req.session?.id;
      const startTime = Date.now();

      try {
        const { error, value } = schemas.listContacts.validate(req.query);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const options = {
          top: value.limit || 50,
          orderby: value.orderby,
          filter: value.filter
        };

        const contacts = await contactsModule.listContacts(options, req, userId, sessionId);

        MonitoringService.trackMetric('contacts.listContacts.duration', Date.now() - startTime, {
          count: contacts.length,
          success: true,
          userId
        });

        res.json({ success: true, data: contacts, count: contacts.length });
      } catch (error) {
        MonitoringService.logError(ErrorService.createError('contacts', 'Failed to list contacts', 'error', {
          error: error.message,
          userId,
          timestamp: new Date().toISOString()
        }));

        res.status(500).json({
          error: 'CONTACTS_OPERATION_FAILED',
          error_description: 'Failed to list contacts'
        });
      }
    },

    /**
     * Get a specific contact by ID.
     */
    async getContact(req, res) {
      const { userId = null } = req.user || {};
      const sessionId = req.session?.id;
      const startTime = Date.now();

      try {
        const { error, value } = schemas.getContact.validate(req.params);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const contact = await contactsModule.getContact(value.contactId, req, userId, sessionId);

        MonitoringService.trackMetric('contacts.getContact.duration', Date.now() - startTime, {
          success: true,
          userId
        });

        res.json({ success: true, data: contact });
      } catch (error) {
        MonitoringService.logError(ErrorService.createError('contacts', 'Failed to get contact', 'error', {
          contactId: req.params.contactId,
          error: error.message,
          userId,
          timestamp: new Date().toISOString()
        }));

        res.status(500).json({
          error: 'CONTACTS_OPERATION_FAILED',
          error_description: 'Failed to get contact'
        });
      }
    },

    /**
     * Create a new contact.
     */
    async createContact(req, res) {
      const { userId = null } = req.user || {};
      const sessionId = req.session?.id;
      const startTime = Date.now();

      try {
        const { error, value } = schemas.createContact.validate(req.body);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const contact = await contactsModule.createContact(value, req, userId, sessionId);

        MonitoringService.trackMetric('contacts.createContact.duration', Date.now() - startTime, {
          success: true,
          userId
        });

        res.status(201).json({ success: true, data: contact, created: true });
      } catch (error) {
        MonitoringService.logError(ErrorService.createError('contacts', 'Failed to create contact', 'error', {
          error: error.message,
          userId,
          timestamp: new Date().toISOString()
        }));

        res.status(500).json({
          error: 'CONTACTS_OPERATION_FAILED',
          error_description: 'Failed to create contact'
        });
      }
    },

    /**
     * Update a contact.
     */
    async updateContact(req, res) {
      const { userId = null } = req.user || {};
      const sessionId = req.session?.id;
      const startTime = Date.now();

      try {
        const combined = { contactId: req.params.contactId, ...req.body };
        const { error, value } = schemas.updateContact.validate(combined);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const { contactId, ...updates } = value;
        const contact = await contactsModule.updateContact(contactId, updates, req, userId, sessionId);

        MonitoringService.trackMetric('contacts.updateContact.duration', Date.now() - startTime, {
          success: true,
          userId
        });

        res.json({ success: true, data: contact, updated: true });
      } catch (error) {
        MonitoringService.logError(ErrorService.createError('contacts', 'Failed to update contact', 'error', {
          contactId: req.params.contactId,
          error: error.message,
          userId,
          timestamp: new Date().toISOString()
        }));

        res.status(500).json({
          error: 'CONTACTS_OPERATION_FAILED',
          error_description: 'Failed to update contact'
        });
      }
    },

    /**
     * Delete a contact.
     */
    async deleteContact(req, res) {
      const { userId = null } = req.user || {};
      const sessionId = req.session?.id;
      const startTime = Date.now();

      try {
        const { error, value } = schemas.deleteContact.validate(req.params);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        await contactsModule.deleteContact(value.contactId, req, userId, sessionId);

        MonitoringService.trackMetric('contacts.deleteContact.duration', Date.now() - startTime, {
          success: true,
          userId
        });

        res.json({ success: true, deleted: true, contactId: value.contactId });
      } catch (error) {
        MonitoringService.logError(ErrorService.createError('contacts', 'Failed to delete contact', 'error', {
          contactId: req.params.contactId,
          error: error.message,
          userId,
          timestamp: new Date().toISOString()
        }));

        res.status(500).json({
          error: 'CONTACTS_OPERATION_FAILED',
          error_description: 'Failed to delete contact'
        });
      }
    },

    /**
     * Search contacts.
     */
    async searchContacts(req, res) {
      const { userId = null } = req.user || {};
      const sessionId = req.session?.id;
      const startTime = Date.now();

      try {
        const { error, value } = schemas.searchContacts.validate(req.query);
        if (error) {
          return res.status(400).json({ error: 'Invalid request', details: error.details });
        }

        const options = { top: value.limit || 25 };
        const contacts = await contactsModule.searchContacts(value.query, options, req, userId, sessionId);

        MonitoringService.trackMetric('contacts.searchContacts.duration', Date.now() - startTime, {
          count: contacts.length,
          success: true,
          userId
        });

        res.json({ success: true, data: contacts, count: contacts.length, query: value.query });
      } catch (error) {
        MonitoringService.logError(ErrorService.createError('contacts', 'Failed to search contacts', 'error', {
          query: req.query.query,
          error: error.message,
          userId,
          timestamp: new Date().toISOString()
        }));

        res.status(500).json({
          error: 'CONTACTS_OPERATION_FAILED',
          error_description: 'Failed to search contacts'
        });
      }
    }
  };
}

module.exports = createContactsController;
