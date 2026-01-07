/**
 * @fileoverview ContactsService - Microsoft Graph Contacts API operations.
 * All methods are async, modular, and use GraphClient for requests.
 * Follows project error handling, validation, and normalization rules.
 */

const graphClientFactory = require('./graph-client.cjs');
const MonitoringService = require('../core/monitoring-service.cjs');
const ErrorService = require('../core/error-service.cjs');

/**
 * Normalizes a contact object from Graph API response.
 * @param {object} contact - Raw contact object from Graph API
 * @returns {object} Normalized contact object
 */
function normalizeContact(contact) {
  return {
    id: contact.id,
    displayName: contact.displayName,
    givenName: contact.givenName,
    surname: contact.surname,
    emailAddresses: (contact.emailAddresses || []).map(e => ({
      address: e.address,
      name: e.name
    })),
    businessPhones: contact.businessPhones || [],
    mobilePhone: contact.mobilePhone,
    homePhones: contact.homePhones || [],
    jobTitle: contact.jobTitle,
    companyName: contact.companyName,
    department: contact.department,
    officeLocation: contact.officeLocation,
    businessAddress: contact.businessAddress ? {
      street: contact.businessAddress.street,
      city: contact.businessAddress.city,
      state: contact.businessAddress.state,
      postalCode: contact.businessAddress.postalCode,
      countryOrRegion: contact.businessAddress.countryOrRegion
    } : null,
    homeAddress: contact.homeAddress ? {
      street: contact.homeAddress.street,
      city: contact.homeAddress.city,
      state: contact.homeAddress.state,
      postalCode: contact.homeAddress.postalCode,
      countryOrRegion: contact.homeAddress.countryOrRegion
    } : null,
    birthday: contact.birthday,
    personalNotes: contact.personalNotes,
    categories: contact.categories || [],
    createdDateTime: contact.createdDateTime,
    lastModifiedDateTime: contact.lastModifiedDateTime
  };
}

/**
 * Gets all contacts for the current user.
 * @param {object} options - Query options
 * @param {number} [options.top=50] - Number of contacts to retrieve
 * @param {string} [options.orderby] - Order by field (e.g., 'displayName')
 * @param {string} [options.filter] - OData filter
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<Array<object>>} Normalized contact objects
 */
async function listContacts(options = {}, req, userId, sessionId) {
  const startTime = new Date();

  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;

  try {
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Listing contacts', {
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString(),
        options: { top: options.top || 50 }
      }, 'contacts');
    }

    const client = await graphClientFactory.createClient(req, resolvedUserId, resolvedSessionId);
    const top = options.top || 50;

    let queryParams = [`$top=${top}`];
    if (options.orderby) {
      queryParams.push(`$orderby=${encodeURIComponent(options.orderby)}`);
    }
    if (options.filter) {
      queryParams.push(`$filter=${encodeURIComponent(options.filter)}`);
    }

    const url = `/me/contacts?${queryParams.join('&')}`;
    const res = await client.api(url, resolvedUserId, resolvedSessionId).get();

    const endTime = new Date();
    const duration = endTime - startTime;

    const contacts = (res.value || []).map(normalizeContact);

    if (resolvedUserId) {
      MonitoringService.info('Retrieved contacts successfully', {
        count: contacts.length,
        duration: `${duration}ms`,
        timestamp: new Date().toISOString()
      }, 'contacts', null, resolvedUserId);
    }

    return contacts;
  } catch (error) {
    const endTime = new Date();
    const duration = endTime - startTime;

    const mcpError = ErrorService.createError(
      'contacts',
      'Failed to list contacts',
      'error',
      {
        endpoint: '/me/contacts',
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
 * Gets a specific contact by ID.
 * @param {string} contactId - Contact ID
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<object>} Normalized contact object
 */
async function getContact(contactId, req, userId, sessionId) {
  const startTime = new Date();

  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;

  try {
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Getting contact', {
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString(),
        contactId
      }, 'contacts');
    }

    const client = await graphClientFactory.createClient(req, resolvedUserId, resolvedSessionId);

    const url = `/me/contacts/${contactId}`;
    const res = await client.api(url, resolvedUserId, resolvedSessionId).get();

    const endTime = new Date();
    const duration = endTime - startTime;

    if (resolvedUserId) {
      MonitoringService.info('Retrieved contact successfully', {
        contactId,
        duration: `${duration}ms`,
        timestamp: new Date().toISOString()
      }, 'contacts', null, resolvedUserId);
    }

    return normalizeContact(res);
  } catch (error) {
    const endTime = new Date();
    const duration = endTime - startTime;

    const mcpError = ErrorService.createError(
      'contacts',
      'Failed to get contact',
      'error',
      {
        endpoint: `/me/contacts/${contactId}`,
        contactId,
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
 * Creates a new contact.
 * @param {object} contactData - Contact data
 * @param {string} [contactData.givenName] - First name
 * @param {string} [contactData.surname] - Last name
 * @param {string} [contactData.displayName] - Display name
 * @param {Array} [contactData.emailAddresses] - Email addresses
 * @param {Array} [contactData.businessPhones] - Business phone numbers
 * @param {string} [contactData.mobilePhone] - Mobile phone number
 * @param {string} [contactData.jobTitle] - Job title
 * @param {string} [contactData.companyName] - Company name
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<object>} Created contact object
 */
async function createContact(contactData, req, userId, sessionId) {
  const startTime = new Date();

  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;

  try {
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Creating contact', {
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString(),
        displayName: contactData.displayName || `${contactData.givenName} ${contactData.surname}`
      }, 'contacts');
    }

    const client = await graphClientFactory.createClient(req, resolvedUserId, resolvedSessionId);

    // Build contact object
    const contact = {};
    if (contactData.givenName) contact.givenName = contactData.givenName;
    if (contactData.surname) contact.surname = contactData.surname;
    if (contactData.displayName) contact.displayName = contactData.displayName;
    if (contactData.emailAddresses) contact.emailAddresses = contactData.emailAddresses;
    if (contactData.businessPhones) contact.businessPhones = contactData.businessPhones;
    if (contactData.homePhones) contact.homePhones = contactData.homePhones;
    if (contactData.mobilePhone) contact.mobilePhone = contactData.mobilePhone;
    if (contactData.jobTitle) contact.jobTitle = contactData.jobTitle;
    if (contactData.companyName) contact.companyName = contactData.companyName;
    if (contactData.department) contact.department = contactData.department;
    if (contactData.officeLocation) contact.officeLocation = contactData.officeLocation;
    if (contactData.businessAddress) contact.businessAddress = contactData.businessAddress;
    if (contactData.homeAddress) contact.homeAddress = contactData.homeAddress;
    if (contactData.birthday) contact.birthday = contactData.birthday;
    if (contactData.personalNotes) contact.personalNotes = contactData.personalNotes;
    if (contactData.categories) contact.categories = contactData.categories;

    const url = '/me/contacts';
    const res = await client.api(url, resolvedUserId, resolvedSessionId).post(contact);

    const endTime = new Date();
    const duration = endTime - startTime;

    if (resolvedUserId) {
      MonitoringService.info('Created contact successfully', {
        contactId: res.id,
        duration: `${duration}ms`,
        timestamp: new Date().toISOString()
      }, 'contacts', null, resolvedUserId);
    }

    return normalizeContact(res);
  } catch (error) {
    const endTime = new Date();
    const duration = endTime - startTime;

    const mcpError = ErrorService.createError(
      'contacts',
      'Failed to create contact',
      'error',
      {
        endpoint: '/me/contacts',
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
 * Updates a contact.
 * @param {string} contactId - Contact ID
 * @param {object} updates - Update data
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<object>} Updated contact object
 */
async function updateContact(contactId, updates, req, userId, sessionId) {
  const startTime = new Date();

  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;

  try {
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Updating contact', {
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString(),
        contactId,
        updates: Object.keys(updates)
      }, 'contacts');
    }

    const client = await graphClientFactory.createClient(req, resolvedUserId, resolvedSessionId);

    const url = `/me/contacts/${contactId}`;
    const res = await client.api(url, resolvedUserId, resolvedSessionId).patch(updates);

    const endTime = new Date();
    const duration = endTime - startTime;

    if (resolvedUserId) {
      MonitoringService.info('Updated contact successfully', {
        contactId,
        duration: `${duration}ms`,
        timestamp: new Date().toISOString()
      }, 'contacts', null, resolvedUserId);
    }

    return normalizeContact(res);
  } catch (error) {
    const endTime = new Date();
    const duration = endTime - startTime;

    const mcpError = ErrorService.createError(
      'contacts',
      'Failed to update contact',
      'error',
      {
        endpoint: `/me/contacts/${contactId}`,
        contactId,
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
 * Deletes a contact.
 * @param {string} contactId - Contact ID
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<boolean>} True if deleted successfully
 */
async function deleteContact(contactId, req, userId, sessionId) {
  const startTime = new Date();

  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;

  try {
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Deleting contact', {
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString(),
        contactId
      }, 'contacts');
    }

    const client = await graphClientFactory.createClient(req, resolvedUserId, resolvedSessionId);

    const url = `/me/contacts/${contactId}`;
    await client.api(url, resolvedUserId, resolvedSessionId).delete();

    const endTime = new Date();
    const duration = endTime - startTime;

    if (resolvedUserId) {
      MonitoringService.info('Deleted contact successfully', {
        contactId,
        duration: `${duration}ms`,
        timestamp: new Date().toISOString()
      }, 'contacts', null, resolvedUserId);
    }

    return true;
  } catch (error) {
    const endTime = new Date();
    const duration = endTime - startTime;

    const mcpError = ErrorService.createError(
      'contacts',
      'Failed to delete contact',
      'error',
      {
        endpoint: `/me/contacts/${contactId}`,
        contactId,
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
 * Searches contacts by name or email.
 * @param {string} searchQuery - Search query string
 * @param {object} options - Query options
 * @param {number} [options.top=25] - Number of contacts to retrieve
 * @param {object} req - Express request object
 * @param {string} userId - User ID for logging context
 * @param {string} sessionId - Session ID for logging context
 * @returns {Promise<Array<object>>} Normalized contact objects
 */
async function searchContacts(searchQuery, options = {}, req, userId, sessionId) {
  const startTime = new Date();

  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;

  try {
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Searching contacts', {
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString(),
        query: searchQuery?.substring(0, 50)
      }, 'contacts');
    }

    const client = await graphClientFactory.createClient(req, resolvedUserId, resolvedSessionId);
    const top = options.top || 25;

    // Use $search for contacts search
    const url = `/me/contacts?$search="${encodeURIComponent(searchQuery)}"&$top=${top}`;
    const res = await client.api(url, resolvedUserId, resolvedSessionId).get();

    const endTime = new Date();
    const duration = endTime - startTime;

    const contacts = (res.value || []).map(normalizeContact);

    if (resolvedUserId) {
      MonitoringService.info('Searched contacts successfully', {
        query: searchQuery?.substring(0, 50),
        count: contacts.length,
        duration: `${duration}ms`,
        timestamp: new Date().toISOString()
      }, 'contacts', null, resolvedUserId);
    }

    return contacts;
  } catch (error) {
    const endTime = new Date();
    const duration = endTime - startTime;

    const mcpError = ErrorService.createError(
      'contacts',
      'Failed to search contacts',
      'error',
      {
        endpoint: '/me/contacts',
        searchQuery: searchQuery?.substring(0, 50),
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
  listContacts,
  getContact,
  createContact,
  updateContact,
  deleteContact,
  searchContacts
};
