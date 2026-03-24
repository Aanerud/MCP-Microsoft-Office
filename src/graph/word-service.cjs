/**
 * @fileoverview WordService - Microsoft Graph Word document operations.
 * Uses docx for creation, mammoth for reading, and JSZip for metadata extraction.
 * Follows project error handling, validation, and normalization rules.
 */

const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, AlignmentType, ImageRun } = require('docx');
const mammoth = require('mammoth');
const JSZip = require('jszip');
const { parseStringPromise } = require('xml2js');
const graphClientFactory = require('./graph-client.cjs');
const filesService = require('./files-service.cjs');
const ErrorService = require('../core/error-service.cjs');
const MonitoringService = require('../core/monitoring-service.cjs');

const MAX_FILE_SIZE = 25 * 1024 * 1024; // 25MB

/**
 * Builds docx Paragraph objects from a content section.
 * @param {Object} section - Section definition with type and properties
 * @returns {Paragraph[]} Array of docx Paragraph objects
 * @private
 */
function buildSectionParagraphs(section) {
  switch (section.type) {
    case 'heading':
      return [new Paragraph({
        text: section.text,
        heading: HeadingLevel[`HEADING_${section.level || 1}`]
      })];

    case 'paragraph':
      return [new Paragraph({
        children: [new TextRun(section.text)]
      })];

    case 'table': {
      const rows = [];
      if (section.headers && section.headers.length > 0) {
        rows.push(new TableRow({
          children: section.headers.map(header => new TableCell({
            children: [new Paragraph({ children: [new TextRun({ text: header, bold: true })] })]
          }))
        }));
      }
      if (section.rows) {
        for (const row of section.rows) {
          rows.push(new TableRow({
            children: row.map(cell => new TableCell({
              children: [new Paragraph({ children: [new TextRun(String(cell))] })]
            }))
          }));
        }
      }
      return [new Table({ rows })];
    }

    case 'list':
      return (section.items || []).map(item => {
        if (section.ordered) {
          return new Paragraph({
            text: item,
            numbering: { reference: 'default-numbering', level: 0 }
          });
        }
        return new Paragraph({
          text: item,
          bullet: { level: 0 }
        });
      });

    case 'image':
      return [new Paragraph({
        children: [new ImageRun({
          data: Buffer.from(section.data, 'base64'),
          transformation: {
            width: section.width || 300,
            height: section.height || 200
          }
        })]
      })];

    default:
      return [];
  }
}

/**
 * Creates a Word document from structured content and uploads it to OneDrive.
 * @param {string} fileName - Name for the document (should end in .docx)
 * @param {Object} content - Document content with sections array
 * @param {string} [folderId] - Target folder ID (defaults to root)
 * @param {Object} req - Express request object
 * @param {string} [userId] - User ID for logging context
 * @param {string} [sessionId] - Session ID for logging context
 * @returns {Promise<Object>} Uploaded file metadata from Graph API
 */
async function createWordDocument(fileName, content, folderId, req, userId, sessionId) {
  const startTime = Date.now();
  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;

  try {
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Word document creation requested', {
        fileName,
        folderId: folderId || 'root',
        sectionCount: content?.sections?.length || 0,
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString()
      }, 'word');
    }

    const paragraphs = [];
    for (const section of (content.sections || [])) {
      paragraphs.push(...buildSectionParagraphs(section));
    }

    const doc = new Document({
      sections: [{
        children: paragraphs
      }]
    });

    const buffer = await Packer.toBuffer(doc);

    // Use filesService.uploadFile which handles binary Buffers correctly
    const result = await filesService.uploadFile(fileName, buffer, req, userId, sessionId);

    const executionTime = Date.now() - startTime;

    if (resolvedUserId) {
      MonitoringService.info('Word document created successfully', {
        fileName,
        fileId: result.id,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'word', null, resolvedUserId);
    }

    MonitoringService.trackMetric('word_create_success', executionTime, {
      fileName,
      sectionCount: content?.sections?.length || 0,
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });

    return result;
  } catch (error) {
    const executionTime = Date.now() - startTime;
    const mcpError = ErrorService.createError(
      'word',
      `Failed to create Word document: ${error.message}`,
      'error',
      {
        fileName,
        folderId: folderId || 'root',
        error: error.message,
        stack: error.stack,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }
    );
    MonitoringService.logError(mcpError);

    if (resolvedUserId) {
      MonitoringService.error('Word document creation failed', {
        error: error.message,
        fileName,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'word', null, resolvedUserId);
    }

    MonitoringService.trackMetric('word_create_failure', executionTime, {
      errorType: error.code || 'unknown',
      fileName,
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });

    throw mcpError;
  }
}

/**
 * Reads a Word document and extracts HTML and plain text content.
 * @param {string} fileId - OneDrive file ID
 * @param {Object} req - Express request object
 * @param {string} [userId] - User ID for logging context
 * @param {string} [sessionId] - Session ID for logging context
 * @returns {Promise<{html: string, text: string, warnings: Array}>}
 */
async function readWordDocument(fileId, req, userId, sessionId) {
  const startTime = Date.now();
  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;

  try {
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Word document read requested', {
        fileId,
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString()
      }, 'word');
    }

    const size = await filesService.getFileSize(fileId, req, userId, sessionId);
    if (size > MAX_FILE_SIZE) {
      throw new Error(`File size ${size} bytes exceeds maximum allowed size of ${MAX_FILE_SIZE} bytes`);
    }

    const client = await graphClientFactory.createClient(req, resolvedUserId, resolvedSessionId);
    const downloadResult = await client.api(`/me/drive/items/${fileId}/content`, resolvedUserId, resolvedSessionId).get();

    // Ensure we have a Buffer for mammoth/jszip
    let buffer;
    if (Buffer.isBuffer(downloadResult)) {
      buffer = downloadResult;
    } else if (typeof downloadResult === 'string') {
      buffer = Buffer.from(downloadResult, 'binary');
    } else {
      throw new Error(`Unexpected download result type: ${typeof downloadResult}, keys: ${Object.keys(downloadResult || {})}, preview: ${JSON.stringify(downloadResult).substring(0, 100)}`);
    }

    // Debug: log buffer info to help diagnose zip parse issues
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Word doc download buffer info', {
        isBuffer: Buffer.isBuffer(buffer),
        length: buffer.length,
        firstBytes: buffer.slice(0, 4).toString('hex'),
        isZip: buffer.slice(0, 2).toString() === 'PK'
      }, 'word');
    }

    const result = await mammoth.convertToHtml({ buffer });
    const textResult = await mammoth.extractRawText({ buffer });

    const executionTime = Date.now() - startTime;

    if (resolvedUserId) {
      MonitoringService.info('Word document read successfully', {
        fileId,
        htmlLength: result.value.length,
        textLength: textResult.value.length,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'word', null, resolvedUserId);
    }

    MonitoringService.trackMetric('word_read_success', executionTime, {
      fileId,
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });

    return {
      html: result.value,
      text: textResult.value,
      warnings: result.messages
    };
  } catch (error) {
    const executionTime = Date.now() - startTime;
    const mcpError = ErrorService.createError(
      'word',
      `Failed to read Word document: ${error.message}`,
      'error',
      {
        fileId,
        error: error.message,
        stack: error.stack,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }
    );
    MonitoringService.logError(mcpError);

    if (resolvedUserId) {
      MonitoringService.error('Word document read failed', {
        error: error.message,
        fileId,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'word', null, resolvedUserId);
    }

    MonitoringService.trackMetric('word_read_failure', executionTime, {
      errorType: error.code || 'unknown',
      fileId,
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });

    throw mcpError;
  }
}

/**
 * Converts a Word document to PDF using Graph API.
 * @param {string} fileId - OneDrive file ID
 * @param {Object} req - Express request object
 * @param {string} [userId] - User ID for logging context
 * @param {string} [sessionId] - Session ID for logging context
 * @returns {Promise<Buffer>} PDF content as a Buffer
 */
async function convertDocumentToPdf(fileId, req, userId, sessionId) {
  const startTime = Date.now();
  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;

  try {
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Word to PDF conversion requested', {
        fileId,
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString()
      }, 'word');
    }

    const client = await graphClientFactory.createClient(req, resolvedUserId, resolvedSessionId);
    const result = await client.api(`/me/drive/items/${fileId}/content?format=pdf`, resolvedUserId, resolvedSessionId).get();

    const executionTime = Date.now() - startTime;

    if (resolvedUserId) {
      MonitoringService.info('Word document converted to PDF', {
        fileId,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'word', null, resolvedUserId);
    }

    MonitoringService.trackMetric('word_convert_pdf_success', executionTime, {
      fileId,
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });

    return result;
  } catch (error) {
    const executionTime = Date.now() - startTime;
    const mcpError = ErrorService.createError(
      'word',
      `Failed to convert Word document to PDF: ${error.message}`,
      'error',
      {
        fileId,
        error: error.message,
        stack: error.stack,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }
    );
    MonitoringService.logError(mcpError);

    if (resolvedUserId) {
      MonitoringService.error('Word to PDF conversion failed', {
        error: error.message,
        fileId,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'word', null, resolvedUserId);
    }

    MonitoringService.trackMetric('word_convert_pdf_failure', executionTime, {
      errorType: error.code || 'unknown',
      fileId,
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });

    throw mcpError;
  }
}

/**
 * Extracts metadata from a Word document (docProps + Graph metadata).
 * @param {string} fileId - OneDrive file ID
 * @param {Object} req - Express request object
 * @param {string} [userId] - User ID for logging context
 * @param {string} [sessionId] - Session ID for logging context
 * @returns {Promise<Object>} Merged metadata object
 */
async function getWordDocumentMetadata(fileId, req, userId, sessionId) {
  const startTime = Date.now();
  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;

  try {
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Word document metadata requested', {
        fileId,
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString()
      }, 'word');
    }

    const size = await filesService.getFileSize(fileId, req, userId, sessionId);
    if (size > MAX_FILE_SIZE) {
      throw new Error(`File size ${size} bytes exceeds maximum allowed size of ${MAX_FILE_SIZE} bytes`);
    }

    // Get Graph file metadata and download content
    const client = await graphClientFactory.createClient(req, resolvedUserId, resolvedSessionId);
    const graphMeta = await client.api(`/me/drive/items/${fileId}`, resolvedUserId, resolvedSessionId).get();
    const buffer = await client.api(`/me/drive/items/${fileId}/content`, resolvedUserId, resolvedSessionId).get();
    const zip = await JSZip.loadAsync(buffer);
    const coreXml = await zip.file('docProps/core.xml')?.async('string');

    let docProps = {};
    if (coreXml) {
      const parsed = await parseStringPromise(coreXml);
      const props = parsed['cp:coreProperties'] || parsed['coreProperties'] || {};
      docProps = {
        title: props['dc:title']?.[0] || props['title']?.[0] || null,
        creator: props['dc:creator']?.[0] || props['creator']?.[0] || null,
        created: props['dcterms:created']?.[0]?._ || props['dcterms:created']?.[0] || null,
        modified: props['dcterms:modified']?.[0]?._ || props['dcterms:modified']?.[0] || null,
        description: props['dc:description']?.[0] || props['description']?.[0] || null,
        keywords: props['cp:keywords']?.[0] || props['keywords']?.[0] || null
      };
    }

    const executionTime = Date.now() - startTime;

    if (resolvedUserId) {
      MonitoringService.info('Word document metadata retrieved', {
        fileId,
        hasDocProps: !!coreXml,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'word', null, resolvedUserId);
    }

    MonitoringService.trackMetric('word_metadata_success', executionTime, {
      fileId,
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });

    return {
      ...docProps,
      id: graphMeta.id,
      name: graphMeta.name,
      size: graphMeta.size,
      webUrl: graphMeta.webUrl,
      lastModifiedDateTime: graphMeta.lastModifiedDateTime,
      createdDateTime: graphMeta.createdDateTime
    };
  } catch (error) {
    const executionTime = Date.now() - startTime;
    const mcpError = ErrorService.createError(
      'word',
      `Failed to get Word document metadata: ${error.message}`,
      'error',
      {
        fileId,
        error: error.message,
        stack: error.stack,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }
    );
    MonitoringService.logError(mcpError);

    if (resolvedUserId) {
      MonitoringService.error('Word document metadata retrieval failed', {
        error: error.message,
        fileId,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'word', null, resolvedUserId);
    }

    MonitoringService.trackMetric('word_metadata_failure', executionTime, {
      errorType: error.code || 'unknown',
      fileId,
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });

    throw mcpError;
  }
}

/**
 * Converts a Word document to HTML.
 * @param {string} fileId - OneDrive file ID
 * @param {Object} req - Express request object
 * @param {string} [userId] - User ID for logging context
 * @param {string} [sessionId] - Session ID for logging context
 * @returns {Promise<{html: string, warnings: Array}>}
 */
async function getWordDocumentAsHtml(fileId, req, userId, sessionId) {
  const startTime = Date.now();
  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;

  try {
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('Word document HTML conversion requested', {
        fileId,
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString()
      }, 'word');
    }

    const size = await filesService.getFileSize(fileId, req, userId, sessionId);
    if (size > MAX_FILE_SIZE) {
      throw new Error(`File size ${size} bytes exceeds maximum allowed size of ${MAX_FILE_SIZE} bytes`);
    }

    const client = await graphClientFactory.createClient(req, resolvedUserId, resolvedSessionId);
    const buffer = await client.api(`/me/drive/items/${fileId}/content`, resolvedUserId, resolvedSessionId).get();
    const result = await mammoth.convertToHtml({ buffer });

    const executionTime = Date.now() - startTime;

    if (resolvedUserId) {
      MonitoringService.info('Word document converted to HTML', {
        fileId,
        htmlLength: result.value.length,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'word', null, resolvedUserId);
    }

    MonitoringService.trackMetric('word_html_success', executionTime, {
      fileId,
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });

    return {
      html: result.value,
      warnings: result.messages
    };
  } catch (error) {
    const executionTime = Date.now() - startTime;
    const mcpError = ErrorService.createError(
      'word',
      `Failed to convert Word document to HTML: ${error.message}`,
      'error',
      {
        fileId,
        error: error.message,
        stack: error.stack,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }
    );
    MonitoringService.logError(mcpError);

    if (resolvedUserId) {
      MonitoringService.error('Word to HTML conversion failed', {
        error: error.message,
        fileId,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'word', null, resolvedUserId);
    }

    MonitoringService.trackMetric('word_html_failure', executionTime, {
      errorType: error.code || 'unknown',
      fileId,
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });

    throw mcpError;
  }
}

module.exports = {
  createWordDocument,
  readWordDocument,
  convertDocumentToPdf,
  getWordDocumentMetadata,
  getWordDocumentAsHtml
};
