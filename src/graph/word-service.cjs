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
    } else if (downloadResult && typeof downloadResult === 'object') {
      // Graph might return a JSON error or redirect response instead of binary
      throw new Error(`Download returned object instead of binary: ${JSON.stringify(downloadResult).substring(0, 200)}`);
    } else {
      throw new Error(`Unexpected download result type: ${typeof downloadResult}`);
    }

    // Check if the file is a valid Open XML zip (PK header)
    const isOpenXml = buffer.length >= 4 && buffer.slice(0, 2).toString() === 'PK';

    if (!isOpenXml) {
      // Not a .docx (Open XML) — likely a legacy .doc, encrypted file, or other format
      // Fall back to Graph's server-side HTML conversion which handles all Office formats
      MonitoringService.info('Word doc is not Open XML (header: ' + buffer.slice(0, 4).toString('hex') + '), using Graph HTML conversion fallback', {
        fileId, headerHex: buffer.slice(0, 4).toString('hex')
      }, 'word');

      let html = '', text = '', warnings = ['File is not in Open XML (.docx) format'];
      try {
        const htmlResult = await client.api(`/me/drive/items/${fileId}/content?format=html`, resolvedUserId, resolvedSessionId).get();
        html = Buffer.isBuffer(htmlResult) ? htmlResult.toString('utf8') : (typeof htmlResult === 'string' ? htmlResult : '');
        text = html.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
        warnings.push('Read via Graph HTML conversion');
      } catch (convErr) {
        // Graph can't convert this format (406 notSupported) — return what we know from metadata
        const meta = await client.api(`/me/drive/items/${fileId}`, resolvedUserId, resolvedSessionId).get();
        html = `<p>Unable to extract content from this file format. File: ${meta.name}, Size: ${meta.size} bytes. <a href="${meta.webUrl}">Open in browser</a></p>`;
        text = `Unable to extract content from this file format. File: ${meta.name}, Size: ${meta.size} bytes. Open in browser: ${meta.webUrl}`;
        warnings.push('Graph conversion not supported for this file format — use the webUrl to open in browser');
      }

      const executionTime = Date.now() - startTime;
      if (resolvedUserId) {
        MonitoringService.info('Word document read via fallback', {
          fileId, htmlLength: html.length, executionTimeMs: executionTime
        }, 'word', null, resolvedUserId);
      }

      return { html, text, warnings: ['File was not in .docx format — read via Graph server-side conversion'] };
    }

    MonitoringService.debug('Word doc download buffer info', {
      isBuffer: true, length: buffer.length, firstBytes: buffer.slice(0, 4).toString('hex'), isZip: true
    }, 'word');

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
    // Try to extract docProps from Open XML, fall back to Graph metadata for non-zip files
    let docProps = {};
    const downloadResult = await client.api(`/me/drive/items/${fileId}/content`, resolvedUserId, resolvedSessionId).get();
    const buffer = Buffer.isBuffer(downloadResult) ? downloadResult : (typeof downloadResult === 'string' ? Buffer.from(downloadResult, 'binary') : null);
    const isOpenXml = buffer && buffer.length >= 4 && buffer.slice(0, 2).toString() === 'PK';

    if (isOpenXml) {
      const zip = await JSZip.loadAsync(buffer);
      const coreXml = await zip.file('docProps/core.xml')?.async('string');
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
    } else {
      // Legacy format — extract what we can from Graph driveItem metadata
      docProps = {
        title: graphMeta.name?.replace(/\.[^.]+$/, '') || null,
        creator: graphMeta.createdBy?.user?.displayName || null,
        created: graphMeta.createdDateTime || null,
        modified: graphMeta.lastModifiedDateTime || null,
        description: null,
        keywords: null,
        _note: 'Metadata extracted from Graph (file is not Open XML format)'
      };
    }

    const executionTime = Date.now() - startTime;

    if (resolvedUserId) {
      MonitoringService.info('Word document metadata retrieved', {
        fileId,
        isOpenXml,
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
    const downloadResult = await client.api(`/me/drive/items/${fileId}/content`, resolvedUserId, resolvedSessionId).get();
    const buffer = Buffer.isBuffer(downloadResult) ? downloadResult : (typeof downloadResult === 'string' ? Buffer.from(downloadResult, 'binary') : downloadResult);
    const isOpenXml = Buffer.isBuffer(buffer) && buffer.length >= 4 && buffer.slice(0, 2).toString() === 'PK';

    let html, warnings = [];
    if (isOpenXml) {
      const result = await mammoth.convertToHtml({ buffer });
      html = result.value;
      warnings = result.messages;
    } else {
      // Fall back to Graph HTML conversion for legacy/non-zip formats
      try {
        const htmlResult = await client.api(`/me/drive/items/${fileId}/content?format=html`, resolvedUserId, resolvedSessionId).get();
        html = Buffer.isBuffer(htmlResult) ? htmlResult.toString('utf8') : (typeof htmlResult === 'string' ? htmlResult : '');
        warnings = [{ message: 'Converted via Graph (file is not Open XML format)' }];
      } catch (convErr) {
        const meta = await client.api(`/me/drive/items/${fileId}`, resolvedUserId, resolvedSessionId).get();
        html = `<p>Unable to convert this file format to HTML. File: ${meta.name}. <a href="${meta.webUrl}">Open in browser</a></p>`;
        warnings = [{ message: 'Conversion not supported — use webUrl to open in browser' }];
      }
    }

    const executionTime = Date.now() - startTime;

    if (resolvedUserId) {
      MonitoringService.info('Word document converted to HTML', {
        fileId, htmlLength: html.length, fallback: !isOpenXml, executionTimeMs: executionTime
      }, 'word', null, resolvedUserId);
    }

    MonitoringService.trackMetric('word_html_success', executionTime, {
      fileId,
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });

    return { html, warnings };
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
