/**
 * @fileoverview PowerPointService - Microsoft Graph PowerPoint operations.
 * Uses pptxgenjs for creation and JSZip for reading/metadata extraction.
 * Follows project error handling, validation, and normalization rules.
 */

const PptxGenJS = require('pptxgenjs');
const JSZip = require('jszip');
const { parseStringPromise } = require('xml2js');
const graphClientFactory = require('./graph-client.cjs');
const filesService = require('./files-service.cjs');
const ErrorService = require('../core/error-service.cjs');
const MonitoringService = require('../core/monitoring-service.cjs');

const MAX_FILE_SIZE = 25 * 1024 * 1024; // 25MB

/**
 * Builds a single slide in the presentation from a slide definition.
 * @param {Object} pptx - PptxGenJS instance
 * @param {Object} slideDef - Slide definition with layout, title, subtitle, body
 * @private
 */
function buildSlide(pptx, slideDef) {
  const slide = pptx.addSlide();

  switch (slideDef.layout) {
    case 'title':
      slide.addText(slideDef.title || '', {
        x: 0.5, y: 1, w: 9, h: 1.5,
        fontSize: 36, bold: true
      });
      if (slideDef.subtitle) {
        slide.addText(slideDef.subtitle, {
          x: 0.5, y: 3, w: 9, h: 1,
          fontSize: 20, color: '666666'
        });
      }
      break;

    case 'content': {
      slide.addText(slideDef.title || '', {
        x: 0.5, y: 0.3, w: 9, h: 0.8,
        fontSize: 28, bold: true
      });
      let currentY = 1.5;
      if (slideDef.body && Array.isArray(slideDef.body)) {
        for (const item of slideDef.body) {
          if (item.type === 'text') {
            slide.addText(item.text || '', {
              x: 0.5, y: currentY, w: 9, h: 0.5,
              fontSize: 16
            });
            currentY += 0.6;
          } else if (item.type === 'image') {
            slide.addImage({
              data: 'data:image/png;base64,' + item.data,
              x: item.x || 0.5,
              y: item.y || currentY,
              w: item.width || 4,
              h: item.height || 3
            });
            currentY += (item.height || 3) + 0.2;
          }
        }
      }
      break;
    }

    case 'blank':
    default:
      // Empty slide
      break;
  }

  return slide;
}

/**
 * Recursively extracts text from XML elements by finding all <a:t> nodes.
 * @param {Object} element - Parsed XML element
 * @returns {string[]} Array of text strings found
 * @private
 */
function extractTexts(element) {
  const texts = [];
  if (!element || typeof element !== 'object') return texts;

  if (element['a:t']) {
    for (const t of Array.isArray(element['a:t']) ? element['a:t'] : [element['a:t']]) {
      const text = typeof t === 'string' ? t : t._ || '';
      if (text.trim()) texts.push(text.trim());
    }
  }

  for (const key of Object.keys(element)) {
    const child = element[key];
    if (Array.isArray(child)) {
      for (const item of child) {
        texts.push(...extractTexts(item));
      }
    } else if (typeof child === 'object' && child !== null) {
      texts.push(...extractTexts(child));
    }
  }

  return texts;
}

/**
 * Creates a PowerPoint presentation from slide definitions and uploads to OneDrive.
 * @param {string} fileName - Name for the file (should end in .pptx)
 * @param {Array<Object>} slides - Array of slide definitions
 * @param {string} [folderId] - Target folder ID (defaults to root)
 * @param {Object} req - Express request object
 * @param {string} [userId] - User ID for logging context
 * @param {string} [sessionId] - Session ID for logging context
 * @returns {Promise<Object>} Uploaded file metadata from Graph API
 */
async function createPresentation(fileName, slides, folderId, req, userId, sessionId) {
  const startTime = Date.now();
  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;

  try {
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('PowerPoint creation requested', {
        fileName,
        folderId: folderId || 'root',
        slideCount: slides?.length || 0,
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString()
      }, 'powerpoint');
    }

    const pptx = new PptxGenJS();

    for (const slideDef of (slides || [])) {
      buildSlide(pptx, slideDef);
    }

    const buffer = Buffer.from(await pptx.write({ outputType: 'arraybuffer' }));

    // Use filesService.uploadFile which handles binary Buffers correctly
    const result = await filesService.uploadFile(fileName, buffer, req, userId, sessionId);

    const executionTime = Date.now() - startTime;

    if (resolvedUserId) {
      MonitoringService.info('PowerPoint created successfully', {
        fileName,
        fileId: result.id,
        slideCount: slides?.length || 0,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'powerpoint', null, resolvedUserId);
    }

    MonitoringService.trackMetric('powerpoint_create_success', executionTime, {
      fileName,
      slideCount: slides?.length || 0,
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });

    return result;
  } catch (error) {
    const executionTime = Date.now() - startTime;
    const mcpError = ErrorService.createError(
      'powerpoint',
      `Failed to create PowerPoint: ${error.message}`,
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
      MonitoringService.error('PowerPoint creation failed', {
        error: error.message,
        fileName,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'powerpoint', null, resolvedUserId);
    }

    MonitoringService.trackMetric('powerpoint_create_failure', executionTime, {
      errorType: error.code || 'unknown',
      fileName,
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });

    throw mcpError;
  }
}

/**
 * Reads a PowerPoint file and extracts text content from all slides.
 * @param {string} fileId - OneDrive file ID
 * @param {Object} req - Express request object
 * @param {string} [userId] - User ID for logging context
 * @param {string} [sessionId] - Session ID for logging context
 * @returns {Promise<{slideCount: number, slides: Array<{index: number, texts: string[]}>}>}
 */
async function readPresentation(fileId, req, userId, sessionId) {
  const startTime = Date.now();
  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;

  try {
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('PowerPoint read requested', {
        fileId,
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString()
      }, 'powerpoint');
    }

    const size = await filesService.getFileSize(fileId, req, userId, sessionId);
    if (size > MAX_FILE_SIZE) {
      throw new Error(`File size ${size} bytes exceeds maximum allowed size of ${MAX_FILE_SIZE} bytes`);
    }

    // Strategy: Try Graph HTML conversion first (works for ALL formats),
    // fall back to jszip binary parsing only if Graph fails.
    const client = await graphClientFactory.createClient(req, resolvedUserId, resolvedSessionId);

    // Attempt 1: Graph server-side HTML conversion
    try {
      const htmlResult = await client.api(`/me/drive/items/${fileId}/content?format=html`, resolvedUserId, resolvedSessionId).get();
      const html = Buffer.isBuffer(htmlResult) ? htmlResult.toString('utf8') : (typeof htmlResult === 'string' ? htmlResult : '');
      if (html && html.length > 0) {
        const text = html.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
        // Try to count slides from the HTML (each slide typically generates a section)
        const slideMatches = html.match(/<div[^>]*class="[^"]*slide[^"]*"/gi) || [];
        MonitoringService.trackMetric('powerpoint_read_success', Date.now() - startTime, { fileId, method: 'graph-html', userId: resolvedUserId });
        return {
          slideCount: slideMatches.length || null,
          slides: [{ index: 1, texts: [text.substring(0, 10000)] }],
          _format: 'graph-html'
        };
      }
    } catch (graphErr) {
      MonitoringService.debug('Graph HTML conversion failed, trying jszip fallback', { fileId, error: graphErr.message }, 'powerpoint');
    }

    // Attempt 2: Download binary via @microsoft.graph.downloadUrl and parse with jszip
    try {
      const driveItem = await client.api(`/me/drive/items/${fileId}`, resolvedUserId, resolvedSessionId).get();
      const downloadUrl = driveItem['@microsoft.graph.downloadUrl'];
      let buffer = null;
      if (downloadUrl) {
        const fetch = require('node-fetch');
        const dlResponse = await fetch(downloadUrl);
        buffer = await dlResponse.buffer();
      }

      if (buffer && buffer.length >= 4 && buffer.slice(0, 2).toString() === 'PK') {
        const zip = await JSZip.loadAsync(buffer);

        // Find and sort slide XML files
        const slideFiles = Object.keys(zip.files)
      .filter(f => f.match(/^ppt\/slides\/slide\d+\.xml$/))
      .sort((a, b) => {
        const numA = parseInt(a.match(/slide(\d+)\.xml$/)[1], 10);
        const numB = parseInt(b.match(/slide(\d+)\.xml$/)[1], 10);
        return numA - numB;
      });

        const slides = [];
        for (let i = 0; i < slideFiles.length; i++) {
          const xml = await zip.file(slideFiles[i]).async('string');
          const parsed = await parseStringPromise(xml);
          const texts = extractTexts(parsed);
          slides.push({ index: i + 1, texts });
        }
        MonitoringService.trackMetric('powerpoint_read_success', Date.now() - startTime, { fileId, method: 'jszip', slideCount: slides.length, userId: resolvedUserId });
        return { slideCount: slides.length, slides };
      }
    } catch (jszipErr) {
      MonitoringService.debug('jszip fallback also failed', { fileId, error: jszipErr.message }, 'powerpoint');
    }

    // Attempt 3: Return file metadata with webUrl
    const meta = await client.api(`/me/drive/items/${fileId}`, resolvedUserId, resolvedSessionId).get();
    MonitoringService.trackMetric('powerpoint_read_fallback', Date.now() - startTime, { fileId, userId: resolvedUserId });
    return {
      slideCount: null,
      slides: [{ index: 1, texts: [`Unable to extract content. Open in browser: ${meta.webUrl}`] }],
      _format: 'fallback'
    };
  } catch (error) {
    const executionTime = Date.now() - startTime;
    const mcpError = ErrorService.createError(
      'powerpoint',
      `Failed to read PowerPoint: ${error.message}`,
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
      MonitoringService.error('PowerPoint read failed', {
        error: error.message,
        fileId,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'powerpoint', null, resolvedUserId);
    }

    MonitoringService.trackMetric('powerpoint_read_failure', executionTime, {
      errorType: error.code || 'unknown',
      fileId,
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });

    throw mcpError;
  }
}

/**
 * Converts a PowerPoint to PDF using Graph API.
 * @param {string} fileId - OneDrive file ID
 * @param {Object} req - Express request object
 * @param {string} [userId] - User ID for logging context
 * @param {string} [sessionId] - Session ID for logging context
 * @returns {Promise<Buffer>} PDF content as a Buffer
 */
async function convertPresentationToPdf(fileId, req, userId, sessionId) {
  const startTime = Date.now();
  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;

  try {
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('PowerPoint to PDF conversion requested', {
        fileId,
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString()
      }, 'powerpoint');
    }

    const client = await graphClientFactory.createClient(req, resolvedUserId, resolvedSessionId);
    const result = await client.api(`/me/drive/items/${fileId}/content?format=pdf`, resolvedUserId, resolvedSessionId).get();

    const executionTime = Date.now() - startTime;

    if (resolvedUserId) {
      MonitoringService.info('PowerPoint converted to PDF', {
        fileId,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'powerpoint', null, resolvedUserId);
    }

    MonitoringService.trackMetric('powerpoint_convert_pdf_success', executionTime, {
      fileId,
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });

    return result;
  } catch (error) {
    const executionTime = Date.now() - startTime;
    const mcpError = ErrorService.createError(
      'powerpoint',
      `Failed to convert PowerPoint to PDF: ${error.message}`,
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
      MonitoringService.error('PowerPoint to PDF conversion failed', {
        error: error.message,
        fileId,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'powerpoint', null, resolvedUserId);
    }

    MonitoringService.trackMetric('powerpoint_convert_pdf_failure', executionTime, {
      errorType: error.code || 'unknown',
      fileId,
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });

    throw mcpError;
  }
}

/**
 * Extracts metadata from a PowerPoint file (docProps + Graph metadata + slide count).
 * @param {string} fileId - OneDrive file ID
 * @param {Object} req - Express request object
 * @param {string} [userId] - User ID for logging context
 * @param {string} [sessionId] - Session ID for logging context
 * @returns {Promise<Object>} Merged metadata object
 */
async function getPresentationMetadata(fileId, req, userId, sessionId) {
  const startTime = Date.now();
  const resolvedUserId = userId || req?.user?.userId;
  const resolvedSessionId = sessionId || req?.session?.id;

  try {
    if (process.env.NODE_ENV === 'development') {
      MonitoringService.debug('PowerPoint metadata requested', {
        fileId,
        sessionId: resolvedSessionId,
        timestamp: new Date().toISOString()
      }, 'powerpoint');
    }

    const size = await filesService.getFileSize(fileId, req, userId, sessionId);
    if (size > MAX_FILE_SIZE) {
      throw new Error(`File size ${size} bytes exceeds maximum allowed size of ${MAX_FILE_SIZE} bytes`);
    }

    // Get Graph file metadata and download content
    const client = await graphClientFactory.createClient(req, resolvedUserId, resolvedSessionId);
    const graphMeta = await client.api(`/me/drive/items/${fileId}`, resolvedUserId, resolvedSessionId).get();
    const downloadResult = await client.api(`/me/drive/items/${fileId}/content`, resolvedUserId, resolvedSessionId).get();
    const buffer = Buffer.isBuffer(downloadResult) ? downloadResult : (typeof downloadResult === 'string' ? Buffer.from(downloadResult, 'binary') : null);
    const isOpenXml = buffer && buffer.length >= 4 && buffer.slice(0, 2).toString() === 'PK';

    let slideCount = null;
    let docProps = {};

    if (isOpenXml) {
      const zip = await JSZip.loadAsync(buffer);
      slideCount = Object.keys(zip.files).filter(f => f.match(/^ppt\/slides\/slide\d+\.xml$/)).length;
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
      // Legacy format — use Graph metadata
      docProps = {
        title: graphMeta.name?.replace(/\.[^.]+$/, '') || null,
        creator: graphMeta.createdBy?.user?.displayName || null,
        created: graphMeta.createdDateTime || null,
        modified: graphMeta.lastModifiedDateTime || null,
        description: null,
        keywords: null,
        _note: 'Metadata from Graph (file is not Open XML format)'
      };
    }

    const executionTime = Date.now() - startTime;

    if (resolvedUserId) {
      MonitoringService.info('PowerPoint metadata retrieved', {
        fileId,
        slideCount,
        isOpenXml,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'powerpoint', null, resolvedUserId);
    }

    MonitoringService.trackMetric('powerpoint_metadata_success', executionTime, {
      fileId,
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });

    return {
      ...docProps,
      slideCount,
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
      'powerpoint',
      `Failed to get PowerPoint metadata: ${error.message}`,
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
      MonitoringService.error('PowerPoint metadata retrieval failed', {
        error: error.message,
        fileId,
        executionTimeMs: executionTime,
        timestamp: new Date().toISOString()
      }, 'powerpoint', null, resolvedUserId);
    }

    MonitoringService.trackMetric('powerpoint_metadata_failure', executionTime, {
      errorType: error.code || 'unknown',
      fileId,
      userId: resolvedUserId,
      timestamp: new Date().toISOString()
    });

    throw mcpError;
  }
}

module.exports = {
  createPresentation,
  readPresentation,
  convertPresentationToPdf,
  getPresentationMetadata
};
