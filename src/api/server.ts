import express from 'express';
import cors from 'cors';
import multer from 'multer';
import { config } from 'dotenv';
import { fileURLToPath } from 'node:url';
import { dirname, join, resolve } from 'node:path';
import { createServer } from 'node:https';
import { readFileSync, existsSync } from 'node:fs';
import { DeepSeekProvider } from '../llm/deepseekProvider.js';
import { runAgent } from '../agent/loop.js';
import { runAuto, type AutoEvent } from '../agent/autoLoop.js';
import { extractReference, composeIntent, type ExtractedReference } from '../utils/extractText.js';
import { convertToOoxml, libreofficeStatus } from '../utils/convertOoxml.js';
import { normalizeUploadFilename } from '../utils/filename.js';
import { generateDocxManifest } from '../office/docx/docxManifest.js';
import { generatePptxManifest } from '../office/pptx/pptxManifest.js';
import {
  createSession,
  getSession,
  bumpSession,
  statusOf,
  type DocumentSession,
} from '../utils/sessionCache.js';
import { logger } from '../utils/logger.js';

config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const PROJECT_ROOT = resolve(__dirname, '../../..');

const app = express();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

app.use(cors());
app.use(express.json({ limit: '50mb' }));

const isOfficeMode = process.argv.includes('--office');
const isHttpsMode = process.argv.includes('--https');

if (isOfficeMode) {
  const officeAddinPath = join(PROJECT_ROOT, 'apps/office-addin');
  app.use('/addin', express.static(officeAddinPath));
  logger.info(`Office Add-in files served from ${officeAddinPath}`);
}

// Web app (primary entry point)
const webAppPath = join(PROJECT_ROOT, 'apps/web');
app.use('/', express.static(webAppPath));
logger.info(`Web app served from ${webAppPath}`);

const defaultProvider = new DeepSeekProvider();

function pickProvider(req: express.Request): DeepSeekProvider {
  const userKey = (req.header('x-deepseek-key') ?? '').trim();
  if (userKey && userKey.startsWith('sk-')) {
    logger.info('Using user-supplied API key for this request');
    return new DeepSeekProvider({ apiKey: userKey });
  }
  return defaultProvider;
}

app.post('/api/agent/run', upload.single('file'), async (req, res) => {
  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    const { task, fileType: requestedType } = req.body;
    if (!task) {
      return res.status(400).json({ error: 'No task provided' });
    }

    const uploadName = normalizeUploadFilename(file.originalname);
    const ext = uploadName.split('.').pop()?.toLowerCase();
    const fileType = requestedType ?? (ext === 'docx' ? 'docx' : ext === 'pptx' ? 'pptx' : null);

    if (fileType !== 'docx' && fileType !== 'pptx') {
      return res.status(400).json({ error: 'Unsupported file type. Use .docx or .pptx' });
    }

    logger.info('Processing file', {
      filename: file.originalname,
      displayFilename: uploadName,
      fileType,
      size: file.size,
    });

    const result = await runAgent(pickProvider(req), file.buffer, fileType, task);

    if (result.success && result.modifiedFile) {
      res.setHeader('Content-Type', 'application/octet-stream');
      res.setHeader(
        'Content-Disposition',
        `attachment; filename="modified-document"; filename*=UTF-8''${encodeURIComponent(`modified_${uploadName}`)}`
      );
      res.send(result.modifiedFile);
    } else {
      res.json({
        success: false,
        summary: result.summary,
        iterations: result.iterations,
        usage: result.usage,
      });
    }
  } catch (err) {
    logger.error('Agent execution failed', { error: err instanceof Error ? err.message : String(err) });
    res.status(500).json({ error: err instanceof Error ? err.message : 'Internal server error' });
  }
});

const autoUpload = upload.fields([
  { name: 'file', maxCount: 1 },
  { name: 'references', maxCount: 6 },
]);

function modifiedFilename(filename: string): string {
  const ext = filename.split('.').pop()?.toLowerCase();
  if (ext === 'doc') return `modified_${filename.replace(/\.doc$/i, '.docx')}`;
  if (ext === 'ppt') return `modified_${filename.replace(/\.ppt$/i, '.pptx')}`;
  return `modified_${filename}`;
}

function downloadUrlFor(sessionId: string): string {
  return `/api/session/${encodeURIComponent(sessionId)}/download`;
}

app.post('/api/document/inspect', upload.single('file'), async (req, res) => {
  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    const uploadName = normalizeUploadFilename(file.originalname);
    const ext = uploadName.split('.').pop()?.toLowerCase();
    let buffer = file.buffer;
    let filename = uploadName;
    let fileType: 'docx' | 'pptx';

    if (ext === 'doc' || ext === 'ppt') {
      const converted = await convertToOoxml(file.buffer, uploadName);
      buffer = converted.buffer;
      filename = converted.filename;
      fileType = converted.kind;
    } else if (ext === 'docx' || ext === 'pptx') {
      fileType = ext;
    } else {
      return res.status(400).json({
        error: 'Unsupported file type. Use .docx / .pptx / .doc / .ppt',
      });
    }

    if (fileType === 'docx') {
      const manifest = await generateDocxManifest(buffer);
      return res.json({
        fileType,
        filename,
        preview: {
          type: 'docx',
          paragraphs: manifest.paragraphs.slice(0, 180).map((p) => ({
            id: p.paragraphId,
            text: p.textPreview,
            styleId: p.styleId,
            styleName: p.styleName,
            headingLevel: p.headingLevel,
            isListItem: p.isListItem,
            numbering: p.numbering,
            format: p.format,
            commentIds: p.commentIds,
          })),
          tables: manifest.tables.slice(0, 30).map((table) => ({
            id: table.tableId,
            rows: table.rows.slice(0, 40).map((row) => ({
              id: row.rowId,
              cells: row.cells.slice(0, 12).map((cell) => ({
                id: cell.cellId,
                text: cell.textPreview,
                columnSpan: cell.columnSpan,
                verticalMerge: cell.verticalMerge,
              })),
            })),
          })),
        },
        summary: {
          paragraphs: manifest.paragraphs.length,
          tables: manifest.tables.length,
          headers: manifest.headers.length,
          footers: manifest.footers.length,
          comments: manifest.comments?.length ?? 0,
          hasComments: manifest.hasComments,
          hasFootnotes: manifest.hasFootnotes,
          hasEndnotes: manifest.hasEndnotes,
          hasNumbering: manifest.hasNumbering,
          validationErrors: manifest.validation.errors.length,
          validationWarnings: manifest.validation.warnings.length,
        },
        comments: (manifest.comments ?? []).map((comment) => ({
          commentId: comment.commentId,
          author: comment.author,
          date: comment.date,
          textPreview: comment.textPreview,
          anchoredParagraphIds: comment.anchoredParagraphIds,
        })),
        reviewTargets: manifest.paragraphs
          .filter((p) => p.commentIds && p.commentIds.length > 0)
          .slice(0, 20)
          .map((p) => ({
            id: p.paragraphId,
            textPreview: p.textPreview,
            styleId: p.styleId,
            commentIds: p.commentIds,
          })),
        warnings: manifest.validation.warnings,
        errors: manifest.validation.errors,
      });
    }

    const manifest = await generatePptxManifest(buffer);
    return res.json({
      fileType,
      filename,
      preview: {
        type: 'pptx',
        slides: manifest.slides.slice(0, 80).map((slide) => ({
          id: slide.slideId,
          index: slide.slideIndex,
          title: slide.titleCandidate,
          hasNotes: slide.hasNotes,
          hasComments: slide.hasComments,
          shapes: slide.shapes.slice(0, 30).map((shape) => ({
            id: shape.shapeId,
            type: shape.shapeType,
            placeholderType: shape.placeholderType,
            text: shape.textPreview,
            position: shape.position,
          })),
        })),
      },
      summary: {
        slides: manifest.slides.length,
        shapes: manifest.slides.reduce((sum, slide) => sum + slide.shapes.length, 0),
        layouts: manifest.layouts.length,
        masters: manifest.masters.length,
        comments: manifest.slides.filter((slide) => slide.hasComments).length,
        speakerNotes: manifest.slides.filter((slide) => slide.hasNotes).length,
        validationErrors: manifest.validation.errors.length,
        validationWarnings: manifest.validation.warnings.length,
      },
      comments: [],
      reviewTargets: manifest.slides.slice(0, 20).map((slide) => ({
        id: slide.slideId,
        textPreview: slide.titleCandidate ?? '',
        shapeCount: slide.shapes.length,
        hasNotes: slide.hasNotes,
        hasComments: slide.hasComments,
      })),
      warnings: manifest.validation.warnings,
      errors: manifest.validation.errors,
    });
  } catch (err) {
    logger.error('Document inspect failed', {
      error: err instanceof Error ? err.message : String(err),
    });
    res.status(500).json({ error: err instanceof Error ? err.message : 'Internal server error' });
  }
});

app.post('/api/agent/auto', autoUpload, async (req, res) => {
  // NDJSON streaming response: one JSON event per line
  res.setHeader('Content-Type', 'application/x-ndjson; charset=utf-8');
  res.setHeader('Cache-Control', 'no-cache, no-transform');
  res.setHeader('X-Accel-Buffering', 'no');
  // Flush headers immediately so the client can start reading
  if (typeof (res as any).flushHeaders === 'function') {
    (res as any).flushHeaders();
  }

  const writeEvent = (event: AutoEvent): void => {
    try {
      res.write(JSON.stringify(event) + '\n');
    } catch (err) {
      logger.warn('Failed to write auto event', {
        error: err instanceof Error ? err.message : String(err),
      });
    }
  };

  try {
    const files = req.files as
      | { file?: Express.Multer.File[]; references?: Express.Multer.File[] }
      | undefined;
    const uploadedFile = files?.file?.[0];
    const refFiles = files?.references ?? [];
    const freeText = (req.body.intent ?? req.body.task ?? '').toString().trim();
    const sessionIdInput = (req.body.sessionId ?? '').toString().trim();

    // Resolve target buffer: either reuse session OR consume fresh upload.
    let buffer: Buffer;
    let fileType: 'docx' | 'pptx';
    let filename: string;
    let session: DocumentSession | null = null;

    if (sessionIdInput) {
      session = getSession(sessionIdInput);
      if (!session) {
        writeEvent({
          type: 'error',
          error: 'Session expired or unknown. Please re-upload the document.',
        });
        return res.end();
      }
      buffer = session.buffer;
      fileType = session.fileType;
      filename = session.filename;
    } else if (uploadedFile) {
      const uploadName = normalizeUploadFilename(uploadedFile.originalname);
      const ext = uploadName.split('.').pop()?.toLowerCase();

      if (ext === 'doc' || ext === 'ppt') {
        // Legacy binary → convert to OOXML on the fly
        writeEvent({
          type: 'converting',
          from: ext,
          filename: uploadName,
        });
        try {
          const converted = await convertToOoxml(
            uploadedFile.buffer,
            uploadName
          );
          buffer = converted.buffer;
          filename = converted.filename;
          fileType = converted.kind;
        } catch (err) {
          writeEvent({
            type: 'error',
            error: err instanceof Error ? err.message : String(err),
          });
          return res.end();
        }
      } else if (ext === 'docx' || ext === 'pptx') {
        buffer = uploadedFile.buffer;
        fileType = ext;
        filename = uploadName;
      } else {
        writeEvent({
          type: 'error',
          error:
            '不支持的目标文档类型。支持：.docx / .pptx / .doc / .ppt（.doc/.ppt 需要 libreoffice）',
        });
        return res.end();
      }
    } else {
      writeEvent({ type: 'error', error: 'No file uploaded and no sessionId provided' });
      return res.end();
    }

    if (!freeText && refFiles.length === 0) {
      writeEvent({
        type: 'error',
        error: 'No intent provided. Type instructions or attach a reference.',
      });
      return res.end();
    }

    // Extract reference texts (sequential is fine; each is fast)
    const references: ExtractedReference[] = [];
    for (const r of refFiles) {
      try {
        const ref = await extractReference(r.buffer, normalizeUploadFilename(r.originalname));
        references.push(ref);
      } catch (err) {
        writeEvent({
          type: 'error',
          error: err instanceof Error ? err.message : String(err),
        });
        return res.end();
      }
    }

    const intent = composeIntent(freeText, references);

    // Create session on first upload, bump on reuse
    if (!session) {
      session = createSession(buffer, fileType, filename);
    } else {
      bumpSession(session.id);
    }

    writeEvent({ type: 'session', session: statusOf(session) });

    logger.info('Auto run requested', {
      filename,
      fileType,
      size: buffer.length,
      reuse: !!sessionIdInput,
      reuseCount: session.reuseCount,
      freeTextLength: freeText.length,
      referenceCount: references.length,
      composedIntentLength: intent.length,
    });

    const writeAutoEvent = (event: AutoEvent): void => {
      if (event.type === 'complete' && event.modifiedFileBase64) {
        session.buffer = Buffer.from(event.modifiedFileBase64, 'base64');
        session.fileType = fileType;
        session.filename = filename;
        writeEvent({
          type: 'complete',
          success: event.success,
          totalUsage: event.totalUsage,
          downloadUrl: downloadUrlFor(session.id),
          modifiedFilename: modifiedFilename(filename),
        });
        return;
      }
      writeEvent(event);
    };

    await runAuto(pickProvider(req), buffer, fileType, intent, writeAutoEvent);
    res.end();
  } catch (err) {
    logger.error('Auto endpoint failed', {
      error: err instanceof Error ? err.message : String(err),
    });
    writeEvent({
      type: 'error',
      error: err instanceof Error ? err.message : 'Internal server error',
    });
    res.end();
  }
});

app.get('/api/session/:id/download', (req, res) => {
  try {
    const session = getSession(req.params.id);
    if (!session) {
      return res.status(404).json({ error: 'Session expired or unknown. Please re-run the edit.' });
    }

    const filename = modifiedFilename(session.filename);
    const mime = session.fileType === 'docx'
      ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      : 'application/vnd.openxmlformats-officedocument.presentationml.presentation';

    res.setHeader('Content-Type', mime);
    res.setHeader(
      'Content-Disposition',
      `attachment; filename="modified-document.${session.fileType}"; filename*=UTF-8''${encodeURIComponent(filename)}`
    );
    res.send(session.buffer);
  } catch (err) {
    logger.error('Session download failed', {
      error: err instanceof Error ? err.message : String(err),
    });
    res.status(500).json({ error: err instanceof Error ? err.message : 'Internal server error' });
  }
});

app.get('/api/health', async (_req, res) => {
  const lo = await libreofficeStatus();
  res.json({
    status: 'ok',
    timestamp: new Date().toISOString(),
    libreoffice: lo,
  });
});

const PORT = parseInt(process.env.PORT ?? '3000', 10);
const HOST = process.env.HOST ?? '0.0.0.0';

if (isHttpsMode) {
  const keyPath = join(PROJECT_ROOT, 'certs/localhost.key');
  const certPath = join(PROJECT_ROOT, 'certs/localhost.crt');

  if (!existsSync(keyPath) || !existsSync(certPath)) {
    logger.error('SSL certificates not found. Run: npm run generate-certs');
    process.exit(1);
  }

  const httpsServer = createServer(
    {
      key: readFileSync(keyPath),
      cert: readFileSync(certPath),
    },
    app
  );

  httpsServer.listen(PORT, HOST, () => {
    logger.info(`HTTPS Server running at https://${HOST}:${PORT}`);
    logger.info('Office Add-in URL: https://localhost:3000/addin/taskpane.html');
  });
} else {
  app.listen(PORT, HOST, () => {
    logger.info(`Server running at http://${HOST}:${PORT}`);
  });
}

export { app };
