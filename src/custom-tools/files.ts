import { unzipSync, strFromU8 } from 'fflate';
import { PDFParse } from 'pdf-parse';
import { z } from 'zod';
import {
  CustomToolDefinition,
  CustomToolHandlerContext,
  DriveItemSummary,
  GraphCollectionResponse,
  GraphToolError,
  JsonRecord,
  asRecord,
  buildDrivePathEndpoint,
  decodeCursor,
  encodeCursor,
  encodePathSegment,
  errorResult,
  getArray,
  getOptionalBoolean,
  getOptionalNumber,
  getOptionalString,
  getRequiredString,
  getString,
  graphRequest,
  normalizeDriveItem,
  normalizePath,
  resolveDriveItemByPath,
  success,
} from './shared.js';

const textEditOperationSchema = z.discriminatedUnion('type', [
  z.object({
    type: z.literal('replaceAll'),
    find: z.string(),
    replace: z.string(),
  }),
  z.object({
    type: z.literal('replaceFirst'),
    find: z.string(),
    replace: z.string(),
  }),
  z.object({
    type: z.literal('append'),
    text: z.string(),
  }),
  z.object({
    type: z.literal('prepend'),
    text: z.string(),
  }),
  z.object({
    type: z.literal('overwrite'),
    text: z.string(),
  }),
]);

type TextEditOperation = z.infer<typeof textEditOperationSchema>;

const TEXT_MIME_PREFIXES = ['text/'];
const TEXT_MIME_TYPES = new Set([
  'application/json',
  'application/ld+json',
  'application/xml',
  'application/javascript',
  'application/x-javascript',
  'application/typescript',
  'application/x-typescript',
  'application/yaml',
  'application/x-yaml',
  'application/csv',
  'application/sql',
  'application/x-sh',
]);
const TEXT_EXTENSIONS = new Set([
  'txt',
  'md',
  'markdown',
  'json',
  'xml',
  'yaml',
  'yml',
  'csv',
  'tsv',
  'js',
  'mjs',
  'cjs',
  'ts',
  'tsx',
  'jsx',
  'html',
  'css',
  'sql',
  'log',
  'py',
  'rb',
  'go',
  'rs',
  'java',
  'c',
  'cpp',
  'h',
  'sh',
  'env',
]);

type ExtractedFileText = {
  id: string;
  driveId: string | null;
  name: string;
  path: string;
  label: string;
  text: string | null;
  truncated: boolean;
  extractionStatus: 'ok' | 'unsupported' | 'empty';
  sourceMimeType: string | null;
  webUrl: string | null;
};

function getExtension(name: string): string | null {
  const segments = name.toLowerCase().split('.');
  return segments.length > 1 ? (segments.at(-1) ?? null) : null;
}

function decodeXmlEntities(value: string): string {
  return value
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-f]+);/gi, (_, hex: string) => String.fromCodePoint(parseInt(hex, 16)))
    .replace(/&#([0-9]+);/g, (_, decimal: string) => String.fromCodePoint(parseInt(decimal, 10)));
}

function normalizeExtractedText(
  text: string,
  maxChars: number
): {
  text: string | null;
  truncated: boolean;
  extractionStatus: 'ok' | 'empty';
} {
  const normalized = text
    .replace(/\r\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
  if (!normalized) {
    return {
      text: null,
      truncated: false,
      extractionStatus: 'empty',
    };
  }
  return {
    text: normalized.slice(0, maxChars),
    truncated: normalized.length > maxChars,
    extractionStatus: 'ok',
  };
}

function extractXmlTagText(
  xml: string,
  textTag: RegExp,
  breakPatterns: Array<[RegExp, string]> = []
): string {
  let normalized = xml;
  for (const [pattern, replacement] of breakPatterns) {
    normalized = normalized.replace(pattern, replacement);
  }
  const values: string[] = [];
  let match: RegExpExecArray | null;
  while ((match = textTag.exec(normalized)) !== null) {
    values.push(decodeXmlEntities(match[1] ?? ''));
  }
  return values.join(' ').replace(/\s+\n/g, '\n').replace(/\n\s+/g, '\n');
}

function extractDocxText(buffer: Buffer): string {
  const archive = unzipSync(new Uint8Array(buffer));
  const parts = Object.keys(archive)
    .filter((name) =>
      /word\/(document|header\d+|footer\d+|footnotes|endnotes|comments)\.xml$/i.test(name)
    )
    .sort();
  return parts
    .map((name) =>
      extractXmlTagText(strFromU8(archive[name]!), /<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g, [
        [/<w:tab\/>/g, '\t'],
        [/<w:br\/>/g, '\n'],
        [/<\/w:p>/g, '\n'],
        [/<\/w:tr>/g, '\n'],
      ])
    )
    .join('\n')
    .trim();
}

function extractPptxText(buffer: Buffer): string {
  const archive = unzipSync(new Uint8Array(buffer));
  const slides = Object.keys(archive)
    .filter((name) => /ppt\/slides\/slide\d+\.xml$/i.test(name))
    .sort((left, right) => left.localeCompare(right, undefined, { numeric: true }));
  return slides
    .map((name, index) => {
      const slideText = extractXmlTagText(strFromU8(archive[name]!), /<a:t>([\s\S]*?)<\/a:t>/g, [
        [/<a:br\/>/g, '\n'],
      ]).trim();
      return slideText ? `Slide ${index + 1}\n${slideText}` : '';
    })
    .filter(Boolean)
    .join('\n\n')
    .trim();
}

function extractXlsxText(buffer: Buffer): string {
  const archive = unzipSync(new Uint8Array(buffer));
  const sharedStringsXml = archive['xl/sharedStrings.xml']
    ? strFromU8(archive['xl/sharedStrings.xml']!)
    : '';
  const sharedStrings = Array.from(sharedStringsXml.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)).map(
    (match) => decodeXmlEntities(match[1] ?? '')
  );
  const sheets = Object.keys(archive)
    .filter((name) => /xl\/worksheets\/sheet\d+\.xml$/i.test(name))
    .sort((left, right) => left.localeCompare(right, undefined, { numeric: true }));

  return sheets
    .map((name, index) => {
      const xml = strFromU8(archive[name]!);
      const rows = Array.from(xml.matchAll(/<row\b[^>]*>([\s\S]*?)<\/row>/g)).map((rowMatch) => {
        const rowXml = rowMatch[1] ?? '';
        const cells = Array.from(rowXml.matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/g)).map(
          (cellMatch) => {
            const attrs = cellMatch[1] ?? '';
            const cellXml = cellMatch[2] ?? '';
            const rawValue = cellXml.match(/<v>([\s\S]*?)<\/v>/)?.[1] ?? '';
            if (/\bt="s"/.test(attrs)) {
              const sharedIndex = Number.parseInt(rawValue, 10);
              return sharedStrings[sharedIndex] ?? '';
            }
            return decodeXmlEntities(rawValue);
          }
        );
        return cells.filter(Boolean).join('\t').trim();
      });
      const body = rows.filter(Boolean).join('\n').trim();
      return body ? `Sheet ${index + 1}\n${body}` : '';
    })
    .filter(Boolean)
    .join('\n\n')
    .trim();
}

function isTextLikeFile(item: DriveItemSummary): boolean {
  const mimeType = item.mimeType?.toLowerCase();
  if (mimeType) {
    if (TEXT_MIME_PREFIXES.some((prefix) => mimeType.startsWith(prefix))) {
      return true;
    }
    if (TEXT_MIME_TYPES.has(mimeType)) {
      return true;
    }
  }
  const extension = getExtension(item.name);
  return extension ? TEXT_EXTENSIONS.has(extension) : false;
}

async function getSiteWebUrl(
  context: CustomToolHandlerContext,
  siteId: string
): Promise<string | null> {
  const response = await graphRequest<JsonRecord>(
    context,
    `/sites/${encodePathSegment(siteId)}?$select=webUrl`
  );
  return getString(response, 'webUrl');
}

async function downloadFileBuffer(
  context: CustomToolHandlerContext,
  item: DriveItemSummary
): Promise<Buffer> {
  const response = await context.graphClient.makeRequest(
    `/drives/${encodePathSegment(item.driveId ?? '')}/items/${encodePathSegment(item.id)}/content`,
    {
      accessToken: context.accessToken,
      responseType: 'buffer',
    }
  );
  return Buffer.isBuffer(response) ? response : Buffer.from([]);
}

function previewText(value: string | null, maxChars: number): string | null {
  if (!value) {
    return null;
  }
  return value.slice(0, Math.min(maxChars, 500));
}

function applyTextOperations(
  original: string,
  operations: TextEditOperation[]
): { content: string; appliedOperations: string[] } {
  let content = original;
  const appliedOperations: string[] = [];

  for (const operation of operations) {
    switch (operation.type) {
      case 'replaceAll':
        content = content.split(operation.find).join(operation.replace);
        appliedOperations.push('replaceAll');
        break;
      case 'replaceFirst':
        content = content.replace(operation.find, operation.replace);
        appliedOperations.push('replaceFirst');
        break;
      case 'append':
        content = `${content}${operation.text}`;
        appliedOperations.push('append');
        break;
      case 'prepend':
        content = `${operation.text}${content}`;
        appliedOperations.push('prepend');
        break;
      case 'overwrite':
        content = operation.text;
        appliedOperations.push('overwrite');
        break;
    }
  }

  return { content, appliedOperations };
}

function buildMutationPreview(
  operation: string,
  target: DriveItemSummary,
  extra: JsonRecord = {}
): JsonRecord {
  return {
    previewOnly: true,
    operation,
    wouldMutate: true,
    target,
    ...extra,
  };
}

async function listFolderChildren(
  context: CustomToolHandlerContext,
  path: string,
  siteId?: string,
  driveId?: string,
  cursor?: string,
  limit: number = 100
): Promise<{ items: DriveItemSummary[]; nextCursor?: string }> {
  const normalizedPath = normalizePath(path);
  let endpoint: string;

  if (cursor) {
    endpoint = decodeCursor(cursor);
  } else {
    const item = await resolveDriveItemByPath(context, normalizedPath, siteId, driveId);
    if (!item.isFolder) {
      throw new GraphToolError('unsupported_operation', `Path is not a folder: ${normalizedPath}`);
    }
    endpoint =
      `/drives/${encodePathSegment(item.driveId ?? '')}/items/${encodePathSegment(item.id)}/children` +
      `?$top=${limit}&$select=id,name,size,lastModifiedDateTime,webUrl,parentReference,folder,file,lastModifiedBy`;
  }

  const response = await graphRequest<GraphCollectionResponse<JsonRecord>>(context, endpoint);
  return {
    items: (response.value ?? []).map((entry) => normalizeDriveItem(entry)),
    nextCursor: response['@odata.nextLink'] ? encodeCursor(response['@odata.nextLink']) : undefined,
  };
}

export async function searchFilesInternal(
  context: CustomToolHandlerContext,
  query: string,
  limit: number,
  path?: string,
  siteId?: string,
  fileType?: string
): Promise<DriveItemSummary[]> {
  const response = await graphRequest<JsonRecord>(context, '/search/query', {
    method: 'POST',
    body: JSON.stringify({
      requests: [
        {
          entityTypes: ['driveItem'],
          query: { queryString: query },
          from: 0,
          size: limit * 3,
        },
      ],
    }),
  });

  const normalizedPathPrefix = path ? normalizePath(path) : undefined;
  const siteWebUrl = siteId ? await getSiteWebUrl(context, siteId) : null;
  return getArray(response, 'value')
    .flatMap((entry) => getArray(asRecord(entry), 'hitsContainers'))
    .flatMap((container) => getArray(asRecord(container), 'hits'))
    .map(asRecord)
    .filter((hit): hit is JsonRecord => !!hit)
    .map((hit) => asRecord(hit.resource))
    .filter((resource): resource is JsonRecord => !!resource)
    .filter((resource) => {
      if (!siteId) {
        return true;
      }
      const parentReference = asRecord(resource.parentReference);
      return (
        getString(parentReference, 'siteId') === siteId ||
        (typeof resource.webUrl === 'string' &&
          !!siteWebUrl &&
          resource.webUrl.startsWith(siteWebUrl))
      );
    })
    .map((resource) => normalizeDriveItem(resource))
    .filter((item) => !normalizedPathPrefix || item.path.startsWith(normalizedPathPrefix))
    .filter((item) => !fileType || item.name.toLowerCase().endsWith(`.${fileType.toLowerCase()}`))
    .slice(0, limit);
}

async function getFileContentInternal(
  context: CustomToolHandlerContext,
  item: DriveItemSummary
): Promise<{
  id: string;
  driveId: string | null;
  name: string;
  path: string;
  label: string;
  content: string | null;
  contentType: string | null;
  webUrl: string | null;
}> {
  if (item.isFolder) {
    throw new GraphToolError('unsupported_operation', 'Cannot read content from a folder.');
  }

  const content = isTextLikeFile(item)
    ? (await downloadFileBuffer(context, item)).toString('utf8')
    : null;

  return {
    id: item.id,
    driveId: item.driveId,
    name: item.name,
    path: item.path,
    label: item.label,
    content,
    contentType: item.mimeType,
    webUrl: item.webUrl,
  };
}

export async function getFileTextInternal(
  context: CustomToolHandlerContext,
  item: DriveItemSummary,
  maxChars: number
): Promise<ExtractedFileText> {
  if (item.isFolder) {
    throw new GraphToolError('unsupported_operation', 'Cannot extract text from a folder.');
  }

  const buffer = await downloadFileBuffer(context, item);
  const extension = getExtension(item.name);

  let extractedText = '';
  if (isTextLikeFile(item)) {
    extractedText = buffer.toString('utf8');
  } else if (extension === 'pdf') {
    const parser = new PDFParse({ data: buffer });
    try {
      const parsed = await parser.getText();
      extractedText = parsed.text ?? '';
    } finally {
      await parser.destroy();
    }
  } else if (extension === 'docx') {
    extractedText = extractDocxText(buffer);
  } else if (extension === 'pptx') {
    extractedText = extractPptxText(buffer);
  } else if (extension === 'xlsx') {
    extractedText = extractXlsxText(buffer);
  } else {
    return {
      id: item.id,
      driveId: item.driveId,
      name: item.name,
      path: item.path,
      label: item.label,
      text: null,
      truncated: false,
      extractionStatus: 'unsupported',
      sourceMimeType: item.mimeType,
      webUrl: item.webUrl,
    };
  }

  const normalized = normalizeExtractedText(extractedText, maxChars);
  return {
    id: item.id,
    driveId: item.driveId,
    name: item.name,
    path: item.path,
    label: item.label,
    text: normalized.text,
    truncated: normalized.truncated,
    extractionStatus: normalized.extractionStatus,
    sourceMimeType: item.mimeType,
    webUrl: item.webUrl,
  };
}

async function writeTextFile(
  context: CustomToolHandlerContext,
  path: string,
  content: string,
  siteId?: string,
  driveId?: string
): Promise<DriveItemSummary> {
  const normalizedPath = normalizePath(path);
  const reference = await buildDrivePathEndpoint(context, normalizedPath, siteId, driveId);
  const endpoint = `${reference}/content`;
  const response = await graphRequest<JsonRecord>(context, endpoint, {
    method: 'PUT',
    headers: { 'Content-Type': 'text/plain; charset=utf-8' },
    body: content,
  });
  return normalizeDriveItem(response, normalizedPath);
}

export const fileToolDefinitions: CustomToolDefinition[] = [
  {
    name: 'resolve-drive-path',
    description: 'Resolve a path-first OneDrive or SharePoint path into a concrete drive item.',
    method: 'GET',
    path: '/me/drive/root:{path}',
    requiresOrgMode: false,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Files.Read', 'Sites.Read.All'],
    schema: {
      path: z.string().min(1).describe('Absolute path such as /Projects/Q1/plan.docx'),
      siteId: z.string().min(1).optional().describe('Optional SharePoint site ID'),
      driveId: z.string().min(1).optional().describe('Optional explicit drive ID'),
    },
    handler: async (params, context) => {
      const path = getRequiredString(params, 'path');
      const siteId = getOptionalString(params, 'siteId');
      const driveId = getOptionalString(params, 'driveId');
      return success(
        context.graphClient,
        await resolveDriveItemByPath(context, path, siteId, driveId)
      );
    },
  },
  {
    name: 'list-folder',
    description: 'List a folder by path without forcing Canvas to know Graph IDs first.',
    method: 'GET',
    path: '/me/drive/root:{path}/children',
    requiresOrgMode: false,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Files.Read', 'Sites.Read.All'],
    schema: {
      path: z.string().min(1).describe('Absolute folder path'),
      siteId: z.string().min(1).optional().describe('Optional SharePoint site ID'),
      driveId: z.string().min(1).optional().describe('Optional explicit drive ID'),
      limit: z.number().int().min(1).max(100).default(100).describe('Maximum results'),
      cursor: z.string().optional().describe('Opaque pagination cursor'),
    },
    handler: async (params, context) => {
      const path = getRequiredString(params, 'path');
      const siteId = getOptionalString(params, 'siteId');
      const driveId = getOptionalString(params, 'driveId');
      const limit = getOptionalNumber(params, 'limit', 100);
      const cursor = getOptionalString(params, 'cursor');
      return success(
        context.graphClient,
        await listFolderChildren(context, path, siteId, driveId, cursor, limit)
      );
    },
  },
  {
    name: 'search-files',
    description: 'Search files by query and optionally narrow results by path or site.',
    method: 'POST',
    path: '/search/query',
    requiresOrgMode: false,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Files.Read.All', 'Sites.Read.All'],
    schema: {
      query: z.string().min(1).describe('Search query'),
      path: z.string().min(1).optional().describe('Optional path prefix filter'),
      siteId: z.string().min(1).optional().describe('Optional site ID filter'),
      fileType: z.string().min(1).optional().describe('Optional extension such as docx'),
      limit: z.number().int().min(1).max(25).default(25).describe('Maximum results'),
    },
    handler: async (params, context) => {
      const query = getRequiredString(params, 'query');
      const path = getOptionalString(params, 'path');
      const siteId = getOptionalString(params, 'siteId');
      const fileType = getOptionalString(params, 'fileType');
      const limit = getOptionalNumber(params, 'limit', 25);
      return success(context.graphClient, {
        items: await searchFilesInternal(context, query, limit, path, siteId, fileType),
      });
    },
  },
  {
    name: 'get-file-metadata',
    description: 'Get normalized file metadata from a path-first lookup.',
    method: 'GET',
    path: '/me/drive/root:{path}',
    requiresOrgMode: false,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Files.Read', 'Sites.Read.All'],
    schema: {
      path: z.string().min(1).describe('Absolute path'),
      siteId: z.string().min(1).optional(),
      driveId: z.string().min(1).optional(),
    },
    handler: async (params, context) => {
      const path = getRequiredString(params, 'path');
      const siteId = getOptionalString(params, 'siteId');
      const driveId = getOptionalString(params, 'driveId');
      return success(
        context.graphClient,
        await resolveDriveItemByPath(context, path, siteId, driveId)
      );
    },
  },
  {
    name: 'get-file-content',
    description: 'Fetch file content from a path-first lookup.',
    method: 'GET',
    path: '/me/drive/root:{path}/content',
    requiresOrgMode: false,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Files.Read', 'Sites.Read.All'],
    schema: {
      path: z.string().min(1).describe('Absolute file path'),
      siteId: z.string().min(1).optional(),
      driveId: z.string().min(1).optional(),
    },
    handler: async (params, context) => {
      const path = getRequiredString(params, 'path');
      const siteId = getOptionalString(params, 'siteId');
      const driveId = getOptionalString(params, 'driveId');
      const item = await resolveDriveItemByPath(context, path, siteId, driveId);
      return success(context.graphClient, await getFileContentInternal(context, item));
    },
  },
  {
    name: 'get-file-text',
    description: 'Extract normalized text from a file when the format supports text extraction.',
    method: 'GET',
    path: '/me/drive/root:{path}/content',
    requiresOrgMode: false,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Files.Read', 'Sites.Read.All'],
    schema: {
      path: z.string().min(1).describe('Absolute file path'),
      siteId: z.string().min(1).optional(),
      driveId: z.string().min(1).optional(),
      maxChars: z
        .number()
        .int()
        .min(100)
        .max(200000)
        .default(20000)
        .describe('Maximum extracted characters'),
    },
    handler: async (params, context) => {
      const path = getRequiredString(params, 'path');
      const siteId = getOptionalString(params, 'siteId');
      const driveId = getOptionalString(params, 'driveId');
      const maxChars = getOptionalNumber(params, 'maxChars', 20000);
      const item = await resolveDriveItemByPath(context, path, siteId, driveId);
      return success(context.graphClient, await getFileTextInternal(context, item, maxChars));
    },
  },
  {
    name: 'get-file-context',
    description: 'Return file metadata plus extracted text and a preview in one call.',
    method: 'GET',
    path: '/me/drive/root:{path}/content',
    requiresOrgMode: false,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Files.Read', 'Sites.Read.All'],
    schema: {
      path: z.string().min(1).describe('Absolute file path'),
      siteId: z.string().min(1).optional(),
      driveId: z.string().min(1).optional(),
      maxChars: z
        .number()
        .int()
        .min(100)
        .max(200000)
        .default(20000)
        .describe('Maximum extracted characters'),
    },
    handler: async (params, context) => {
      const path = getRequiredString(params, 'path');
      const siteId = getOptionalString(params, 'siteId');
      const driveId = getOptionalString(params, 'driveId');
      const maxChars = getOptionalNumber(params, 'maxChars', 20000);
      const file = await resolveDriveItemByPath(context, path, siteId, driveId);
      const extracted = await getFileTextInternal(context, file, maxChars);
      return success(context.graphClient, {
        file,
        ...(extracted.text ? { text: extracted.text } : {}),
        previewText: previewText(extracted.text, maxChars),
        extractionStatus: extracted.extractionStatus,
        truncated: extracted.truncated,
      });
    },
  },
  {
    name: 'create-text-file',
    description: 'Create a UTF-8 text file at a path-first location.',
    method: 'PUT',
    path: '/me/drive/root:{path}/content',
    requiresOrgMode: false,
    readOnlyHint: false,
    destructiveHint: true,
    openWorldHint: true,
    scopes: ['Files.ReadWrite', 'Sites.ReadWrite.All'],
    schema: {
      path: z.string().min(1).describe('Absolute file path'),
      content: z.string().describe('UTF-8 text content'),
      siteId: z.string().min(1).optional(),
      driveId: z.string().min(1).optional(),
      previewOnly: z.boolean().optional().describe('Preview without creating the file'),
    },
    handler: async (params, context) => {
      const path = getRequiredString(params, 'path');
      const siteId = getOptionalString(params, 'siteId');
      const driveId = getOptionalString(params, 'driveId');
      const previewOnly = getOptionalBoolean(params, 'previewOnly', false);
      try {
        const existing = await resolveDriveItemByPath(context, path, siteId, driveId);
        return errorResult(context.graphClient, 'unsupported_operation', {
          message: `File already exists: ${path}`,
          target: existing,
        });
      } catch (error) {
        if (
          (error instanceof GraphToolError && error.code !== 'not_found') ||
          (!(error instanceof GraphToolError) && !(error as Error).message.includes('404'))
        ) {
          throw error;
        }
        const previewTarget = {
          id: '',
          driveId: driveId ?? null,
          name: path.split('/').at(-1) ?? path,
          path: normalizePath(path),
          isFolder: false,
          mimeType: 'text/plain',
          size: getRequiredString(params, 'content').length,
          lastModifiedDateTime: null,
          modifiedBy: null,
          webUrl: null,
          label: path.split('/').at(-1) ?? path,
        } satisfies DriveItemSummary;
        if (previewOnly) {
          return success(
            context.graphClient,
            buildMutationPreview('create-text-file', previewTarget, {
              size: getRequiredString(params, 'content').length,
            })
          );
        }
        const item = await writeTextFile(
          context,
          path,
          getRequiredString(params, 'content'),
          siteId,
          driveId
        );
        return success(context.graphClient, item);
      }
    },
  },
  {
    name: 'update-text-file',
    description: 'Update a UTF-8 text file at a path-first location.',
    method: 'PUT',
    path: '/me/drive/root:{path}/content',
    requiresOrgMode: false,
    readOnlyHint: false,
    destructiveHint: true,
    openWorldHint: true,
    scopes: ['Files.ReadWrite', 'Sites.ReadWrite.All'],
    schema: {
      path: z.string().min(1).describe('Absolute file path'),
      content: z.string().describe('UTF-8 text content'),
      siteId: z.string().min(1).optional(),
      driveId: z.string().min(1).optional(),
      previewOnly: z.boolean().optional().describe('Preview without writing the file'),
    },
    handler: async (params, context) => {
      const path = getRequiredString(params, 'path');
      const siteId = getOptionalString(params, 'siteId');
      const driveId = getOptionalString(params, 'driveId');
      const previewOnly = getOptionalBoolean(params, 'previewOnly', false);
      const item = await resolveDriveItemByPath(context, path, siteId, driveId);
      if (!isTextLikeFile(item)) {
        throw new GraphToolError(
          'unsupported_operation',
          'update-text-file only supports text-like files.'
        );
      }
      if (previewOnly) {
        return success(
          context.graphClient,
          buildMutationPreview('update-text-file', item, {
            size: getRequiredString(params, 'content').length,
          })
        );
      }
      const updated = await writeTextFile(
        context,
        path,
        getRequiredString(params, 'content'),
        siteId,
        driveId
      );
      return success(context.graphClient, updated);
    },
  },
  {
    name: 'edit-text-file',
    description: 'Apply structured text edit operations to a text file with optional preview mode.',
    method: 'PUT',
    path: '/me/drive/root:{path}/content',
    requiresOrgMode: false,
    readOnlyHint: false,
    destructiveHint: true,
    openWorldHint: true,
    scopes: ['Files.ReadWrite', 'Sites.ReadWrite.All'],
    schema: {
      path: z.string().min(1).describe('Absolute file path'),
      siteId: z.string().min(1).optional(),
      driveId: z.string().min(1).optional(),
      operations: z
        .array(textEditOperationSchema)
        .min(1)
        .describe('Structured text edit operations'),
      previewOnly: z.boolean().optional().describe('Preview the edit without writing the file'),
    },
    handler: async (params, context) => {
      const path = getRequiredString(params, 'path');
      const siteId = getOptionalString(params, 'siteId');
      const driveId = getOptionalString(params, 'driveId');
      const previewOnly = getOptionalBoolean(params, 'previewOnly', false);
      const item = await resolveDriveItemByPath(context, path, siteId, driveId);
      if (!isTextLikeFile(item)) {
        throw new GraphToolError(
          'unsupported_operation',
          'edit-text-file only supports text-like files.'
        );
      }
      const current = await getFileTextInternal(context, item, 1000000);
      const original = current.text ?? '';
      const operations = textEditOperationSchema
        .array()
        .parse(params.operations) as TextEditOperation[];
      const result = applyTextOperations(original, operations);
      const changedCharacters = Math.abs(result.content.length - original.length);

      if (previewOnly) {
        return success(context.graphClient, {
          previewOnly: true,
          file: item,
          operationsApplied: result.appliedOperations,
          changedCharacters,
          previewText: previewText(result.content, 4000),
        });
      }

      const updated = await writeTextFile(context, path, result.content, siteId, driveId);
      return success(context.graphClient, {
        file: updated,
        operationsApplied: result.appliedOperations,
        changedCharacters,
      });
    },
  },
  {
    name: 'delete-file',
    description: 'Delete a file or folder from a path-first location.',
    method: 'DELETE',
    path: '/drives/{driveId}/items/{itemId}',
    requiresOrgMode: false,
    readOnlyHint: false,
    destructiveHint: true,
    openWorldHint: true,
    scopes: ['Files.ReadWrite', 'Sites.ReadWrite.All'],
    schema: {
      path: z.string().min(1).describe('Absolute path'),
      siteId: z.string().min(1).optional(),
      driveId: z.string().min(1).optional(),
      previewOnly: z.boolean().optional().describe('Preview without deleting the file'),
    },
    handler: async (params, context) => {
      const path = getRequiredString(params, 'path');
      const siteId = getOptionalString(params, 'siteId');
      const driveId = getOptionalString(params, 'driveId');
      const previewOnly = getOptionalBoolean(params, 'previewOnly', false);
      const item = await resolveDriveItemByPath(context, path, siteId, driveId);
      if (previewOnly) {
        return success(context.graphClient, buildMutationPreview('delete-file', item));
      }
      await graphRequest<JsonRecord>(
        context,
        `/drives/${encodePathSegment(item.driveId ?? '')}/items/${encodePathSegment(item.id)}`,
        { method: 'DELETE' }
      );
      return success(context.graphClient, {
        success: true,
        id: item.id,
        driveId: item.driveId,
        path: item.path,
        label: item.label,
      });
    },
  },
];
