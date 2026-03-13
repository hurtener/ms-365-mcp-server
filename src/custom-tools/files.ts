import { z } from 'zod';
import {
  CustomToolDefinition,
  CustomToolHandlerContext,
  DriveItemSummary,
  GraphCollectionResponse,
  GraphToolError,
  JsonRecord,
  buildDrivePathEndpoint,
  decodeCursor,
  encodeCursor,
  encodePathSegment,
  errorResult,
  asRecord,
  getArray,
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

async function searchFilesInternal(
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
  const results = getArray(response, 'value')
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

  return results;
}

async function getFileContentInternal(
  context: CustomToolHandlerContext,
  item: DriveItemSummary
): Promise<{
  id: string;
  driveId: string | null;
  name: string;
  path: string;
  content: string | null;
  contentType: string | null;
  webUrl: string | null;
}> {
  if (item.isFolder) {
    throw new GraphToolError('unsupported_operation', 'Cannot read content from a folder.');
  }

  const response = await graphRequest<{ rawResponse?: string; message?: string }>(
    context,
    `/drives/${encodePathSegment(item.driveId ?? '')}/items/${encodePathSegment(item.id)}/content`
  );

  return {
    id: item.id,
    driveId: item.driveId,
    name: item.name,
    path: item.path,
    content: response.rawResponse ?? null,
    contentType: item.mimeType,
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
      return success(context.graphClient, await resolveDriveItemByPath(context, path, siteId, driveId));
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
      return success(context.graphClient, await resolveDriveItemByPath(context, path, siteId, driveId));
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
    },
    handler: async (params, context) => {
      const path = getRequiredString(params, 'path');
      const siteId = getOptionalString(params, 'siteId');
      const driveId = getOptionalString(params, 'driveId');
      try {
        await resolveDriveItemByPath(context, path, siteId, driveId);
        return errorResult(context.graphClient, 'unsupported_operation', {
          message: `File already exists: ${path}`,
        });
      } catch (error) {
        if (
          (error instanceof GraphToolError && error.code !== 'not_found') ||
          (!(error instanceof GraphToolError) && !(error as Error).message.includes('404'))
        ) {
          throw error;
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
    },
    handler: async (params, context) => {
      const path = getRequiredString(params, 'path');
      const siteId = getOptionalString(params, 'siteId');
      const driveId = getOptionalString(params, 'driveId');
      await resolveDriveItemByPath(context, path, siteId, driveId);
      const item = await writeTextFile(
        context,
        path,
        getRequiredString(params, 'content'),
        siteId,
        driveId
      );
      return success(context.graphClient, item);
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
    },
    handler: async (params, context) => {
      const path = getRequiredString(params, 'path');
      const siteId = getOptionalString(params, 'siteId');
      const driveId = getOptionalString(params, 'driveId');
      const item = await resolveDriveItemByPath(context, path, siteId, driveId);
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
      });
    },
  },
];
