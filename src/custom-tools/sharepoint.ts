import { z } from 'zod';
import {
  CustomToolDefinition,
  CustomToolHandlerContext,
  DriveItemSummary,
  GraphCollectionResponse,
  GraphToolError,
  JsonRecord,
  SiteSummary,
  asRecord,
  buildBodyPreview,
  decodeCursor,
  encodeCursor,
  encodePathSegment,
  getArray,
  getOptionalNumber,
  getOptionalString,
  getRequiredString,
  getString,
  graphRequest,
  normalizeDriveItem,
  normalizePath,
  normalizeSite,
  success,
} from './shared.js';

type SiteDriveSummary = {
  driveId: string;
  name: string | null;
  displayName: string | null;
  driveType: string | null;
  webUrl: string | null;
  label: string;
};

type SharePointPageSummary = {
  pageId: string;
  title: string | null;
  name: string | null;
  webUrl: string | null;
  lastModifiedDateTime: string | null;
  label: string;
};

type SharePointListSummary = {
  listId: string;
  name: string | null;
  displayName: string | null;
  webUrl: string | null;
  label: string;
};

type SharePointListItemSummary = {
  itemId: string;
  webUrl: string | null;
  fields: JsonRecord;
  label: string;
};

type SharePointContentSummary = {
  id: string;
  siteId: string | null;
  title: string | null;
  webUrl: string | null;
  contentType: string | null;
  bodyPreview: string | null;
  label: string;
};

type SearchCursorState = {
  kind: 'site_files' | 'sharepoint_content';
  query: string;
  limit: number;
  offset: number;
  siteId?: string;
  path?: string;
};

function buildLibraryLabel(
  entry: Pick<SiteDriveSummary, 'displayName' | 'name' | 'driveId'>
): string {
  return entry.displayName ?? entry.name ?? entry.driveId;
}

function buildPageLabel(entry: Pick<SharePointPageSummary, 'title' | 'name' | 'pageId'>): string {
  return entry.title ?? entry.name ?? entry.pageId;
}

function buildListLabel(
  entry: Pick<SharePointListSummary, 'displayName' | 'name' | 'listId'>
): string {
  return entry.displayName ?? entry.name ?? entry.listId;
}

function buildListItemLabel(fields: JsonRecord, itemId: string): string {
  return (
    getString(fields, 'Title') ??
    getString(fields, 'LinkTitle') ??
    getString(fields, 'FileLeafRef') ??
    getString(fields, 'Name') ??
    itemId
  );
}

function buildSharePointContentLabel(
  entry: Pick<SharePointContentSummary, 'title' | 'webUrl' | 'id'>
): string {
  return entry.title ?? entry.webUrl ?? entry.id;
}

function encodeSearchCursor(state: SearchCursorState): string {
  return Buffer.from(JSON.stringify(state), 'utf8').toString('base64url');
}

function decodeSearchCursor(cursor: string): SearchCursorState {
  try {
    const parsed = JSON.parse(
      Buffer.from(cursor, 'base64url').toString('utf8')
    ) as SearchCursorState;
    if (
      !parsed ||
      (parsed.kind !== 'site_files' && parsed.kind !== 'sharepoint_content') ||
      typeof parsed.query !== 'string' ||
      typeof parsed.limit !== 'number' ||
      typeof parsed.offset !== 'number'
    ) {
      throw new Error('invalid');
    }
    return parsed;
  } catch {
    throw new GraphToolError('invalid_cursor', 'Invalid pagination cursor.');
  }
}

function normalizeLibrary(entry: JsonRecord): SiteDriveSummary {
  const normalized = {
    driveId: getString(entry, 'id') ?? '',
    name: getString(entry, 'name'),
    displayName: getString(entry, 'displayName') ?? getString(entry, 'name'),
    driveType: getString(entry, 'driveType'),
    webUrl: getString(entry, 'webUrl'),
  };
  return {
    ...normalized,
    label: buildLibraryLabel(normalized),
  };
}

function normalizePage(entry: JsonRecord): SharePointPageSummary {
  const normalized = {
    pageId: getString(entry, 'id') ?? '',
    title: getString(entry, 'title') ?? getString(entry, 'name'),
    name: getString(entry, 'name'),
    webUrl: getString(entry, 'webUrl'),
    lastModifiedDateTime: getString(entry, 'lastModifiedDateTime'),
  };
  return {
    ...normalized,
    label: buildPageLabel(normalized),
  };
}

function normalizeSharePointList(entry: JsonRecord): SharePointListSummary {
  const normalized = {
    listId: getString(entry, 'id') ?? '',
    name: getString(entry, 'name'),
    displayName: getString(entry, 'displayName') ?? getString(entry, 'name'),
    webUrl: getString(entry, 'webUrl'),
  };
  return {
    ...normalized,
    label: buildListLabel(normalized),
  };
}

function normalizeSharePointListItem(entry: JsonRecord): SharePointListItemSummary {
  const fields = asRecord(entry.fields) ?? {};
  return {
    itemId: getString(entry, 'id') ?? '',
    webUrl: getString(entry, 'webUrl'),
    fields,
    label: buildListItemLabel(fields, getString(entry, 'id') ?? ''),
  };
}

function normalizeSharePointContent(resource: JsonRecord): SharePointContentSummary {
  const normalized = {
    id: getString(resource, 'id') ?? '',
    siteId: getString(resource, 'siteId'),
    title: getString(resource, 'title') ?? getString(resource, 'name'),
    webUrl: getString(resource, 'webUrl'),
    contentType: getString(resource, 'contentclass') ?? getString(resource, '@odata.type'),
    bodyPreview: buildBodyPreview(resource),
  };
  return {
    ...normalized,
    label: buildSharePointContentLabel(normalized),
  };
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

async function listSitesInternal(
  context: CustomToolHandlerContext,
  query: string | undefined,
  limit: number,
  cursor?: string
): Promise<{ items: SiteSummary[]; nextCursor?: string }> {
  const endpoint = cursor
    ? decodeCursor(cursor)
    : query
      ? `/sites?$search=${encodeURIComponent(`"${query}"`)}&$top=${limit}&$select=id,name,displayName,webUrl`
      : `/sites?$search=${encodeURIComponent('"*"')}&$top=${limit}&$select=id,name,displayName,webUrl`;
  const response = await graphRequest<GraphCollectionResponse<JsonRecord>>(context, endpoint, {
    headers: { ConsistencyLevel: 'eventual' },
  });

  const items = (response.value ?? []).map(normalizeSite).slice(0, limit);
  return {
    items,
    nextCursor: response['@odata.nextLink'] ? encodeCursor(response['@odata.nextLink']) : undefined,
  };
}

async function listSiteDrivesInternal(
  context: CustomToolHandlerContext,
  siteId: string,
  limit: number,
  cursor?: string
): Promise<{
  items: SiteDriveSummary[];
  nextCursor?: string;
}> {
  const endpoint = cursor
    ? decodeCursor(cursor)
    : `/sites/${encodePathSegment(siteId)}/drives?$top=${limit}&$select=id,name,driveType,webUrl`;
  const response = await graphRequest<GraphCollectionResponse<JsonRecord>>(context, endpoint);
  return {
    items: (response.value ?? []).map(normalizeLibrary),
    nextCursor: response['@odata.nextLink'] ? encodeCursor(response['@odata.nextLink']) : undefined,
  };
}

async function searchDriveItemsPage(
  context: CustomToolHandlerContext,
  query: string,
  limit: number,
  siteId?: string,
  path?: string,
  cursor?: string
): Promise<{ items: DriveItemSummary[]; nextCursor?: string }> {
  const normalizedPath = path ? normalizePath(path) : undefined;
  const cursorState = cursor ? decodeSearchCursor(cursor) : undefined;
  if (cursorState && cursorState.kind !== 'site_files') {
    throw new GraphToolError('invalid_cursor', 'Invalid pagination cursor.');
  }
  const offset = cursorState?.offset ?? 0;
  const searchLimit = cursorState?.limit ?? limit;
  const pageSize = Math.max(searchLimit * 3, 25);
  const response = await graphRequest<JsonRecord>(context, '/search/query', {
    method: 'POST',
    body: JSON.stringify({
      requests: [
        {
          entityTypes: ['driveItem'],
          query: { queryString: query },
          from: offset,
          size: pageSize,
        },
      ],
    }),
  });
  const siteWebUrl = siteId ? await getSiteWebUrl(context, siteId) : null;
  const containers = getArray(response, 'value').flatMap((entry) =>
    getArray(asRecord(entry), 'hitsContainers')
  );
  const hits = containers
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
    .filter((item) => !normalizedPath || item.path.startsWith(normalizedPath))
    .slice(0, searchLimit);

  const hasMore = containers.some(
    (container) => asRecord(container)?.moreResultsAvailable === true
  );
  return {
    items: hits,
    nextCursor: hasMore
      ? encodeSearchCursor({
          kind: 'site_files',
          query,
          limit: searchLimit,
          offset: offset + pageSize,
          siteId,
          path: normalizedPath,
        })
      : undefined,
  };
}

export async function searchSharePointContentInternal(
  context: CustomToolHandlerContext,
  query: string,
  limit: number,
  siteId?: string
): Promise<SharePointContentSummary[]> {
  const { items } = await searchSharePointContentPage(context, query, limit, siteId);
  return items;
}

async function searchSharePointContentPage(
  context: CustomToolHandlerContext,
  query: string,
  limit: number,
  siteId?: string,
  cursor?: string
): Promise<{ items: SharePointContentSummary[]; nextCursor?: string }> {
  const cursorState = cursor ? decodeSearchCursor(cursor) : undefined;
  if (cursorState && cursorState.kind !== 'sharepoint_content') {
    throw new GraphToolError('invalid_cursor', 'Invalid pagination cursor.');
  }
  const offset = cursorState?.offset ?? 0;
  const searchLimit = cursorState?.limit ?? limit;
  const pageSize = Math.max(searchLimit * 3, 25);
  const response = await graphRequest<JsonRecord>(context, '/search/query', {
    method: 'POST',
    body: JSON.stringify({
      requests: [
        {
          entityTypes: ['listItem', 'site', 'driveItem'],
          query: { queryString: query },
          from: offset,
          size: pageSize,
        },
      ],
    }),
  });

  const siteWebUrl = siteId ? await getSiteWebUrl(context, siteId) : null;
  const containers = getArray(response, 'value').flatMap((entry) =>
    getArray(asRecord(entry), 'hitsContainers')
  );
  const items = containers
    .flatMap((container) => getArray(asRecord(container), 'hits'))
    .map(asRecord)
    .filter((hit): hit is JsonRecord => !!hit)
    .map((hit) => asRecord(hit.resource))
    .filter((resource): resource is JsonRecord => !!resource)
    .filter((resource) => {
      if (!siteId) {
        return true;
      }
      return (
        getString(resource, 'siteId') === siteId ||
        (typeof resource.webUrl === 'string' &&
          !!siteWebUrl &&
          resource.webUrl.startsWith(siteWebUrl))
      );
    })
    .map(normalizeSharePointContent)
    .slice(0, searchLimit);
  const hasMore = containers.some(
    (container) => asRecord(container)?.moreResultsAvailable === true
  );

  return {
    items,
    nextCursor: hasMore
      ? encodeSearchCursor({
          kind: 'sharepoint_content',
          query,
          limit: searchLimit,
          offset: offset + pageSize,
          siteId,
        })
      : undefined,
  };
}

async function listSharePointPagesInternal(
  context: CustomToolHandlerContext,
  siteId: string,
  limit: number,
  cursor?: string
): Promise<{ items: SharePointPageSummary[]; nextCursor?: string }> {
  const endpoint = cursor
    ? decodeCursor(cursor)
    : `/sites/${encodePathSegment(siteId)}/pages/microsoft.graph.sitePage?$top=${limit}&$select=id,name,title,webUrl,lastModifiedDateTime`;
  const response = await graphRequest<GraphCollectionResponse<JsonRecord>>(context, endpoint);

  return {
    items: (response.value ?? []).map(normalizePage),
    nextCursor: response['@odata.nextLink'] ? encodeCursor(response['@odata.nextLink']) : undefined,
  };
}

async function listSharePointListsInternal(
  context: CustomToolHandlerContext,
  siteId: string,
  limit: number,
  cursor?: string
): Promise<{
  items: SharePointListSummary[];
  nextCursor?: string;
}> {
  const endpoint = cursor
    ? decodeCursor(cursor)
    : `/sites/${encodePathSegment(siteId)}/lists?$top=${limit}&$select=id,name,displayName,webUrl`;
  const response = await graphRequest<GraphCollectionResponse<JsonRecord>>(context, endpoint);
  return {
    items: (response.value ?? []).map(normalizeSharePointList),
    nextCursor: response['@odata.nextLink'] ? encodeCursor(response['@odata.nextLink']) : undefined,
  };
}

async function listSharePointListItemsInternal(
  context: CustomToolHandlerContext,
  siteId: string,
  listId: string,
  limit: number,
  cursor?: string
): Promise<{
  items: SharePointListItemSummary[];
  nextCursor?: string;
}> {
  const endpoint = cursor
    ? decodeCursor(cursor)
    : `/sites/${encodePathSegment(siteId)}/lists/${encodePathSegment(listId)}/items?$expand=fields&$top=${limit}`;
  const response = await graphRequest<GraphCollectionResponse<JsonRecord>>(context, endpoint);
  return {
    items: (response.value ?? []).map(normalizeSharePointListItem),
    nextCursor: response['@odata.nextLink'] ? encodeCursor(response['@odata.nextLink']) : undefined,
  };
}

export const sharepointToolDefinitions: CustomToolDefinition[] = [
  {
    name: 'list-sites',
    description: 'List accessible SharePoint sites with normalized site summaries.',
    method: 'GET',
    path: '/sites',
    requiresOrgMode: true,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Sites.Read.All'],
    schema: {
      limit: z.number().int().min(1).max(100).default(25).describe('Maximum results'),
      cursor: z.string().optional().describe('Opaque pagination cursor'),
    },
    handler: async (params, context) => {
      const limit = getOptionalNumber(params, 'limit', 25);
      const cursor = getOptionalString(params, 'cursor');
      return success(
        context.graphClient,
        await listSitesInternal(context, undefined, limit, cursor)
      );
    },
  },
  {
    name: 'search-sites',
    description: 'Search SharePoint sites by name or query text.',
    method: 'GET',
    path: '/sites',
    requiresOrgMode: true,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Sites.Read.All'],
    schema: {
      query: z.string().min(1).describe('Search query'),
      limit: z.number().int().min(1).max(100).default(25).describe('Maximum results'),
      cursor: z.string().optional().describe('Opaque pagination cursor'),
    },
    handler: async (params, context) => {
      const query = getRequiredString(params, 'query');
      const limit = getOptionalNumber(params, 'limit', 25);
      const cursor = getOptionalString(params, 'cursor');
      return success(context.graphClient, await listSitesInternal(context, query, limit, cursor));
    },
  },
  {
    name: 'list-site-drives',
    description: 'List document libraries or drives available on a SharePoint site.',
    method: 'GET',
    path: '/sites/{siteId}/drives',
    requiresOrgMode: true,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Sites.Read.All'],
    schema: {
      siteId: z.string().min(1).describe('SharePoint site ID'),
      limit: z.number().int().min(1).max(100).default(50).describe('Maximum results'),
      cursor: z.string().optional().describe('Opaque pagination cursor'),
    },
    handler: async (params, context) => {
      const siteId = getRequiredString(params, 'siteId');
      const limit = getOptionalNumber(params, 'limit', 50);
      const cursor = getOptionalString(params, 'cursor');
      return success(
        context.graphClient,
        await listSiteDrivesInternal(context, siteId, limit, cursor)
      );
    },
  },
  {
    name: 'list-document-libraries',
    description: 'List SharePoint document libraries with drive metadata.',
    method: 'GET',
    path: '/sites/{siteId}/drives',
    requiresOrgMode: true,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Sites.Read.All'],
    schema: {
      siteId: z.string().min(1).describe('SharePoint site ID'),
      limit: z.number().int().min(1).max(100).default(50).describe('Maximum results'),
      cursor: z.string().optional().describe('Opaque pagination cursor'),
    },
    handler: async (params, context) => {
      const siteId = getRequiredString(params, 'siteId');
      const limit = getOptionalNumber(params, 'limit', 50);
      const cursor = getOptionalString(params, 'cursor');
      return success(
        context.graphClient,
        await listSiteDrivesInternal(context, siteId, limit, cursor)
      );
    },
  },
  {
    name: 'search-site-files',
    description: 'Search files scoped to one SharePoint site and return normalized drive items.',
    method: 'POST',
    path: '/search/query',
    requiresOrgMode: true,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Sites.Read.All'],
    schema: {
      siteId: z.string().min(1).describe('SharePoint site ID'),
      query: z.string().min(1).describe('Search query'),
      path: z.string().min(1).optional().describe('Optional path prefix filter'),
      limit: z.number().int().min(1).max(50).default(25).describe('Maximum results'),
      cursor: z.string().optional().describe('Opaque pagination cursor'),
    },
    handler: async (params, context) => {
      const siteId = getRequiredString(params, 'siteId');
      const query = getRequiredString(params, 'query');
      const path = getOptionalString(params, 'path');
      const limit = getOptionalNumber(params, 'limit', 25);
      const cursor = getOptionalString(params, 'cursor');
      return success(
        context.graphClient,
        await searchDriveItemsPage(context, query, limit, siteId, path, cursor)
      );
    },
  },
  {
    name: 'search-sharepoint-content',
    description: 'Search SharePoint content across files, lists, and sites.',
    method: 'POST',
    path: '/search/query',
    requiresOrgMode: true,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Sites.Read.All'],
    schema: {
      query: z.string().min(1).describe('Search query'),
      siteId: z.string().min(1).optional().describe('Optional SharePoint site ID filter'),
      limit: z.number().int().min(1).max(50).default(25).describe('Maximum results'),
      cursor: z.string().optional().describe('Opaque pagination cursor'),
    },
    handler: async (params, context) => {
      const query = getRequiredString(params, 'query');
      const siteId = getOptionalString(params, 'siteId');
      const limit = getOptionalNumber(params, 'limit', 25);
      const cursor = getOptionalString(params, 'cursor');
      return success(
        context.graphClient,
        await searchSharePointContentPage(context, query, limit, siteId, cursor)
      );
    },
  },
  {
    name: 'list-sharepoint-pages',
    description: 'List modern SharePoint pages on a site.',
    method: 'GET',
    path: '/sites/{siteId}/pages/microsoft.graph.sitePage',
    requiresOrgMode: true,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Sites.Read.All'],
    schema: {
      siteId: z.string().min(1).describe('SharePoint site ID'),
      limit: z.number().int().min(1).max(100).default(25).describe('Maximum results'),
      cursor: z.string().optional().describe('Opaque pagination cursor'),
    },
    handler: async (params, context) => {
      const siteId = getRequiredString(params, 'siteId');
      const limit = getOptionalNumber(params, 'limit', 25);
      const cursor = getOptionalString(params, 'cursor');
      return success(
        context.graphClient,
        await listSharePointPagesInternal(context, siteId, limit, cursor)
      );
    },
  },
  {
    name: 'list-sharepoint-lists',
    description: 'List SharePoint lists for a site.',
    method: 'GET',
    path: '/sites/{siteId}/lists',
    requiresOrgMode: true,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Sites.Read.All'],
    schema: {
      siteId: z.string().min(1).describe('SharePoint site ID'),
      limit: z.number().int().min(1).max(100).default(25).describe('Maximum results'),
      cursor: z.string().optional().describe('Opaque pagination cursor'),
    },
    handler: async (params, context) => {
      const siteId = getRequiredString(params, 'siteId');
      const limit = getOptionalNumber(params, 'limit', 25);
      const cursor = getOptionalString(params, 'cursor');
      return success(
        context.graphClient,
        await listSharePointListsInternal(context, siteId, limit, cursor)
      );
    },
  },
  {
    name: 'list-sharepoint-list-items',
    description: 'List items from a SharePoint list with expanded fields.',
    method: 'GET',
    path: '/sites/{siteId}/lists/{listId}/items',
    requiresOrgMode: true,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Sites.Read.All'],
    schema: {
      siteId: z.string().min(1).describe('SharePoint site ID'),
      listId: z.string().min(1).describe('SharePoint list ID'),
      limit: z.number().int().min(1).max(100).default(25).describe('Maximum results'),
      cursor: z.string().optional().describe('Opaque pagination cursor'),
    },
    handler: async (params, context) => {
      const siteId = getRequiredString(params, 'siteId');
      const listId = getRequiredString(params, 'listId');
      const limit = getOptionalNumber(params, 'limit', 25);
      const cursor = getOptionalString(params, 'cursor');
      return success(
        context.graphClient,
        await listSharePointListItemsInternal(context, siteId, listId, limit, cursor)
      );
    },
  },
  {
    name: 'search-sharepoint-pages',
    description: 'Search SharePoint pages by title and preview text.',
    method: 'POST',
    path: '/search/query',
    requiresOrgMode: true,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Sites.Read.All'],
    schema: {
      query: z.string().min(1).describe('Search query'),
      siteId: z.string().min(1).optional().describe('Optional SharePoint site ID filter'),
      limit: z.number().int().min(1).max(50).default(25).describe('Maximum results'),
    },
    handler: async (params, context) => {
      const query = getRequiredString(params, 'query');
      const siteId = getOptionalString(params, 'siteId');
      const limit = getOptionalNumber(params, 'limit', 25);
      if (siteId) {
        const pages = await listSharePointPagesInternal(context, siteId, 100);
        const loweredQuery = query.toLowerCase();
        return success(context.graphClient, {
          items: pages.items
            .filter(
              (page) =>
                page.title?.toLowerCase().includes(loweredQuery) ||
                page.name?.toLowerCase().includes(loweredQuery)
            )
            .slice(0, limit),
        });
      }

      const items = await searchSharePointContentInternal(context, query, limit);
      return success(context.graphClient, {
        items: items.filter((item) => item.contentType?.toLowerCase().includes('page')),
      });
    },
  },
];
