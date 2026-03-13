import { z } from 'zod';
import {
  CustomToolDefinition,
  CustomToolHandlerContext,
  DriveItemSummary,
  GraphCollectionResponse,
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

type SharePointPageSummary = {
  pageId: string;
  title: string | null;
  name: string | null;
  webUrl: string | null;
  lastModifiedDateTime: string | null;
};

function normalizePage(entry: JsonRecord): SharePointPageSummary {
  return {
    pageId: getString(entry, 'id') ?? '',
    title: getString(entry, 'title') ?? getString(entry, 'name'),
    name: getString(entry, 'name'),
    webUrl: getString(entry, 'webUrl'),
    lastModifiedDateTime: getString(entry, 'lastModifiedDateTime'),
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
  items: Array<{
    driveId: string;
    name: string | null;
    driveType: string | null;
    webUrl: string | null;
  }>;
  nextCursor?: string;
}> {
  const endpoint = cursor
    ? decodeCursor(cursor)
    : `/sites/${encodePathSegment(siteId)}/drives?$top=${limit}&$select=id,name,driveType,webUrl`;
  const response = await graphRequest<GraphCollectionResponse<JsonRecord>>(context, endpoint);
  return {
    items: (response.value ?? []).map((entry) => ({
      driveId: getString(entry, 'id') ?? '',
      name: getString(entry, 'name'),
      driveType: getString(entry, 'driveType'),
      webUrl: getString(entry, 'webUrl'),
    })),
    nextCursor: response['@odata.nextLink'] ? encodeCursor(response['@odata.nextLink']) : undefined,
  };
}

async function searchDriveItems(
  context: CustomToolHandlerContext,
  query: string,
  limit: number,
  siteId?: string,
  path?: string
): Promise<DriveItemSummary[]> {
  const response = await graphRequest<JsonRecord>(context, '/search/query', {
    method: 'POST',
    body: JSON.stringify({
      requests: [
        {
          entityTypes: ['driveItem'],
          query: { queryString: query },
          from: 0,
          size: Math.max(limit * 3, 25),
        },
      ],
    }),
  });

  const normalizedPath = path ? normalizePath(path) : undefined;
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
    .filter((item) => !normalizedPath || item.path.startsWith(normalizedPath))
    .slice(0, limit);
}

async function searchSharePointContentInternal(
  context: CustomToolHandlerContext,
  query: string,
  limit: number,
  siteId?: string
): Promise<
  Array<{
    id: string;
    siteId: string | null;
    title: string | null;
    webUrl: string | null;
    contentType: string | null;
    bodyPreview: string | null;
  }>
> {
  const response = await graphRequest<JsonRecord>(context, '/search/query', {
    method: 'POST',
    body: JSON.stringify({
      requests: [
        {
          entityTypes: ['listItem', 'site', 'driveItem'],
          query: { queryString: query },
          from: 0,
          size: Math.max(limit * 3, 25),
        },
      ],
    }),
  });

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
      return (
        getString(resource, 'siteId') === siteId ||
        (typeof resource.webUrl === 'string' &&
          !!siteWebUrl &&
          resource.webUrl.startsWith(siteWebUrl))
      );
    })
    .map((resource) => ({
      id: getString(resource, 'id') ?? '',
      siteId: getString(resource, 'siteId'),
      title: getString(resource, 'title') ?? getString(resource, 'name'),
      webUrl: getString(resource, 'webUrl'),
      contentType: getString(resource, 'contentclass') ?? getString(resource, '@odata.type'),
      bodyPreview: buildBodyPreview(resource),
    }))
    .slice(0, limit);
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
  items: Array<{
    listId: string;
    name: string | null;
    displayName: string | null;
    webUrl: string | null;
  }>;
  nextCursor?: string;
}> {
  const endpoint = cursor
    ? decodeCursor(cursor)
    : `/sites/${encodePathSegment(siteId)}/lists?$top=${limit}&$select=id,name,displayName,webUrl`;
  const response = await graphRequest<GraphCollectionResponse<JsonRecord>>(context, endpoint);
  return {
    items: (response.value ?? []).map((entry) => ({
      listId: getString(entry, 'id') ?? '',
      name: getString(entry, 'name'),
      displayName: getString(entry, 'displayName') ?? getString(entry, 'name'),
      webUrl: getString(entry, 'webUrl'),
    })),
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
  items: Array<{
    itemId: string;
    webUrl: string | null;
    fields: JsonRecord;
  }>;
  nextCursor?: string;
}> {
  const endpoint = cursor
    ? decodeCursor(cursor)
    : `/sites/${encodePathSegment(siteId)}/lists/${encodePathSegment(listId)}/items?$expand=fields&$top=${limit}`;
  const response = await graphRequest<GraphCollectionResponse<JsonRecord>>(context, endpoint);
  return {
    items: (response.value ?? []).map((entry) => ({
      itemId: getString(entry, 'id') ?? '',
      webUrl: getString(entry, 'webUrl'),
      fields: asRecord(entry.fields) ?? {},
    })),
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
    },
    handler: async (params, context) => {
      const siteId = getRequiredString(params, 'siteId');
      const query = getRequiredString(params, 'query');
      const path = getOptionalString(params, 'path');
      const limit = getOptionalNumber(params, 'limit', 25);
      return success(context.graphClient, {
        items: await searchDriveItems(context, query, limit, siteId, path),
      });
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
    },
    handler: async (params, context) => {
      const query = getRequiredString(params, 'query');
      const siteId = getOptionalString(params, 'siteId');
      const limit = getOptionalNumber(params, 'limit', 25);
      return success(context.graphClient, {
        items: await searchSharePointContentInternal(context, query, limit, siteId),
      });
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
