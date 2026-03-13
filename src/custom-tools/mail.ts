import { z } from 'zod';
import {
  AttachmentSummary,
  CustomToolDefinition,
  CustomToolHandlerContext,
  GraphCollectionResponse,
  JsonRecord,
  decodeCursor,
  encodeCursor,
  encodePathSegment,
  getOptionalNumber,
  getOptionalString,
  getRequiredString,
  getRequiredStringArray,
  getString,
  graphRequest,
  normalizeAttachment,
  normalizeMail,
  resolveQueryToBestUser,
  searchUsersInternal,
  success,
} from './shared.js';

const MAIL_SELECT =
  '$select=id,subject,from,toRecipients,receivedDateTime,bodyPreview,body,hasAttachments,conversationId';

const WELL_KNOWN_FOLDERS: Record<string, string> = {
  inbox: 'inbox',
  drafts: 'drafts',
  sent: 'sentitems',
  sentitems: 'sentitems',
  archive: 'archive',
  deleted: 'deleteditems',
  deleteditems: 'deleteditems',
  junk: 'junkemail',
  junkemail: 'junkemail',
};

async function resolveMailFolderId(
  context: CustomToolHandlerContext,
  folder: string
): Promise<string> {
  const normalized = folder.trim().toLowerCase().replace(/\s+/g, '');
  if (WELL_KNOWN_FOLDERS[normalized]) {
    return WELL_KNOWN_FOLDERS[normalized];
  }

  const response = await graphRequest<GraphCollectionResponse<JsonRecord>>(
    context,
    '/me/mailFolders?$top=200&$select=id,displayName'
  );
  const exact = (response.value ?? []).find(
    (entry) => getString(entry, 'displayName')?.toLowerCase() === folder.toLowerCase()
  );
  return getString(exact, 'id') ?? folder;
}

async function searchMailInternal(
  context: CustomToolHandlerContext,
  query: string,
  limit: number,
  folder?: string,
  cursor?: string
): Promise<{ items: ReturnType<typeof normalizeMail>[]; nextCursor?: string }> {
  const encodedQuery = encodeURIComponent(`"${query}"`);
  const basePath = folder
    ? `/me/mailFolders/${encodePathSegment(await resolveMailFolderId(context, folder))}/messages`
    : '/me/messages';
  const endpoint = cursor
    ? decodeCursor(cursor)
    : `${basePath}?$search=${encodedQuery}&$top=${limit}&${MAIL_SELECT}`;
  const response = await graphRequest<GraphCollectionResponse<JsonRecord>>(context, endpoint, {
    headers: { ConsistencyLevel: 'eventual' },
  });
  return {
    items: (response.value ?? []).map(normalizeMail),
    nextCursor: response['@odata.nextLink'] ? encodeCursor(response['@odata.nextLink']) : undefined,
  };
}

async function listAttachmentsInternal(
  context: CustomToolHandlerContext,
  messageId: string,
  limit: number,
  cursor?: string
): Promise<{ items: AttachmentSummary[]; nextCursor?: string }> {
  const endpoint = cursor
    ? decodeCursor(cursor)
    : `/me/messages/${encodePathSegment(messageId)}/attachments?$top=${limit}`;
  const response = await graphRequest<GraphCollectionResponse<JsonRecord>>(context, endpoint);
  return {
    items: (response.value ?? []).map(normalizeAttachment),
    nextCursor: response['@odata.nextLink'] ? encodeCursor(response['@odata.nextLink']) : undefined,
  };
}

async function getAttachmentInternal(
  context: CustomToolHandlerContext,
  messageId: string,
  attachmentId: string
): Promise<AttachmentSummary & { contentBytes?: string | null; contentId?: string | null }> {
  const response = await graphRequest<JsonRecord>(
    context,
    `/me/messages/${encodePathSegment(messageId)}/attachments/${encodePathSegment(attachmentId)}`
  );
  return {
    ...normalizeAttachment(response),
    contentBytes: getString(response, 'contentBytes'),
    contentId: getString(response, 'contentId'),
  };
}

async function getMailMessageInternal(
  context: CustomToolHandlerContext,
  messageId: string
): Promise<JsonRecord> {
  return graphRequest<JsonRecord>(
    context,
    `/me/messages/${encodePathSegment(messageId)}?${MAIL_SELECT}`
  );
}

export const mailToolDefinitions: CustomToolDefinition[] = [
  {
    name: 'search-mail',
    description: 'Search mail with optional folder scoping and normalized message summaries.',
    method: 'GET',
    path: '/me/messages',
    requiresOrgMode: false,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Mail.Read'],
    schema: {
      query: z.string().min(1).describe('Mail search query'),
      folder: z.string().min(1).optional().describe('Optional folder name, such as Inbox'),
      limit: z.number().int().min(1).max(50).default(25).describe('Maximum results'),
      cursor: z.string().optional().describe('Opaque pagination cursor'),
    },
    handler: async (params, context) => {
      const query = getRequiredString(params, 'query');
      const folder = getOptionalString(params, 'folder');
      const limit = getOptionalNumber(params, 'limit', 25);
      const cursor = getOptionalString(params, 'cursor');
      return success(
        context.graphClient,
        await searchMailInternal(context, query, limit, folder, cursor)
      );
    },
  },
  {
    name: 'resolve-mail-recipients',
    description: 'Resolve recipient names or emails into candidate Microsoft users.',
    method: 'GET',
    path: '/users',
    requiresOrgMode: true,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['User.Read.All'],
    schema: {
      queries: z.array(z.string().min(1)).min(1).max(25).describe('Names or emails to resolve'),
    },
    handler: async (params, context) => {
      const queries = getRequiredStringArray(params, 'queries');
      return success(context.graphClient, {
        items: await Promise.all(
          queries.map(async (query) => ({
            query,
            matches: await searchUsersInternal(context, query, 5),
          }))
        ),
      });
    },
  },
  {
    name: 'list-attachments',
    description: 'List attachments for one mail message.',
    method: 'GET',
    path: '/me/messages/{messageId}/attachments',
    requiresOrgMode: false,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Mail.Read'],
    schema: {
      messageId: z.string().min(1).describe('Mail message ID'),
      limit: z.number().int().min(1).max(100).default(100).describe('Maximum results'),
      cursor: z.string().optional().describe('Opaque pagination cursor'),
    },
    handler: async (params, context) => {
      const messageId = getRequiredString(params, 'messageId');
      const limit = getOptionalNumber(params, 'limit', 100);
      const cursor = getOptionalString(params, 'cursor');
      return success(
        context.graphClient,
        await listAttachmentsInternal(context, messageId, limit, cursor)
      );
    },
  },
  {
    name: 'get-attachment-content',
    description: 'Get attachment metadata and inline content bytes when available.',
    method: 'GET',
    path: '/me/messages/{messageId}/attachments/{attachmentId}',
    requiresOrgMode: false,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Mail.Read'],
    schema: {
      messageId: z.string().min(1).describe('Mail message ID'),
      attachmentId: z.string().min(1).describe('Attachment ID'),
    },
    handler: async (params, context) => {
      const messageId = getRequiredString(params, 'messageId');
      const attachmentId = getRequiredString(params, 'attachmentId');
      return success(
        context.graphClient,
        await getAttachmentInternal(context, messageId, attachmentId)
      );
    },
  },
  {
    name: 'thread-mail',
    description: 'Retrieve the mail thread for a message by conversation ID.',
    method: 'GET',
    path: '/me/messages',
    requiresOrgMode: false,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Mail.Read'],
    schema: {
      messageId: z.string().min(1).describe('Mail message ID'),
      limit: z.number().int().min(1).max(100).default(25).describe('Maximum results'),
    },
    handler: async (params, context) => {
      const messageId = getRequiredString(params, 'messageId');
      const limit = getOptionalNumber(params, 'limit', 25);
      const message = await getMailMessageInternal(context, messageId);
      const conversationId = getString(message, 'conversationId');
      if (!conversationId) {
        return success(context.graphClient, { items: [normalizeMail(message)] });
      }
      const response = await graphRequest<GraphCollectionResponse<JsonRecord>>(
        context,
        `/me/messages?$filter=${encodeURIComponent(`conversationId eq '${conversationId.replace(/'/g, "''")}'`)}&$top=${limit}&${MAIL_SELECT}`
      );
      return success(context.graphClient, {
        items: (response.value ?? []).map(normalizeMail),
      });
    },
  },
  {
    name: 'find-related-mail',
    description: 'Find related mail using a seed message or free-text query.',
    method: 'GET',
    path: '/me/messages',
    requiresOrgMode: false,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Mail.Read'],
    schema: {
      messageId: z.string().min(1).optional().describe('Optional seed message ID'),
      query: z.string().min(1).optional().describe('Optional free-text query'),
      limit: z.number().int().min(1).max(50).default(25).describe('Maximum results'),
    },
    handler: async (params, context) => {
      const messageId = getOptionalString(params, 'messageId');
      const explicitQuery = getOptionalString(params, 'query');
      const limit = getOptionalNumber(params, 'limit', 25);

      let query = explicitQuery;
      if (!query && messageId) {
        const message = await getMailMessageInternal(context, messageId);
        const normalized = normalizeMail(message);
        query = [normalized.subject, normalized.from?.email, normalized.from?.displayName]
          .filter((value): value is string => !!value)
          .join(' ');
      }

      if (!query) {
        throw new Error('Either messageId or query is required.');
      }

      return success(context.graphClient, await searchMailInternal(context, query, limit));
    },
  },
  {
    name: 'search-mail-by-person',
    description: 'Search mail involving one resolved person.',
    method: 'GET',
    path: '/me/messages',
    requiresOrgMode: true,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Mail.Read', 'User.Read.All'],
    schema: {
      query: z.string().min(1).describe('Person name or email'),
      folder: z.string().min(1).optional().describe('Optional folder name'),
      limit: z.number().int().min(1).max(50).default(25).describe('Maximum results'),
    },
    handler: async (params, context) => {
      const query = getRequiredString(params, 'query');
      const folder = getOptionalString(params, 'folder');
      const limit = getOptionalNumber(params, 'limit', 25);
      const person = await resolveQueryToBestUser(context, query);
      const email = person.email ?? person.userPrincipalName ?? query;
      const searchQuery = `from:${email} OR to:${email}`;
      return success(
        context.graphClient,
        await searchMailInternal(context, searchQuery, limit, folder)
      );
    },
  },
];
