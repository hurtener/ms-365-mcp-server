import { z } from 'zod';
import {
  AttachmentSummary,
  CustomToolDefinition,
  CustomToolHandlerContext,
  GraphCollectionResponse,
  JsonRecord,
  ambiguousMatchResult,
  decodeCursor,
  encodeCursor,
  encodePathSegment,
  getArray,
  getOptionalBoolean,
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
  '$select=id,subject,from,toRecipients,receivedDateTime,bodyPreview,body,hasAttachments,conversationId,isDraft,webLink';

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

type ResolvedRecipient = {
  email: string;
  displayName: string | null;
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

export async function searchMailInternal(
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

export async function listAttachmentsInternal(
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

export async function getMailMessageInternal(
  context: CustomToolHandlerContext,
  messageId: string
): Promise<JsonRecord> {
  return graphRequest<JsonRecord>(
    context,
    `/me/messages/${encodePathSegment(messageId)}?${MAIL_SELECT}`
  );
}

async function resolveMailRecipientsInternal(
  context: CustomToolHandlerContext,
  entries: unknown[]
): Promise<ResolvedRecipient[]> {
  const resolved: ResolvedRecipient[] = [];
  for (const entry of entries) {
    if (typeof entry === 'string') {
      if (entry.includes('@')) {
        try {
          const user = await resolveQueryToBestUser(context, entry);
          resolved.push({
            email: user.email ?? user.userPrincipalName ?? entry,
            displayName: user.displayName,
          });
        } catch {
          resolved.push({
            email: entry,
            displayName: null,
          });
        }
        continue;
      }
      const user = await resolveQueryToBestUser(context, entry);
      resolved.push({
        email: user.email ?? user.userPrincipalName ?? entry,
        displayName: user.displayName,
      });
      continue;
    }
    const record = typeof entry === 'object' && entry ? (entry as JsonRecord) : undefined;
    const email = record ? getString(record, 'email') : undefined;
    if (!email) {
      continue;
    }
    resolved.push({
      email,
      displayName: getString(record, 'displayName') ?? null,
    });
  }
  return resolved;
}

async function getThreadMessagesInternal(
  context: CustomToolHandlerContext,
  messageId: string,
  limit: number
): Promise<Array<ReturnType<typeof normalizeMail>>> {
  const message = await getMailMessageInternal(context, messageId);
  const conversationId = getString(message, 'conversationId');
  if (!conversationId) {
    return [normalizeMail(message)];
  }
  const response = await graphRequest<GraphCollectionResponse<JsonRecord>>(
    context,
    `/me/messages?$filter=${encodeURIComponent(`conversationId eq '${conversationId.replace(/'/g, "''")}'`)}&$top=${limit}&${MAIL_SELECT}`
  );
  return (response.value ?? []).map(normalizeMail);
}

function buildReplySubject(originalSubject: string | null, prefix: string): string {
  const subject = originalSubject ?? '(no subject)';
  const lower = subject.toLowerCase();
  if (lower.startsWith(`${prefix.toLowerCase()}:`)) {
    return subject;
  }
  return `${prefix}: ${subject}`;
}

async function runMailAction(
  context: CustomToolHandlerContext,
  action: 'reply' | 'replyAll' | 'forward' | 'createReply' | 'createReplyAll' | 'createForward',
  params: Record<string, unknown>,
  options: {
    messageId: string;
    comment?: string;
    recipients?: ResolvedRecipient[];
  }
) {
  const message = normalizeMail(await getMailMessageInternal(context, options.messageId));

  if (getOptionalBoolean(params, 'previewOnly', false)) {
    const subjectPreview =
      action === 'forward' || action === 'createForward'
        ? buildReplySubject(message.subject, 'FW')
        : buildReplySubject(message.subject, 'RE');
    return success(context.graphClient, {
      action,
      targetMessage: message,
      resolvedRecipients: options.recipients?.map((recipient) => ({
        email: recipient.email,
        displayName: recipient.displayName,
        label: recipient.displayName
          ? `${recipient.displayName} <${recipient.email}>`
          : recipient.email,
      })),
      subjectPreview,
      bodyPreview: options.comment ?? null,
      wouldSend: action === 'reply' || action === 'replyAll' || action === 'forward',
      wouldCreateDraft:
        action === 'createReply' || action === 'createReplyAll' || action === 'createForward',
    });
  }

  const endpoint = `/me/messages/${encodePathSegment(options.messageId)}/${action}`;
  const body: JsonRecord = {};
  if (options.comment) {
    body.comment = options.comment;
  }
  if (options.recipients) {
    body.toRecipients = options.recipients.map((recipient) => ({
      emailAddress: {
        address: recipient.email,
        name: recipient.displayName ?? recipient.email,
      },
    }));
  }

  const response = await graphRequest<JsonRecord>(context, endpoint, {
    method: 'POST',
    body: JSON.stringify(body),
  });

  if (action.startsWith('create')) {
    const draft = normalizeMail(response);
    return success(context.graphClient, {
      action,
      draft,
      draftId: draft.messageId,
    });
  }

  return success(context.graphClient, {
    action,
    messageId: options.messageId,
    targetMessage: message,
    sent: true,
  });
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
      return success(context.graphClient, {
        items: await getThreadMessagesInternal(context, messageId, limit),
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
  {
    name: 'get-mail-context',
    description: 'Return a message plus optional thread and attachment context in one call.',
    method: 'GET',
    path: '/me/messages/{messageId}',
    requiresOrgMode: false,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Mail.Read'],
    schema: {
      messageId: z.string().min(1).describe('Mail message ID'),
      includeThread: z.boolean().default(true).describe('Include thread messages'),
      includeAttachments: z.boolean().default(false).describe('Include attachments'),
      threadLimit: z.number().int().min(1).max(50).default(10).describe('Maximum thread messages'),
    },
    handler: async (params, context) => {
      const messageId = getRequiredString(params, 'messageId');
      const includeThread = getOptionalBoolean(params, 'includeThread', true);
      const includeAttachments = getOptionalBoolean(params, 'includeAttachments', false);
      const threadLimit = getOptionalNumber(params, 'threadLimit', 10);
      const message = normalizeMail(await getMailMessageInternal(context, messageId));
      const result: JsonRecord = { message };
      if (includeThread) {
        result.thread = await getThreadMessagesInternal(context, messageId, threadLimit);
      }
      if (includeAttachments) {
        result.attachments = (await listAttachmentsInternal(context, messageId, 100)).items;
      }
      return success(context.graphClient, result);
    },
  },
  {
    name: 'reply-mail',
    description: 'Reply to one mail message with optional preview mode.',
    method: 'POST',
    path: '/me/messages/{messageId}/reply',
    requiresOrgMode: false,
    readOnlyHint: false,
    destructiveHint: true,
    openWorldHint: true,
    scopes: ['Mail.ReadWrite'],
    schema: {
      messageId: z.string().min(1).describe('Mail message ID'),
      comment: z.string().min(1).describe('Reply comment'),
      previewOnly: z.boolean().optional().describe('Preview without sending'),
    },
    handler: async (params, context) =>
      runMailAction(context, 'reply', params, {
        messageId: getRequiredString(params, 'messageId'),
        comment: getRequiredString(params, 'comment'),
      }),
  },
  {
    name: 'reply-all-mail',
    description: 'Reply all to one mail message with optional preview mode.',
    method: 'POST',
    path: '/me/messages/{messageId}/replyAll',
    requiresOrgMode: false,
    readOnlyHint: false,
    destructiveHint: true,
    openWorldHint: true,
    scopes: ['Mail.ReadWrite'],
    schema: {
      messageId: z.string().min(1).describe('Mail message ID'),
      comment: z.string().min(1).describe('Reply-all comment'),
      previewOnly: z.boolean().optional().describe('Preview without sending'),
    },
    handler: async (params, context) =>
      runMailAction(context, 'replyAll', params, {
        messageId: getRequiredString(params, 'messageId'),
        comment: getRequiredString(params, 'comment'),
      }),
  },
  {
    name: 'forward-mail',
    description: 'Forward one mail message with optional preview mode.',
    method: 'POST',
    path: '/me/messages/{messageId}/forward',
    requiresOrgMode: false,
    readOnlyHint: false,
    destructiveHint: true,
    openWorldHint: true,
    scopes: ['Mail.ReadWrite', 'User.Read.All'],
    schema: {
      messageId: z.string().min(1).describe('Mail message ID'),
      to: z
        .array(
          z.union([
            z.string(),
            z.object({ email: z.string().email(), displayName: z.string().optional() }),
          ])
        )
        .min(1)
        .describe('Recipients to forward to'),
      comment: z.string().optional().describe('Optional forward comment'),
      previewOnly: z.boolean().optional().describe('Preview without sending'),
    },
    handler: async (params, context) =>
      runMailAction(context, 'forward', params, {
        messageId: getRequiredString(params, 'messageId'),
        comment: getOptionalString(params, 'comment'),
        recipients: await resolveMailRecipientsInternal(
          context,
          getArray(params as JsonRecord, 'to')
        ),
      }),
  },
  {
    name: 'draft-mail-reply',
    description: 'Create a draft reply for one mail message.',
    method: 'POST',
    path: '/me/messages/{messageId}/createReply',
    requiresOrgMode: false,
    readOnlyHint: false,
    destructiveHint: true,
    openWorldHint: true,
    scopes: ['Mail.ReadWrite'],
    schema: {
      messageId: z.string().min(1).describe('Mail message ID'),
      comment: z.string().optional().describe('Optional draft comment'),
    },
    handler: async (params, context) =>
      runMailAction(context, 'createReply', params, {
        messageId: getRequiredString(params, 'messageId'),
        comment: getOptionalString(params, 'comment'),
      }),
  },
  {
    name: 'draft-mail-reply-all',
    description: 'Create a draft reply-all for one mail message.',
    method: 'POST',
    path: '/me/messages/{messageId}/createReplyAll',
    requiresOrgMode: false,
    readOnlyHint: false,
    destructiveHint: true,
    openWorldHint: true,
    scopes: ['Mail.ReadWrite'],
    schema: {
      messageId: z.string().min(1).describe('Mail message ID'),
      comment: z.string().optional().describe('Optional draft comment'),
    },
    handler: async (params, context) =>
      runMailAction(context, 'createReplyAll', params, {
        messageId: getRequiredString(params, 'messageId'),
        comment: getOptionalString(params, 'comment'),
      }),
  },
  {
    name: 'draft-mail-forward',
    description: 'Create a draft forward for one mail message.',
    method: 'POST',
    path: '/me/messages/{messageId}/createForward',
    requiresOrgMode: false,
    readOnlyHint: false,
    destructiveHint: true,
    openWorldHint: true,
    scopes: ['Mail.ReadWrite', 'User.Read.All'],
    schema: {
      messageId: z.string().min(1).describe('Mail message ID'),
      to: z
        .array(
          z.union([
            z.string(),
            z.object({ email: z.string().email(), displayName: z.string().optional() }),
          ])
        )
        .min(1)
        .describe('Recipients to include on the draft'),
      comment: z.string().optional().describe('Optional draft comment'),
    },
    handler: async (params, context) =>
      runMailAction(context, 'createForward', params, {
        messageId: getRequiredString(params, 'messageId'),
        comment: getOptionalString(params, 'comment'),
        recipients: await resolveMailRecipientsInternal(
          context,
          getArray(params as JsonRecord, 'to')
        ),
      }),
  },
];
