import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import GraphClient from './graph-client.js';
import AuthManager from './auth.js';
import logger from './logger.js';
import { getRequestTokens } from './request-context.js';

type ToolResult = Awaited<ReturnType<GraphClient['graphRequest']>>;

type JsonRecord = Record<string, unknown>;

type ToolSchema = Record<string, z.ZodTypeAny>;

type ToolRegistryEntry = {
  name: string;
  description: string;
  method: string;
  path: string;
  readOnlyHint: boolean;
  openWorldHint: boolean;
  destructiveHint?: boolean;
  handler: (params: Record<string, unknown>) => Promise<ToolResult>;
};

type CustomToolContext = {
  graphClient: GraphClient;
  authManager?: AuthManager;
};

type CustomToolHandlerContext = {
  graphClient: GraphClient;
  authManager?: AuthManager;
  accessToken?: string;
};

type CustomToolDefinition = {
  name: string;
  description: string;
  method: 'GET' | 'POST';
  path: string;
  requiresOrgMode: boolean;
  readOnlyHint: boolean;
  openWorldHint: boolean;
  destructiveHint?: boolean;
  scopes: string[];
  schema: ToolSchema;
  handler: (
    params: Record<string, unknown>,
    context: CustomToolHandlerContext
  ) => Promise<ToolResult>;
};

type NormalizedUser = {
  userId: string;
  displayName: string | null;
  email: string | null;
  userPrincipalName: string | null;
  givenName: string | null;
  surname: string | null;
  jobTitle: string | null;
  department: string | null;
};

type NormalizedChatMember = {
  memberId: string;
  userId: string | null;
  displayName: string | null;
  email: string | null;
  userPrincipalName: string | null;
  roles: string[];
  membershipType: string | null;
  membershipOrigins?: string[];
};

type NormalizedChat = {
  id: string;
  chatType: string;
  topic: string | null;
  lastUpdatedDateTime: string | null;
  webUrl?: string | null;
  memberCount: number;
};

type RankedUser = NormalizedUser & {
  matchScore: number;
  matchReasons: string[];
};

type ChatCandidate = {
  chat: NormalizedChat;
  members: NormalizedChatMember[];
  matchType: string;
  rank: number;
  matchReasons: string[];
  score: number;
  exactness: number;
};

const CHAT_TYPES = new Set(['oneOnOne', 'group', 'meeting']);

const userSelect =
  '$select=id,displayName,givenName,surname,mail,userPrincipalName,jobTitle,department';
const chatSelect = '$select=id,chatType,topic,lastUpdatedDateTime,webUrl';
const messageSelect = '$select=id,createdDateTime,from,body,summary';

const customToolDefinitions: CustomToolDefinition[] = [
  {
    name: 'search-users',
    description:
      'Resolve a human name or email into likely Microsoft users with deterministic ranking.',
    method: 'GET',
    path: '/users',
    requiresOrgMode: true,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['User.Read.All'],
    schema: {
      query: z.string().min(1).describe('Human name, email address, or UPN to resolve'),
      limit: z
        .number()
        .int()
        .min(1)
        .max(25)
        .default(10)
        .describe('Maximum results to return'),
    },
    handler: async (params, context) => {
      const query = getRequiredString(params, 'query');
      const limit = getOptionalNumber(params, 'limit', 10);
      const items = await searchUsersInternal(context, query, limit);
      return success(context.graphClient, { items });
    },
  },
  {
    name: 'list-chat-members',
    description: 'List the actual participants of a Teams chat with normalized identity fields.',
    method: 'GET',
    path: '/chats/{chatId}/members',
    requiresOrgMode: true,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Chat.Read'],
    schema: {
      chatId: z.string().min(1).describe('Chat ID (for example 19:...@thread.v2)'),
    },
    handler: async (params, context) => {
      const chatId = getRequiredString(params, 'chatId');
      const members = await fetchChatMembers(context, chatId);
      return success(context.graphClient, { chatId, members });
    },
  },
  {
    name: 'get-chat-details',
    description: 'Get a normalized chat object and optionally inline the chat member roster.',
    method: 'GET',
    path: '/chats/{chatId}',
    requiresOrgMode: true,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Chat.Read'],
    schema: {
      chatId: z.string().min(1).describe('Chat ID (for example 19:...@thread.v2)'),
      includeMembers: z
        .boolean()
        .default(true)
        .describe('Include normalized members inline in the response'),
    },
    handler: async (params, context) => {
      const chatId = getRequiredString(params, 'chatId');
      const includeMembers = getOptionalBoolean(params, 'includeMembers', true);
      const { chat, members } = await getChatDetailsInternal(context, chatId, includeMembers);
      return success(
        context.graphClient,
        includeMembers ? { ...chat, members } : { ...chat }
      );
    },
  },
  {
    name: 'find-chats-by-participant',
    description:
      'Find and rank the best chats for a participant so the LLM does not have to disambiguate raw chat lists.',
    method: 'GET',
    path: '/me/chats',
    requiresOrgMode: true,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Chat.Read', 'User.Read.All'],
    schema: {
      userId: z.string().min(1).optional().describe('Resolved Microsoft user ID'),
      email: z.string().email().optional().describe('Participant email address'),
      includeGroupChats: z
        .boolean()
        .default(true)
        .describe('Include group and meeting chats in the ranked results'),
      limit: z
        .number()
        .int()
        .min(1)
        .max(25)
        .default(10)
        .describe('Maximum results to return'),
    },
    handler: async (params, context) => {
      const includeGroupChats = getOptionalBoolean(params, 'includeGroupChats', true);
      const limit = getOptionalNumber(params, 'limit', 10);
      const resolvedUser = await resolveParticipantInput(context, params);
      const items = await findChatsByParticipantInternal(
        context,
        resolvedUser,
        includeGroupChats,
        limit
      );
      return success(context.graphClient, { items });
    },
  },
  {
    name: 'list-recent-chats',
    description:
      'List recent chats in recency order with opaque pagination and optional inline members.',
    method: 'GET',
    path: '/me/chats',
    requiresOrgMode: true,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Chat.Read'],
    schema: {
      limit: z
        .number()
        .int()
        .min(1)
        .max(50)
        .default(20)
        .describe('Maximum chats to return'),
      cursor: z.string().optional().describe('Opaque pagination cursor returned by nextCursor'),
      includeMembers: z
        .boolean()
        .default(true)
        .describe('Include normalized members inline in each chat item'),
    },
    handler: async (params, context) => {
      const limit = getOptionalNumber(params, 'limit', 20);
      const cursor = getOptionalString(params, 'cursor');
      const includeMembers = getOptionalBoolean(params, 'includeMembers', true);
      const result = await listRecentChatsInternal(context, limit, cursor, includeMembers);
      return success(context.graphClient, result);
    },
  },
  {
    name: 'get-chat-context',
    description: 'Get one chat, its members, and a preview window of recent messages in one call.',
    method: 'GET',
    path: '/chats/{chatId}',
    requiresOrgMode: true,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Chat.Read', 'ChatMessage.Read'],
    schema: {
      chatId: z.string().min(1).describe('Chat ID (for example 19:...@thread.v2)'),
      messageLimit: z
        .number()
        .int()
        .min(1)
        .max(25)
        .default(10)
        .describe('Number of recent messages to include'),
    },
    handler: async (params, context) => {
      const chatId = getRequiredString(params, 'chatId');
      const messageLimit = getOptionalNumber(params, 'messageLimit', 10);
      const { chat, members } = await getChatDetailsInternal(context, chatId, true);
      const recentMessages = await fetchChatMessages(context, chatId, messageLimit);
      return success(context.graphClient, { chat, members, recentMessages });
    },
  },
  {
    name: 'search-messages',
    description: 'Search Teams chat messages by text and optionally narrow results by participant.',
    method: 'POST',
    path: '/search/query',
    requiresOrgMode: true,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Chat.Read', 'ChannelMessage.Read.All'],
    schema: {
      query: z.string().min(1).describe('Free-text search query'),
      participantUserId: z.string().min(1).optional().describe('Only return hits from chats containing this user'),
      scope: z.literal('chats').default('chats').describe('Currently only chats are supported'),
      limit: z
        .number()
        .int()
        .min(1)
        .max(25)
        .default(20)
        .describe('Maximum search hits to return'),
    },
    handler: async (params, context) => {
      const scope = getOptionalString(params, 'scope') ?? 'chats';
      if (scope !== 'chats') {
        return errorResult(context.graphClient, 'unsupported_scope', {
          message: 'Only scope="chats" is currently supported.',
        });
      }

      const query = getRequiredString(params, 'query');
      const participantUserId = getOptionalString(params, 'participantUserId');
      const limit = getOptionalNumber(params, 'limit', 20);
      const items = await searchMessagesInternal(context, query, participantUserId, limit);
      return success(context.graphClient, { items });
    },
  },
  {
    name: 'resolve-person',
    description:
      'Resolve a person query to one high-confidence Microsoft user or return structured ambiguity.',
    method: 'GET',
    path: '/users',
    requiresOrgMode: true,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['User.Read.All'],
    schema: {
      query: z.string().min(1).describe('Human name, email address, or UPN to resolve'),
    },
    handler: async (params, context) => {
      const query = getRequiredString(params, 'query');
      const items = await searchUsersInternal(context, query, 10);

      if (items.length === 0) {
        return errorResult(context.graphClient, 'user_not_found', { query });
      }

      const [best, second] = items;
      if (best.matchScore < 0.85 || (second && best.matchScore - second.matchScore < 0.1)) {
        return ambiguousMatchResult(
          context.graphClient,
          items.slice(0, 5).map((item, index) => ({
            userId: item.userId,
            label: buildUserCandidateLabel(item),
            rank: index + 1,
          }))
        );
      }

      return success(context.graphClient, best);
    },
  },
  {
    name: 'list-user-presence',
    description: 'Get Teams presence for up to 100 users in a single request.',
    method: 'POST',
    path: '/communications/getPresencesByUserId',
    requiresOrgMode: true,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Presence.Read.All'],
    schema: {
      userIds: z
        .array(z.string().min(1))
        .min(1)
        .max(100)
        .describe('Microsoft user IDs to look up'),
    },
    handler: async (params, context) => {
      const userIds = getRequiredStringArray(params, 'userIds');
      const response = await graphRequest<JsonRecord[] | GraphCollectionResponse<JsonRecord>>(
        context,
        '/communications/getPresencesByUserId',
        {
          method: 'POST',
          body: JSON.stringify({ ids: userIds }),
        }
      );

      const source = Array.isArray(response) ? response : (response.value ?? []);
      const items = source.map((entry) => ({
        userId: getString(entry, 'id') ?? '',
        availability: getString(entry, 'availability'),
        activity: getString(entry, 'activity'),
        statusMessage: extractStatusMessage(entry),
      }));

      return success(context.graphClient, { items });
    },
  },
  {
    name: 'list-channel-members',
    description: 'List normalized members for a Teams channel.',
    method: 'GET',
    path: '/teams/{teamId}/channels/{channelId}/allMembers',
    requiresOrgMode: true,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['ChannelMember.Read.All'],
    schema: {
      teamId: z.string().min(1).describe('Teams team ID'),
      channelId: z.string().min(1).describe('Teams channel ID'),
    },
    handler: async (params, context) => {
      const teamId = getRequiredString(params, 'teamId');
      const channelId = getRequiredString(params, 'channelId');
      const response = await graphRequest<GraphCollectionResponse<JsonRecord>>(
        context,
        `/teams/${encodePathSegment(teamId)}/channels/${encodePathSegment(channelId)}/allMembers`
      );
      const members = normalizeMembers(response.value ?? []);
      return success(context.graphClient, { teamId, channelId, members });
    },
  },
];

type GraphCollectionResponse<T> = {
  value?: T[];
  '@odata.nextLink'?: string;
};

export function getCustomToolDefinitions(): readonly CustomToolDefinition[] {
  return customToolDefinitions;
}

export function buildCustomToolRegistry(
  graphClient: GraphClient,
  authManager?: AuthManager,
  multiAccount: boolean = false,
  accountNames: string[] = []
): Map<string, ToolRegistryEntry> {
  const registry = new Map<string, ToolRegistryEntry>();
  const context: CustomToolContext = { graphClient, authManager };

  for (const definition of customToolDefinitions) {
    registry.set(definition.name, {
      name: definition.name,
      description: definition.description,
      method: definition.method,
      path: definition.path,
      readOnlyHint: definition.readOnlyHint,
      openWorldHint: definition.openWorldHint,
      destructiveHint: definition.destructiveHint,
      handler: buildCustomToolHandler(definition, context),
    });
  }

  if (multiAccount && accountNames.length > 0) {
    logger.info(
      `Multi-account mode available for custom tools with accounts: ${accountNames.join(', ')}`
    );
  }

  return registry;
}

export function registerCustomTools(
  server: McpServer,
  graphClient: GraphClient,
  enabledToolsRegex: RegExp | undefined,
  orgMode: boolean = false,
  authManager?: AuthManager,
  multiAccount: boolean = false,
  accountNames: string[] = []
): number {
  const context: CustomToolContext = { graphClient, authManager };
  let registeredCount = 0;

  for (const definition of customToolDefinitions) {
    if (definition.requiresOrgMode && !orgMode) {
      continue;
    }

    if (enabledToolsRegex && !enabledToolsRegex.test(definition.name)) {
      continue;
    }

    const schema = buildSchemaWithCommonParams(definition.schema, multiAccount, accountNames);
    server.tool(
      definition.name,
      definition.description,
      schema,
      {
        title: definition.name,
        readOnlyHint: definition.readOnlyHint,
        destructiveHint: definition.destructiveHint,
        openWorldHint: definition.openWorldHint,
      },
      buildCustomToolHandler(definition, context)
    );
    registeredCount++;
  }

  return registeredCount;
}

function buildCustomToolHandler(
  definition: CustomToolDefinition,
  context: CustomToolContext
): (params: Record<string, unknown>) => Promise<ToolResult> {
  return async (params) => {
    try {
      const accountAccessToken = await resolveAccessToken(params, context.authManager);
      return await definition.handler(params, {
        ...context,
        accessToken: accountAccessToken,
      });
    } catch (error) {
      logger.error(`Error in custom tool ${definition.name}: ${(error as Error).message}`);
      if (error instanceof GraphToolError) {
        return errorResult(context.graphClient, error.code, {
          message: error.message,
          ...(error.details ?? {}),
        });
      }
      return errorResult(
        context.graphClient,
        inferErrorCode(error, 'Custom tool request failed'),
        {
          message: (error as Error).message,
        }
      );
    }
  };
}

function buildSchemaWithCommonParams(
  schema: ToolSchema,
  multiAccount: boolean,
  accountNames: string[]
): ToolSchema {
  const finalSchema: ToolSchema = { ...schema };

  if (multiAccount) {
    const accountHint =
      accountNames.length > 0 ? `Known accounts: ${accountNames.join(', ')}. ` : '';
    finalSchema.account = z
      .string()
      .describe(
        `${accountHint}Microsoft account email to use for this request. ` +
          `Required when multiple accounts are configured.`
      )
      .optional();
  }

  return finalSchema;
}

async function resolveAccessToken(
  params: Record<string, unknown>,
  authManager?: AuthManager
): Promise<string | undefined> {
  if (!authManager || authManager.isOAuthModeEnabled() || getRequestTokens()) {
    return undefined;
  }

  const accountParam = typeof params.account === 'string' ? params.account : undefined;
  return authManager.getTokenForAccount(accountParam);
}

async function searchUsersInternal(
  context: CustomToolHandlerContext,
  query: string,
  limit: number
): Promise<RankedUser[]> {
  const searchExpression = buildUserSearchExpression(query);
  const endpoint =
    `/users?${userSelect}` +
    `&$top=${Math.max(limit * 3, 25)}` +
    `&$search=${encodeURIComponent(searchExpression)}`;

  const response = await graphRequest<GraphCollectionResponse<JsonRecord>>(context, endpoint, {
    headers: {
      ConsistencyLevel: 'eventual',
    },
  });

  const ranked = (response.value ?? [])
    .map((entry) => {
      const normalized = normalizeUser(entry);
      const ranking = rankUser(normalized, query);
      return {
        ...normalized,
        matchScore: ranking.score,
        matchReasons: ranking.reasons,
      };
    })
    .filter((entry) => entry.matchScore > 0)
    .sort(compareRankedUsers)
    .slice(0, limit);

  return ranked;
}

async function resolveParticipantInput(
  context: CustomToolHandlerContext,
  params: Record<string, unknown>
): Promise<NormalizedUser> {
  const userId = getOptionalString(params, 'userId');
  if (userId) {
    return {
      userId,
      displayName: null,
      email: null,
      userPrincipalName: null,
      givenName: null,
      surname: null,
      jobTitle: null,
      department: null,
    };
  }

  const email = getOptionalString(params, 'email');
  if (!email) {
    throw new Error('Either userId or email is required.');
  }

  const items = await searchUsersInternal(context, email, 5);
  if (items.length === 0) {
    throw new GraphToolError('user_not_found', 'No matching user was found.', {
      email,
    });
  }

  const exact = items.filter((item) => {
    const normalizedEmail = normalizeString(email);
    return (
      normalizeString(item.email) === normalizedEmail ||
      normalizeString(item.userPrincipalName) === normalizedEmail
    );
  });

  if (exact.length === 1) {
    return exact[0];
  }

  if (exact.length > 1) {
    throw new GraphToolError('ambiguous_match', 'Multiple users matched the email address.', {
      candidates: exact.map((item, index) => ({
        userId: item.userId,
        label: buildUserLabel(item),
        rank: index + 1,
      })),
    });
  }

  const [best, second] = items;
  if (!best || best.matchScore < 0.85 || (second && best.matchScore - second.matchScore < 0.1)) {
    throw new GraphToolError('ambiguous_match', 'Multiple users matched the email address.', {
      candidates: items.map((item, index) => ({
        userId: item.userId,
        label: buildUserLabel(item),
        rank: index + 1,
      })),
    });
  }

  return best;
}

async function findChatsByParticipantInternal(
  context: CustomToolHandlerContext,
  participant: NormalizedUser,
  includeGroupChats: boolean,
  limit: number
): Promise<Array<Omit<ChatCandidate, 'score' | 'exactness'>>> {
  const recentChats = await listRecentChatsInternal(
    context,
    Math.max(limit * 5, 25),
    undefined,
    true
  );

  const candidates = recentChats.items
    .filter((entry) => includeGroupChats || entry.chatType === 'oneOnOne')
    .filter((entry) => !entry.membersUnavailable)
    .map((entry) => ({
      chat: pickChatFields(entry),
      members: entry.members ?? [],
    }))
    .filter((entry) =>
      entry.members.some(
        (member) =>
          member.userId === participant.userId || sameEmail(member.email, participant.email)
      )
    )
    .map(({ chat, members }) => rankChatCandidate(chat, members, participant))
    .sort(compareChatCandidates)
    .slice(0, limit)
    .map(({ score: _score, exactness: _exactness, ...candidate }, index) => ({
      ...candidate,
      rank: index + 1,
    }));

  return candidates;
}

async function listRecentChatsInternal(
  context: CustomToolHandlerContext,
  limit: number,
  cursor?: string,
  includeMembers: boolean = true
): Promise<{
  items: Array<
    NormalizedChat & {
      members?: NormalizedChatMember[];
      membersUnavailable?: boolean;
      membersError?: string;
    }
  >;
  nextCursor?: string;
}> {
  const endpoint = cursor
    ? decodeCursor(cursor)
    : `/me/chats?${chatSelect}&$top=${limit}&$orderby=${encodeURIComponent('lastUpdatedDateTime desc')}`;
  const response = await graphRequest<GraphCollectionResponse<JsonRecord>>(context, endpoint);
  const chats = [...(response.value ?? [])].sort(compareRawChatsByRecency);

  const items: Array<
    NormalizedChat & {
      members?: NormalizedChatMember[];
      membersUnavailable?: boolean;
      membersError?: string;
    }
  > = [];

  for (const rawChat of chats.slice(0, limit)) {
    const normalizedChat = normalizeChat(rawChat);
    if (!CHAT_TYPES.has(normalizedChat.chatType)) {
      continue;
    }

    if (!includeMembers) {
      items.push(normalizedChat);
      continue;
    }

    const membersResult = await tryFetchChatMembers(context, normalizedChat.id);
    items.push({
      ...normalizedChat,
      memberCount: membersResult.members.length,
      members: membersResult.members,
      membersUnavailable: membersResult.membersUnavailable || undefined,
      membersError: membersResult.membersError,
    });
  }

  return {
    items,
    nextCursor: response['@odata.nextLink'] ? encodeCursor(response['@odata.nextLink']) : undefined,
  };
}

async function getChatDetailsInternal(
  context: CustomToolHandlerContext,
  chatId: string,
  includeMembers: boolean
): Promise<{ chat: NormalizedChat; members: NormalizedChatMember[] }> {
  const rawChat = await graphRequest<JsonRecord>(
    context,
    `/chats/${encodePathSegment(chatId)}?${chatSelect}`
  );
  const chat = normalizeChat(rawChat);

  if (!CHAT_TYPES.has(chat.chatType)) {
    throw new GraphToolError('unsupported_chat_type', `Unsupported chat type: ${chat.chatType}`);
  }

  const members = includeMembers
    ? await fetchChatMembers(context, chatId, 'chat_members_unavailable')
    : [];

  return {
    chat: {
      ...chat,
      memberCount: includeMembers ? members.length : chat.memberCount,
    },
    members,
  };
}

async function fetchChatMembers(
  context: CustomToolHandlerContext,
  chatId: string,
  notFoundCode: string = 'chat_not_found'
): Promise<NormalizedChatMember[]> {
  try {
    const response = await graphRequest<GraphCollectionResponse<JsonRecord>>(
      context,
      `/chats/${encodePathSegment(chatId)}/members`
    );
    return normalizeMembers(response.value ?? []);
  } catch (error) {
    const inferred = inferErrorCode(error, 'Failed to fetch chat members');
    if (inferred === 'insufficient_scope') {
      throw new GraphToolError('insufficient_scope', (error as Error).message);
    }
    if (inferred === 'chat_not_found') {
      throw new GraphToolError(notFoundCode, `Chat not found: ${chatId}`);
    }
    throw new GraphToolError('chat_members_unavailable', (error as Error).message);
  }
}

async function tryFetchChatMembers(
  context: CustomToolHandlerContext,
  chatId: string
): Promise<{
  members: NormalizedChatMember[];
  membersUnavailable?: boolean;
  membersError?: string;
}> {
  try {
    return {
      members: await fetchChatMembers(context, chatId),
    };
  } catch (error) {
    if (error instanceof GraphToolError) {
      if (error.code === 'chat_not_found') {
        throw error;
      }

      return {
        members: [],
        membersUnavailable: true,
        membersError: error.code,
      };
    }

    return {
      members: [],
      membersUnavailable: true,
      membersError: inferErrorCode(error, 'Failed to fetch chat members'),
    };
  }
}

async function fetchChatMessages(
  context: CustomToolHandlerContext,
  chatId: string,
  limit: number
): Promise<
  Array<{
    id: string;
    from: { userId: string | null; displayName: string | null };
    createdDateTime: string | null;
    bodyPreview: string | null;
  }>
> {
  const response = await graphRequest<GraphCollectionResponse<JsonRecord>>(
    context,
    `/chats/${encodePathSegment(chatId)}/messages?${messageSelect}&$top=${limit}`
  );

  return (response.value ?? []).map((message) => ({
    id: getString(message, 'id') ?? '',
    from: extractMessageSender(message),
    createdDateTime: getString(message, 'createdDateTime'),
    bodyPreview: extractMessagePreview(message),
  }));
}

async function searchMessagesInternal(
  context: CustomToolHandlerContext,
  query: string,
  participantUserId: string | undefined,
  limit: number
): Promise<
  Array<{
    chatId: string | null;
    messageId: string | null;
    chatType: string | null;
    topic: string | null;
    from: { userId: string | null; displayName: string | null };
    createdDateTime: string | null;
    bodyPreview: string | null;
  }>
> {
  const response = await graphRequest<JsonRecord>(context, '/search/query', {
    method: 'POST',
    body: JSON.stringify({
      requests: [
        {
          entityTypes: ['chatMessage'],
          query: {
            queryString: query,
          },
          from: 0,
          size: limit,
        },
      ],
    }),
  });

  const hitsContainers = getArray(response, 'value')
    .flatMap((entry) => getArray(asRecord(entry), 'hitsContainers'))
    .map(asRecord)
    .filter((entry): entry is JsonRecord => !!entry);

  let items = hitsContainers
    .flatMap((container) => getArray(container, 'hits'))
    .map(asRecord)
    .filter((entry): entry is JsonRecord => !!entry)
    .map((hit) => {
      const resource = asRecord(hit.resource) ?? {};
      return {
        chatId: getString(resource, 'chatId'),
        messageId: getString(resource, 'id'),
        chatType: getString(resource, 'chatType'),
        topic: getString(resource, 'topic'),
        from: extractMessageSender(resource),
        createdDateTime: getString(resource, 'createdDateTime'),
        bodyPreview: extractMessagePreview(resource),
      };
    });

  if (participantUserId) {
    const rosterCache = new Map<string, NormalizedChatMember[]>();
    const filtered: typeof items = [];

    for (const item of items) {
      if (!item.chatId) {
        continue;
      }

      if (!rosterCache.has(item.chatId)) {
        const membersResult = await tryFetchChatMembers(context, item.chatId);
        if (membersResult.membersUnavailable) {
          continue;
        }
        rosterCache.set(item.chatId, membersResult.members);
      }

      const members = rosterCache.get(item.chatId) ?? [];
      if (members.some((member) => member.userId === participantUserId)) {
        filtered.push(item);
      }
    }

    items = filtered;
  }

  return items.slice(0, limit);
}

async function graphRequest<T>(
  context: CustomToolHandlerContext,
  endpoint: string,
  options: {
    method?: string;
    headers?: Record<string, string>;
    body?: string;
  } = {}
): Promise<T> {
  return (await context.graphClient.makeRequest(endpoint, {
    ...options,
    accessToken: context.accessToken,
  })) as T;
}

function normalizeUser(entry: JsonRecord): NormalizedUser {
  return {
    userId: getString(entry, 'id') ?? '',
    displayName: getString(entry, 'displayName'),
    email: getString(entry, 'mail'),
    userPrincipalName: getString(entry, 'userPrincipalName'),
    givenName: getString(entry, 'givenName'),
    surname: getString(entry, 'surname'),
    jobTitle: getString(entry, 'jobTitle'),
    department: getString(entry, 'department'),
  };
}

function normalizeMembers(entries: JsonRecord[]): NormalizedChatMember[] {
  const deduped = new Map<string, NormalizedChatMember>();

  for (const entry of entries) {
    const user = asRecord(entry.user);
    const roles = getStringArray(entry, 'roles');
    const email =
      getString(entry, 'email') ??
      getString(entry, 'mail') ??
      getString(user, 'email') ??
      getString(user, 'mail') ??
      getString(entry, 'userPrincipalName') ??
      getString(user, 'userPrincipalName');
    const userPrincipalName =
      getString(entry, 'userPrincipalName') ?? getString(user, 'userPrincipalName');
    const displayName = getString(entry, 'displayName') ?? getString(user, 'displayName') ?? null;
    const userId = getString(entry, 'userId') ?? getString(user, 'id');
    const memberId = getString(entry, 'id') ?? `${userId ?? displayName ?? email ?? 'member'}`;
    const membershipType =
      roles.includes('owner') ? 'owner' : roles[0] ?? getString(entry, '@odata.type') ?? 'member';
    const membershipOrigins = getMembershipOrigins(entry);

    if (!userId && !displayName && !email && !userPrincipalName) {
      continue;
    }

    const normalized: NormalizedChatMember = {
      memberId,
      userId: userId ?? null,
      displayName,
      email: email ?? null,
      userPrincipalName: userPrincipalName ?? null,
      roles,
      membershipType,
      membershipOrigins: membershipOrigins.length > 0 ? membershipOrigins : undefined,
    };

    const dedupeKey =
      normalized.userId ??
      normalized.email ??
      normalized.userPrincipalName ??
      `${normalized.displayName ?? 'unknown'}:${normalized.memberId}`;

    const existing = deduped.get(dedupeKey);
    if (!existing) {
      deduped.set(dedupeKey, normalized);
      continue;
    }

    existing.roles = Array.from(new Set([...existing.roles, ...normalized.roles])).sort();
    existing.membershipOrigins = Array.from(
      new Set([...(existing.membershipOrigins ?? []), ...(normalized.membershipOrigins ?? [])])
    ).sort();
    existing.membershipType = existing.membershipType ?? normalized.membershipType;
  }

  return Array.from(deduped.values());
}

function normalizeChat(entry: JsonRecord): NormalizedChat {
  const members = getArray(entry, 'members').map(asRecord).filter((item): item is JsonRecord => !!item);

  return {
    id: getString(entry, 'id') ?? '',
    chatType: getString(entry, 'chatType') ?? 'unknown',
    topic: getString(entry, 'topic'),
    lastUpdatedDateTime: getString(entry, 'lastUpdatedDateTime'),
    webUrl: getString(entry, 'webUrl'),
    memberCount: members.length,
  };
}

function pickChatFields(chat: NormalizedChat & { members?: NormalizedChatMember[] }): NormalizedChat {
  return {
    id: chat.id,
    chatType: chat.chatType,
    topic: chat.topic,
    lastUpdatedDateTime: chat.lastUpdatedDateTime,
    webUrl: chat.webUrl ?? null,
    memberCount: chat.memberCount,
  };
}

function rankUser(
  user: NormalizedUser,
  query: string
): { score: number; reasons: string[] } {
  const normalizedQuery = normalizeString(query);
  const queryTokens = tokenize(query);
  const displayName = normalizeString(user.displayName);
  const email = normalizeString(user.email);
  const userPrincipalName = normalizeString(user.userPrincipalName);

  let score = 0;
  const reasons: string[] = [];

  if (!normalizedQuery) {
    return { score: 0, reasons };
  }

  if (email === normalizedQuery) {
    score += 1;
    reasons.push('mail_exact');
  }

  if (userPrincipalName === normalizedQuery) {
    score = Math.max(score, 0.99);
    reasons.push('userPrincipalName_exact');
  }

  if (displayName === normalizedQuery) {
    score = Math.max(score, 0.98);
    reasons.push('displayName_exact');
  }

  if (!reasons.length && displayName.startsWith(normalizedQuery)) {
    score = Math.max(score, 0.9);
    reasons.push('displayName_prefix');
  }

  if (!reasons.length && (email.startsWith(normalizedQuery) || userPrincipalName.startsWith(normalizedQuery))) {
    score = Math.max(score, 0.89);
    reasons.push('mail_prefix');
  }

  const overlap = queryTokens.length > 0 ? queryTokens.filter((token) => displayName.includes(token)).length / queryTokens.length : 0;
  if (overlap > 0) {
    score = Math.max(score, 0.6 + overlap * 0.25);
    reasons.push('displayName_token_overlap');
  }

  if (!reasons.length && (displayName.includes(normalizedQuery) || email.includes(normalizedQuery) || userPrincipalName.includes(normalizedQuery))) {
    score = Math.max(score, 0.6);
    reasons.push('substring_match');
  }

  return {
    score: Number(score.toFixed(2)),
    reasons,
  };
}

function rankChatCandidate(
  chat: NormalizedChat,
  members: NormalizedChatMember[],
  participant: NormalizedUser
): ChatCandidate {
  const isOneOnOne = chat.chatType === 'oneOnOne';
  const participantExact = members.some((member) => member.userId === participant.userId);
  const reasons = ['participant_exact'];
  let score = participantExact ? 100 : 80;
  let exactness = participantExact ? 2 : 1;
  let matchType = 'participant_match';

  if (isOneOnOne && members.length <= 2) {
    score += 50;
    exactness = 3;
    matchType = 'exact_one_on_one';
    reasons.push('one_on_one');
  } else if (chat.chatType === 'group' || chat.chatType === 'meeting') {
    score += 15;
    matchType = 'group_participant_match';
    reasons.push(chat.chatType === 'meeting' ? 'meeting_chat' : 'group_chat');
  }

  const recencyBonus = calculateRecencyBonus(chat.lastUpdatedDateTime);
  if (recencyBonus > 0) {
    reasons.push('recent');
  }
  score += recencyBonus;

  return {
    chat,
    members,
    matchType,
    rank: 0,
    matchReasons: reasons,
    score,
    exactness,
  };
}

function compareChatCandidates(left: ChatCandidate, right: ChatCandidate): number {
  const scoreDelta = right.score - left.score;
  if (scoreDelta !== 0) {
    return scoreDelta;
  }

  const recencyDelta = compareDates(right.chat.lastUpdatedDateTime, left.chat.lastUpdatedDateTime);
  if (recencyDelta !== 0) {
    return recencyDelta;
  }

  const exactnessDelta = right.exactness - left.exactness;
  if (exactnessDelta !== 0) {
    return exactnessDelta;
  }

  return left.chat.id.localeCompare(right.chat.id);
}

function compareRankedUsers(left: RankedUser, right: RankedUser): number {
  const scoreDelta = right.matchScore - left.matchScore;
  if (scoreDelta !== 0) {
    return scoreDelta;
  }

  const exactnessDelta =
    userExactnessScore(right.matchReasons) - userExactnessScore(left.matchReasons);
  if (exactnessDelta !== 0) {
    return exactnessDelta;
  }

  return buildUserLabel(left).localeCompare(buildUserLabel(right));
}

function compareRawChatsByRecency(left: JsonRecord, right: JsonRecord): number {
  return compareDates(getString(right, 'lastUpdatedDateTime'), getString(left, 'lastUpdatedDateTime'));
}

function compareDates(left: string | null | undefined, right: string | null | undefined): number {
  const leftMs = left ? Date.parse(left) : 0;
  const rightMs = right ? Date.parse(right) : 0;
  return leftMs - rightMs;
}

function calculateRecencyBonus(lastUpdatedDateTime: string | null): number {
  if (!lastUpdatedDateTime) {
    return 0;
  }

  const now = Date.now();
  const updated = Date.parse(lastUpdatedDateTime);
  if (Number.isNaN(updated)) {
    return 0;
  }

  const ageDays = (now - updated) / (1000 * 60 * 60 * 24);
  if (ageDays <= 7) {
    return 15;
  }
  if (ageDays <= 30) {
    return 8;
  }
  if (ageDays <= 90) {
    return 3;
  }
  return 0;
}

function buildUserSearchExpression(query: string): string {
  const escaped = query.replace(/"/g, '\\"');
  return `"displayName:${escaped}" OR "mail:${escaped}" OR "userPrincipalName:${escaped}"`;
}

function encodeCursor(nextLink: string): string {
  return Buffer.from(nextLink, 'utf8').toString('base64url');
}

function decodeCursor(cursor: string): string {
  const nextLink = Buffer.from(cursor, 'base64url').toString('utf8');
  try {
    const url = new URL(nextLink);
    return `${url.pathname.replace('/v1.0', '')}${url.search}`;
  } catch {
    throw new GraphToolError('invalid_cursor', 'Invalid pagination cursor.');
  }
}

function extractMessageSender(message: JsonRecord): { userId: string | null; displayName: string | null } {
  const from = asRecord(message.from) ?? {};
  const user = asRecord(from.user) ?? {};
  return {
    userId: getString(user, 'id') ?? getString(from, 'userId') ?? null,
    displayName: getString(user, 'displayName') ?? getString(from, 'displayName') ?? null,
  };
}

function extractMessagePreview(message: JsonRecord): string | null {
  const summary = getString(message, 'summary');
  if (summary) {
    return summary;
  }

  const body = asRecord(message.body);
  const content = getString(body, 'content');
  if (!content) {
    return null;
  }

  return stripHtml(content).slice(0, 280) || null;
}

function extractStatusMessage(entry: JsonRecord): string | null {
  const statusMessage = asRecord(entry.statusMessage);
  const itemBody = asRecord(statusMessage?.message);
  return getString(itemBody, 'content');
}

function stripHtml(value: string): string {
  return value.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
}

function success(graphClient: GraphClient, data: unknown): ToolResult {
  return graphClient.formatJsonResponse(data);
}

function errorResult(
  graphClient: GraphClient,
  error: string,
  extra: JsonRecord = {}
): ToolResult {
  const result = graphClient.formatJsonResponse({
    error,
    ...extra,
  });
  return {
    ...result,
    isError: true,
  };
}

function ambiguousMatchResult(
  graphClient: GraphClient,
  candidates: Array<{ userId?: string; label: string; rank: number }>
): ToolResult {
  return errorResult(graphClient, 'ambiguous_match', { candidates });
}

function inferErrorCode(error: unknown, fallbackMessage: string): string {
  if (error instanceof GraphToolError) {
    return error.code;
  }

  const message = (error as Error)?.message ?? fallbackMessage;
  if (message.includes('scope error') || message.includes('403')) {
    return 'insufficient_scope';
  }
  if (message.includes('404')) {
    return 'chat_not_found';
  }
  return 'graph_request_failed';
}

class GraphToolError extends Error {
  code: string;
  details?: JsonRecord;

  constructor(code: string, message: string, details?: JsonRecord) {
    super(message);
    this.code = code;
    this.details = details;
  }
}

function buildUserLabel(user: Pick<NormalizedUser, 'displayName' | 'email' | 'userPrincipalName'>): string {
  return user.displayName ?? user.email ?? user.userPrincipalName ?? 'Unknown user';
}

function buildUserCandidateLabel(
  user: Pick<NormalizedUser, 'displayName' | 'email' | 'userPrincipalName'>
): string {
  if (user.displayName && user.email) {
    return `${user.displayName} <${user.email}>`;
  }

  return buildUserLabel(user);
}

function userExactnessScore(reasons: string[]): number {
  if (reasons.includes('mail_exact') || reasons.includes('userPrincipalName_exact')) {
    return 3;
  }
  if (reasons.includes('displayName_exact')) {
    return 2;
  }
  if (reasons.includes('displayName_prefix') || reasons.includes('mail_prefix')) {
    return 1;
  }
  return 0;
}

function getMembershipOrigins(entry: JsonRecord): string[] {
  const origins = [
    getString(entry, 'source'),
    getString(entry, 'membershipSource'),
    getString(entry, 'origin'),
  ].filter((value): value is string => !!value);

  if (getArray(entry, 'indirectMembership').length > 0) {
    origins.push('indirect');
  }

  return Array.from(new Set(origins)).sort();
}

function sameEmail(left: string | null, right: string | null): boolean {
  if (!left || !right) {
    return false;
  }
  return normalizeString(left) === normalizeString(right);
}

function normalizeString(value: string | null | undefined): string {
  return (value ?? '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function tokenize(value: string): string[] {
  return normalizeString(value)
    .split(/[^a-z0-9@._-]+/i)
    .map((token) => token.trim())
    .filter(Boolean);
}

function encodePathSegment(value: string): string {
  return encodeURIComponent(value).replace(/%3D/g, '=');
}

function getRequiredString(params: Record<string, unknown>, key: string): string {
  const value = getOptionalString(params, key);
  if (!value) {
    throw new Error(`Missing required parameter: ${key}`);
  }
  return value;
}

function getOptionalString(params: Record<string, unknown>, key: string): string | undefined {
  const value = params[key];
  return typeof value === 'string' ? value : undefined;
}

function getOptionalNumber(params: Record<string, unknown>, key: string, defaultValue: number): number {
  const value = params[key];
  return typeof value === 'number' ? value : defaultValue;
}

function getOptionalBoolean(
  params: Record<string, unknown>,
  key: string,
  defaultValue: boolean
): boolean {
  const value = params[key];
  return typeof value === 'boolean' ? value : defaultValue;
}

function getRequiredStringArray(params: Record<string, unknown>, key: string): string[] {
  const value = params[key];
  if (!Array.isArray(value) || value.some((entry) => typeof entry !== 'string')) {
    throw new Error(`Missing required parameter: ${key}`);
  }
  return value as string[];
}

function asRecord(value: unknown): JsonRecord | undefined {
  return value && typeof value === 'object' && !Array.isArray(value) ? (value as JsonRecord) : undefined;
}

function getString(record: JsonRecord | undefined, key: string): string | null {
  if (!record) {
    return null;
  }
  const value = record[key];
  return typeof value === 'string' ? value : null;
}

function getArray(record: JsonRecord | undefined, key: string): unknown[] {
  if (!record) {
    return [];
  }
  const value = record[key];
  return Array.isArray(value) ? value : [];
}

function getStringArray(record: JsonRecord | undefined, key: string): string[] {
  return getArray(record, key).filter((entry): entry is string => typeof entry === 'string');
}
