import { z } from 'zod';
import {
  CHAT_TYPES,
  ChatCandidate,
  CustomToolDefinition,
  CustomToolHandlerContext,
  GraphCollectionResponse,
  GraphToolError,
  JsonRecord,
  NormalizedChat,
  NormalizedChatMember,
  NormalizedUser,
  asRecord,
  calculateRecencyBonus,
  chatSelect,
  compareDates,
  decodeCursor,
  encodeCursor,
  encodePathSegment,
  errorResult,
  extractMessageSender,
  getArray,
  getOptionalBoolean,
  getOptionalNumber,
  getOptionalString,
  getRequiredString,
  graphRequest,
  normalizeChat,
  normalizeMembers,
  resolveQueryToBestUser,
  sameEmail,
  searchUsersInternal,
  success,
} from './shared.js';

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

function pickChatFields(chat: NormalizedChat): NormalizedChat {
  return {
    id: chat.id,
    chatType: chat.chatType,
    topic: chat.topic,
    lastUpdatedDateTime: chat.lastUpdatedDateTime,
    webUrl: chat.webUrl ?? null,
    memberCount: chat.memberCount,
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
    const message = (error as Error).message;
    if (message.includes('403')) {
      throw new GraphToolError('insufficient_scope', message);
    }
    if (message.includes('404')) {
      throw new GraphToolError(notFoundCode, `Chat not found: ${chatId}`);
    }
    throw new GraphToolError('chat_members_unavailable', message);
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
    return { members: await fetchChatMembers(context, chatId) };
  } catch (error) {
    if (error instanceof GraphToolError && error.code === 'chat_not_found') {
      throw error;
    }
    return {
      members: [],
      membersUnavailable: true,
      membersError: error instanceof GraphToolError ? error.code : 'chat_members_unavailable',
    };
  }
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
    throw new GraphToolError('unsupported_operation', `Unsupported chat type: ${chat.chatType}`);
  }
  const members = includeMembers ? await fetchChatMembers(context, chatId) : [];
  return {
    chat: {
      ...chat,
      memberCount: includeMembers ? members.length : chat.memberCount,
    },
    members,
  };
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
  const chats = [...(response.value ?? [])].sort((left, right) =>
    compareDates(
      typeof right.lastUpdatedDateTime === 'string' ? right.lastUpdatedDateTime : null,
      typeof left.lastUpdatedDateTime === 'string' ? left.lastUpdatedDateTime : null
    )
  );

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

async function resolveParticipantInput(
  context: CustomToolHandlerContext,
  params: Record<string, unknown>
): Promise<NormalizedUser> {
  const userId = getOptionalString(params, 'userId');
  if (userId) {
    const items = await searchUsersInternal(context, userId, 5);
    const exact = items.find((item) => item.userId === userId);
    return (
      exact ?? {
        userId,
        displayName: null,
        email: null,
        userPrincipalName: null,
        givenName: null,
        surname: null,
        jobTitle: null,
        department: null,
      }
    );
  }

  const email = getOptionalString(params, 'email');
  if (!email) {
    throw new Error('Either userId or email is required.');
  }
  return resolveQueryToBestUser(context, email);
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
    `/chats/${encodePathSegment(chatId)}/messages?$select=id,createdDateTime,from,body,summary&$top=${limit}`
  );
  return (response.value ?? []).map((message) => ({
    id: typeof message.id === 'string' ? message.id : '',
    from: extractMessageSender(message),
    createdDateTime: typeof message.createdDateTime === 'string' ? message.createdDateTime : null,
    bodyPreview:
      typeof asRecord(message.body)?.content === 'string'
        ? (asRecord(message.body)?.content as string).replace(/<[^>]+>/g, ' ').trim()
        : typeof message.summary === 'string'
          ? message.summary
          : null,
  }));
}

export async function searchMessagesInternal(
  context: CustomToolHandlerContext,
  query: string,
  participantUserId: string | undefined,
  limit: number
): Promise<JsonRecord[]> {
  const response = await graphRequest<JsonRecord>(context, '/search/query', {
    method: 'POST',
    body: JSON.stringify({
      requests: [
        {
          entityTypes: ['chatMessage'],
          query: { queryString: query },
          from: 0,
          size: limit,
        },
      ],
    }),
  });

  let items = getArray(response, 'value')
    .flatMap((entry) => getArray(asRecord(entry), 'hitsContainers'))
    .flatMap((container) => getArray(asRecord(container), 'hits'))
    .map(asRecord)
    .filter((hit): hit is JsonRecord => !!hit)
    .map((hit) => asRecord(hit.resource) ?? {})
    .map((resource) => ({
      chatId: typeof resource.chatId === 'string' ? resource.chatId : null,
      messageId: typeof resource.id === 'string' ? resource.id : null,
      chatType: typeof resource.chatType === 'string' ? resource.chatType : null,
      topic: typeof resource.topic === 'string' ? resource.topic : null,
      from: extractMessageSender(resource),
      createdDateTime:
        typeof resource.createdDateTime === 'string' ? resource.createdDateTime : null,
      bodyPreview:
        typeof asRecord(resource.body)?.content === 'string'
          ? (asRecord(resource.body)?.content as string).replace(/<[^>]+>/g, ' ').trim()
          : null,
    }));

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

export const teamsToolDefinitions: CustomToolDefinition[] = [
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
      includeMembers: z.boolean().default(true).describe('Include normalized members'),
    },
    handler: async (params, context) => {
      const chatId = getRequiredString(params, 'chatId');
      const includeMembers = getOptionalBoolean(params, 'includeMembers', true);
      const { chat, members } = await getChatDetailsInternal(context, chatId, includeMembers);
      return success(context.graphClient, includeMembers ? { ...chat, members } : chat);
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
      includeGroupChats: z.boolean().default(true).describe('Include group and meeting chats'),
      limit: z.number().int().min(1).max(25).default(10).describe('Maximum results to return'),
    },
    handler: async (params, context) => {
      const includeGroupChats = getOptionalBoolean(params, 'includeGroupChats', true);
      const limit = getOptionalNumber(params, 'limit', 10);
      const participant = await resolveParticipantInput(context, params);
      const recentChats = await listRecentChatsInternal(
        context,
        Math.max(limit * 5, 25),
        undefined,
        true
      );
      const items = recentChats.items
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
      limit: z.number().int().min(1).max(50).default(20).describe('Maximum chats to return'),
      cursor: z.string().optional().describe('Opaque pagination cursor'),
      includeMembers: z.boolean().default(true).describe('Include normalized members inline'),
    },
    handler: async (params, context) => {
      const limit = getOptionalNumber(params, 'limit', 20);
      const cursor = getOptionalString(params, 'cursor');
      const includeMembers = getOptionalBoolean(params, 'includeMembers', true);
      return success(
        context.graphClient,
        await listRecentChatsInternal(context, limit, cursor, includeMembers)
      );
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
      chatId: z.string().min(1).describe('Chat ID'),
      messageLimit: z.number().int().min(1).max(25).default(10).describe('Number of messages'),
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
      participantUserId: z
        .string()
        .min(1)
        .optional()
        .describe('Only return hits from matching chats'),
      scope: z.literal('chats').default('chats').describe('Currently only chats are supported'),
      limit: z.number().int().min(1).max(25).default(20).describe('Maximum search hits to return'),
    },
    handler: async (params, context) => {
      const scope = getOptionalString(params, 'scope') ?? 'chats';
      if (scope !== 'chats') {
        return errorResult(context.graphClient, 'unsupported_operation', {
          message: 'Only scope="chats" is currently supported.',
        });
      }
      const query = getRequiredString(params, 'query');
      const participantUserId = getOptionalString(params, 'participantUserId');
      const limit = getOptionalNumber(params, 'limit', 20);
      return success(context.graphClient, {
        items: await searchMessagesInternal(context, query, participantUserId, limit),
      });
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
      return success(context.graphClient, {
        teamId,
        channelId,
        members: normalizeMembers(response.value ?? []),
      });
    },
  },
];
