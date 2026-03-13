import { z } from 'zod';
import {
  CustomToolDefinition,
  CustomToolHandlerContext,
  GraphCollectionResponse,
  JsonRecord,
  ambiguousMatchResult,
  asRecord,
  buildUserCandidateLabel,
  errorResult,
  extractStatusMessage,
  getOptionalString,
  getRequiredString,
  getRequiredStringArray,
  graphRequest,
  normalizeUser,
  resolveQueryToBestUser,
  searchUsersInternal,
  success,
  userSelect,
} from './shared.js';

async function getUserById(
  context: CustomToolHandlerContext,
  userId: string
): Promise<ReturnType<typeof normalizeUser>> {
  const response = await graphRequest<JsonRecord>(
    context,
    `/users/${encodeURIComponent(userId)}?${userSelect}`
  );
  return normalizeUser(response);
}

export const peopleToolDefinitions: CustomToolDefinition[] = [
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
      limit: z.number().int().min(1).max(25).default(10).describe('Maximum results to return'),
    },
    handler: async (params, context) => {
      const query = getRequiredString(params, 'query');
      const limit = typeof params.limit === 'number' ? params.limit : 10;
      const items = await searchUsersInternal(context, query, limit);
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
        return errorResult(context.graphClient, 'not_found', { query });
      }

      const [best, second] = items;
      if (!best || best.matchScore < 0.85 || (second && best.matchScore - second.matchScore < 0.1)) {
        return ambiguousMatchResult(
          context.graphClient,
          items.slice(0, 5).map((item, index) => ({
            userId: item.userId,
            label: buildUserCandidateLabel(item),
            rank: index + 1,
          }))
        );
      }

      return success(context.graphClient, { items: [best] });
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
      userIds: z.array(z.string().min(1)).min(1).max(100).describe('Microsoft user IDs to look up'),
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
      return success(context.graphClient, {
        items: source.map((entry) => ({
          userId: typeof entry.id === 'string' ? entry.id : '',
          availability: typeof entry.availability === 'string' ? entry.availability : null,
          activity: typeof entry.activity === 'string' ? entry.activity : null,
          statusMessage: extractStatusMessage(entry),
        })),
      });
    },
  },
  {
    name: 'get-user',
    description:
      'Retrieve one normalized Microsoft user by userId or resolve a user by email/name query.',
    method: 'GET',
    path: '/users/{userId}',
    requiresOrgMode: true,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['User.Read.All'],
    schema: {
      userId: z.string().min(1).optional().describe('Microsoft user ID'),
      query: z.string().min(1).optional().describe('Email address or human name to resolve'),
    },
    handler: async (params, context) => {
      const userId = getOptionalString(params, 'userId');
      const query = getOptionalString(params, 'query');

      if (!userId && !query) {
        throw new Error('Either userId or query is required.');
      }

      const user = userId
        ? await getUserById(context, userId)
        : await resolveQueryToBestUser(context, query!);
      return success(context.graphClient, { items: [user] });
    },
  },
  {
    name: 'get-manager',
    description: 'Get the manager of a Microsoft user.',
    method: 'GET',
    path: '/users/{userId}/manager',
    requiresOrgMode: true,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['User.Read.All'],
    schema: {
      userId: z.string().min(1).describe('Microsoft user ID'),
    },
    handler: async (params, context) => {
      const userId = getRequiredString(params, 'userId');
      const response = await graphRequest<JsonRecord>(
        context,
        `/users/${encodeURIComponent(userId)}/manager?${userSelect}`
      );
      return success(context.graphClient, { items: [normalizeUser(response)] });
    },
  },
  {
    name: 'get-direct-reports',
    description: 'Get the direct reports of a Microsoft user.',
    method: 'GET',
    path: '/users/{userId}/directReports',
    requiresOrgMode: true,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['User.Read.All'],
    schema: {
      userId: z.string().min(1).describe('Microsoft user ID'),
      limit: z.number().int().min(1).max(50).default(25).describe('Maximum results to return'),
    },
    handler: async (params, context) => {
      const userId = getRequiredString(params, 'userId');
      const limit = typeof params.limit === 'number' ? params.limit : 25;
      const response = await graphRequest<GraphCollectionResponse<JsonRecord>>(
        context,
        `/users/${encodeURIComponent(userId)}/directReports?${userSelect}&$top=${limit}`
      );
      return success(context.graphClient, {
        items: (response.value ?? [])
          .map(asRecord)
          .filter((entry): entry is JsonRecord => !!entry)
          .map(normalizeUser),
      });
    },
  },
  {
    name: 'search-people-across-surfaces',
    description:
      'Search people across directory and lightweight activity signals for Canvas disambiguation.',
    method: 'GET',
    path: '/users',
    requiresOrgMode: true,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['User.Read.All', 'People.Read'],
    schema: {
      query: z.string().min(1).describe('Name or email query'),
      limit: z.number().int().min(1).max(25).default(10).describe('Maximum results to return'),
    },
    handler: async (params, context) => {
      const query = getRequiredString(params, 'query');
      const limit = typeof params.limit === 'number' ? params.limit : 10;
      const items = await searchUsersInternal(context, query, limit);
      return success(context.graphClient, {
        items: items.map((item) => ({
          ...item,
          labels: [item.department, item.jobTitle].filter(Boolean),
        })),
      });
    },
  },
];
