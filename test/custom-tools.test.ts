import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { registerDiscoveryTools, registerGraphTools } from '../src/graph-tools.js';
import GraphClient from '../src/graph-client.js';
import { buildScopesFromEndpoints } from '../src/auth.js';

vi.mock('../src/logger.js', () => ({
  default: {
    info: vi.fn(),
    error: vi.fn(),
    warn: vi.fn(),
  },
}));

vi.mock('../src/generated/client.js', () => ({
  api: {
    endpoints: [
      {
        alias: 'list-mail-messages',
        method: 'GET',
        path: '/me/messages',
        description: 'List mail messages',
        parameters: [],
      },
    ],
  },
}));

type ToolHandler = (params: Record<string, unknown>) => Promise<{
  content: Array<{ type: string; text: string }>;
  isError?: boolean;
}>;

function createMockGraphClient(
  makeRequest: (endpoint: string, options?: Record<string, unknown>) => Promise<unknown>
): GraphClient {
  return {
    makeRequest: vi.fn(makeRequest),
    formatJsonResponse: vi.fn((data: unknown) => ({
      content: [{ type: 'text', text: JSON.stringify(data) }],
    })),
  } as unknown as GraphClient;
}

function captureHandlers(server: McpServer): Map<string, ToolHandler> {
  const handlers = new Map<string, ToolHandler>();
  vi.spyOn(server, 'tool').mockImplementation(((...args: unknown[]) => {
    const name = args[0] as string;
    const handler = args[args.length - 1];
    if (typeof handler === 'function') {
      handlers.set(name, handler as ToolHandler);
    }
  }) as never);
  return handlers;
}

function parseResult(result: Awaited<ReturnType<ToolHandler>>): unknown {
  return JSON.parse(result.content[0].text);
}

describe('custom Canvas-oriented tools', () => {
  beforeEach(() => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date('2026-03-12T12:00:00Z'));
  });

  afterEach(() => {
    vi.useRealTimers();
    vi.restoreAllMocks();
  });

  it('registers custom tools in org mode and respects filtering', () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const graphClient = createMockGraphClient(async () => ({ value: [] }));

    registerGraphTools(server, graphClient, false, 'search-users|get-chat-details', true);

    expect(Array.from(handlers.keys())).toEqual(['search-users', 'get-chat-details']);
  });

  it('includes custom tools in discovery mode results', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const graphClient = createMockGraphClient(async () => ({ value: [] }));

    registerDiscoveryTools(server, graphClient, false, undefined, true);

    const searchTools = handlers.get('search-tools');
    expect(searchTools).toBeDefined();

    const result = await searchTools!({ limit: 50 });
    const parsed = parseResult(result) as { tools: Array<{ name: string }> };

    expect(parsed.tools.some((tool) => tool.name === 'find-chats-by-participant')).toBe(true);
  });

  it('applies enabled-tools filtering in discovery mode', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const graphClient = createMockGraphClient(async () => ({ value: [] }));

    registerDiscoveryTools(server, graphClient, false, 'search-users', true);

    const searchTools = handlers.get('search-tools');
    const executeTool = handlers.get('execute-tool');
    expect(searchTools).toBeDefined();
    expect(executeTool).toBeDefined();

    const searchResult = await searchTools!({ limit: 20 });
    const parsedSearch = parseResult(searchResult) as { tools: Array<{ name: string }> };
    expect(parsedSearch.tools.map((tool) => tool.name)).toEqual(['search-users']);

    const executeResult = await executeTool!({
      tool_name: 'find-chats-by-participant',
      parameters: {},
    });
    const parsedExecute = parseResult(executeResult) as { error: string };
    expect(executeResult.isError).toBe(true);
    expect(parsedExecute.error).toContain('Tool not found');
  });

  it('adds custom tool scopes when building filtered auth scopes', () => {
    const scopes = buildScopesFromEndpoints(true, 'search-users|list-user-presence');

    expect(scopes).toContain('User.Read.All');
    expect(scopes).toContain('Presence.Read.All');
    expect(scopes).not.toContain('Chat.Read');
  });

  it('search-users ranks exact display name matches first', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const graphClient = createMockGraphClient(async (endpoint, options) => {
      expect(endpoint).toContain('/users?');
      expect((options?.headers as Record<string, string>).ConsistencyLevel).toBe('eventual');
      return {
        value: [
          {
            id: 'user-1',
            displayName: 'Juan Casiraghi',
            mail: 'juan@company.com',
            userPrincipalName: 'juan@company.com',
            givenName: 'Juan',
            surname: 'Casiraghi',
          },
          {
            id: 'user-2',
            displayName: 'Juan Carlos',
            mail: 'juanc@company.com',
            userPrincipalName: 'juanc@company.com',
          },
        ],
      };
    });

    registerGraphTools(server, graphClient, false, 'search-users', true);

    const result = await handlers.get('search-users')!({
      query: 'Juan Casiraghi',
      limit: 10,
    });
    const parsed = parseResult(result) as {
      items: Array<{ userId: string; matchReasons: string[] }>;
    };

    expect(parsed.items[0].userId).toBe('user-1');
    expect(parsed.items[0].matchReasons).toContain('displayName_exact');
  });

  it('list-chat-members normalizes roster entries and maps not found errors', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const graphClient = createMockGraphClient(async (endpoint) => {
      if (endpoint.includes('/missing-chat/')) {
        throw new Error('Microsoft Graph API error: 404 Not Found');
      }

      return {
        value: [
          {
            id: 'membership-1',
            userId: 'user-1',
            displayName: 'Juan Casiraghi',
            email: 'juan@company.com',
            userPrincipalName: 'juan@company.com',
            roles: ['owner'],
          },
        ],
      };
    });

    registerGraphTools(server, graphClient, false, 'list-chat-members', true);

    const successResult = await handlers.get('list-chat-members')!({ chatId: 'chat-1' });
    const parsedSuccess = parseResult(successResult) as {
      members: Array<{ userId: string; membershipType: string }>;
    };
    expect(parsedSuccess.members[0]).toMatchObject({
      userId: 'user-1',
      membershipType: 'owner',
    });

    const errorResult = await handlers.get('list-chat-members')!({ chatId: 'missing-chat' });
    const parsedError = parseResult(errorResult) as { error: string };
    expect(errorResult.isError).toBe(true);
    expect(parsedError.error).toBe('chat_not_found');
  });

  it('find-chats-by-participant ranks exact one-on-one chats before group chats', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const graphClient = createMockGraphClient(async (endpoint) => {
      if (endpoint.startsWith('/users?')) {
        return {
          value: [
            {
              id: 'juan-id',
              displayName: 'Juan Casiraghi',
              mail: 'juan@company.com',
              userPrincipalName: 'juan@company.com',
            },
          ],
        };
      }

      if (endpoint.startsWith('/me/chats?')) {
        return {
          value: [
            {
              id: 'group-chat',
              chatType: 'group',
              topic: 'AI Leads',
              lastUpdatedDateTime: '2026-03-11T10:00:00Z',
            },
            {
              id: 'one-chat',
              chatType: 'oneOnOne',
              topic: null,
              lastUpdatedDateTime: '2026-03-12T11:00:00Z',
            },
          ],
        };
      }

      if (endpoint === '/chats/group-chat/members') {
        return {
          value: [
            { id: 'm1', userId: 'juan-id', displayName: 'Juan Casiraghi', email: 'juan@company.com' },
            { id: 'm2', userId: 'other-id', displayName: 'Other User', email: 'other@company.com' },
            { id: 'm3', userId: 'me-id', displayName: 'Me', email: 'me@company.com' },
          ],
        };
      }

      if (endpoint === '/chats/one-chat/members') {
        return {
          value: [
            { id: 'm1', userId: 'juan-id', displayName: 'Juan Casiraghi', email: 'juan@company.com' },
            { id: 'm2', userId: 'me-id', displayName: 'Me', email: 'me@company.com' },
          ],
        };
      }

      throw new Error(`Unexpected endpoint: ${endpoint}`);
    });

    registerGraphTools(server, graphClient, false, 'find-chats-by-participant', true);

    const result = await handlers.get('find-chats-by-participant')!({
      email: 'juan@company.com',
      includeGroupChats: true,
      limit: 10,
    });
    const parsed = parseResult(result) as {
      items: Array<{ chat: { id: string }; matchType: string; rank: number }>;
    };

    expect(parsed.items[0]).toMatchObject({
      matchType: 'exact_one_on_one',
      rank: 1,
      chat: { id: 'one-chat' },
    });
    expect(parsed.items[1].chat.id).toBe('group-chat');
  });

  it('skips chats whose member rosters are temporarily unavailable in bulk resolution flows', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const graphClient = createMockGraphClient(async (endpoint) => {
      if (endpoint.startsWith('/users?')) {
        return {
          value: [
            {
              id: 'juan-id',
              displayName: 'Juan Casiraghi',
              mail: 'juan@company.com',
              userPrincipalName: 'juan@company.com',
            },
          ],
        };
      }

      if (endpoint.startsWith('/me/chats?')) {
        return {
          value: [
            {
              id: 'bad-chat',
              chatType: 'group',
              topic: 'Meeting',
              lastUpdatedDateTime: '2026-03-12T11:30:00Z',
            },
            {
              id: 'good-chat',
              chatType: 'oneOnOne',
              topic: null,
              lastUpdatedDateTime: '2026-03-12T11:00:00Z',
            },
          ],
        };
      }

      if (endpoint === '/chats/bad-chat/members') {
        throw new Error('Microsoft Graph API error: 403 Forbidden - scope error');
      }

      if (endpoint === '/chats/good-chat/members') {
        return {
          value: [
            { id: 'm1', userId: 'juan-id', displayName: 'Juan Casiraghi', email: 'juan@company.com' },
            { id: 'm2', userId: 'me-id', displayName: 'Me', email: 'me@company.com' },
          ],
        };
      }

      throw new Error(`Unexpected endpoint: ${endpoint}`);
    });

    registerGraphTools(server, graphClient, false, 'find-chats-by-participant|list-recent-chats', true);

    const recentResult = await handlers.get('list-recent-chats')!({
      limit: 10,
      includeMembers: true,
    });
    const parsedRecent = parseResult(recentResult) as {
      items: Array<{ id: string; membersUnavailable?: boolean; members?: Array<unknown> }>;
    };
    expect(parsedRecent.items).toHaveLength(2);
    expect(parsedRecent.items[0]).toMatchObject({
      id: 'bad-chat',
      membersUnavailable: true,
      members: [],
    });

    const findResult = await handlers.get('find-chats-by-participant')!({
      email: 'juan@company.com',
      includeGroupChats: true,
      limit: 10,
    });
    const parsedFind = parseResult(findResult) as { items: Array<{ chat: { id: string } }> };
    expect(parsedFind.items).toHaveLength(1);
    expect(parsedFind.items[0].chat.id).toBe('good-chat');
  });

  it('get-chat-context returns chat, members, and recent messages', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const graphClient = createMockGraphClient(async (endpoint) => {
      if (endpoint.startsWith('/chats/chat-1?')) {
        return {
          id: 'chat-1',
          chatType: 'oneOnOne',
          topic: null,
          lastUpdatedDateTime: '2026-03-12T11:00:00Z',
          webUrl: 'https://teams.microsoft.com/l/chat/0/0?users=juan@company.com',
        };
      }

      if (endpoint === '/chats/chat-1/members') {
        return {
          value: [
            { id: 'm1', userId: 'juan-id', displayName: 'Juan Casiraghi', email: 'juan@company.com' },
            { id: 'm2', userId: 'me-id', displayName: 'Me', email: 'me@company.com' },
          ],
        };
      }

      if (endpoint.startsWith('/chats/chat-1/messages?')) {
        return {
          value: [
            {
              id: 'msg-1',
              createdDateTime: '2026-03-12T10:59:00Z',
              from: {
                user: { id: 'juan-id', displayName: 'Juan Casiraghi' },
              },
              body: { content: '<div>deploy prod tonight</div>' },
            },
          ],
        };
      }

      throw new Error(`Unexpected endpoint: ${endpoint}`);
    });

    registerGraphTools(server, graphClient, false, 'get-chat-context', true);

    const result = await handlers.get('get-chat-context')!({
      chatId: 'chat-1',
      messageLimit: 5,
    });
    const parsed = parseResult(result) as {
      chat: { id: string };
      members: Array<{ userId: string }>;
      recentMessages: Array<{ bodyPreview: string }>;
    };

    expect(parsed.chat.id).toBe('chat-1');
    expect(parsed.members).toHaveLength(2);
    expect(parsed.recentMessages[0].bodyPreview).toBe('deploy prod tonight');
  });

  it('search-messages narrows search hits by participant membership', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const graphClient = createMockGraphClient(async (endpoint, options) => {
      if (endpoint === '/search/query') {
        expect(options?.method).toBe('POST');
        return {
          value: [
            {
              hitsContainers: [
                {
                  hits: [
                    {
                      resource: {
                        id: 'msg-1',
                        chatId: 'chat-1',
                        chatType: 'group',
                        topic: 'AI Leads',
                        createdDateTime: '2026-03-12T10:59:00Z',
                        from: {
                          user: { id: 'juan-id', displayName: 'Juan Casiraghi' },
                        },
                        body: { content: '<p>deploy prod tonight</p>' },
                      },
                    },
                    {
                      resource: {
                        id: 'msg-2',
                        chatId: 'chat-2',
                        chatType: 'group',
                        topic: 'Infra',
                        createdDateTime: '2026-03-12T09:59:00Z',
                        from: {
                          user: { id: 'other-id', displayName: 'Other User' },
                        },
                        body: { content: '<p>deploy prod tomorrow</p>' },
                      },
                    },
                  ],
                },
              ],
            },
          ],
        };
      }

      if (endpoint === '/chats/chat-1/members') {
        return {
          value: [{ id: 'm1', userId: 'juan-id', displayName: 'Juan Casiraghi' }],
        };
      }

      if (endpoint === '/chats/chat-2/members') {
        return {
          value: [{ id: 'm2', userId: 'other-id', displayName: 'Other User' }],
        };
      }

      throw new Error(`Unexpected endpoint: ${endpoint}`);
    });

    registerGraphTools(server, graphClient, false, 'search-messages', true);

    const result = await handlers.get('search-messages')!({
      query: 'deploy prod',
      participantUserId: 'juan-id',
      scope: 'chats',
      limit: 20,
    });
    const parsed = parseResult(result) as {
      items: Array<{ chatId: string; messageId: string }>;
    };

    expect(parsed.items).toHaveLength(1);
    expect(parsed.items[0]).toMatchObject({ chatId: 'chat-1', messageId: 'msg-1' });
  });

  it('resolve-person returns structured ambiguity when matches are too close', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const graphClient = createMockGraphClient(async () => ({
      value: [
        {
          id: 'user-1',
          displayName: 'Juan Casiraghi',
          mail: 'juan@company.com',
          userPrincipalName: 'juan@company.com',
        },
        {
          id: 'user-2',
          displayName: 'Juan Casiraghi',
          mail: 'juan.casiraghi@company.com',
          userPrincipalName: 'juan.casiraghi@company.com',
        },
      ],
    }));

    registerGraphTools(server, graphClient, false, 'resolve-person', true);

    const result = await handlers.get('resolve-person')!({ query: 'Juan Casiraghi' });
    const parsed = parseResult(result) as { error: string; candidates: Array<{ rank: number }> };

    expect(result.isError).toBe(true);
    expect(parsed.error).toBe('ambiguous_match');
    expect(parsed.candidates[0].rank).toBe(1);
  });

  it('deduplicates channel members and preserves membership origin hints', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const graphClient = createMockGraphClient(async (endpoint) => {
      if (endpoint === '/teams/team-1/channels/channel-1/allMembers') {
        return {
          value: [
            {
              id: 'member-1',
              userId: 'juan-id',
              displayName: 'Juan Casiraghi',
              email: 'juan@company.com',
              roles: ['owner'],
              source: 'direct',
            },
            {
              id: 'member-2',
              userId: 'juan-id',
              displayName: 'Juan Casiraghi',
              email: 'juan@company.com',
              roles: ['member'],
              source: 'indirect',
            },
          ],
        };
      }

      throw new Error(`Unexpected endpoint: ${endpoint}`);
    });

    registerGraphTools(server, graphClient, false, 'list-channel-members', true);

    const result = await handlers.get('list-channel-members')!({
      teamId: 'team-1',
      channelId: 'channel-1',
    });
    const parsed = parseResult(result) as {
      members: Array<{ userId: string; roles: string[]; membershipOrigins?: string[] }>;
    };

    expect(parsed.members).toHaveLength(1);
    expect(parsed.members[0]).toMatchObject({
      userId: 'juan-id',
      roles: ['member', 'owner'],
      membershipOrigins: ['direct', 'indirect'],
    });
  });
});
