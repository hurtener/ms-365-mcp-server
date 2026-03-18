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

  it('exposes the availability-based calendar creator as an event write tool in discovery mode', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const graphClient = createMockGraphClient(async () => ({ value: [] }));

    registerDiscoveryTools(
      server,
      graphClient,
      false,
      'create-calendar-event-from-availability',
      false
    );

    const searchTools = handlers.get('search-tools');
    expect(searchTools).toBeDefined();

    const result = await searchTools!({ limit: 10 });
    const parsed = parseResult(result) as {
      tools: Array<{ name: string; method: string; path: string }>;
    };

    expect(parsed.tools).toEqual([
      {
        name: 'create-calendar-event-from-availability',
        method: 'POST',
        path: '/me/events',
        description:
          'Find the best slot where attendees are available, then create the calendar event there.',
      },
    ]);
  });

  it('includes personal-surface custom tools in discovery mode without org mode', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const graphClient = createMockGraphClient(async () => ({ value: [] }));

    registerDiscoveryTools(server, graphClient, false, 'resolve-drive-path|search-mail', false);

    const searchTools = handlers.get('search-tools');
    expect(searchTools).toBeDefined();

    const result = await searchTools!({ limit: 20 });
    const parsed = parseResult(result) as { tools: Array<{ name: string }> };

    expect(parsed.tools.map((tool) => tool.name).sort()).toEqual([
      'resolve-drive-path',
      'search-mail',
    ]);
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

  it('does not register mutating custom tools in read-only mode', () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const graphClient = createMockGraphClient(async () => ({ value: [] }));

    registerGraphTools(
      server,
      graphClient,
      true,
      'create-text-file|search-mail|complete-task|list-tasks',
      false
    );

    expect(Array.from(handlers.keys()).sort()).toEqual(['list-tasks', 'search-mail']);
  });

  it('does not include deactivated tools in filtered auth scopes', () => {
    const scopes = buildScopesFromEndpoints(true, 'search-users|list-user-presence');

    expect(scopes).toContain('User.Read.All');
    expect(scopes).not.toContain('Presence.Read.All');
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

      if (endpoint === '/me/chats?$top=50') {
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
            {
              id: 'm1',
              userId: 'juan-id',
              displayName: 'Juan Casiraghi',
              email: 'juan@company.com',
            },
            { id: 'm2', userId: 'other-id', displayName: 'Other User', email: 'other@company.com' },
            { id: 'm3', userId: 'me-id', displayName: 'Me', email: 'me@company.com' },
          ],
        };
      }

      if (endpoint === '/chats/one-chat/members') {
        return {
          value: [
            {
              id: 'm1',
              userId: 'juan-id',
              displayName: 'Juan Casiraghi',
              email: 'juan@company.com',
            },
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

      if (endpoint === '/me/chats?$top=50' || endpoint === '/me/chats?$top=10') {
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
            {
              id: 'm1',
              userId: 'juan-id',
              displayName: 'Juan Casiraghi',
              email: 'juan@company.com',
            },
            { id: 'm2', userId: 'me-id', displayName: 'Me', email: 'me@company.com' },
          ],
        };
      }

      throw new Error(`Unexpected endpoint: ${endpoint}`);
    });

    registerGraphTools(
      server,
      graphClient,
      false,
      'find-chats-by-participant|list-recent-chats',
      true
    );

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
      if (endpoint === '/chats/chat-1') {
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
            {
              id: 'm1',
              userId: 'juan-id',
              displayName: 'Juan Casiraghi',
              email: 'juan@company.com',
            },
            { id: 'm2', userId: 'me-id', displayName: 'Me', email: 'me@company.com' },
          ],
        };
      }

      if (endpoint === '/chats/chat-1/messages?$top=5') {
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

  it('search-messages hydrates sparse Graph search hits with chat metadata and hit summaries', async () => {
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
                      summary: 'Mauro Beneitez mentioned in thread',
                      resource: {
                        id: 'msg-1',
                        chatId: 'chat-1',
                        createdDateTime: '2026-03-12T10:59:00Z',
                      },
                    },
                  ],
                },
              ],
            },
          ],
        };
      }

      if (endpoint === '/chats/chat-1') {
        return {
          id: 'chat-1',
          chatType: 'oneOnOne',
          topic: null,
          lastUpdatedDateTime: '2026-03-12T11:00:00Z',
        };
      }

      throw new Error(`Unexpected endpoint: ${endpoint}`);
    });

    registerGraphTools(server, graphClient, false, 'search-messages', true);

    const result = await handlers.get('search-messages')!({
      query: 'Mauro Beneitez',
      scope: 'chats',
      limit: 10,
    });
    const parsed = parseResult(result) as {
      items: Array<{
        chatId: string;
        chatType: string | null;
        topic: string | null;
        bodyPreview: string | null;
      }>;
    };

    expect(parsed.items).toEqual([
      {
        chatId: 'chat-1',
        messageId: 'msg-1',
        chatType: 'oneOnOne',
        topic: null,
        from: {
          userId: null,
          displayName: null,
        },
        createdDateTime: '2026-03-12T10:59:00Z',
        bodyPreview: 'Mauro Beneitez mentioned in thread',
      },
    ]);
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

  it('adds expanded custom tool scopes for files, sites, calendar, tasks, mail, and people', () => {
    const scopes = buildScopesFromEndpoints(
      true,
      'resolve-drive-path|list-sites|search-calendar-events|search-tasks|search-mail|get-user'
    );

    expect(scopes).toContain('Files.Read');
    expect(scopes).toContain('Sites.Read.All');
    expect(scopes).toContain('Calendars.Read');
    expect(scopes).toContain('Tasks.Read');
    expect(scopes).toContain('Mail.Read');
    expect(scopes).toContain('User.Read.All');
  });

  it('resolve-drive-path prefers explicit driveId and returns normalized file metadata', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const graphClient = createMockGraphClient(async (endpoint) => {
      expect(endpoint).toBe(
        '/drives/drive-1/root:/Projects/Q1/plan.docx:?$select=id,name,size,lastModifiedDateTime,webUrl,parentReference,folder,file,lastModifiedBy'
      );
      return {
        id: 'item-1',
        name: 'plan.docx',
        size: 12345,
        webUrl: 'https://contoso.sharepoint.com/plan.docx',
        file: {
          mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        },
        parentReference: {
          driveId: 'drive-1',
          path: '/drives/drive-1/root:/Projects/Q1',
        },
      };
    });

    registerGraphTools(server, graphClient, false, 'resolve-drive-path', false);

    const result = await handlers.get('resolve-drive-path')!({
      path: '/Projects/Q1/plan.docx',
      driveId: 'drive-1',
    });
    const parsed = parseResult(result) as {
      id: string;
      driveId: string;
      path: string;
      name: string;
      isFolder: boolean;
    };

    expect(parsed).toMatchObject({
      id: 'item-1',
      driveId: 'drive-1',
      path: '/Projects/Q1/plan.docx',
      name: 'plan.docx',
      isFolder: false,
    });
  });

  it('list-sites returns normalized site summaries in org mode', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const graphClient = createMockGraphClient(async (endpoint, options) => {
      expect(endpoint).toContain('/sites?');
      expect((options?.headers as Record<string, string>).ConsistencyLevel).toBe('eventual');
      return {
        value: [
          {
            id: 'site-1',
            name: 'marketing',
            displayName: 'Marketing Team',
            webUrl: 'https://contoso.sharepoint.com/sites/marketing',
          },
        ],
      };
    });

    registerGraphTools(server, graphClient, false, 'list-sites', true);

    const result = await handlers.get('list-sites')!({ limit: 10 });
    const parsed = parseResult(result) as {
      items: Array<{ siteId: string; displayName: string; isPersonalSite: boolean }>;
    };

    expect(parsed.items[0]).toMatchObject({
      siteId: 'site-1',
      displayName: 'Marketing Team',
      isPersonalSite: false,
      label: 'Marketing Team',
    });
  });

  it('search-sites supports query pagination and stable labels', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const graphClient = createMockGraphClient(async (endpoint, options) => {
      expect((options?.headers as Record<string, string>).ConsistencyLevel).toBe('eventual');
      if (endpoint === '/sites?$search=%22Marketing%22&$skiptoken=next') {
        return {
          value: [
            {
              id: 'site-2',
              name: 'marketing-ops',
              displayName: 'Marketing Ops',
              webUrl: 'https://contoso.sharepoint.com/sites/marketing-ops',
            },
          ],
        };
      }

      if (
        endpoint === '/sites?$search=%22Marketing%22&$top=10&$select=id,name,displayName,webUrl'
      ) {
        return {
          value: [
            {
              id: 'site-1',
              name: 'marketing',
              displayName: 'Marketing Team',
              webUrl: 'https://contoso.sharepoint.com/sites/marketing',
            },
          ],
          '@odata.nextLink':
            'https://graph.microsoft.com/v1.0/sites?$search=%22Marketing%22&$skiptoken=next',
        };
      }

      throw new Error(`Unexpected endpoint: ${endpoint}`);
    });

    registerGraphTools(server, graphClient, false, 'search-sites', true);

    const first = parseResult(
      await handlers.get('search-sites')!({ query: 'Marketing', limit: 10 })
    ) as {
      items: Array<{ siteId: string; label: string }>;
      nextCursor?: string;
    };

    expect(first.items[0]).toMatchObject({
      siteId: 'site-1',
      label: 'Marketing Team',
    });
    expect(first.nextCursor).toBeTruthy();

    const second = parseResult(
      await handlers.get('search-sites')!({
        query: 'Marketing',
        limit: 10,
        cursor: first.nextCursor,
      })
    ) as {
      items: Array<{ siteId: string; label: string }>;
    };

    expect(second.items).toEqual([
      {
        siteId: 'site-2',
        name: 'marketing-ops',
        displayName: 'Marketing Ops',
        webUrl: 'https://contoso.sharepoint.com/sites/marketing-ops',
        isPersonalSite: false,
        label: 'Marketing Ops',
      },
    ]);
  });

  it('lists document libraries with labels and pagination', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const graphClient = createMockGraphClient(async (endpoint) => {
      if (endpoint === '/sites/site-1/drives?$top=2&$select=id,name,driveType,webUrl') {
        return {
          value: [
            {
              id: 'drive-1',
              name: 'Shared Documents',
              driveType: 'documentLibrary',
              webUrl: 'https://contoso.sharepoint.com/sites/marketing/Shared%20Documents',
            },
          ],
          '@odata.nextLink': 'https://graph.microsoft.com/v1.0/sites/site-1/drives?$skiptoken=next',
        };
      }
      if (endpoint === '/sites/site-1/drives?$skiptoken=next') {
        return {
          value: [
            {
              id: 'drive-2',
              name: 'Brand Assets',
              driveType: 'documentLibrary',
              webUrl: 'https://contoso.sharepoint.com/sites/marketing/Brand%20Assets',
            },
          ],
        };
      }
      throw new Error(`Unexpected endpoint: ${endpoint}`);
    });

    registerGraphTools(server, graphClient, false, 'list-document-libraries', true);

    const first = parseResult(
      await handlers.get('list-document-libraries')!({ siteId: 'site-1', limit: 2 })
    ) as {
      items: Array<{ driveId: string; label: string }>;
      nextCursor?: string;
    };

    expect(first.items[0]).toMatchObject({
      driveId: 'drive-1',
      label: 'Shared Documents',
    });
    expect(first.nextCursor).toBeTruthy();

    const second = parseResult(
      await handlers.get('list-document-libraries')!({
        siteId: 'site-1',
        limit: 2,
        cursor: first.nextCursor,
      })
    ) as {
      items: Array<{ driveId: string; label: string }>;
    };

    expect(second.items[0]).toMatchObject({
      driveId: 'drive-2',
      label: 'Brand Assets',
    });
  });

  it('lists a SharePoint-backed folder by siteId and path', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const graphClient = createMockGraphClient(async (endpoint) => {
      if (
        endpoint ===
        '/sites/site-1/drive/root:/Shared Documents/Plans:?$select=id,name,size,lastModifiedDateTime,webUrl,parentReference,folder,file,lastModifiedBy'
      ) {
        return {
          id: 'folder-1',
          name: 'Plans',
          parentReference: {
            driveId: 'drive-1',
            path: '/drives/drive-1/root:/Shared Documents',
          },
          folder: {},
          webUrl: 'https://contoso.sharepoint.com/sites/marketing/Shared%20Documents/Plans',
        };
      }

      if (
        endpoint ===
        '/drives/drive-1/items/folder-1/children?$top=2&$select=id,name,size,lastModifiedDateTime,webUrl,parentReference,folder,file,lastModifiedBy'
      ) {
        return {
          value: [
            {
              id: 'item-1',
              name: 'plan.docx',
              parentReference: {
                driveId: 'drive-1',
                path: '/drives/drive-1/root:/Shared Documents/Plans',
                siteId: 'site-1',
              },
              file: {
                mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
              },
              webUrl:
                'https://contoso.sharepoint.com/sites/marketing/Shared%20Documents/Plans/plan.docx',
            },
          ],
        };
      }

      throw new Error(`Unexpected endpoint: ${endpoint}`);
    });

    registerGraphTools(server, graphClient, false, 'list-folder', false);

    const result = parseResult(
      await handlers.get('list-folder')!({
        siteId: 'site-1',
        path: '/Shared Documents/Plans',
        limit: 2,
      })
    ) as {
      items: Array<{ id: string; path: string; label: string }>;
    };

    expect(result.items[0]).toMatchObject({
      id: 'item-1',
      path: '/Shared Documents/Plans/plan.docx',
      label: 'plan.docx (/Shared Documents/Plans)',
    });
  });

  it('search-site-files supports site/path narrowing with opaque cursors', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const graphClient = createMockGraphClient(async (endpoint, options) => {
      if (endpoint === '/sites/site-1?$select=webUrl') {
        return { webUrl: 'https://contoso.sharepoint.com/sites/marketing' };
      }
      if (endpoint === '/search/query') {
        const body = JSON.parse((options?.body as string) ?? '{}') as {
          requests: Array<{ from: number; size: number; query: { queryString: string } }>;
        };
        if (body.requests[0]?.from === 0) {
          return {
            value: [
              {
                hitsContainers: [
                  {
                    moreResultsAvailable: true,
                    hits: [
                      {
                        resource: {
                          id: 'item-1',
                          name: 'plan.docx',
                          webUrl:
                            'https://contoso.sharepoint.com/sites/marketing/Shared%20Documents/Plans/plan.docx',
                          parentReference: {
                            driveId: 'drive-1',
                            path: '/drives/drive-1/root:/Shared Documents/Plans',
                            siteId: 'site-1',
                          },
                          file: {
                            mimeType:
                              'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                          },
                        },
                      },
                    ],
                  },
                ],
              },
            ],
          };
        }
        if (body.requests[0]?.from === 25) {
          return {
            value: [
              {
                hitsContainers: [
                  {
                    hits: [
                      {
                        resource: {
                          id: 'item-2',
                          name: 'roadmap.docx',
                          webUrl:
                            'https://contoso.sharepoint.com/sites/marketing/Shared%20Documents/Plans/roadmap.docx',
                          parentReference: {
                            driveId: 'drive-1',
                            path: '/drives/drive-1/root:/Shared Documents/Plans',
                            siteId: 'site-1',
                          },
                          file: {
                            mimeType:
                              'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                          },
                        },
                      },
                    ],
                  },
                ],
              },
            ],
          };
        }
      }
      throw new Error(`Unexpected endpoint: ${endpoint}`);
    });

    registerGraphTools(server, graphClient, false, 'search-site-files', true);

    const first = parseResult(
      await handlers.get('search-site-files')!({
        siteId: 'site-1',
        query: 'plan',
        path: '/Shared Documents/Plans',
        limit: 2,
      })
    ) as {
      items: Array<{ id: string; path: string }>;
      nextCursor?: string;
    };

    expect(first.items).toHaveLength(1);
    expect(first.items[0].path).toBe('/Shared Documents/Plans/plan.docx');
    expect(first.nextCursor).toBeTruthy();

    const second = parseResult(
      await handlers.get('search-site-files')!({
        siteId: 'site-1',
        query: 'plan',
        path: '/Shared Documents/Plans',
        limit: 2,
        cursor: first.nextCursor,
      })
    ) as {
      items: Array<{ id: string }>;
    };

    expect(second.items[0].id).toBe('item-2');
  });

  it('reads SharePoint-backed file text and context through path-first site resolution', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const graphClient = createMockGraphClient(async (endpoint, options) => {
      if (
        endpoint ===
        '/sites/site-1/drive/root:/Shared Documents/Plans/plan.md:?$select=id,name,size,lastModifiedDateTime,webUrl,parentReference,folder,file,lastModifiedBy'
      ) {
        return {
          id: 'plan-1',
          name: 'plan.md',
          size: 42,
          webUrl: 'https://contoso.sharepoint.com/sites/marketing/Shared%20Documents/Plans/plan.md',
          parentReference: {
            driveId: 'drive-1',
            path: '/drives/drive-1/root:/Shared Documents/Plans',
            siteId: 'site-1',
          },
          file: { mimeType: 'text/markdown' },
        };
      }
      if (endpoint === '/drives/drive-1/items/plan-1/content') {
        expect(options?.responseType).toBe('buffer');
        return Buffer.from('# Launch Plan\nShip it.');
      }
      throw new Error(`Unexpected endpoint: ${endpoint}`);
    });

    registerGraphTools(server, graphClient, false, 'get-file-text|get-file-context', false);

    const textResult = parseResult(
      await handlers.get('get-file-text')!({
        siteId: 'site-1',
        path: '/Shared Documents/Plans/plan.md',
      })
    ) as { text: string; label: string };

    const contextResult = parseResult(
      await handlers.get('get-file-context')!({
        siteId: 'site-1',
        path: '/Shared Documents/Plans/plan.md',
      })
    ) as { file: { path: string }; previewText: string; extractionStatus: string };

    expect(textResult).toMatchObject({
      text: '# Launch Plan\nShip it.',
      label: 'plan.md (/Shared Documents/Plans)',
    });
    expect(contextResult).toMatchObject({
      file: { path: '/Shared Documents/Plans/plan.md' },
      extractionStatus: 'ok',
    });
    expect(contextResult.previewText).toContain('Launch Plan');
  });

  it('lists SharePoint pages, lists, and list items with labels and pagination', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const graphClient = createMockGraphClient(async (endpoint) => {
      if (
        endpoint ===
        '/sites/site-1/pages/microsoft.graph.sitePage?$top=2&$select=id,name,title,webUrl,lastModifiedDateTime'
      ) {
        return {
          value: [
            {
              id: 'page-1',
              name: 'Launch.aspx',
              title: 'Launch Home',
              webUrl: 'https://contoso.sharepoint.com/sites/marketing/SitePages/Launch.aspx',
              lastModifiedDateTime: '2026-03-12T10:00:00Z',
            },
          ],
          '@odata.nextLink':
            'https://graph.microsoft.com/v1.0/sites/site-1/pages/microsoft.graph.sitePage?$skiptoken=next',
        };
      }
      if (endpoint === '/sites/site-1/lists?$top=2&$select=id,name,displayName,webUrl') {
        return {
          value: [
            {
              id: 'list-1',
              name: 'launchTasks',
              displayName: 'Launch Tasks',
              webUrl: 'https://contoso.sharepoint.com/sites/marketing/Lists/LaunchTasks',
            },
          ],
        };
      }
      if (endpoint === '/sites/site-1/lists/list-1/items?$expand=fields&$top=2') {
        return {
          value: [
            {
              id: 'item-1',
              webUrl:
                'https://contoso.sharepoint.com/sites/marketing/Lists/LaunchTasks/DispForm.aspx?ID=1',
              fields: {
                Title: 'Confirm launch email',
                Status: 'Not Started',
              },
            },
          ],
        };
      }
      throw new Error(`Unexpected endpoint: ${endpoint}`);
    });

    registerGraphTools(
      server,
      graphClient,
      false,
      'list-sharepoint-pages|list-sharepoint-lists|list-sharepoint-list-items',
      true
    );

    const pages = parseResult(
      await handlers.get('list-sharepoint-pages')!({ siteId: 'site-1', limit: 2 })
    ) as {
      items: Array<{ pageId: string; label: string }>;
      nextCursor?: string;
    };
    const lists = parseResult(
      await handlers.get('list-sharepoint-lists')!({ siteId: 'site-1', limit: 2 })
    ) as {
      items: Array<{ listId: string; label: string }>;
    };
    const items = parseResult(
      await handlers.get('list-sharepoint-list-items')!({
        siteId: 'site-1',
        listId: 'list-1',
        limit: 2,
      })
    ) as {
      items: Array<{ itemId: string; label: string; fields: { Status: string } }>;
    };

    expect(pages.items[0]).toMatchObject({ pageId: 'page-1', label: 'Launch Home' });
    expect(pages.nextCursor).toBeTruthy();
    expect(lists.items[0]).toMatchObject({ listId: 'list-1', label: 'Launch Tasks' });
    expect(items.items[0]).toMatchObject({
      itemId: 'item-1',
      label: 'Confirm launch email',
      fields: { Status: 'Not Started' },
    });
  });

  it('supports a SharePoint discovery flow from site to library to file context', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const graphClient = createMockGraphClient(async (endpoint, options) => {
      if (endpoint === '/sites?$search=%22Marketing%22&$top=5&$select=id,name,displayName,webUrl') {
        return {
          value: [
            {
              id: 'site-1',
              name: 'marketing',
              displayName: 'Marketing Team',
              webUrl: 'https://contoso.sharepoint.com/sites/marketing',
            },
          ],
        };
      }
      if (endpoint === '/sites/site-1/drives?$top=5&$select=id,name,driveType,webUrl') {
        return {
          value: [
            {
              id: 'drive-1',
              name: 'Shared Documents',
              driveType: 'documentLibrary',
              webUrl: 'https://contoso.sharepoint.com/sites/marketing/Shared%20Documents',
            },
          ],
        };
      }
      if (
        endpoint ===
        '/drives/drive-1/root:/Plans:?$select=id,name,size,lastModifiedDateTime,webUrl,parentReference,folder,file,lastModifiedBy'
      ) {
        return {
          id: 'folder-1',
          name: 'Plans',
          parentReference: { driveId: 'drive-1', path: '/drives/drive-1/root:' },
          folder: {},
          webUrl: 'https://contoso.sharepoint.com/sites/marketing/Shared%20Documents/Plans',
        };
      }
      if (
        endpoint ===
        '/drives/drive-1/items/folder-1/children?$top=5&$select=id,name,size,lastModifiedDateTime,webUrl,parentReference,folder,file,lastModifiedBy'
      ) {
        return {
          value: [
            {
              id: 'plan-1',
              name: 'launch.md',
              parentReference: { driveId: 'drive-1', path: '/drives/drive-1/root:/Plans' },
              file: { mimeType: 'text/markdown' },
              webUrl:
                'https://contoso.sharepoint.com/sites/marketing/Shared%20Documents/Plans/launch.md',
            },
          ],
        };
      }
      if (
        endpoint ===
        '/drives/drive-1/root:/Plans/launch.md:?$select=id,name,size,lastModifiedDateTime,webUrl,parentReference,folder,file,lastModifiedBy'
      ) {
        return {
          id: 'plan-1',
          name: 'launch.md',
          parentReference: { driveId: 'drive-1', path: '/drives/drive-1/root:/Plans' },
          file: { mimeType: 'text/markdown' },
          webUrl:
            'https://contoso.sharepoint.com/sites/marketing/Shared%20Documents/Plans/launch.md',
        };
      }
      if (endpoint === '/drives/drive-1/items/plan-1/content') {
        expect(options?.responseType).toBe('buffer');
        return Buffer.from('Launch checklist');
      }
      throw new Error(`Unexpected endpoint: ${endpoint}`);
    });

    registerGraphTools(
      server,
      graphClient,
      false,
      'search-sites|list-document-libraries|list-folder|get-file-context',
      true
    );

    const sites = parseResult(
      await handlers.get('search-sites')!({ query: 'Marketing', limit: 5 })
    ) as {
      items: Array<{ siteId: string }>;
    };
    const libraries = parseResult(
      await handlers.get('list-document-libraries')!({ siteId: sites.items[0].siteId, limit: 5 })
    ) as {
      items: Array<{ driveId: string }>;
    };
    const folder = parseResult(
      await handlers.get('list-folder')!({
        driveId: libraries.items[0].driveId,
        path: '/Plans',
        limit: 5,
      })
    ) as {
      items: Array<{ path: string }>;
    };
    const context = parseResult(
      await handlers.get('get-file-context')!({
        driveId: libraries.items[0].driveId,
        path: folder.items[0].path,
      })
    ) as {
      file: { path: string };
      text: string;
    };

    expect(context).toMatchObject({
      file: { path: '/Plans/launch.md' },
      text: 'Launch checklist',
    });
  });

  it('search-calendar-events filters calendarView results by query text', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const graphClient = createMockGraphClient(async (endpoint) => {
      expect(endpoint).toContain('/me/calendarView?');
      return {
        value: [
          {
            id: 'event-1',
            subject: 'Project kickoff',
            start: { dateTime: '2026-03-12T14:00:00Z' },
            end: { dateTime: '2026-03-12T15:00:00Z' },
            organizer: {
              emailAddress: { name: 'Juan Casiraghi', address: 'juan@company.com' },
            },
          },
          {
            id: 'event-2',
            subject: 'Random sync',
            start: { dateTime: '2026-03-12T16:00:00Z' },
            end: { dateTime: '2026-03-12T16:30:00Z' },
          },
        ],
      };
    });

    registerGraphTools(server, graphClient, false, 'search-calendar-events', false);

    const result = await handlers.get('search-calendar-events')!({
      query: 'kickoff',
      start: '2026-03-01T00:00:00Z',
      end: '2026-03-31T23:59:59Z',
      limit: 10,
    });
    const parsed = parseResult(result) as { items: Array<{ id: string; subject: string }> };

    expect(parsed.items).toHaveLength(1);
    expect(parsed.items[0]).toMatchObject({ id: 'event-1', subject: 'Project kickoff' });
  });

  it('find-availability allows external email attendees without directory resolution', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const graphClient = createMockGraphClient(async (endpoint, options) => {
      if (endpoint.startsWith('/users?')) {
        throw new Error('Microsoft Graph API error: 404 Not Found');
      }

      if (endpoint === '/me/findMeetingTimes') {
        const body = JSON.parse((options?.body as string) ?? '{}') as {
          attendees: Array<{ emailAddress: { address: string } }>;
        };
        expect(body.attendees[0].emailAddress.address).toBe('external@example.com');
        return {
          meetingTimeSuggestions: [
            {
              meetingTimeSlot: {
                start: { dateTime: '2026-03-12T14:00:00Z' },
                end: { dateTime: '2026-03-12T14:30:00Z' },
              },
              confidence: 0.92,
            },
          ],
        };
      }

      throw new Error(`Unexpected endpoint: ${endpoint}`);
    });

    registerGraphTools(server, graphClient, false, 'find-availability', false);

    const result = await handlers.get('find-availability')!({
      participants: ['external@example.com'],
      start: '2026-03-12T09:00:00Z',
      end: '2026-03-12T18:00:00Z',
      durationMinutes: 30,
    });
    const parsed = parseResult(result) as {
      slots: Array<{ start: string; score: number }>;
    };

    expect(parsed.slots[0]).toMatchObject({
      start: '2026-03-12T14:00:00Z',
      score: 0.92,
    });
  });

  it('responds to calendar events through explicit RSVP tools', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const graphClient = createMockGraphClient(async (endpoint, options) => {
      const body = JSON.parse((options?.body as string) ?? '{}') as {
        sendResponse: boolean;
        comment?: string;
      };

      if (endpoint === '/me/events/event-1/accept') {
        expect(body).toMatchObject({ sendResponse: true, comment: 'See you there' });
        return {};
      }

      if (endpoint === '/me/events/event-1/decline') {
        expect(body).toMatchObject({ sendResponse: false });
        return {};
      }

      if (endpoint === '/me/events/event-1/tentativelyAccept') {
        expect(body).toMatchObject({ sendResponse: true });
        return {};
      }

      throw new Error(`Unexpected endpoint: ${endpoint}`);
    });

    registerGraphTools(
      server,
      graphClient,
      false,
      'accept-calendar-event|decline-calendar-event|tentatively-accept-calendar-event',
      false
    );

    const acceptResult = await handlers.get('accept-calendar-event')!({
      eventId: 'event-1',
      sendResponse: true,
      comment: 'See you there',
    });
    const declineResult = await handlers.get('decline-calendar-event')!({
      eventId: 'event-1',
      sendResponse: false,
    });
    const tentativeResult = await handlers.get('tentatively-accept-calendar-event')!({
      eventId: 'event-1',
    });

    expect(parseResult(acceptResult)).toMatchObject({
      eventId: 'event-1',
      response: 'accepted',
      sendResponse: true,
      comment: 'See you there',
    });
    expect(parseResult(declineResult)).toMatchObject({
      eventId: 'event-1',
      response: 'declined',
      sendResponse: false,
    });
    expect(parseResult(tentativeResult)).toMatchObject({
      eventId: 'event-1',
      response: 'tentativelyAccepted',
      sendResponse: true,
    });
  });

  it('create-calendar-event-from-availability chooses a fully-available slot before creating', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const graphClient = createMockGraphClient(async (endpoint, options) => {
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

      if (endpoint === '/me/findMeetingTimes') {
        return {
          meetingTimeSuggestions: [
            {
              meetingTimeSlot: {
                start: { dateTime: '2026-03-12T15:00:00Z' },
                end: { dateTime: '2026-03-12T15:30:00Z' },
              },
              confidence: 0.99,
              attendeeAvailability: [
                {
                  availability: 'busy',
                  attendee: {
                    emailAddress: { name: 'Juan Casiraghi', address: 'juan@company.com' },
                  },
                },
              ],
            },
            {
              meetingTimeSlot: {
                start: { dateTime: '2026-03-12T14:00:00Z' },
                end: { dateTime: '2026-03-12T14:30:00Z' },
              },
              confidence: 0.85,
              attendeeAvailability: [
                {
                  availability: 'free',
                  attendee: {
                    emailAddress: { name: 'Juan Casiraghi', address: 'juan@company.com' },
                  },
                },
              ],
            },
          ],
        };
      }

      if (endpoint === '/me/events') {
        const body = JSON.parse((options?.body as string) ?? '{}') as {
          subject: string;
          start: { dateTime: string };
          end: { dateTime: string };
          attendees: Array<{ emailAddress: { address: string } }>;
        };
        expect(body.subject).toBe('Planning sync');
        expect(body.start.dateTime).toBe('2026-03-12T14:00:00Z');
        expect(body.end.dateTime).toBe('2026-03-12T14:30:00Z');
        expect(body.attendees[0].emailAddress.address).toBe('juan@company.com');
        return {
          id: 'event-1',
          subject: body.subject,
          start: body.start,
          end: body.end,
          attendees: [
            {
              emailAddress: { name: 'Juan Casiraghi', address: 'juan@company.com' },
            },
          ],
        };
      }

      throw new Error(`Unexpected endpoint: ${endpoint}`);
    });

    registerGraphTools(
      server,
      graphClient,
      false,
      'create-calendar-event-from-availability',
      false
    );

    const result = await handlers.get('create-calendar-event-from-availability')!({
      subject: 'Planning sync',
      attendees: ['juan@company.com'],
      windowStart: '2026-03-12T09:00:00Z',
      windowEnd: '2026-03-12T18:00:00Z',
      durationMinutes: 30,
      isOnlineMeeting: true,
    });
    const parsed = parseResult(result) as {
      event: { id: string; start: string; end: string };
      chosenSlot: { start: string; allAttendeesAvailable: boolean };
    };

    expect(parsed.event).toMatchObject({
      id: 'event-1',
      start: '2026-03-12T14:00:00Z',
      end: '2026-03-12T14:30:00Z',
    });
    expect(parsed.chosenSlot).toMatchObject({
      start: '2026-03-12T14:00:00Z',
      allAttendeesAvailable: true,
    });
  });

  it('create-calendar-event-from-availability returns a structured error when no fully-free slot exists', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const graphClient = createMockGraphClient(async (endpoint) => {
      if (endpoint === '/me/findMeetingTimes') {
        return {
          meetingTimeSuggestions: [
            {
              meetingTimeSlot: {
                start: { dateTime: '2026-03-12T15:00:00Z' },
                end: { dateTime: '2026-03-12T15:30:00Z' },
              },
              confidence: 0.95,
              attendeeAvailability: [
                {
                  availability: 'busy',
                  attendee: {
                    emailAddress: { name: 'External', address: 'external@example.com' },
                  },
                },
              ],
            },
          ],
        };
      }

      throw new Error(`Unexpected endpoint: ${endpoint}`);
    });

    registerGraphTools(
      server,
      graphClient,
      false,
      'create-calendar-event-from-availability',
      false
    );

    const result = await handlers.get('create-calendar-event-from-availability')!({
      subject: 'Blocked sync',
      attendees: ['external@example.com'],
      windowStart: '2026-03-12T09:00:00Z',
      windowEnd: '2026-03-12T18:00:00Z',
      durationMinutes: 30,
    });
    const parsed = parseResult(result) as {
      error: string;
      suggestions: Array<{ blockedAttendeeCount: number }>;
    };

    expect(result.isError).toBe(true);
    expect(parsed.error).toBe('not_found');
    expect(parsed.suggestions[0].blockedAttendeeCount).toBe(1);
  });

  it('create-calendar-event-from-availability preserves ambiguous attendee candidates', async () => {
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

    registerGraphTools(server, graphClient, false, 'create-calendar-event-from-availability', true);

    const result = await handlers.get('create-calendar-event-from-availability')!({
      subject: 'Planning sync',
      attendees: ['Juan Casiraghi'],
      windowStart: '2026-03-12T09:00:00Z',
      windowEnd: '2026-03-12T18:00:00Z',
      durationMinutes: 30,
    });
    const parsed = parseResult(result) as {
      error: string;
      candidates: Array<{ rank: number; label: string }>;
    };

    expect(result.isError).toBe(true);
    expect(parsed.error).toBe('ambiguous_match');
    expect(parsed.candidates[0].rank).toBe(1);
    expect(parsed.candidates[0].label).toContain('Juan Casiraghi');
  });

  it('search-tasks applies text and status filters across task lists', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const graphClient = createMockGraphClient(async (endpoint) => {
      if (endpoint.startsWith('/me/todo/lists?')) {
        return {
          value: [{ id: 'list-1', displayName: 'Tasks', wellknownListName: 'defaultList' }],
        };
      }

      if (endpoint === '/me/todo/lists/list-1/tasks?$top=200&$select=id,status') {
        return {
          value: [
            { id: 'task-1', status: 'notStarted' },
            { id: 'task-2', status: 'completed' },
          ],
        };
      }

      if (
        endpoint ===
        '/me/todo/lists/list-1/tasks?$top=100&$select=id,title,status,importance,dueDateTime,body'
      ) {
        return {
          value: [
            {
              id: 'task-1',
              title: 'Deploy prod',
              status: 'notStarted',
              importance: 'high',
              body: { content: 'Tonight after QA signoff' },
            },
            {
              id: 'task-2',
              title: 'Write notes',
              status: 'completed',
              importance: 'normal',
              body: { content: 'Postmortem' },
            },
          ],
        };
      }

      throw new Error(`Unexpected endpoint: ${endpoint}`);
    });

    registerGraphTools(server, graphClient, false, 'search-tasks', false);

    const result = await handlers.get('search-tasks')!({
      query: 'deploy',
      status: 'notStarted',
      limit: 10,
    });
    const parsed = parseResult(result) as { items: Array<{ taskId: string; title: string }> };

    expect(parsed.items).toHaveLength(1);
    expect(parsed.items[0]).toMatchObject({ taskId: 'task-1', title: 'Deploy prod' });
  });

  it('list-tasks returns aggregate nextCursor when paging across all task lists', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const graphClient = createMockGraphClient(async (endpoint) => {
      if (endpoint.startsWith('/me/todo/lists?')) {
        return {
          value: [
            { id: 'list-1', displayName: 'Tasks', wellknownListName: 'defaultList' },
            { id: 'list-2', displayName: 'Backlog', wellknownListName: 'none' },
          ],
        };
      }

      if (endpoint === '/me/todo/lists/list-1/tasks?$top=200&$select=id,status') {
        return { value: [{ id: 'task-1', status: 'notStarted' }] };
      }

      if (endpoint === '/me/todo/lists/list-2/tasks?$top=200&$select=id,status') {
        return { value: [{ id: 'task-2', status: 'notStarted' }] };
      }

      if (
        endpoint ===
        '/me/todo/lists/list-1/tasks?$top=100&$select=id,title,status,importance,dueDateTime,body'
      ) {
        return {
          value: [
            {
              id: 'task-1',
              title: 'First',
              status: 'notStarted',
              dueDateTime: { dateTime: '2026-03-13T10:00:00Z' },
            },
          ],
        };
      }

      if (
        endpoint ===
        '/me/todo/lists/list-2/tasks?$top=100&$select=id,title,status,importance,dueDateTime,body'
      ) {
        return {
          value: [
            {
              id: 'task-2',
              title: 'Second',
              status: 'notStarted',
              dueDateTime: { dateTime: '2026-03-14T10:00:00Z' },
            },
          ],
        };
      }

      throw new Error(`Unexpected endpoint: ${endpoint}`);
    });

    registerGraphTools(server, graphClient, false, 'list-tasks', false);

    const firstResult = await handlers.get('list-tasks')!({ limit: 1 });
    const firstParsed = parseResult(firstResult) as {
      items: Array<{ taskId: string }>;
      nextCursor?: string;
    };

    expect(firstParsed.items[0].taskId).toBe('task-1');
    expect(firstParsed.nextCursor).toBeTruthy();

    const secondResult = await handlers.get('list-tasks')!({
      limit: 1,
      cursor: firstParsed.nextCursor,
    });
    const secondParsed = parseResult(secondResult) as { items: Array<{ taskId: string }> };
    expect(secondParsed.items[0].taskId).toBe('task-2');
  });

  it('search-mail scopes to the requested folder and returns normalized messages', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const graphClient = createMockGraphClient(async (endpoint, options) => {
      expect(endpoint).toContain('/me/mailFolders/inbox/messages?');
      expect((options?.headers as Record<string, string>).ConsistencyLevel).toBe('eventual');
      return {
        value: [
          {
            id: 'msg-1',
            subject: 'Invoice',
            from: {
              emailAddress: { name: 'Mauro Benitez', address: 'mauro@company.com' },
            },
            toRecipients: [
              {
                emailAddress: { name: 'Santiago', address: 'santiago@company.com' },
              },
            ],
            receivedDateTime: '2026-03-12T10:00:00Z',
            bodyPreview: 'Invoice attached',
            hasAttachments: true,
          },
        ],
      };
    });

    registerGraphTools(server, graphClient, false, 'search-mail', false);

    const result = await handlers.get('search-mail')!({
      query: 'invoice Mauro',
      folder: 'Inbox',
      limit: 10,
    });
    const parsed = parseResult(result) as {
      items: Array<{ messageId: string; subject: string; hasAttachments: boolean }>;
    };

    expect(parsed.items[0]).toMatchObject({
      messageId: 'msg-1',
      subject: 'Invoice',
      hasAttachments: true,
    });
  });

  it('create-text-file preserves non-not-found lookup failures instead of attempting a write', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const graphClient = createMockGraphClient(async (endpoint) => {
      if (endpoint.includes('/me/drive/root:/blocked.txt:')) {
        throw new Error('Microsoft Graph API scope error: 403 Forbidden - scope error');
      }

      if (endpoint.endsWith('/content')) {
        throw new Error('Should not attempt write after lookup scope failure');
      }

      throw new Error(`Unexpected endpoint: ${endpoint}`);
    });

    registerGraphTools(server, graphClient, false, 'create-text-file', false);

    const result = await handlers.get('create-text-file')!({
      path: '/blocked.txt',
      content: 'hello',
    });
    const parsed = parseResult(result) as { error: string };

    expect(result.isError).toBe(true);
    expect(parsed.error).toBe('insufficient_scope');
  });

  it('search-sharepoint-content filters by site webUrl instead of opaque siteId string matching', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const graphClient = createMockGraphClient(async (endpoint) => {
      if (endpoint === '/sites/site-1?$select=webUrl') {
        return { webUrl: 'https://contoso.sharepoint.com/sites/marketing' };
      }

      if (endpoint === '/search/query') {
        return {
          value: [
            {
              hitsContainers: [
                {
                  hits: [
                    {
                      resource: {
                        id: 'doc-1',
                        title: 'Brand deck',
                        webUrl:
                          'https://contoso.sharepoint.com/sites/marketing/Shared%20Documents/brand.pptx',
                      },
                    },
                    {
                      resource: {
                        id: 'doc-2',
                        title: 'Wrong site',
                        webUrl:
                          'https://contoso.sharepoint.com/sites/sales/Shared%20Documents/qbr.pptx',
                      },
                    },
                  ],
                },
              ],
            },
          ],
        };
      }

      throw new Error(`Unexpected endpoint: ${endpoint}`);
    });

    registerGraphTools(server, graphClient, false, 'search-sharepoint-content', true);

    const result = await handlers.get('search-sharepoint-content')!({
      query: 'brand',
      siteId: 'site-1',
      limit: 10,
    });
    const parsed = parseResult(result) as { items: Array<{ id: string }> };

    expect(parsed.items).toHaveLength(1);
    expect(parsed.items[0].id).toBe('doc-1');
  });
});
