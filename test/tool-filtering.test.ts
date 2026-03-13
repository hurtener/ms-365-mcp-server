import { beforeEach, describe, expect, it, vi } from 'vitest';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { registerGraphTools } from '../src/graph-tools.js';
import GraphClient from '../src/graph-client.js';

vi.mock('../src/logger.js', () => ({
  default: {
    info: vi.fn(),
    error: vi.fn(),
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
      },
      { alias: 'send-mail', method: 'POST', path: '/me/sendMail', description: 'Send mail' },
      {
        alias: 'list-calendar-events',
        method: 'GET',
        path: '/me/events',
        description: 'List calendar events',
      },
      {
        alias: 'list-excel-worksheets',
        method: 'GET',
        path: '/workbook/worksheets',
        description: 'List Excel worksheets',
      },
      { alias: 'get-current-user', method: 'GET', path: '/me', description: 'Get current user' },
    ],
  },
}));

describe('Tool Filtering', () => {
  let server: McpServer;
  let graphClient: GraphClient;
  let toolSpy: ReturnType<typeof vi.spyOn>;

  beforeEach(() => {
    server = new McpServer({ name: 'test', version: '1.0.0' });
    graphClient = {} as GraphClient;
    toolSpy = vi.spyOn(server, 'tool').mockImplementation(() => {});
  });

  it('should register all tools when no filter is provided', () => {
    registerGraphTools(server, graphClient, false);

    const toolNames = toolSpy.mock.calls.map((call: unknown[]) => call[0] as string);
    expect(toolNames.length).toBeGreaterThanOrEqual(5);
    expect(toolSpy).toHaveBeenCalledWith(
      'list-mail-messages',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
    expect(toolSpy).toHaveBeenCalledWith(
      'send-mail',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
    expect(toolSpy).toHaveBeenCalledWith(
      'list-calendar-events',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
    expect(toolSpy).toHaveBeenCalledWith(
      'list-excel-worksheets',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
    expect(toolSpy).toHaveBeenCalledWith(
      'get-current-user',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
    expect(toolNames).toContain('search-mail');
  });

  it('should filter tools by regex pattern - mail only', () => {
    registerGraphTools(server, graphClient, false, 'mail');

    const toolNames = toolSpy.mock.calls.map((call: unknown[]) => call[0] as string);
    expect(toolSpy).toHaveBeenCalledWith(
      'list-mail-messages',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
    expect(toolSpy).toHaveBeenCalledWith(
      'send-mail',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
    expect(toolNames).toContain('search-mail');
    expect(toolNames.every((name) => /mail/i.test(name))).toBe(true);
  });

  it('should filter tools by regex pattern - calendar or excel', () => {
    registerGraphTools(server, graphClient, false, 'calendar|excel');

    const toolNames = toolSpy.mock.calls.map((call: unknown[]) => call[0] as string);
    expect(toolSpy).toHaveBeenCalledWith(
      'list-calendar-events',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
    expect(toolSpy).toHaveBeenCalledWith(
      'list-excel-worksheets',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
    expect(toolNames.every((name) => /calendar|excel/i.test(name))).toBe(true);
  });

  it('should handle invalid regex patterns gracefully', () => {
    registerGraphTools(server, graphClient, false, '[invalid regex');

    const toolNames = toolSpy.mock.calls.map((call: unknown[]) => call[0] as string);
    expect(toolNames.length).toBeGreaterThanOrEqual(5);
    expect(toolNames).toContain('search-mail');
  });

  it('should combine read-only and filtering correctly', () => {
    registerGraphTools(server, graphClient, true, 'mail');

    const toolNames = toolSpy.mock.calls.map((call: unknown[]) => call[0] as string);
    expect(toolSpy).toHaveBeenCalledWith(
      'list-mail-messages',
      expect.any(String),
      expect.any(Object),
      expect.any(Object),
      expect.any(Function)
    );
    expect(toolNames).not.toContain('send-mail');
    expect(toolNames.every((name) => /mail/i.test(name))).toBe(true);
  });

  it('should register no tools when pattern matches nothing', () => {
    registerGraphTools(server, graphClient, false, 'nonexistent');

    expect(toolSpy).toHaveBeenCalledTimes(0);
  });
});
