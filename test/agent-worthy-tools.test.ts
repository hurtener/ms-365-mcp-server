import { afterEach, describe, expect, it, vi } from 'vitest';
import { strToU8, zipSync } from 'fflate';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';

import { registerGraphTools } from '../src/graph-tools.js';
import GraphClient from '../src/graph-client.js';

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

function createJwtWithScopes(scopes: string[]): string {
  const header = Buffer.from(JSON.stringify({ alg: 'none', typ: 'JWT' })).toString('base64url');
  const payload = Buffer.from(JSON.stringify({ scp: scopes.join(' ') })).toString('base64url');
  return `${header}.${payload}.signature`;
}

function createSimplePdf(text: string): Buffer {
  const objects = [
    '1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n',
    '2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n',
    '3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 144] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>\nendobj\n',
    `4 0 obj\n<< /Length ${text.length + 31} >>\nstream\nBT\n/F1 24 Tf\n72 72 Td\n(${text}) Tj\nET\nendstream\nendobj\n`,
    '5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj\n',
  ];

  let pdf = '%PDF-1.4\n';
  const offsets = [0];
  for (const object of objects) {
    offsets.push(Buffer.byteLength(pdf, 'utf8'));
    pdf += object;
  }
  const xrefOffset = Buffer.byteLength(pdf, 'utf8');
  pdf += `xref\n0 ${objects.length + 1}\n`;
  pdf += '0000000000 65535 f \n';
  for (let index = 1; index < offsets.length; index += 1) {
    pdf += `${String(offsets[index]).padStart(10, '0')} 00000 n \n`;
  }
  pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF`;
  return Buffer.from(pdf, 'utf8');
}

describe('agent-worthy tool surfaces', () => {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  it('extracts text from Office formats and marks unsupported binaries', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);

    const docxBuffer = Buffer.from(
      zipSync({
        'word/document.xml': strToU8(
          '<w:document><w:body><w:p><w:r><w:t>Hello DOCX</w:t></w:r></w:p></w:body></w:document>'
        ),
      })
    );
    const pptxBuffer = Buffer.from(
      zipSync({
        'ppt/slides/slide1.xml': strToU8('<p:sld><a:t>Hello PPTX</a:t></p:sld>'),
      })
    );
    const xlsxBuffer = Buffer.from(
      zipSync({
        'xl/sharedStrings.xml': strToU8('<sst><si><t>Hello XLSX</t></si></sst>'),
        'xl/worksheets/sheet1.xml': strToU8(
          '<worksheet><sheetData><row><c t="s"><v>0</v></c></row></sheetData></worksheet>'
        ),
      })
    );
    const graphClient = createMockGraphClient(async (endpoint, options) => {
      if (endpoint.startsWith('/drives/drive-1/root:/Docs/hello.docx:')) {
        return {
          id: 'docx-1',
          name: 'hello.docx',
          parentReference: { driveId: 'drive-1', path: '/drives/drive-1/root:/Docs' },
          file: {
            mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
          },
        };
      }
      if (endpoint.startsWith('/drives/drive-1/root:/Docs/slides.pptx:')) {
        return {
          id: 'pptx-1',
          name: 'slides.pptx',
          parentReference: { driveId: 'drive-1', path: '/drives/drive-1/root:/Docs' },
          file: {
            mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
          },
        };
      }
      if (endpoint.startsWith('/drives/drive-1/root:/Docs/sheet.xlsx:')) {
        return {
          id: 'xlsx-1',
          name: 'sheet.xlsx',
          parentReference: { driveId: 'drive-1', path: '/drives/drive-1/root:/Docs' },
          file: {
            mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          },
        };
      }
      if (endpoint.startsWith('/drives/drive-1/root:/Docs/file.bin:')) {
        return {
          id: 'bin-1',
          name: 'file.bin',
          parentReference: { driveId: 'drive-1', path: '/drives/drive-1/root:/Docs' },
          file: { mimeType: 'application/octet-stream' },
        };
      }
      if (endpoint === '/drives/drive-1/items/docx-1/content') {
        expect(options?.responseType).toBe('buffer');
        return docxBuffer;
      }
      if (endpoint === '/drives/drive-1/items/pptx-1/content') {
        return pptxBuffer;
      }
      if (endpoint === '/drives/drive-1/items/xlsx-1/content') {
        return xlsxBuffer;
      }
      if (endpoint === '/drives/drive-1/items/bin-1/content') {
        return Buffer.from([0xde, 0xad, 0xbe, 0xef]);
      }
      throw new Error(`Unexpected endpoint: ${endpoint}`);
    });

    registerGraphTools(server, graphClient, false, 'get-file-text|get-file-context', false);

    const docxResult = parseResult(
      await handlers.get('get-file-text')!({ path: '/Docs/hello.docx', driveId: 'drive-1' })
    ) as { text: string; label: string };
    const pptxResult = parseResult(
      await handlers.get('get-file-text')!({ path: '/Docs/slides.pptx', driveId: 'drive-1' })
    ) as { text: string };
    const xlsxResult = parseResult(
      await handlers.get('get-file-text')!({ path: '/Docs/sheet.xlsx', driveId: 'drive-1' })
    ) as { text: string };
    const unsupportedResult = parseResult(
      await handlers.get('get-file-context')!({ path: '/Docs/file.bin', driveId: 'drive-1' })
    ) as { extractionStatus: string; previewText: string | null };

    expect(docxResult.text).toContain('Hello DOCX');
    expect(docxResult.label).toContain('hello.docx');
    expect(pptxResult.text).toContain('Hello PPTX');
    expect(xlsxResult.text).toContain('Hello XLSX');
    expect(unsupportedResult.extractionStatus).toBe('unsupported');
    expect(unsupportedResult.previewText).toBeNull();
  });

  it('extracts PDF text and supports preview and execute for edit-text-file', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const graphClient = createMockGraphClient(async (endpoint, options) => {
      if (endpoint.startsWith('/drives/drive-1/root:/Docs/report.pdf:')) {
        return {
          id: 'pdf-1',
          name: 'report.pdf',
          parentReference: { driveId: 'drive-1', path: '/drives/drive-1/root:/Docs' },
          file: { mimeType: 'application/pdf' },
        };
      }
      if (endpoint.startsWith('/drives/drive-1/root:/Docs/readme.md:')) {
        return {
          id: 'md-1',
          name: 'readme.md',
          parentReference: { driveId: 'drive-1', path: '/drives/drive-1/root:/Docs' },
          file: { mimeType: 'text/markdown' },
        };
      }
      if (endpoint === '/drives/drive-1/items/pdf-1/content') {
        return createSimplePdf('PDF extracted text');
      }
      if (endpoint === '/drives/drive-1/items/md-1/content') {
        return Buffer.from('hello world');
      }
      if (endpoint === '/drives/drive-1/root:/Docs/readme.md:/content') {
        expect(options?.body).toBe('hello codex');
        return {
          id: 'md-1',
          name: 'readme.md',
          parentReference: { driveId: 'drive-1', path: '/drives/drive-1/root:/Docs' },
          file: { mimeType: 'text/markdown' },
        };
      }
      throw new Error(`Unexpected endpoint: ${endpoint}`);
    });

    registerGraphTools(server, graphClient, false, 'get-file-text|edit-text-file', false);

    const pdfResult = parseResult(
      await handlers.get('get-file-text')!({ path: '/Docs/report.pdf', driveId: 'drive-1' })
    ) as { text: string };
    const previewResult = parseResult(
      await handlers.get('edit-text-file')!({
        path: '/Docs/readme.md',
        driveId: 'drive-1',
        operations: [{ type: 'replaceAll', find: 'world', replace: 'codex' }],
        previewOnly: true,
      })
    ) as { previewOnly: boolean; previewText: string };
    const executeResult = parseResult(
      await handlers.get('edit-text-file')!({
        path: '/Docs/readme.md',
        driveId: 'drive-1',
        operations: [{ type: 'replaceAll', find: 'world', replace: 'codex' }],
      })
    ) as { file: { label: string } };

    expect(pdfResult.text).toContain('PDF extracted text');
    expect(previewResult.previewOnly).toBe(true);
    expect(previewResult.previewText).toContain('hello codex');
    expect(executeResult.file.label).toContain('readme.md');
  });

  it('supports mail action wrappers and mail context', async () => {
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
      if (endpoint.startsWith('/me/messages/msg-1?')) {
        return {
          id: 'msg-1',
          subject: 'Quarterly update',
          from: { emailAddress: { name: 'Mauro', address: 'mauro@company.com' } },
          toRecipients: [{ emailAddress: { name: 'Santiago', address: 'santiago@company.com' } }],
          receivedDateTime: '2026-03-12T10:00:00Z',
          conversationId: 'conv-1',
          hasAttachments: true,
        };
      }
      if (endpoint.startsWith('/me/messages?$filter=')) {
        return {
          value: [
            {
              id: 'msg-1',
              subject: 'Quarterly update',
              from: { emailAddress: { name: 'Mauro', address: 'mauro@company.com' } },
              toRecipients: [
                { emailAddress: { name: 'Santiago', address: 'santiago@company.com' } },
              ],
              receivedDateTime: '2026-03-12T10:00:00Z',
            },
          ],
        };
      }
      if (endpoint === '/me/messages/msg-1/attachments?$top=100') {
        return {
          value: [{ id: 'att-1', name: 'report.pdf', contentType: 'application/pdf', size: 123 }],
        };
      }
      if (endpoint === '/me/messages/msg-1/reply') {
        expect(options?.method).toBe('POST');
        return {};
      }
      if (endpoint === '/me/messages/msg-1/createForward') {
        return {
          id: 'draft-1',
          subject: 'FW: Quarterly update',
          from: { emailAddress: { name: 'Santiago', address: 'santiago@company.com' } },
          toRecipients: [{ emailAddress: { name: 'Juan', address: 'juan@company.com' } }],
          receivedDateTime: '2026-03-12T10:05:00Z',
          isDraft: true,
        };
      }
      throw new Error(`Unexpected endpoint: ${endpoint}`);
    });

    registerGraphTools(
      server,
      graphClient,
      false,
      'reply-mail|draft-mail-forward|get-mail-context',
      true
    );

    const previewReply = parseResult(
      await handlers.get('reply-mail')!({
        messageId: 'msg-1',
        comment: 'Looks good',
        previewOnly: true,
      })
    ) as { wouldSend: boolean; subjectPreview: string };
    const sentReply = parseResult(
      await handlers.get('reply-mail')!({
        messageId: 'msg-1',
        comment: 'Looks good',
      })
    ) as { sent: boolean };
    const draftForward = parseResult(
      await handlers.get('draft-mail-forward')!({
        messageId: 'msg-1',
        to: ['juan@company.com'],
        comment: 'FYI',
      })
    ) as { draftId: string; draft: { label: string } };
    const mailContext = parseResult(
      await handlers.get('get-mail-context')!({
        messageId: 'msg-1',
        includeThread: true,
        includeAttachments: true,
      })
    ) as { message: { label: string }; thread: Array<unknown>; attachments: Array<unknown> };

    expect(previewReply.wouldSend).toBe(true);
    expect(previewReply.subjectPreview).toContain('RE:');
    expect(sentReply.sent).toBe(true);
    expect(draftForward.draftId).toBe('draft-1');
    expect(draftForward.draft.label).toContain('FW: Quarterly update');
    expect(mailContext.message.label).toContain('Quarterly update');
    expect(mailContext.thread).toHaveLength(1);
    expect(mailContext.attachments).toHaveLength(1);
  });

  it('reschedules calendar events with preview and execution', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const graphClient = createMockGraphClient(async (endpoint, options) => {
      if (endpoint.startsWith('/me/events/event-1?')) {
        return {
          id: 'event-1',
          subject: 'Planning sync',
          start: { dateTime: '2026-03-12T13:00:00Z' },
          end: { dateTime: '2026-03-12T13:30:00Z' },
          attendees: [{ emailAddress: { name: 'Juan', address: 'juan@company.com' } }],
        };
      }
      if (endpoint === '/me/findMeetingTimes') {
        return {
          meetingTimeSuggestions: [
            {
              meetingTimeSlot: {
                start: { dateTime: '2026-03-12T14:00:00Z' },
                end: { dateTime: '2026-03-12T14:30:00Z' },
              },
              confidence: 0.9,
              attendeeAvailability: [
                {
                  availability: 'free',
                  attendee: { emailAddress: { name: 'Juan', address: 'juan@company.com' } },
                },
              ],
            },
          ],
        };
      }
      if (endpoint === '/me/events/event-1') {
        expect(options?.method).toBe('PATCH');
        return {
          id: 'event-1',
          subject: 'Planning sync',
          start: { dateTime: '2026-03-12T14:00:00Z' },
          end: { dateTime: '2026-03-12T14:30:00Z' },
          attendees: [{ emailAddress: { name: 'Juan', address: 'juan@company.com' } }],
        };
      }
      throw new Error(`Unexpected endpoint: ${endpoint}`);
    });

    registerGraphTools(server, graphClient, false, 'reschedule-calendar-event', false);

    const previewResult = parseResult(
      await handlers.get('reschedule-calendar-event')!({
        eventId: 'event-1',
        previewOnly: true,
      })
    ) as { previewOnly: boolean; chosenSlot: { start: string } };
    const executeResult = parseResult(
      await handlers.get('reschedule-calendar-event')!({
        eventId: 'event-1',
      })
    ) as { updatedEvent: { start: string; label: string } };

    expect(previewResult.previewOnly).toBe(true);
    expect(previewResult.chosenSlot.start).toBe('2026-03-12T14:00:00Z');
    expect(executeResult.updatedEvent.start).toBe('2026-03-12T14:00:00Z');
    expect(executeResult.updatedEvent.label).toContain('Planning sync');
  });

  it('searches M365 content across surfaces with normalized labels', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const graphClient = createMockGraphClient(async (endpoint, options) => {
      if (endpoint === '/search/query') {
        const body = JSON.parse((options?.body as string) ?? '{}') as {
          requests?: Array<{ entityTypes?: string[] }>;
        };
        const entityTypes = body.requests?.[0]?.entityTypes ?? [];
        if (entityTypes.includes('chatMessage')) {
          return {
            value: [
              {
                hitsContainers: [
                  {
                    hits: [
                      {
                        resource: {
                          id: 'chat-msg-1',
                          chatId: 'chat-1',
                          chatType: 'group',
                          topic: 'AI Leads',
                          createdDateTime: '2026-03-12T11:00:00Z',
                          from: { user: { id: 'juan-id', displayName: 'Juan Casiraghi' } },
                          body: { content: '<p>deploy prod tonight</p>' },
                        },
                      },
                    ],
                  },
                ],
              },
            ],
          };
        }
        if (entityTypes.includes('driveItem') && entityTypes.length === 1) {
          return {
            value: [
              {
                hitsContainers: [
                  {
                    hits: [
                      {
                        resource: {
                          id: 'file-1',
                          name: 'deploy-plan.md',
                          webUrl: 'https://contoso.sharepoint.com/deploy-plan',
                          parentReference: {
                            driveId: 'drive-1',
                            path: '/drives/drive-1/root:/Docs',
                          },
                          lastModifiedDateTime: '2026-03-12T10:00:00Z',
                          file: { mimeType: 'text/markdown' },
                        },
                      },
                    ],
                  },
                ],
              },
            ],
          };
        }
        return {
          value: [
            {
              hitsContainers: [
                {
                  hits: [
                    {
                      resource: {
                        id: 'page-1',
                        title: 'Deploy runbook',
                        webUrl: 'https://contoso.sharepoint.com/sites/ops/SitePages/Deploy.aspx',
                        siteId: 'site-1',
                        contentclass: 'STS_SitePage',
                        body: { content: 'deploy prod checklist' },
                      },
                    },
                  ],
                },
              ],
            },
          ],
        };
      }
      if (endpoint.startsWith('/me/messages?')) {
        return {
          value: [
            {
              id: 'mail-1',
              subject: 'Deploy prod',
              from: { emailAddress: { name: 'Mauro', address: 'mauro@company.com' } },
              toRecipients: [
                { emailAddress: { name: 'Santiago', address: 'santiago@company.com' } },
              ],
              receivedDateTime: '2026-03-12T12:00:00Z',
              bodyPreview: 'deploy prod tonight',
            },
          ],
        };
      }
      throw new Error(`Unexpected endpoint: ${endpoint}`);
    });

    registerGraphTools(server, graphClient, false, 'search-m365-content', true);

    const result = parseResult(
      await handlers.get('search-m365-content')!({
        query: 'deploy prod',
        limit: 10,
      })
    ) as { items: Array<{ surface: string; label: string }> };

    expect(result.items.map((item) => item.surface)).toEqual([
      'mail',
      'chatMessage',
      'file',
      'sharepointPage',
    ]);
    expect(result.items[0]?.label).toContain('Deploy prod');
    expect(result.items[2]?.label).toContain('deploy-plan.md');
  });

  it('includes missingScopes when a custom tool fails with insufficient scope', async () => {
    const server = new McpServer({ name: 'test', version: '1.0.0' });
    const handlers = captureHandlers(server);
    const authManager = {
      isOAuthModeEnabled: () => false,
      getTokenForAccount: async () => createJwtWithScopes(['Mail.Read']),
    };
    const graphClient = createMockGraphClient(async (endpoint) => {
      if (endpoint.startsWith('/me/messages/msg-1?')) {
        return {
          id: 'msg-1',
          subject: 'Quarterly update',
          from: { emailAddress: { name: 'Mauro', address: 'mauro@company.com' } },
          toRecipients: [{ emailAddress: { name: 'Santiago', address: 'santiago@company.com' } }],
        };
      }
      if (endpoint === '/me/messages/msg-1/reply') {
        throw new Error('Microsoft Graph API scope error: 403 Forbidden - scope error');
      }
      throw new Error(`Unexpected endpoint: ${endpoint}`);
    });

    registerGraphTools(server, graphClient, false, 'reply-mail', false, authManager as never);

    const result = await handlers.get('reply-mail')!({
      messageId: 'msg-1',
      comment: 'Thanks',
      account: 'me@company.com',
    });
    const parsed = parseResult(result) as {
      error: string;
      requiredScopes: string[];
      missingScopes: string[];
    };

    expect(result.isError).toBe(true);
    expect(parsed.error).toBe('insufficient_scope');
    expect(parsed.requiredScopes).toContain('Mail.ReadWrite');
    expect(parsed.missingScopes).toContain('Mail.ReadWrite');
  });
});
