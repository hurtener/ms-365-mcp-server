import { z } from 'zod';
import {
  CustomToolDefinition,
  CustomToolHandlerContext,
  GraphToolError,
  compareDates,
  getOptionalNumber,
  getOptionalString,
  getRequiredString,
  normalizeString,
  resolveQueryToBestUser,
  sameEmail,
  success,
} from './shared.js';
import { searchFilesInternal } from './files.js';
import { searchMailInternal } from './mail.js';
import { searchSharePointContentInternal } from './sharepoint.js';
import { searchMessagesInternal } from './teams.js';

type SearchSurface = 'mail' | 'chatMessage' | 'file' | 'sharepointPage';

type SearchHit = {
  surface: SearchSurface;
  id: string;
  label: string;
  title: string | null;
  bodyPreview: string | null;
  webUrl: string | null;
  timestamp: string | null;
  containerLabel: string | null;
  participants?: Array<{
    displayName: string | null;
    email?: string | null;
    userId?: string | null;
  }>;
  path?: string | null;
  siteId?: string | null;
  score: number;
  exactness: number;
};

function encodeSearchCursor(offset: number): string {
  return Buffer.from(JSON.stringify({ offset }), 'utf8').toString('base64url');
}

function decodeSearchCursor(cursor: string | undefined): number {
  if (!cursor) {
    return 0;
  }
  try {
    const parsed = JSON.parse(Buffer.from(cursor, 'base64url').toString('utf8')) as {
      offset?: number;
    };
    return typeof parsed.offset === 'number' ? parsed.offset : 0;
  } catch {
    throw new GraphToolError('invalid_cursor', 'Invalid pagination cursor.');
  }
}

function scoreExactness(query: string, values: Array<string | null | undefined>): number {
  const normalizedQuery = normalizeString(query);
  if (!normalizedQuery) {
    return 0;
  }
  let best = 0;
  for (const value of values) {
    const normalizedValue = normalizeString(value);
    if (!normalizedValue) {
      continue;
    }
    if (normalizedValue === normalizedQuery) {
      best = Math.max(best, 3);
    } else if (normalizedValue.startsWith(normalizedQuery)) {
      best = Math.max(best, 2);
    } else if (normalizedValue.includes(normalizedQuery)) {
      best = Math.max(best, 1);
    }
  }
  return best;
}

function compareHits(left: SearchHit, right: SearchHit): number {
  const scoreDelta = right.score - left.score;
  if (scoreDelta !== 0) {
    return scoreDelta;
  }
  const recencyDelta = compareDates(right.timestamp, left.timestamp);
  if (recencyDelta !== 0) {
    return recencyDelta;
  }
  const exactnessDelta = right.exactness - left.exactness;
  if (exactnessDelta !== 0) {
    return exactnessDelta;
  }
  return left.label.localeCompare(right.label);
}

export const searchToolDefinitions: CustomToolDefinition[] = [
  {
    name: 'search-m365-content',
    description:
      'Search mail, chat messages, files, and SharePoint pages through one normalized surface.',
    method: 'POST',
    path: '/search/query',
    requiresOrgMode: true,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: [
      'Mail.Read',
      'Chat.Read',
      'ChannelMessage.Read.All',
      'Files.Read.All',
      'Sites.Read.All',
      'User.Read.All',
    ],
    schema: {
      query: z.string().min(1).describe('Cross-surface search query'),
      surfaces: z
        .array(z.enum(['mail', 'chatMessage', 'file', 'sharepointPage']))
        .optional()
        .describe('Optional surface filter'),
      limit: z.number().int().min(1).max(50).default(20).describe('Maximum results'),
      cursor: z.string().optional().describe('Opaque pagination cursor'),
      participantQuery: z
        .string()
        .min(1)
        .optional()
        .describe('Optional person filter for mail/chat results'),
      siteId: z
        .string()
        .min(1)
        .optional()
        .describe('Optional SharePoint site filter for file/page hits'),
      path: z.string().min(1).optional().describe('Optional path prefix for file hits'),
    },
    handler: async (params, context) => {
      const query = getRequiredString(params, 'query');
      const surfaces = (params.surfaces as SearchSurface[] | undefined) ?? [
        'mail',
        'chatMessage',
        'file',
        'sharepointPage',
      ];
      const limit = getOptionalNumber(params, 'limit', 20);
      const cursorOffset = decodeSearchCursor(getOptionalString(params, 'cursor'));
      const participantQuery = getOptionalString(params, 'participantQuery');
      const siteId = getOptionalString(params, 'siteId');
      const path = getOptionalString(params, 'path');
      const perSurfaceLimit = Math.max(limit * 3, 25);
      const participant = participantQuery
        ? await resolveQueryToBestUser(context, participantQuery)
        : undefined;
      const hits: SearchHit[] = [];

      if (surfaces.includes('mail')) {
        const mailResult = await searchMailInternal(context, query, perSurfaceLimit);
        hits.push(
          ...mailResult.items
            .filter((item) => {
              if (!participant) {
                return true;
              }
              const participantEmail = participant.email ?? participant.userPrincipalName;
              return (
                sameEmail(item.from?.email ?? null, participantEmail ?? null) ||
                item.to.some((recipient) =>
                  sameEmail(recipient.email ?? null, participantEmail ?? null)
                )
              );
            })
            .map((item, index) => ({
              surface: 'mail' as const,
              id: item.messageId,
              label: item.label,
              title: item.subject,
              bodyPreview: item.bodyPreview,
              webUrl: null,
              timestamp: item.receivedDateTime,
              containerLabel: item.from?.label ?? null,
              participants: [item.from, ...item.to].filter(Boolean),
              score: 90 - index,
              exactness: scoreExactness(query, [item.subject, item.bodyPreview, item.label]),
            }))
        );
      }

      if (surfaces.includes('chatMessage')) {
        const messageHits = await searchMessagesInternal(
          context,
          query,
          participant?.userId,
          perSurfaceLimit
        );
        hits.push(
          ...messageHits.map((item, index) => ({
            surface: 'chatMessage' as const,
            id: String(item.messageId ?? ''),
            label: String(item.topic ?? item.bodyPreview ?? 'Chat message'),
            title: (item.topic as string | null) ?? null,
            bodyPreview: (item.bodyPreview as string | null) ?? null,
            webUrl: null,
            timestamp: (item.createdDateTime as string | null) ?? null,
            containerLabel:
              (item.topic as string | null) ?? (item.chatType as string | null) ?? 'Chat',
            participants: item.from
              ? [item.from as { userId?: string | null; displayName?: string | null }]
              : undefined,
            score: 80 - index,
            exactness: scoreExactness(query, [
              item.topic as string | null,
              item.bodyPreview as string | null,
            ]),
          }))
        );
      }

      if (!participant && surfaces.includes('file')) {
        const fileHits = await searchFilesInternal(context, query, perSurfaceLimit, path, siteId);
        hits.push(
          ...fileHits.map((item, index) => ({
            surface: 'file' as const,
            id: item.id,
            label: item.label,
            title: item.name,
            bodyPreview: null,
            webUrl: item.webUrl,
            timestamp: item.lastModifiedDateTime,
            containerLabel: item.path.split('/').slice(0, -1).join('/') || '/',
            path: item.path,
            score: 70 - index,
            exactness: scoreExactness(query, [item.name, item.path, item.label]),
          }))
        );
      }

      if (!participant && surfaces.includes('sharepointPage')) {
        const pageHits = await searchSharePointContentInternal(
          context,
          query,
          perSurfaceLimit,
          siteId
        );
        hits.push(
          ...pageHits
            .filter(
              (item) =>
                normalizeString(item.contentType).includes('sitepage') ||
                normalizeString(item.contentType).includes('page')
            )
            .map((item, index) => ({
              surface: 'sharepointPage' as const,
              id: item.id,
              label: item.title ?? item.webUrl ?? item.id,
              title: item.title,
              bodyPreview: item.bodyPreview,
              webUrl: item.webUrl,
              timestamp: null,
              containerLabel: item.siteId,
              siteId: item.siteId,
              score: 60 - index,
              exactness: scoreExactness(query, [item.title, item.bodyPreview]),
            }))
        );
      }

      const sorted = hits.sort(compareHits);
      const items = sorted
        .slice(cursorOffset, cursorOffset + limit)
        .map(({ score: _score, exactness: _exactness, ...item }) => item);
      const nextCursor =
        cursorOffset + limit < sorted.length ? encodeSearchCursor(cursorOffset + limit) : undefined;

      return success(context.graphClient, { items, nextCursor });
    },
  },
];
