import { z } from 'zod';
import GraphClient from '../graph-client.js';
import AuthManager from '../auth.js';
import logger from '../logger.js';
import { getRequestTokens } from '../request-context.js';

export type ToolResult = Awaited<ReturnType<GraphClient['graphRequest']>>;
export type JsonRecord = Record<string, unknown>;
export type ToolSchema = Record<string, z.ZodTypeAny>;

export type ToolRegistryEntry = {
  name: string;
  description: string;
  method: string;
  path: string;
  requiresOrgMode: boolean;
  readOnlyHint: boolean;
  openWorldHint: boolean;
  destructiveHint?: boolean;
  handler: (params: Record<string, unknown>) => Promise<ToolResult>;
};

export type CustomToolContext = {
  graphClient: GraphClient;
  authManager?: AuthManager;
};

export type CustomToolHandlerContext = {
  graphClient: GraphClient;
  authManager?: AuthManager;
  accessToken?: string;
};

export type CustomToolDefinition = {
  name: string;
  description: string;
  method: 'GET' | 'POST' | 'PUT' | 'PATCH' | 'DELETE';
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

export type GraphCollectionResponse<T> = {
  value?: T[];
  '@odata.nextLink'?: string;
};

export type IdentitySummary = {
  userId: string | null;
  displayName: string | null;
  email?: string | null;
};

export type NormalizedUser = {
  userId: string;
  displayName: string | null;
  email: string | null;
  userPrincipalName: string | null;
  givenName: string | null;
  surname: string | null;
  jobTitle: string | null;
  department: string | null;
};

export type RankedUser = NormalizedUser & {
  matchScore: number;
  matchReasons: string[];
};

export type NormalizedChatMember = {
  memberId: string;
  userId: string | null;
  displayName: string | null;
  email: string | null;
  userPrincipalName: string | null;
  roles: string[];
  membershipType: string | null;
  membershipOrigins?: string[];
};

export type NormalizedChat = {
  id: string;
  chatType: string;
  topic: string | null;
  lastUpdatedDateTime: string | null;
  webUrl?: string | null;
  memberCount: number;
};

export type DriveItemSummary = {
  id: string;
  driveId: string | null;
  name: string;
  path: string;
  isFolder: boolean;
  mimeType: string | null;
  size: number | null;
  lastModifiedDateTime: string | null;
  modifiedBy: IdentitySummary | null;
  webUrl: string | null;
};

export type SiteSummary = {
  siteId: string;
  name: string | null;
  displayName: string | null;
  webUrl: string | null;
  isPersonalSite: boolean;
};

export type CalendarParticipant = {
  userId: string | null;
  displayName: string | null;
  email: string | null;
  matchScore?: number;
};

export type AvailabilitySlot = {
  start: string;
  end: string;
  score: number;
};

export type TaskSummary = {
  taskId: string;
  listId: string;
  title: string;
  status: string | null;
  importance: string | null;
  dueDateTime: string | null;
  bodyPreview: string | null;
};

export type MailSummary = {
  messageId: string;
  subject: string | null;
  from: IdentitySummary | null;
  to: IdentitySummary[];
  receivedDateTime: string | null;
  bodyPreview: string | null;
  hasAttachments: boolean;
};

export type AttachmentSummary = {
  attachmentId: string;
  name: string | null;
  contentType: string | null;
  size: number | null;
  isInline: boolean;
};

export type ChatCandidate = {
  chat: NormalizedChat;
  members: NormalizedChatMember[];
  matchType: string;
  rank: number;
  matchReasons: string[];
  score: number;
  exactness: number;
};

export const CHAT_TYPES = new Set(['oneOnOne', 'group', 'meeting']);

export const userSelect =
  '$select=id,displayName,givenName,surname,mail,userPrincipalName,jobTitle,department';
export const chatSelect = '$select=id,chatType,topic,lastUpdatedDateTime,webUrl';
export const messageSelect = '$select=id,createdDateTime,from,body,summary';
export const driveItemSelect =
  '$select=id,name,size,lastModifiedDateTime,webUrl,parentReference,folder,file,lastModifiedBy';
export const siteSelect = '$select=id,name,displayName,webUrl';

export class GraphToolError extends Error {
  code: string;
  details?: JsonRecord;

  constructor(code: string, message: string, details?: JsonRecord) {
    super(message);
    this.code = code;
    this.details = details;
  }
}

export async function resolveAccessToken(
  params: Record<string, unknown>,
  authManager?: AuthManager
): Promise<string | undefined> {
  if (!authManager || authManager.isOAuthModeEnabled() || getRequestTokens()) {
    return undefined;
  }

  const accountParam = typeof params.account === 'string' ? params.account : undefined;
  return authManager.getTokenForAccount(accountParam);
}

export async function graphRequest<T>(
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

export function success(graphClient: GraphClient, data: unknown): ToolResult {
  return graphClient.formatJsonResponse(data);
}

export function errorResult(
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

export function ambiguousMatchResult(
  graphClient: GraphClient,
  candidates: Array<Record<string, unknown>>
): ToolResult {
  return errorResult(graphClient, 'ambiguous_match', { candidates });
}

export function inferErrorCode(error: unknown, fallbackMessage: string): string {
  if (error instanceof GraphToolError) {
    return error.code;
  }

  const message = (error as Error)?.message ?? fallbackMessage;
  if (message.includes('scope error') || message.includes('403')) {
    return 'insufficient_scope';
  }
  if (message.includes('404')) {
    return 'not_found';
  }
  return 'unsupported_operation';
}

export function buildSchemaWithCommonParams(
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

export function buildCustomToolHandler(
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
      return errorResult(context.graphClient, inferErrorCode(error, 'Custom tool request failed'), {
        message: (error as Error).message,
      });
    }
  };
}

export function encodeCursor(nextLink: string): string {
  return Buffer.from(nextLink, 'utf8').toString('base64url');
}

export function decodeCursor(cursor: string): string {
  const nextLink = Buffer.from(cursor, 'base64url').toString('utf8');
  try {
    const url = new URL(nextLink);
    return `${url.pathname.replace('/v1.0', '')}${url.search}`;
  } catch {
    throw new GraphToolError('invalid_cursor', 'Invalid pagination cursor.');
  }
}

export function encodePathSegment(value: string): string {
  return encodeURIComponent(value).replace(/%3D/g, '=');
}

export function normalizePath(path: string): string {
  if (!path || path === '/') {
    return '/';
  }

  const normalized = path.startsWith('/') ? path : `/${path}`;
  return normalized.replace(/\/+/g, '/').replace(/\/$/, '') || '/';
}

export function buildGraphPathReference(path: string): string {
  const normalized = normalizePath(path);
  if (normalized === '/') {
    return 'root';
  }
  return `root:${normalized}:`;
}

export async function resolveDriveItemByPath(
  context: CustomToolHandlerContext,
  path: string,
  siteId?: string,
  driveId?: string
): Promise<DriveItemSummary> {
  const normalizedPath = normalizePath(path);
  const endpoint = await buildDrivePathEndpoint(context, normalizedPath, siteId, driveId);
  const item = await graphRequest<JsonRecord>(context, `${endpoint}?${driveItemSelect}`);
  return normalizeDriveItem(item, normalizedPath);
}

export async function buildDrivePathEndpoint(
  context: CustomToolHandlerContext,
  path: string,
  siteId?: string,
  driveId?: string
): Promise<string> {
  const reference = buildGraphPathReference(path);
  if (driveId) {
    return reference === 'root'
      ? `/drives/${encodePathSegment(driveId)}/root`
      : `/drives/${encodePathSegment(driveId)}/${reference}`;
  }

  if (siteId) {
    return reference === 'root'
      ? `/sites/${encodePathSegment(siteId)}/drive/root`
      : `/sites/${encodePathSegment(siteId)}/drive/${reference}`;
  }

  return reference === 'root' ? '/me/drive/root' : `/me/drive/${reference}`;
}

export function normalizeUser(entry: JsonRecord): NormalizedUser {
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

export function normalizeIdentity(value: JsonRecord | undefined): IdentitySummary | null {
  if (!value) {
    return null;
  }
  const user = asRecord(value.user);
  const emailAddress = asRecord(value.emailAddress);
  return {
    userId: getString(user, 'id') ?? getString(value, 'userId'),
    displayName:
      getString(user, 'displayName') ??
      getString(emailAddress, 'name') ??
      getString(value, 'displayName'),
    email:
      getString(emailAddress, 'address') ??
      getString(user, 'mail') ??
      getString(user, 'userPrincipalName'),
  };
}

export function normalizeRecipients(entries: unknown[]): IdentitySummary[] {
  return entries
    .map(asRecord)
    .filter((entry): entry is JsonRecord => !!entry)
    .map((entry) => normalizeIdentity(entry))
    .filter((entry): entry is IdentitySummary => !!entry);
}

export async function searchUsersInternal(
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

  return (response.value ?? [])
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
}

export async function resolveQueryToBestUser(
  context: CustomToolHandlerContext,
  query: string
): Promise<RankedUser> {
  const items = await searchUsersInternal(context, query, 10);
  if (items.length === 0) {
    throw new GraphToolError('not_found', 'No matching user was found.', { query });
  }

  const [best, second] = items;
  if (!best || best.matchScore < 0.85 || (second && best.matchScore - second.matchScore < 0.1)) {
    throw new GraphToolError('ambiguous_match', 'Multiple users matched the query.', {
      candidates: items.slice(0, 5).map((item, index) => ({
        userId: item.userId,
        label: buildUserCandidateLabel(item),
        rank: index + 1,
      })),
    });
  }

  return best;
}

export function normalizeMembers(entries: JsonRecord[]): NormalizedChatMember[] {
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
    const membershipType = roles.includes('owner')
      ? 'owner'
      : (roles[0] ?? getString(entry, '@odata.type') ?? 'member');
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
  }

  return Array.from(deduped.values());
}

export function normalizeChat(entry: JsonRecord): NormalizedChat {
  const members = getArray(entry, 'members')
    .map(asRecord)
    .filter((item): item is JsonRecord => !!item);

  return {
    id: getString(entry, 'id') ?? '',
    chatType: getString(entry, 'chatType') ?? 'unknown',
    topic: getString(entry, 'topic'),
    lastUpdatedDateTime: getString(entry, 'lastUpdatedDateTime'),
    webUrl: getString(entry, 'webUrl'),
    memberCount: members.length,
  };
}

export function normalizeDriveItem(entry: JsonRecord, fallbackPath?: string): DriveItemSummary {
  const parentReference = asRecord(entry.parentReference);
  const itemPath = getString(parentReference, 'path');
  const path = fallbackPath ?? toCanvasPath(itemPath, getString(entry, 'name'));
  return {
    id: getString(entry, 'id') ?? '',
    driveId: getString(parentReference, 'driveId'),
    name: getString(entry, 'name') ?? '',
    path,
    isFolder: !!asRecord(entry.folder),
    mimeType: getString(asRecord(entry.file), 'mimeType'),
    size: getNumber(entry, 'size'),
    lastModifiedDateTime: getString(entry, 'lastModifiedDateTime'),
    modifiedBy: normalizeIdentity(asRecord(entry.lastModifiedBy)),
    webUrl: getString(entry, 'webUrl'),
  };
}

export function normalizeSite(entry: JsonRecord): SiteSummary {
  const name = getString(entry, 'name');
  const displayName = getString(entry, 'displayName') ?? name;
  const webUrl = getString(entry, 'webUrl');
  const lowerUrl = normalizeString(webUrl);
  return {
    siteId: getString(entry, 'id') ?? '',
    name,
    displayName,
    webUrl,
    isPersonalSite: lowerUrl.includes('/personal/') || lowerUrl.includes('-my.sharepoint.com'),
  };
}

export function normalizeCalendarParticipant(entry: JsonRecord): CalendarParticipant {
  const emailAddress = asRecord(entry.emailAddress) ?? asRecord(entry);
  return {
    userId: getString(entry, 'id') ?? null,
    displayName: getString(emailAddress, 'name') ?? getString(entry, 'displayName'),
    email:
      getString(emailAddress, 'address') ??
      getString(entry, 'mail') ??
      getString(entry, 'userPrincipalName'),
  };
}

export function normalizeTask(entry: JsonRecord, listId: string): TaskSummary {
  return {
    taskId: getString(entry, 'id') ?? '',
    listId,
    title: getString(entry, 'title') ?? '',
    status: getString(entry, 'status'),
    importance: getString(entry, 'importance'),
    dueDateTime: getString(asRecord(entry.dueDateTime), 'dateTime'),
    bodyPreview: buildBodyPreview(entry),
  };
}

export function normalizeMail(entry: JsonRecord): MailSummary {
  return {
    messageId: getString(entry, 'id') ?? '',
    subject: getString(entry, 'subject'),
    from: normalizeIdentity(asRecord(entry.from)),
    to: normalizeRecipients(getArray(entry, 'toRecipients')),
    receivedDateTime: getString(entry, 'receivedDateTime'),
    bodyPreview: getString(entry, 'bodyPreview') ?? buildBodyPreview(entry),
    hasAttachments: getBoolean(entry, 'hasAttachments') ?? false,
  };
}

export function normalizeAttachment(entry: JsonRecord): AttachmentSummary {
  return {
    attachmentId: getString(entry, 'id') ?? '',
    name: getString(entry, 'name'),
    contentType: getString(entry, 'contentType'),
    size: getNumber(entry, 'size'),
    isInline: getBoolean(entry, 'isInline') ?? false,
  };
}

export function buildBodyPreview(entry: JsonRecord): string | null {
  const bodyPreview = getString(entry, 'bodyPreview');
  if (bodyPreview) {
    return bodyPreview;
  }
  const summary = getString(entry, 'summary');
  if (summary) {
    return summary;
  }
  const body = asRecord(entry.body);
  const content = getString(body, 'content');
  return content ? stripHtml(content).slice(0, 280) || null : null;
}

export function extractMessageSender(message: JsonRecord): {
  userId: string | null;
  displayName: string | null;
} {
  const from = asRecord(message.from) ?? {};
  const user = asRecord(from.user) ?? {};
  return {
    userId: getString(user, 'id') ?? getString(from, 'userId') ?? null,
    displayName: getString(user, 'displayName') ?? getString(from, 'displayName') ?? null,
  };
}

export function extractStatusMessage(entry: JsonRecord): string | null {
  const statusMessage = asRecord(entry.statusMessage);
  const itemBody = asRecord(statusMessage?.message);
  return getString(itemBody, 'content');
}

export function buildUserLabel(
  user: Pick<NormalizedUser, 'displayName' | 'email' | 'userPrincipalName'>
): string {
  return user.displayName ?? user.email ?? user.userPrincipalName ?? 'Unknown user';
}

export function buildUserCandidateLabel(
  user: Pick<NormalizedUser, 'displayName' | 'email' | 'userPrincipalName'>
): string {
  if (user.displayName && user.email) {
    return `${user.displayName} <${user.email}>`;
  }
  return buildUserLabel(user);
}

export function normalizeString(value: string | null | undefined): string {
  return (value ?? '').trim().toLowerCase().replace(/\s+/g, ' ');
}

export function tokenize(value: string): string[] {
  return normalizeString(value)
    .split(/[^a-z0-9@._-]+/i)
    .map((token) => token.trim())
    .filter(Boolean);
}

export function rankUser(
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
  if (
    !reasons.length &&
    (email.startsWith(normalizedQuery) || userPrincipalName.startsWith(normalizedQuery))
  ) {
    score = Math.max(score, 0.89);
    reasons.push('mail_prefix');
  }

  const overlap =
    queryTokens.length > 0
      ? queryTokens.filter((token) => displayName.includes(token)).length / queryTokens.length
      : 0;
  if (overlap > 0) {
    score = Math.max(score, 0.6 + overlap * 0.25);
    reasons.push('displayName_token_overlap');
  }

  if (
    !reasons.length &&
    (displayName.includes(normalizedQuery) ||
      email.includes(normalizedQuery) ||
      userPrincipalName.includes(normalizedQuery))
  ) {
    score = Math.max(score, 0.6);
    reasons.push('substring_match');
  }

  return {
    score: Number(score.toFixed(2)),
    reasons,
  };
}

export function compareRankedUsers(left: RankedUser, right: RankedUser): number {
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

export function compareDates(
  left: string | null | undefined,
  right: string | null | undefined
): number {
  const leftMs = left ? Date.parse(left) : 0;
  const rightMs = right ? Date.parse(right) : 0;
  return leftMs - rightMs;
}

export function calculateRecencyBonus(lastUpdatedDateTime: string | null): number {
  if (!lastUpdatedDateTime) {
    return 0;
  }
  const now = Date.now();
  const updated = Date.parse(lastUpdatedDateTime);
  if (Number.isNaN(updated)) {
    return 0;
  }
  const ageDays = (now - updated) / (1000 * 60 * 60 * 24);
  if (ageDays <= 7) return 15;
  if (ageDays <= 30) return 8;
  if (ageDays <= 90) return 3;
  return 0;
}

export function buildUserSearchExpression(query: string): string {
  const escaped = query.replace(/"/g, '\\"');
  return `"displayName:${escaped}" OR "mail:${escaped}" OR "userPrincipalName:${escaped}"`;
}

export function stripHtml(value: string): string {
  return value
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

export function getMembershipOrigins(entry: JsonRecord): string[] {
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

export function toCanvasPath(parentPath: string | null, name: string | null): string {
  if (!parentPath) {
    return name ? `/${name}` : '/';
  }
  const cleaned = parentPath.replace(/^\/drives\/[^/]+\/root:?/, '') || '';
  const joined = `${cleaned}${cleaned.endsWith('/') || cleaned === '' ? '' : '/'}${name ?? ''}`;
  return normalizePath(joined || '/');
}

export function sameEmail(left: string | null, right: string | null): boolean {
  if (!left || !right) {
    return false;
  }
  return normalizeString(left) === normalizeString(right);
}

export function getRequiredString(params: Record<string, unknown>, key: string): string {
  const value = getOptionalString(params, key);
  if (!value) {
    throw new Error(`Missing required parameter: ${key}`);
  }
  return value;
}

export function getOptionalString(
  params: Record<string, unknown>,
  key: string
): string | undefined {
  const value = params[key];
  return typeof value === 'string' ? value : undefined;
}

export function getOptionalNumber(
  params: Record<string, unknown>,
  key: string,
  defaultValue: number
): number {
  const value = params[key];
  return typeof value === 'number' ? value : defaultValue;
}

export function getOptionalBoolean(
  params: Record<string, unknown>,
  key: string,
  defaultValue: boolean
): boolean {
  const value = params[key];
  return typeof value === 'boolean' ? value : defaultValue;
}

export function getOptionalStringArray(
  params: Record<string, unknown>,
  key: string
): string[] | undefined {
  const value = params[key];
  if (!Array.isArray(value) || value.some((entry) => typeof entry !== 'string')) {
    return undefined;
  }
  return value as string[];
}

export function getRequiredStringArray(params: Record<string, unknown>, key: string): string[] {
  const value = getOptionalStringArray(params, key);
  if (!value) {
    throw new Error(`Missing required parameter: ${key}`);
  }
  return value;
}

export function asRecord(value: unknown): JsonRecord | undefined {
  return value && typeof value === 'object' && !Array.isArray(value)
    ? (value as JsonRecord)
    : undefined;
}

export function getString(record: JsonRecord | undefined, key: string): string | null {
  if (!record) {
    return null;
  }
  const value = record[key];
  return typeof value === 'string' ? value : null;
}

export function getNumber(record: JsonRecord | undefined, key: string): number | null {
  if (!record) {
    return null;
  }
  const value = record[key];
  return typeof value === 'number' ? value : null;
}

export function getBoolean(record: JsonRecord | undefined, key: string): boolean | null {
  if (!record) {
    return null;
  }
  const value = record[key];
  return typeof value === 'boolean' ? value : null;
}

export function getArray(record: JsonRecord | undefined, key: string): unknown[] {
  if (!record) {
    return [];
  }
  const value = record[key];
  return Array.isArray(value) ? value : [];
}

export function getStringArray(record: JsonRecord | undefined, key: string): string[] {
  return getArray(record, key).filter((entry): entry is string => typeof entry === 'string');
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
