import { z } from 'zod';
import {
  CalendarParticipant,
  CustomToolDefinition,
  CustomToolHandlerContext,
  GraphToolError,
  GraphCollectionResponse,
  JsonRecord,
  asRecord,
  buildBodyPreview,
  compareDates,
  decodeCursor,
  encodeCursor,
  encodePathSegment,
  errorResult,
  getArray,
  getOptionalBoolean,
  getOptionalNumber,
  getOptionalString,
  getRequiredString,
  getRequiredStringArray,
  getString,
  graphRequest,
  inferErrorCode,
  normalizeCalendarParticipant,
  resolveQueryToBestUser,
  searchUsersInternal,
  success,
} from './shared.js';

type CalendarEventSummary = {
  id: string;
  subject: string | null;
  start: string | null;
  end: string | null;
  organizer: CalendarParticipant | null;
  attendees: CalendarParticipant[];
  onlineMeetingUrl: string | null;
  location: string | null;
  bodyPreview: string | null;
  webUrl: string | null;
  isOnlineMeeting: boolean;
};

type ResolvedAttendeeInput = {
  displayName: string | null;
  address: string;
};

type MeetingSuggestionAttendeeStatus = {
  displayName: string | null;
  email: string | null;
  availability: string | null;
  isAvailable: boolean;
};

type MeetingSuggestionSummary = {
  start: string;
  end: string;
  score: number;
  rankingScore: number;
  allAttendeesAvailable: boolean;
  availableAttendeeCount: number;
  blockedAttendeeCount: number;
  attendeeAvailability: MeetingSuggestionAttendeeStatus[];
  locations: unknown[];
};

const AVAILABLE_ATTENDEE_STATUSES = new Set(['free', 'workingElsewhere']);

function calendarErrorResult(
  context: CustomToolHandlerContext,
  error: unknown,
  fallbackMessage: string
) {
  const extra =
    error instanceof GraphToolError
      ? {
          message: error.message || fallbackMessage,
          ...(error.details ?? {}),
        }
      : {
          message: (error as Error)?.message ?? fallbackMessage,
        };
  return errorResult(context.graphClient, inferErrorCode(error, fallbackMessage), extra);
}

function normalizeEvent(entry: JsonRecord): CalendarEventSummary {
  return {
    id: getString(entry, 'id') ?? '',
    subject: getString(entry, 'subject'),
    start: getString(asRecord(entry.start), 'dateTime'),
    end: getString(asRecord(entry.end), 'dateTime'),
    organizer: normalizeCalendarParticipant(asRecord(entry.organizer) ?? {}),
    attendees: getArray(entry, 'attendees')
      .map(asRecord)
      .filter((item): item is JsonRecord => !!item)
      .map(normalizeCalendarParticipant),
    onlineMeetingUrl:
      getString(asRecord(entry.onlineMeeting), 'joinUrl') ?? getString(entry, 'onlineMeetingUrl'),
    location:
      getString(asRecord(entry.location), 'displayName') ??
      getString(asRecord(entry.locations?.[0] as JsonRecord | undefined), 'displayName'),
    bodyPreview: buildBodyPreview(entry),
    webUrl: getString(entry, 'webLink') ?? getString(entry, 'webUrl'),
    isOnlineMeeting: entry.isOnlineMeeting === true,
  };
}

function eventMatchesQuery(event: CalendarEventSummary, query: string): boolean {
  const lowered = query.toLowerCase();
  return [
    event.subject,
    event.organizer?.displayName,
    event.organizer?.email,
    event.location,
    event.bodyPreview,
    ...event.attendees.flatMap((attendee) => [attendee.displayName, attendee.email]),
  ]
    .filter((value): value is string => !!value)
    .some((value) => value.toLowerCase().includes(lowered));
}

async function listCalendarEventsInternal(
  context: CustomToolHandlerContext,
  start: string,
  end: string,
  limit: number,
  calendarId?: string,
  cursor?: string
): Promise<{ items: CalendarEventSummary[]; nextCursor?: string }> {
  const basePath = calendarId
    ? `/me/calendars/${encodePathSegment(calendarId)}/calendarView`
    : '/me/calendarView';
  const select =
    '$select=id,subject,start,end,organizer,attendees,bodyPreview,body,location,onlineMeeting,onlineMeetingUrl,isOnlineMeeting,webLink';
  const endpoint = cursor
    ? decodeCursor(cursor)
    : `${basePath}?startDateTime=${encodeURIComponent(start)}&endDateTime=${encodeURIComponent(end)}&$top=${limit}&${select}`;
  const response = await graphRequest<GraphCollectionResponse<JsonRecord>>(context, endpoint);
  const items = (response.value ?? [])
    .map(normalizeEvent)
    .sort((left, right) => compareDates(left.start, right.start));
  return {
    items,
    nextCursor: response['@odata.nextLink'] ? encodeCursor(response['@odata.nextLink']) : undefined,
  };
}

async function getCalendarEventInternal(
  context: CustomToolHandlerContext,
  eventId: string,
  calendarId?: string
): Promise<CalendarEventSummary> {
  const basePath = calendarId
    ? `/me/calendars/${encodePathSegment(calendarId)}/events/${encodePathSegment(eventId)}`
    : `/me/events/${encodePathSegment(eventId)}`;
  const select =
    '$select=id,subject,start,end,organizer,attendees,bodyPreview,body,location,onlineMeeting,onlineMeetingUrl,isOnlineMeeting,webLink';
  const response = await graphRequest<JsonRecord>(context, `${basePath}?${select}`);
  return normalizeEvent(response);
}

async function resolveAttendeeInputs(
  context: CustomToolHandlerContext,
  attendees: unknown[]
): Promise<ResolvedAttendeeInput[]> {
  const resolved: ResolvedAttendeeInput[] = [];

  for (const attendee of attendees) {
    if (typeof attendee === 'string') {
      if (attendee.includes('@')) {
        try {
          const user = await resolveQueryToBestUser(context, attendee);
          resolved.push({
            displayName: user.displayName,
            address: user.email ?? user.userPrincipalName ?? attendee,
          });
        } catch {
          resolved.push({
            displayName: null,
            address: attendee,
          });
        }
        continue;
      }

      const user = await resolveQueryToBestUser(context, attendee);
      resolved.push({
        displayName: user.displayName,
        address: user.email ?? user.userPrincipalName ?? attendee,
      });
      continue;
    }

    const record = asRecord(attendee);
    if (!record) {
      continue;
    }

    const email = getString(record, 'email');
    if (!email) {
      continue;
    }

    resolved.push({
      displayName: getString(record, 'displayName'),
      address: email,
    });
  }

  return resolved;
}

async function buildAttendeeInputs(
  context: CustomToolHandlerContext,
  participants: string[]
): Promise<ResolvedAttendeeInput[]> {
  return resolveAttendeeInputs(context, participants);
}

function normalizeAttendeeAvailability(entry: JsonRecord): MeetingSuggestionAttendeeStatus {
  const attendee = normalizeCalendarParticipant(asRecord(entry.attendee) ?? {});
  const availability = getString(entry, 'availability');
  return {
    displayName: attendee.displayName,
    email: attendee.email,
    availability,
    isAvailable: availability ? AVAILABLE_ATTENDEE_STATUSES.has(availability) : false,
  };
}

function buildMeetingSuggestionSummary(
  entry: JsonRecord,
  fallbackStart: string,
  fallbackEnd: string
): MeetingSuggestionSummary {
  const attendeeAvailability = getArray(entry, 'attendeeAvailability')
    .map(asRecord)
    .filter((item): item is JsonRecord => !!item)
    .map(normalizeAttendeeAvailability);
  const availableAttendeeCount = attendeeAvailability.filter((item) => item.isAvailable).length;
  const blockedAttendeeCount = attendeeAvailability.length - availableAttendeeCount;
  const score = scoreMeetingTime(entry);
  const allAttendeesAvailable =
    attendeeAvailability.length > 0 && blockedAttendeeCount === 0 && availableAttendeeCount > 0;
  const availabilityRatio =
    attendeeAvailability.length > 0 ? availableAttendeeCount / attendeeAvailability.length : 0;
  const rankingScore = Number(
    Math.max(
      0,
      Math.min(1, score + (allAttendeesAvailable ? 0.2 : 0) + availabilityRatio * 0.1)
    ).toFixed(2)
  );

  return {
    start:
      getString(asRecord(entry.meetingTimeSlot?.start as JsonRecord | undefined), 'dateTime') ??
      fallbackStart,
    end:
      getString(asRecord(entry.meetingTimeSlot?.end as JsonRecord | undefined), 'dateTime') ??
      fallbackEnd,
    score,
    rankingScore,
    allAttendeesAvailable,
    availableAttendeeCount,
    blockedAttendeeCount,
    attendeeAvailability,
    locations: getArray(entry, 'locations'),
  };
}

function compareMeetingSuggestions(
  left: Pick<MeetingSuggestionSummary, 'allAttendeesAvailable' | 'rankingScore' | 'start'>,
  right: Pick<MeetingSuggestionSummary, 'allAttendeesAvailable' | 'rankingScore' | 'start'>
): number {
  if (left.allAttendeesAvailable !== right.allAttendeesAvailable) {
    return left.allAttendeesAvailable ? -1 : 1;
  }
  if (left.rankingScore !== right.rankingScore) {
    return right.rankingScore - left.rankingScore;
  }
  return compareDates(left.start, right.start);
}

async function createCalendarEventInternal(
  context: CustomToolHandlerContext,
  params: Record<string, unknown>,
  options: {
    subject: string;
    start: string;
    end: string;
    attendees: ResolvedAttendeeInput[];
    calendarId?: string;
  }
): Promise<CalendarEventSummary> {
  const endpoint = options.calendarId
    ? `/me/calendars/${encodePathSegment(options.calendarId)}/events`
    : '/me/events';
  const response = await graphRequest<JsonRecord>(context, endpoint, {
    method: 'POST',
    body: JSON.stringify({
      subject: options.subject,
      start: { dateTime: options.start, timeZone: 'UTC' },
      end: { dateTime: options.end, timeZone: 'UTC' },
      body: getOptionalString(params, 'body')
        ? {
            contentType: 'text',
            content: getOptionalString(params, 'body'),
          }
        : undefined,
      location: getOptionalString(params, 'location')
        ? { displayName: getOptionalString(params, 'location') }
        : undefined,
      isOnlineMeeting: getOptionalBoolean(params, 'isOnlineMeeting', false),
      attendees: options.attendees.map((attendee) => ({
        type: 'required',
        emailAddress: {
          address: attendee.address,
          name: attendee.displayName ?? attendee.address,
        },
      })),
    }),
  });
  return normalizeEvent(response);
}

async function respondToCalendarEventInternal(
  context: CustomToolHandlerContext,
  responseAction: 'accept' | 'decline' | 'tentativelyAccept',
  eventId: string,
  sendResponse: boolean,
  comment?: string,
  calendarId?: string
): Promise<{
  eventId: string;
  calendarId: string | null;
  response: 'accepted' | 'declined' | 'tentativelyAccepted';
  sendResponse: boolean;
  comment: string | null;
}> {
  const basePath = calendarId
    ? `/me/calendars/${encodePathSegment(calendarId)}/events/${encodePathSegment(eventId)}`
    : `/me/events/${encodePathSegment(eventId)}`;
  await graphRequest<JsonRecord>(context, `${basePath}/${responseAction}`, {
    method: 'POST',
    body: JSON.stringify({
      sendResponse,
      ...(comment ? { comment } : {}),
    }),
  });

  return {
    eventId,
    calendarId: calendarId ?? null,
    response:
      responseAction === 'accept'
        ? 'accepted'
        : responseAction === 'decline'
          ? 'declined'
          : 'tentativelyAccepted',
    sendResponse,
    comment: comment ?? null,
  };
}

function scoreMeetingTime(entry: JsonRecord): number {
  const confidence = entry.confidence;
  if (typeof confidence === 'number') {
    return confidence;
  }
  const locations = getArray(entry, 'attendeeAvailability');
  if (locations.length === 0) {
    return 0.5;
  }
  const freeCount = locations.filter(
    (item) => getString(asRecord(item), 'availability') === 'free'
  ).length;
  return Number((freeCount / locations.length).toFixed(2));
}

async function findAvailabilityInternal(
  context: CustomToolHandlerContext,
  participants: string[],
  start: string,
  end: string,
  durationMinutes: number,
  returnSuggestions: boolean
): Promise<{
  slots?: Array<{
    start: string;
    end: string;
    score: number;
    rankingScore: number;
    allAttendeesAvailable: boolean;
    availableAttendeeCount: number;
    blockedAttendeeCount: number;
  }>;
  items?: Array<MeetingSuggestionSummary & Record<string, unknown>>;
}> {
  const attendees = await buildAttendeeInputs(context, participants);
  const response = await graphRequest<JsonRecord>(context, '/me/findMeetingTimes', {
    method: 'POST',
    body: JSON.stringify({
      attendees: attendees.map((attendee) => ({
        type: 'required',
        emailAddress: {
          address: attendee.address,
          name: attendee.displayName ?? attendee.address,
        },
      })),
      timeConstraint: {
        activityDomain: 'work',
        timeSlots: [
          {
            start: { dateTime: start, timeZone: 'UTC' },
            end: { dateTime: end, timeZone: 'UTC' },
          },
        ],
      },
      meetingDuration: `PT${durationMinutes}M`,
      maxCandidates: 20,
    }),
  });

  const suggestions = getArray(response, 'meetingTimeSuggestions')
    .map(asRecord)
    .filter((entry): entry is JsonRecord => !!entry)
    .map((entry) => buildMeetingSuggestionSummary(entry, start, end))
    .sort(compareMeetingSuggestions);

  if (returnSuggestions) {
    return {
      items: suggestions.map((entry) => ({
        start: entry.start,
        end: entry.end,
        score: entry.score,
        rankingScore: entry.rankingScore,
        allAttendeesAvailable: entry.allAttendeesAvailable,
        availableAttendeeCount: entry.availableAttendeeCount,
        blockedAttendeeCount: entry.blockedAttendeeCount,
        attendeeAvailability: entry.attendeeAvailability,
        locations: entry.locations,
      })),
    };
  }

  return {
    slots: suggestions.map((entry) => ({
      start: entry.start,
      end: entry.end,
      score: entry.score,
      rankingScore: entry.rankingScore,
      allAttendeesAvailable: entry.allAttendeesAvailable,
      availableAttendeeCount: entry.availableAttendeeCount,
      blockedAttendeeCount: entry.blockedAttendeeCount,
    })),
  };
}

export const calendarToolDefinitions: CustomToolDefinition[] = [
  {
    name: 'search-calendar-events',
    description: 'Search calendar events in a time range with Canvas-friendly event summaries.',
    method: 'GET',
    path: '/me/calendarView',
    requiresOrgMode: false,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Calendars.Read'],
    schema: {
      query: z.string().min(1).describe('Search query'),
      start: z.string().min(1).describe('Range start in ISO 8601'),
      end: z.string().min(1).describe('Range end in ISO 8601'),
      limit: z.number().int().min(1).max(50).default(25).describe('Maximum results'),
      cursor: z.string().optional().describe('Opaque pagination cursor'),
      calendarId: z.string().min(1).optional().describe('Optional explicit calendar ID'),
    },
    handler: async (params, context) => {
      const query = getRequiredString(params, 'query');
      const start = getRequiredString(params, 'start');
      const end = getRequiredString(params, 'end');
      const limit = getOptionalNumber(params, 'limit', 25);
      const cursor = getOptionalString(params, 'cursor');
      const calendarId = getOptionalString(params, 'calendarId');
      const result = await listCalendarEventsInternal(
        context,
        start,
        end,
        Math.max(limit * 3, 25),
        calendarId,
        cursor
      );
      return success(context.graphClient, {
        items: result.items.filter((event) => eventMatchesQuery(event, query)).slice(0, limit),
        nextCursor: result.nextCursor,
      });
    },
  },
  {
    name: 'resolve-attendees',
    description: 'Resolve attendee names or emails into candidate Microsoft users.',
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
      const items = await Promise.all(
        queries.map(async (query) => ({
          query,
          matches: await searchUsersInternal(context, query, 5),
        }))
      );
      return success(context.graphClient, { items });
    },
  },
  {
    name: 'find-availability',
    description: 'Find likely meeting slots for a set of participants.',
    method: 'POST',
    path: '/me/findMeetingTimes',
    requiresOrgMode: false,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Calendars.Read', 'User.Read.All'],
    schema: {
      participants: z.array(z.string().min(1)).min(1).max(20).describe('Names or emails'),
      start: z.string().min(1).describe('Range start in ISO 8601'),
      end: z.string().min(1).describe('Range end in ISO 8601'),
      durationMinutes: z.number().int().min(15).max(480).describe('Desired meeting duration'),
    },
    handler: async (params, context) => {
      const participants = getRequiredStringArray(params, 'participants');
      const start = getRequiredString(params, 'start');
      const end = getRequiredString(params, 'end');
      const durationMinutes = getOptionalNumber(params, 'durationMinutes', 30);
      return success(
        context.graphClient,
        await findAvailabilityInternal(context, participants, start, end, durationMinutes, false)
      );
    },
  },
  {
    name: 'accept-calendar-event',
    description: 'Accept a calendar event invitation.',
    method: 'POST',
    path: '/me/events/{eventId}/accept',
    requiresOrgMode: false,
    readOnlyHint: false,
    destructiveHint: true,
    openWorldHint: true,
    scopes: ['Calendars.ReadWrite'],
    schema: {
      eventId: z.string().min(1).describe('Event ID'),
      sendResponse: z.boolean().default(true).describe('Send the RSVP response to the organizer'),
      comment: z.string().optional().describe('Optional response comment'),
      calendarId: z.string().min(1).optional().describe('Optional explicit calendar ID'),
    },
    handler: async (params, context) => {
      try {
        const eventId = getRequiredString(params, 'eventId');
        const sendResponse = getOptionalBoolean(params, 'sendResponse', true);
        const comment = getOptionalString(params, 'comment');
        const calendarId = getOptionalString(params, 'calendarId');
        return success(
          context.graphClient,
          await respondToCalendarEventInternal(
            context,
            'accept',
            eventId,
            sendResponse,
            comment,
            calendarId
          )
        );
      } catch (error) {
        return calendarErrorResult(context, error, 'Unable to accept calendar event.');
      }
    },
  },
  {
    name: 'decline-calendar-event',
    description: 'Decline a calendar event invitation.',
    method: 'POST',
    path: '/me/events/{eventId}/decline',
    requiresOrgMode: false,
    readOnlyHint: false,
    destructiveHint: true,
    openWorldHint: true,
    scopes: ['Calendars.ReadWrite'],
    schema: {
      eventId: z.string().min(1).describe('Event ID'),
      sendResponse: z.boolean().default(true).describe('Send the RSVP response to the organizer'),
      comment: z.string().optional().describe('Optional response comment'),
      calendarId: z.string().min(1).optional().describe('Optional explicit calendar ID'),
    },
    handler: async (params, context) => {
      try {
        const eventId = getRequiredString(params, 'eventId');
        const sendResponse = getOptionalBoolean(params, 'sendResponse', true);
        const comment = getOptionalString(params, 'comment');
        const calendarId = getOptionalString(params, 'calendarId');
        return success(
          context.graphClient,
          await respondToCalendarEventInternal(
            context,
            'decline',
            eventId,
            sendResponse,
            comment,
            calendarId
          )
        );
      } catch (error) {
        return calendarErrorResult(context, error, 'Unable to decline calendar event.');
      }
    },
  },
  {
    name: 'tentatively-accept-calendar-event',
    description: 'Tentatively accept a calendar event invitation.',
    method: 'POST',
    path: '/me/events/{eventId}/tentativelyAccept',
    requiresOrgMode: false,
    readOnlyHint: false,
    destructiveHint: true,
    openWorldHint: true,
    scopes: ['Calendars.ReadWrite'],
    schema: {
      eventId: z.string().min(1).describe('Event ID'),
      sendResponse: z.boolean().default(true).describe('Send the RSVP response to the organizer'),
      comment: z.string().optional().describe('Optional response comment'),
      calendarId: z.string().min(1).optional().describe('Optional explicit calendar ID'),
    },
    handler: async (params, context) => {
      try {
        const eventId = getRequiredString(params, 'eventId');
        const sendResponse = getOptionalBoolean(params, 'sendResponse', true);
        const comment = getOptionalString(params, 'comment');
        const calendarId = getOptionalString(params, 'calendarId');
        return success(
          context.graphClient,
          await respondToCalendarEventInternal(
            context,
            'tentativelyAccept',
            eventId,
            sendResponse,
            comment,
            calendarId
          )
        );
      } catch (error) {
        return calendarErrorResult(context, error, 'Unable to tentatively accept calendar event.');
      }
    },
  },
  {
    name: 'get-calendar-event-details',
    description:
      'Get a normalized calendar event shape with organizer, attendees, location, and body preview.',
    method: 'GET',
    path: '/me/events/{eventId}',
    requiresOrgMode: false,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Calendars.Read'],
    schema: {
      eventId: z.string().min(1).describe('Event ID'),
      calendarId: z.string().min(1).optional().describe('Optional explicit calendar ID'),
    },
    handler: async (params, context) => {
      const eventId = getRequiredString(params, 'eventId');
      const calendarId = getOptionalString(params, 'calendarId');
      return success(
        context.graphClient,
        await getCalendarEventInternal(context, eventId, calendarId)
      );
    },
  },
  {
    name: 'suggest-meeting-times',
    description: 'Return richer meeting-time suggestions for scheduling flows.',
    method: 'POST',
    path: '/me/events',
    requiresOrgMode: false,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Calendars.Read', 'User.Read.All'],
    schema: {
      participants: z.array(z.string().min(1)).min(1).max(20).describe('Names or emails'),
      start: z.string().min(1).describe('Range start in ISO 8601'),
      end: z.string().min(1).describe('Range end in ISO 8601'),
      durationMinutes: z.number().int().min(15).max(480).describe('Desired meeting duration'),
    },
    handler: async (params, context) => {
      const participants = getRequiredStringArray(params, 'participants');
      const start = getRequiredString(params, 'start');
      const end = getRequiredString(params, 'end');
      const durationMinutes = getOptionalNumber(params, 'durationMinutes', 30);
      return success(
        context.graphClient,
        await findAvailabilityInternal(context, participants, start, end, durationMinutes, true)
      );
    },
  },
  {
    name: 'create-calendar-event-with-attendees',
    description: 'Create a calendar event using resolved attendee identity objects or strings.',
    method: 'POST',
    path: '/me/events',
    requiresOrgMode: false,
    readOnlyHint: false,
    destructiveHint: true,
    openWorldHint: true,
    scopes: ['Calendars.ReadWrite', 'User.Read.All'],
    schema: {
      subject: z.string().min(1).describe('Event subject'),
      start: z.string().min(1).describe('Start date-time in ISO 8601'),
      end: z.string().min(1).describe('End date-time in ISO 8601'),
      attendees: z
        .array(
          z.union([
            z.string(),
            z.object({ email: z.string().email(), displayName: z.string().optional() }),
          ])
        )
        .optional()
        .describe('Optional attendee identities'),
      body: z.string().optional().describe('Optional body content'),
      location: z.string().optional().describe('Optional location display name'),
      isOnlineMeeting: z.boolean().optional().describe('Create as an online meeting'),
      calendarId: z.string().min(1).optional().describe('Optional explicit calendar ID'),
    },
    handler: async (params, context) => {
      const subject = getRequiredString(params, 'subject');
      const start = getRequiredString(params, 'start');
      const end = getRequiredString(params, 'end');
      const calendarId = getOptionalString(params, 'calendarId');
      const attendeesInput = getArray(params as JsonRecord, 'attendees');
      const resolvedAttendees = await resolveAttendeeInputs(context, attendeesInput);

      return success(
        context.graphClient,
        await createCalendarEventInternal(context, params, {
          subject,
          start,
          end,
          attendees: resolvedAttendees,
          calendarId: calendarId ?? undefined,
        })
      );
    },
  },
  {
    name: 'create-calendar-event-from-availability',
    description:
      'Find the best slot where attendees are available, then create the calendar event there.',
    method: 'POST',
    path: '/me/events',
    requiresOrgMode: false,
    readOnlyHint: false,
    destructiveHint: true,
    openWorldHint: true,
    scopes: ['Calendars.ReadWrite', 'User.Read.All'],
    schema: {
      subject: z.string().min(1).describe('Event subject'),
      attendees: z
        .array(
          z.union([
            z.string(),
            z.object({ email: z.string().email(), displayName: z.string().optional() }),
          ])
        )
        .min(1)
        .max(20)
        .describe('Required attendee identities'),
      windowStart: z.string().min(1).describe('Scheduling window start in ISO 8601'),
      windowEnd: z.string().min(1).describe('Scheduling window end in ISO 8601'),
      durationMinutes: z.number().int().min(15).max(480).describe('Desired meeting duration'),
      allowPartialAvailability: z
        .boolean()
        .optional()
        .describe('Allow the best partial slot when no fully-available slot exists'),
      body: z.string().optional().describe('Optional body content'),
      location: z.string().optional().describe('Optional location display name'),
      isOnlineMeeting: z.boolean().optional().describe('Create as an online meeting'),
      calendarId: z.string().min(1).optional().describe('Optional explicit calendar ID'),
    },
    handler: async (params, context) => {
      try {
        const subject = getRequiredString(params, 'subject');
        const attendeesInput = getArray(params as JsonRecord, 'attendees');
        const windowStart = getRequiredString(params, 'windowStart');
        const windowEnd = getRequiredString(params, 'windowEnd');
        const durationMinutes = getOptionalNumber(params, 'durationMinutes', 30);
        const allowPartialAvailability = getOptionalBoolean(
          params,
          'allowPartialAvailability',
          false
        );
        const calendarId = getOptionalString(params, 'calendarId');
        const resolvedAttendees = await resolveAttendeeInputs(context, attendeesInput);

        if (resolvedAttendees.length === 0) {
          throw new GraphToolError('not_found', 'No attendees could be resolved for scheduling.');
        }

        const suggestionsResponse = await findAvailabilityInternal(
          context,
          resolvedAttendees.map((attendee) => attendee.address),
          windowStart,
          windowEnd,
          durationMinutes,
          true
        );
        const items = suggestionsResponse.items ?? [];
        const bestFullAvailability = items.find((item) => item.allAttendeesAvailable === true);
        const selectedSlot =
          bestFullAvailability ?? (allowPartialAvailability ? items[0] : undefined);

        if (!selectedSlot) {
          return errorResult(context.graphClient, 'not_found', {
            message: 'No fully-available meeting slot was found in the requested window.',
            suggestions: items.slice(0, 5).map((item, index) => ({
              rank: index + 1,
              start: item.start,
              end: item.end,
              score: item.score,
              rankingScore: item.rankingScore,
              allAttendeesAvailable: item.allAttendeesAvailable,
              blockedAttendeeCount: item.blockedAttendeeCount,
            })),
          });
        }

        const event = await createCalendarEventInternal(context, params, {
          subject,
          start: selectedSlot.start,
          end: selectedSlot.end,
          attendees: resolvedAttendees,
          calendarId: calendarId ?? undefined,
        });

        return success(context.graphClient, {
          event,
          chosenSlot: {
            start: selectedSlot.start,
            end: selectedSlot.end,
            score: selectedSlot.score,
            rankingScore: selectedSlot.rankingScore,
            allAttendeesAvailable: selectedSlot.allAttendeesAvailable,
            availableAttendeeCount: selectedSlot.availableAttendeeCount,
            blockedAttendeeCount: selectedSlot.blockedAttendeeCount,
            attendeeAvailability: selectedSlot.attendeeAvailability,
          },
        });
      } catch (error) {
        return calendarErrorResult(
          context,
          error,
          'Unable to create calendar event from availability.'
        );
      }
    },
  },
  {
    name: 'list-calendars-details',
    description: 'List calendars with default and calendar-type metadata.',
    method: 'GET',
    path: '/me/calendars',
    requiresOrgMode: false,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Calendars.Read'],
    schema: {
      limit: z.number().int().min(1).max(100).default(50).describe('Maximum results'),
      cursor: z.string().optional().describe('Opaque pagination cursor'),
    },
    handler: async (params, context) => {
      const limit = getOptionalNumber(params, 'limit', 50);
      const cursor = getOptionalString(params, 'cursor');
      const endpoint = cursor
        ? decodeCursor(cursor)
        : `/me/calendars?$top=${limit}&$select=id,name,color,canEdit,canShare,canViewPrivateItems,isDefaultCalendar,hexColor,allowedOnlineMeetingProviders,defaultOnlineMeetingProvider`;
      const response = await graphRequest<GraphCollectionResponse<JsonRecord>>(context, endpoint);
      return success(context.graphClient, {
        items: (response.value ?? []).map((entry) => ({
          calendarId: getString(entry, 'id') ?? '',
          name: getString(entry, 'name'),
          color: getString(entry, 'color') ?? getString(entry, 'hexColor'),
          isDefault: entry.isDefaultCalendar === true,
          calendarType: getString(entry, 'defaultOnlineMeetingProvider') ?? 'standard',
          canEdit: entry.canEdit === true,
        })),
        nextCursor: response['@odata.nextLink']
          ? encodeCursor(response['@odata.nextLink'])
          : undefined,
      });
    },
  },
];
