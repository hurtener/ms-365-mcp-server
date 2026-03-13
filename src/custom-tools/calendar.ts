import { z } from 'zod';
import {
  CalendarParticipant,
  CustomToolDefinition,
  CustomToolHandlerContext,
  GraphCollectionResponse,
  JsonRecord,
  asRecord,
  buildBodyPreview,
  compareDates,
  decodeCursor,
  encodeCursor,
  encodePathSegment,
  getArray,
  getOptionalBoolean,
  getOptionalNumber,
  getOptionalString,
  getRequiredString,
  getRequiredStringArray,
  getString,
  graphRequest,
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

async function buildAttendeeInputs(
  context: CustomToolHandlerContext,
  participants: string[]
): Promise<Array<{ displayName: string | null; address: string }>> {
  const resolved: Array<{ displayName: string | null; address: string }> = [];

  for (const participant of participants) {
    if (participant.includes('@')) {
      try {
        const user = await resolveQueryToBestUser(context, participant);
        resolved.push({
          displayName: user.displayName,
          address: user.email ?? user.userPrincipalName ?? participant,
        });
      } catch {
        resolved.push({
          displayName: null,
          address: participant,
        });
      }
      continue;
    }

    const user = await resolveQueryToBestUser(context, participant);
    resolved.push({
      displayName: user.displayName,
      address: user.email ?? user.userPrincipalName ?? participant,
    });
  }

  return resolved;
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
): Promise<{ slots?: Array<{ start: string; end: string; score: number }>; items?: JsonRecord[] }> {
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
    .filter((entry): entry is JsonRecord => !!entry);

  if (returnSuggestions) {
    return {
      items: suggestions.map((entry) => ({
        start: getString(
          asRecord(entry.meetingTimeSlot?.start as JsonRecord | undefined),
          'dateTime'
        ),
        end: getString(asRecord(entry.meetingTimeSlot?.end as JsonRecord | undefined), 'dateTime'),
        score: scoreMeetingTime(entry),
        attendeeAvailability: getArray(entry, 'attendeeAvailability'),
        locations: getArray(entry, 'locations'),
      })),
    };
  }

  return {
    slots: suggestions.map((entry) => ({
      start:
        getString(asRecord(entry.meetingTimeSlot?.start as JsonRecord | undefined), 'dateTime') ??
        start,
      end:
        getString(asRecord(entry.meetingTimeSlot?.end as JsonRecord | undefined), 'dateTime') ??
        end,
      score: scoreMeetingTime(entry),
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
      const resolvedAttendees = [];

      for (const entry of attendeesInput) {
        if (typeof entry === 'string') {
          const [resolved] = await buildAttendeeInputs(context, [entry]);
          resolvedAttendees.push(resolved);
          continue;
        }
        const record = asRecord(entry);
        if (!record) {
          continue;
        }
        const email = getString(record, 'email');
        if (!email) {
          continue;
        }
        resolvedAttendees.push({
          address: email,
          displayName: getString(record, 'displayName'),
        });
      }

      const endpoint = calendarId
        ? `/me/calendars/${encodePathSegment(calendarId)}/events`
        : '/me/events';
      const response = await graphRequest<JsonRecord>(context, endpoint, {
        method: 'POST',
        body: JSON.stringify({
          subject,
          start: { dateTime: start, timeZone: 'UTC' },
          end: { dateTime: end, timeZone: 'UTC' },
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
          attendees: resolvedAttendees.map((attendee) => ({
            type: 'required',
            emailAddress: {
              address: attendee.address,
              name: attendee.displayName ?? attendee.address,
            },
          })),
        }),
      });
      return success(context.graphClient, normalizeEvent(response));
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
