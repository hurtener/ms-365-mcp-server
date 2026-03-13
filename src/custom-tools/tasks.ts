import { z } from 'zod';
import {
  CustomToolDefinition,
  CustomToolHandlerContext,
  GraphCollectionResponse,
  GraphToolError,
  JsonRecord,
  asRecord,
  compareDates,
  decodeCursor,
  encodeCursor,
  encodePathSegment,
  getOptionalNumber,
  getOptionalString,
  getRequiredString,
  getString,
  graphRequest,
  normalizeTask,
  success,
} from './shared.js';

type TaskListSummary = {
  listId: string;
  displayName: string | null;
  wellknownListName: string | null;
  openCount: number;
  totalCount: number;
};

type TaskContext = {
  task: {
    taskId: string;
    listId: string;
    title: string;
    status: string | null;
    importance: string | null;
    dueDateTime: string | null;
    bodyPreview: string | null;
    body: string | null;
  };
  list: {
    listId: string;
    displayName: string | null;
    wellknownListName: string | null;
  } | null;
};

function encodeAggregateCursor(offset: number): string {
  return Buffer.from(JSON.stringify({ kind: 'aggregate_tasks', offset }), 'utf8').toString(
    'base64url'
  );
}

function decodeAggregateCursor(cursor: string): number {
  try {
    const parsed = JSON.parse(Buffer.from(cursor, 'base64url').toString('utf8')) as {
      kind?: string;
      offset?: unknown;
    };
    if (parsed.kind !== 'aggregate_tasks' || typeof parsed.offset !== 'number') {
      throw new Error('invalid');
    }
    return parsed.offset;
  } catch {
    throw new GraphToolError('invalid_cursor', 'Invalid pagination cursor.');
  }
}

function taskMatchesFilters(
  entry: ReturnType<typeof normalizeTask>,
  filters: {
    status?: string;
    dueBefore?: string;
    dueAfter?: string;
    importance?: string;
    query?: string;
  }
): boolean {
  if (filters.status && entry.status !== filters.status) {
    return false;
  }
  if (filters.importance && entry.importance !== filters.importance) {
    return false;
  }
  if (filters.dueBefore && compareDates(entry.dueDateTime, filters.dueBefore) >= 0) {
    return false;
  }
  if (filters.dueAfter && compareDates(entry.dueDateTime, filters.dueAfter) <= 0) {
    return false;
  }
  if (filters.query) {
    const lowered = filters.query.toLowerCase();
    const haystack = `${entry.title} ${entry.bodyPreview ?? ''}`.toLowerCase();
    if (!haystack.includes(lowered)) {
      return false;
    }
  }
  return true;
}

async function listTaskListsInternal(
  context: CustomToolHandlerContext,
  limit: number,
  cursor?: string
): Promise<{ items: TaskListSummary[]; nextCursor?: string }> {
  const endpoint = cursor
    ? decodeCursor(cursor)
    : `/me/todo/lists?$top=${limit}&$select=id,displayName,wellknownListName`;
  const response = await graphRequest<GraphCollectionResponse<JsonRecord>>(context, endpoint);

  const items: TaskListSummary[] = [];
  for (const entry of response.value ?? []) {
    const listId = getString(entry, 'id') ?? '';
    const taskResponse = await graphRequest<GraphCollectionResponse<JsonRecord>>(
      context,
      `/me/todo/lists/${encodePathSegment(listId)}/tasks?$top=200&$select=id,status`
    );
    const tasks = taskResponse.value ?? [];
    items.push({
      listId,
      displayName: getString(entry, 'displayName'),
      wellknownListName: getString(entry, 'wellknownListName'),
      totalCount: tasks.length,
      openCount: tasks.filter((task) => getString(task, 'status') !== 'completed').length,
    });
  }

  return {
    items,
    nextCursor: response['@odata.nextLink'] ? encodeCursor(response['@odata.nextLink']) : undefined,
  };
}

async function loadTasksForList(
  context: CustomToolHandlerContext,
  listId: string,
  limit: number,
  cursor?: string
): Promise<{ items: ReturnType<typeof normalizeTask>[]; nextCursor?: string }> {
  const endpoint = cursor
    ? decodeCursor(cursor)
    : `/me/todo/lists/${encodePathSegment(listId)}/tasks?$top=${limit}&$select=id,title,status,importance,dueDateTime,body`;
  const response = await graphRequest<GraphCollectionResponse<JsonRecord>>(context, endpoint);
  return {
    items: (response.value ?? []).map((entry) => normalizeTask(entry, listId)),
    nextCursor: response['@odata.nextLink'] ? encodeCursor(response['@odata.nextLink']) : undefined,
  };
}

async function listTasksInternal(
  context: CustomToolHandlerContext,
  params: {
    listId?: string;
    limit: number;
    cursor?: string;
    status?: string;
    dueBefore?: string;
    dueAfter?: string;
    importance?: string;
    query?: string;
  }
): Promise<{ items: ReturnType<typeof normalizeTask>[]; nextCursor?: string }> {
  if (params.listId) {
    const result = await loadTasksForList(context, params.listId, Math.max(params.limit * 3, 50), params.cursor);
    return {
      items: result.items
        .filter((entry) =>
          taskMatchesFilters(entry, {
            status: params.status,
            dueBefore: params.dueBefore,
            dueAfter: params.dueAfter,
            importance: params.importance,
            query: params.query,
          })
        )
        .slice(0, params.limit),
      nextCursor: result.nextCursor,
    };
  }

  const lists = await listTaskListsInternal(context, 100);
  const collected: ReturnType<typeof normalizeTask>[] = [];
  for (const list of lists.items) {
    const result = await loadTasksForList(context, list.listId, 100);
    collected.push(...result.items);
  }

  const sorted = collected
    .filter((entry) =>
      taskMatchesFilters(entry, {
        status: params.status,
        dueBefore: params.dueBefore,
        dueAfter: params.dueAfter,
        importance: params.importance,
        query: params.query,
      })
    )
    .sort((left, right) => compareDates(left.dueDateTime, right.dueDateTime));
  const offset = params.cursor ? decodeAggregateCursor(params.cursor) : 0;
  const items = sorted.slice(offset, offset + params.limit);

  return {
    items,
    nextCursor:
      offset + params.limit < sorted.length
        ? encodeAggregateCursor(offset + params.limit)
        : undefined,
  };
}

async function getTaskContextInternal(
  context: CustomToolHandlerContext,
  taskId: string,
  listId?: string
): Promise<TaskContext> {
  if (listId) {
    const task = await graphRequest<JsonRecord>(
      context,
      `/me/todo/lists/${encodePathSegment(listId)}/tasks/${encodePathSegment(taskId)}`
    );
    const list = await graphRequest<JsonRecord>(
      context,
      `/me/todo/lists/${encodePathSegment(listId)}?$select=id,displayName,wellknownListName`
    );
    return {
      task: {
        ...normalizeTask(task, listId),
        body: getString(asRecord(task.body), 'content'),
      },
      list: {
        listId,
        displayName: getString(list, 'displayName'),
        wellknownListName: getString(list, 'wellknownListName'),
      },
    };
  }

  const lists = await listTaskListsInternal(context, 100);
  for (const list of lists.items) {
    try {
      return await getTaskContextInternal(context, taskId, list.listId);
    } catch {
      continue;
    }
  }

  throw new Error(`Task not found: ${taskId}`);
}

async function updateTaskStatus(
  context: CustomToolHandlerContext,
  listId: string,
  taskId: string,
  status: string
): Promise<JsonRecord> {
  return graphRequest<JsonRecord>(
    context,
    `/me/todo/lists/${encodePathSegment(listId)}/tasks/${encodePathSegment(taskId)}`,
    {
      method: 'PATCH',
      body: JSON.stringify({ status }),
    }
  );
}

export const taskToolDefinitions: CustomToolDefinition[] = [
  {
    name: 'list-task-lists',
    description: 'List Microsoft To Do task lists with open and total counts.',
    method: 'GET',
    path: '/me/todo/lists',
    requiresOrgMode: false,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Tasks.Read'],
    schema: {
      limit: z.number().int().min(1).max(100).default(50).describe('Maximum results'),
      cursor: z.string().optional().describe('Opaque pagination cursor'),
    },
    handler: async (params, context) => {
      const limit = getOptionalNumber(params, 'limit', 50);
      const cursor = getOptionalString(params, 'cursor');
      return success(context.graphClient, await listTaskListsInternal(context, limit, cursor));
    },
  },
  {
    name: 'list-tasks',
    description: 'List To Do tasks with Canvas-friendly filters.',
    method: 'GET',
    path: '/me/todo/lists/{listId}/tasks',
    requiresOrgMode: false,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Tasks.Read'],
    schema: {
      listId: z.string().min(1).optional().describe('Optional task list ID'),
      status: z.string().min(1).optional().describe('Optional status filter'),
      dueBefore: z.string().min(1).optional().describe('Optional due-before ISO 8601'),
      dueAfter: z.string().min(1).optional().describe('Optional due-after ISO 8601'),
      importance: z.string().min(1).optional().describe('Optional importance filter'),
      limit: z.number().int().min(1).max(100).default(25).describe('Maximum results'),
      cursor: z.string().optional().describe('Opaque pagination cursor'),
    },
    handler: async (params, context) => {
      const listId = getOptionalString(params, 'listId');
      const status = getOptionalString(params, 'status');
      const dueBefore = getOptionalString(params, 'dueBefore');
      const dueAfter = getOptionalString(params, 'dueAfter');
      const importance = getOptionalString(params, 'importance');
      const limit = getOptionalNumber(params, 'limit', 25);
      const cursor = getOptionalString(params, 'cursor');
      return success(
        context.graphClient,
        await listTasksInternal(context, {
          listId,
          status,
          dueBefore,
          dueAfter,
          importance,
          limit,
          cursor,
        })
      );
    },
  },
  {
    name: 'search-tasks',
    description: 'Search To Do tasks by title/body with optional filters.',
    method: 'GET',
    path: '/me/todo/lists/{listId}/tasks',
    requiresOrgMode: false,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Tasks.Read'],
    schema: {
      query: z.string().min(1).describe('Search query'),
      status: z.string().min(1).optional().describe('Optional status filter'),
      listId: z.string().min(1).optional().describe('Optional task list ID'),
      dueBefore: z.string().min(1).optional().describe('Optional due-before ISO 8601'),
      dueAfter: z.string().min(1).optional().describe('Optional due-after ISO 8601'),
      importance: z.string().min(1).optional().describe('Optional importance filter'),
      limit: z.number().int().min(1).max(100).default(25).describe('Maximum results'),
    },
    handler: async (params, context) => {
      const query = getRequiredString(params, 'query');
      const listId = getOptionalString(params, 'listId');
      const status = getOptionalString(params, 'status');
      const dueBefore = getOptionalString(params, 'dueBefore');
      const dueAfter = getOptionalString(params, 'dueAfter');
      const importance = getOptionalString(params, 'importance');
      const limit = getOptionalNumber(params, 'limit', 25);
      return success(
        context.graphClient,
        await listTasksInternal(context, {
          listId,
          status,
          dueBefore,
          dueAfter,
          importance,
          query,
          limit,
        })
      );
    },
  },
  {
    name: 'get-task-context',
    description: 'Get one task plus its parent task list metadata.',
    method: 'GET',
    path: '/me/todo/lists/{listId}/tasks/{taskId}',
    requiresOrgMode: false,
    readOnlyHint: true,
    openWorldHint: true,
    scopes: ['Tasks.Read'],
    schema: {
      taskId: z.string().min(1).describe('Task ID'),
      listId: z.string().min(1).optional().describe('Optional task list ID'),
    },
    handler: async (params, context) => {
      const taskId = getRequiredString(params, 'taskId');
      const listId = getOptionalString(params, 'listId');
      return success(context.graphClient, await getTaskContextInternal(context, taskId, listId));
    },
  },
  {
    name: 'complete-task',
    description: 'Mark a To Do task as completed.',
    method: 'PATCH',
    path: '/me/todo/lists/{listId}/tasks/{taskId}',
    requiresOrgMode: false,
    readOnlyHint: false,
    destructiveHint: true,
    openWorldHint: true,
    scopes: ['Tasks.ReadWrite'],
    schema: {
      listId: z.string().min(1).describe('Task list ID'),
      taskId: z.string().min(1).describe('Task ID'),
    },
    handler: async (params, context) => {
      const listId = getRequiredString(params, 'listId');
      const taskId = getRequiredString(params, 'taskId');
      await updateTaskStatus(context, listId, taskId, 'completed');
      return success(context.graphClient, { success: true, listId, taskId, status: 'completed' });
    },
  },
  {
    name: 'reopen-task',
    description: 'Reopen a completed To Do task.',
    method: 'PATCH',
    path: '/me/todo/lists/{listId}/tasks/{taskId}',
    requiresOrgMode: false,
    readOnlyHint: false,
    destructiveHint: true,
    openWorldHint: true,
    scopes: ['Tasks.ReadWrite'],
    schema: {
      listId: z.string().min(1).describe('Task list ID'),
      taskId: z.string().min(1).describe('Task ID'),
    },
    handler: async (params, context) => {
      const listId = getRequiredString(params, 'listId');
      const taskId = getRequiredString(params, 'taskId');
      await updateTaskStatus(context, listId, taskId, 'notStarted');
      return success(context.graphClient, { success: true, listId, taskId, status: 'notStarted' });
    },
  },
];
