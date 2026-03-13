import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import GraphClient from './graph-client.js';
import AuthManager from './auth.js';
import {
  CustomToolContext,
  CustomToolDefinition,
  ToolRegistryEntry,
  buildCustomToolHandler,
  buildSchemaWithCommonParams,
} from './custom-tools/shared.js';
import { peopleToolDefinitions } from './custom-tools/people.js';
import { teamsToolDefinitions } from './custom-tools/teams.js';
import { fileToolDefinitions } from './custom-tools/files.js';
import { sharepointToolDefinitions } from './custom-tools/sharepoint.js';
import { calendarToolDefinitions } from './custom-tools/calendar.js';
import { taskToolDefinitions } from './custom-tools/tasks.js';
import { mailToolDefinitions } from './custom-tools/mail.js';

const customToolDefinitions: CustomToolDefinition[] = [
  ...peopleToolDefinitions,
  ...teamsToolDefinitions,
  ...fileToolDefinitions,
  ...sharepointToolDefinitions,
  ...calendarToolDefinitions,
  ...taskToolDefinitions,
  ...mailToolDefinitions,
];

export function getCustomToolDefinitions(): readonly CustomToolDefinition[] {
  return customToolDefinitions;
}

export function buildCustomToolRegistry(
  graphClient: GraphClient,
  authManager?: AuthManager,
  _multiAccount: boolean = false,
  _accountNames: string[] = []
): Map<string, ToolRegistryEntry> {
  const registry = new Map<string, ToolRegistryEntry>();
  const context: CustomToolContext = { graphClient, authManager };

  for (const definition of customToolDefinitions) {
    registry.set(definition.name, {
      name: definition.name,
      description: definition.description,
      method: definition.method,
      path: definition.path,
      requiresOrgMode: definition.requiresOrgMode,
      readOnlyHint: definition.readOnlyHint,
      openWorldHint: definition.openWorldHint,
      destructiveHint: definition.destructiveHint,
      handler: buildCustomToolHandler(definition, context),
    });
  }

  return registry;
}

export function registerCustomTools(
  server: McpServer,
  graphClient: GraphClient,
  enabledToolsRegex: RegExp | undefined,
  readOnly: boolean = false,
  orgMode: boolean = false,
  authManager?: AuthManager,
  multiAccount: boolean = false,
  accountNames: string[] = []
): number {
  const context: CustomToolContext = { graphClient, authManager };
  let registeredCount = 0;

  for (const definition of customToolDefinitions) {
    if (definition.requiresOrgMode && !orgMode) {
      continue;
    }

    if (readOnly && !definition.readOnlyHint) {
      continue;
    }

    if (enabledToolsRegex && !enabledToolsRegex.test(definition.name)) {
      continue;
    }

    const schema = buildSchemaWithCommonParams(definition.schema, multiAccount, accountNames);
    server.tool(
      definition.name,
      definition.description,
      schema,
      {
        title: definition.name,
        readOnlyHint: definition.readOnlyHint,
        destructiveHint: definition.destructiveHint,
        openWorldHint: definition.openWorldHint,
      },
      buildCustomToolHandler(definition, context)
    );
    registeredCount++;
  }

  return registeredCount;
}
