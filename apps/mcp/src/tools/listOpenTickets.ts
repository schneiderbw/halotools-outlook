import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import {
  listOpenTicketsForClient,
  listOpenTicketsForUser,
  type HaloTicket,
} from "@iusehalo/halo-api";

const inputSchema = {
  client_id: z
    .number()
    .int()
    .positive()
    .optional()
    .describe("Halo client (company) ID. Pass to list all open tickets for that company."),
  user_id: z
    .number()
    .int()
    .positive()
    .optional()
    .describe("Halo user (contact) ID. Pass to list a single contact's open tickets."),
};

function agentNameOf(t: HaloTicket): string | undefined {
  return (
    t.agent_name ??
    t.agentname ??
    t.assignedagent_name ??
    t.agent?.name ??
    undefined
  );
}

export function registerListOpenTickets(server: McpServer): void {
  server.registerTool(
    "listOpenTickets",
    {
      title: "List open HaloPSA tickets",
      description:
        "List currently-open tickets for a given Halo client (company) and/or user (contact). Provide client_id to see all open tickets for a company, or user_id to see only one contact's open tickets. Returns each ticket's id, summary, status, priority, assigned agent, and date opened.",
      inputSchema,
    },
    async ({ client_id, user_id }) => {
      if (!client_id && !user_id) {
        return {
          isError: true,
          content: [
            {
              type: "text",
              text: "Provide at least one of client_id or user_id.",
            },
          ],
        };
      }
      const tickets: HaloTicket[] = user_id
        ? await listOpenTicketsForUser(user_id)
        : await listOpenTicketsForClient(client_id!);

      const filtered =
        user_id && client_id
          ? tickets.filter((t) => t.client_id === client_id)
          : tickets;

      const rows = filtered.map((t) => ({
        id: t.id,
        summary: t.summary,
        status: t.statusname,
        priority: t.priorityname,
        agent: agentNameOf(t),
        client: t.client_name,
        user: t.user_name,
        dateopened: t.dateopened,
      }));

      return {
        content: [{ type: "text", text: JSON.stringify(rows, null, 2) }],
      };
    },
  );
}
