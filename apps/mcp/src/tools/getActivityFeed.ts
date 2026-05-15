import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { listFeed } from "../halo/client.js";
import { getHaloAuth } from "../halo/context.js";

const inputSchema = {
  client_id: z
    .number()
    .int()
    .positive()
    .optional()
    .describe("Halo client (company) ID — fetches feed for everything related to this company."),
  user_id: z
    .number()
    .int()
    .positive()
    .optional()
    .describe("Halo user (contact) ID — fetches feed for this contact's activity."),
  site_id: z
    .number()
    .int()
    .positive()
    .optional()
    .describe("Halo site ID — fetches feed for this client location."),
  count: z
    .number()
    .int()
    .positive()
    .max(200)
    .optional()
    .describe("Maximum number of feed items to return. Default 20."),
};

export function registerGetActivityFeed(server: McpServer): void {
  server.registerTool(
    "getActivityFeed",
    {
      title: "Get HaloPSA activity feed for a client / user / site",
      description:
        "Return Halo's merged activity feed (ticket actions, CRM notes, status changes) for a given client, contact, or site. Use this to summarize recent interactions or understand what's been happening with an account.",
      inputSchema,
    },
    async ({ client_id, user_id, site_id, count }) => {
      if (!client_id && !user_id && !site_id) {
        return {
          isError: true,
          content: [
            {
              type: "text",
              text: "Provide at least one of client_id, user_id, or site_id.",
            },
          ],
        };
      }
      const auth = getHaloAuth();
      const feed = await listFeed(
        auth,
        { client_id, user_id, site_id },
        count ?? 20,
      );
      const rows = feed.map((f) => ({
        id: f.id,
        datetime: f.datetime,
        who: f.who_name,
        outcome: f.outcome,
        note: f.note,
        ticket_id_hint: f.content_id1,
      }));
      return {
        content: [{ type: "text", text: JSON.stringify(rows, null, 2) }],
      };
    },
  );
}
