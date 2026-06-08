import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { createTicket } from "@iusehalo/halo-api";

const inputSchema = {
  summary: z.string().min(1).describe("Short, single-line ticket subject."),
  details: z
    .string()
    .min(1)
    .describe(
      "Full ticket body — HTML is accepted. Include reproduction steps, context, error messages.",
    ),
  client_id: z
    .number()
    .int()
    .positive()
    .describe("Halo client (company) ID this ticket belongs to."),
  user_id: z
    .number()
    .int()
    .positive()
    .optional()
    .describe("Halo user (contact) ID who reported the issue."),
  tickettype_id: z
    .number()
    .int()
    .positive()
    .optional()
    .describe("Halo ticket type ID. Omit to use the tenant default."),
};

export function registerCreateTicket(server: McpServer): void {
  server.registerTool(
    "createTicket",
    {
      title: "Create a HaloPSA ticket",
      description:
        "Create a new ticket in HaloPSA against a specific client (and optionally a specific contact). Use this when the user has confirmed they want a new ticket logged — don't auto-create on every mention of a problem. Returns the new ticket's id and summary.",
      inputSchema,
    },
    async ({ summary, details, client_id, user_id, tickettype_id }) => {
      const ticket = await createTicket({
        summary,
        details,
        client_id,
        user_id,
        tickettype_id,
      });
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(
              {
                id: ticket.id,
                summary: ticket.summary,
                status: ticket.statusname,
                client_id: ticket.client_id,
                user_id: ticket.user_id,
              },
              null,
              2,
            ),
          },
        ],
      };
    },
  );
}
