import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { createCRMNote } from "@iusehalo/halo-api";

const inputSchema = {
  subject: z.string().min(1).describe("Short note subject / title."),
  note: z.string().min(1).describe("Note body — HTML accepted."),
  client_id: z
    .number()
    .int()
    .positive()
    .optional()
    .describe("Attach the note to this Halo client (company)."),
  user_id: z
    .number()
    .int()
    .positive()
    .optional()
    .describe("Attach the note to this Halo user (contact)."),
  site_id: z
    .number()
    .int()
    .positive()
    .optional()
    .describe("Attach the note to this Halo site (client location)."),
  timetaken: z
    .number()
    .nonnegative()
    .optional()
    .describe("Time spent in decimal hours (e.g. 0.0333 = 2 minutes)."),
  hide_time_taken: z.boolean().optional().describe("Hide time recorded from the customer."),
};

export function registerLogNote(server: McpServer): void {
  server.registerTool(
    "logNote",
    {
      title: "Log a CRM note against a Halo client / contact / site",
      description:
        "Create a CRM note attached to a Halo client, user, or site (exactly one scope). Useful for recording out-of-ticket activity: a discovery call, an account-management touch, a vendor update. Don't use this for ticket work — use appendActionToTicket for that.",
      inputSchema,
    },
    async ({ subject, note, client_id, user_id, site_id, timetaken, hide_time_taken }) => {
      const scopes = [client_id, user_id, site_id].filter((v) => v !== undefined);
      if (scopes.length !== 1) {
        return {
          isError: true,
          content: [
            {
              type: "text",
              text: "Provide exactly one of client_id, user_id, or site_id.",
            },
          ],
        };
      }
      const created = await createCRMNote({
        subject,
        note,
        client_id,
        user_id,
        site_id,
        timetaken,
        hide_time_taken,
      });
      return {
        content: [
          {
            type: "text",
            text: JSON.stringify(
              {
                id: created.id,
                client_id: created.client_id,
                user_id: created.user_id,
                site_id: created.site_id,
                subject: created.subject,
                datetime: created.datetime,
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
