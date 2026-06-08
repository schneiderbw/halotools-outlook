import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { listCannedTextGroups, searchCannedText } from "@iusehalo/halo-api";

const inputSchema = {
  query: z
    .string()
    .describe(
      "Free-text query matched against canned-text name and body. Empty string returns the first 50 entries in the (optional) group.",
    ),
  group: z
    .string()
    .optional()
    .describe(
      "Optional canned-text group name (e.g. 'Onboarding', 'Sales replies'). Resolved to a group_id via /Lookup. Case-insensitive substring match.",
    ),
};

export function registerSearchCannedText(server: McpServer): void {
  server.registerTool(
    "searchCannedText",
    {
      title: "Search HaloPSA canned text snippets",
      description:
        "Search Halo's saved canned-text snippets — short, pre-written responses, email templates, and boilerplate. Optionally scope the search to a named group. Returns the snippet name and body so you can quote or paste it.",
      inputSchema,
    },
    async ({ query, group }) => {
      let groupId: number | undefined;
      if (group) {
        const groups = await listCannedTextGroups();
        const needle = group.trim().toLowerCase();
        const match = groups.find((g) => g.name?.toLowerCase().includes(needle));
        if (!match) {
          return {
            content: [
              {
                type: "text",
                text: `No canned-text group matches '${group}'. Available: ${groups
                  .map((g) => g.name)
                  .filter(Boolean)
                  .join(", ")}`,
              },
            ],
          };
        }
        groupId = match.id;
      }
      const results = await searchCannedText(query, groupId);
      const rows = results.map((c) => ({
        id: c.id,
        name: c.name,
        group_id: c.group_id,
        body: c.text ?? c.html,
      }));
      return {
        content: [{ type: "text", text: JSON.stringify(rows, null, 2) }],
      };
    },
  );
}
