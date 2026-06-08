import { z } from "zod";
import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { findUserByEmail, getOpenTicketCount } from "@iusehalo/halo-api";

const inputSchema = {
  email: z
    .string()
    .email()
    .describe("Email address to look up in HaloPSA's contact (user) directory."),
};

export function registerFindContact(server: McpServer): void {
  server.registerTool(
    "findContact",
    {
      title: "Find HaloPSA contact by email",
      description:
        "Find the HaloPSA contact (user) for an email address. Returns the contact's name, associated client, phone, job title, and a count of their open tickets. Use this first when the user mentions someone by email to find their context in Halo.",
      inputSchema,
    },
    async ({ email }) => {
      const user = await findUserByEmail(email);
      if (!user) {
        return {
          content: [{ type: "text", text: `No contact found in Halo for ${email}.` }],
        };
      }
      const openTicketCount = await getOpenTicketCount(user.id);
      const summary = {
        id: user.id,
        name: user.name,
        email: user.emailaddress,
        jobtitle: user.jobtitle,
        phone: user.phonenumber ?? user.mobile_number,
        client_id: user.client_id,
        client_name: user.client_name,
        site_id: user.site_id,
        site_name: user.site_name,
        openTicketCount,
      };
      return {
        content: [{ type: "text", text: JSON.stringify(summary, null, 2) }],
      };
    },
  );
}
