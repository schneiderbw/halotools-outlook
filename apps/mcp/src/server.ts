import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { registerAllTools } from "./tools/index.js";

export function createHaloMcpServer(): McpServer {
  const server = new McpServer(
    {
      name: "halo-mcp-server",
      version: "0.1.0",
    },
    {
      capabilities: {
        tools: {},
      },
      instructions:
        "HaloPSA tools for an MSP. Use findContact when an email is mentioned to anchor context to a Halo client/user. Use searchTickets for free-text lookup, listOpenTickets for a known client/user, appendActionToTicket to log work on an existing ticket, createTicket to open a new one, logNote for non-ticket CRM activity. searchCannedText returns saved snippets the agent can paste; getActivityFeed gives a merged timeline of an account.",
    },
  );

  registerAllTools(server);
  return server;
}
