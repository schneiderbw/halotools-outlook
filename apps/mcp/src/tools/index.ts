import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";

import { registerFindContact } from "./findContact.js";
import { registerListOpenTickets } from "./listOpenTickets.js";
import { registerSearchTickets } from "./searchTickets.js";
import { registerCreateTicket } from "./createTicket.js";
import { registerAppendActionToTicket } from "./appendActionToTicket.js";
import { registerLogNote } from "./logNote.js";
import { registerSearchCannedText } from "./searchCannedText.js";
import { registerGetActivityFeed } from "./getActivityFeed.js";

export function registerAllTools(server: McpServer): void {
  registerFindContact(server);
  registerListOpenTickets(server);
  registerSearchTickets(server);
  registerCreateTicket(server);
  registerAppendActionToTicket(server);
  registerLogNote(server);
  registerSearchCannedText(server);
  registerGetActivityFeed(server);
}
