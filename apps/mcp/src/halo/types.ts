// Minimal HaloPSA response/payload shapes used by the MCP server.
// Mirrors src/types/halo.ts in the parent repo, trimmed to what the tools touch.

export interface HaloUser {
  id: number;
  name: string;
  emailaddress?: string;
  client_id?: number;
  client_name?: string;
  site_id?: number;
  site_name?: string;
  inactive?: boolean;
  phonenumber?: string;
  mobile_number?: string;
  jobtitle?: string;
}

export interface HaloClient {
  id: number;
  name: string;
  inactive?: boolean;
  accountmanager_name?: string;
  accountmanager_id?: number;
  main_site_id?: number;
  main_site_name?: string;
}

export interface HaloTicket {
  id: number;
  summary: string;
  details?: string;
  status_id: number;
  statusname?: string;
  client_id?: number;
  client_name?: string;
  user_id?: number;
  user_name?: string;
  agent_id?: number;
  agent_name?: string;
  agentname?: string;
  assignedagent_id?: number;
  assignedagent_name?: string;
  agent?: { id?: number; name?: string };
  priority_id?: number;
  priorityname?: string;
  dateoccurred?: string;
  dateopened?: string;
  targetdate?: string;
}

export interface HaloAction {
  id: number;
  ticket_id: number;
  outcome: string;
  note: string;
  who?: string;
  datetime?: string;
  actionnumber?: number;
}

export interface CreateActionPayload {
  ticket_id: number;
  outcome: string;
  note: string;
  hiddenfromuser?: boolean;
  time_taken?: number;
  internetmessageid?: string;
  inreplyto?: string;
  references?: string;
}

export interface CreateTicketPayload {
  summary: string;
  details: string;
  client_id?: number;
  user_id?: number;
  site_id?: number;
  tickettype_id?: number;
  agent_id?: number;
  priority_id?: number;
}

export interface HaloCannedText {
  id: number;
  name: string;
  group_id: number;
  text?: string;
  html?: string;
  restriction_type?: number;
  is_favourite?: boolean;
}

export interface HaloCannedTextGroup {
  id: number;
  name: string;
  valueint1?: number;
  sequence?: number;
}

export interface HaloCRMNote {
  id: number;
  client_id?: number;
  site_id?: number;
  user_id?: number;
  datetime: string;
  who_agentid?: number;
  subject?: string;
  note: string;
  timetaken?: number;
  ticketid?: number;
}

export interface CreateCRMNotePayload {
  client_id?: number | string;
  site_id?: number | string;
  user_id?: number | string;
  subject: string;
  note: string;
  timetaken?: number;
  hide_time_taken?: boolean;
}

export interface HaloFeedItem {
  id: number;
  datetime: string;
  entitytype: number;
  agent_id?: number;
  user_id?: number;
  note?: string;
  outcome?: string;
  who_name?: string;
  content_id1?: number;
  content_id2?: number;
}

export interface HaloFeedResponse {
  record_count: number;
  feed: HaloFeedItem[];
}

export interface CRMScope {
  client_id?: number;
  site_id?: number;
  user_id?: number;
}
