// Minimal shapes for the HaloPSA endpoints we use.

export interface HaloUser {
  id: number;
  name: string;
  emailaddress?: string;
  client_id?: number;
  client_name?: string;
  site_id?: number;
  site_name?: string;
  inactive?: boolean;
}

export interface HaloClient {
  id: number;
  name: string;
  inactive?: boolean;
  client_email_address_domain?: string;
}

export interface HaloTicketType {
  id: number;
  name: string;
  use: string;
  inactive?: boolean;
}

export interface HaloStatus {
  id: number;
  name: string;
  type?: string;
  /** When true, this status counts as "ticket is closed" for Halo's reporting */
  isclosed?: boolean;
  inactive?: boolean;
}

export interface HaloAgent {
  id: number;
  name: string;
  email?: string;
  inactive?: boolean;
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
  tickettype_id?: number;
  category_1?: string;
  dateoccurred?: string;
  dateopened?: string;
  customfields?: Array<{ name: string; value: unknown }>;
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

export interface HaloAttachmentInline {
  filename: string;
  data_base64: string;
  contenttype?: string;
  isimage?: boolean;
}

export interface CreateActionPayload {
  ticket_id: number;
  outcome: string;
  note: string;
  hiddenfromuser?: boolean;
  attachments?: HaloAttachmentInline[];
  emailfrom?: string;
  emailfromname?: string;
  emailsubject?: string;
  /** Decimal hours spent on this action (e.g., 0.25 for 15 minutes). */
  time_taken?: number;
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
  category_1?: string;
  attachments?: HaloAttachmentInline[];
  customfields?: Array<{ name: string; value: string | number | boolean }>;
}

/** Partial update payload for an existing ticket. Halo accepts mutated fields only. */
export interface UpdateTicketPayload {
  id: number;
  status_id?: number;
  agent_id?: number;
  priority_id?: number;
  customfields?: Array<{ name: string; value: string | number | boolean }>;
}

/**
 * Knowledge base article shape from /KBArticle.
 * Body lives in `faq_answer` on most tenants; some older tenants surface it under `details`.
 * Callers should try faq_answer first, then fall back to details.
 */
export interface HaloKbArticle {
  id: number;
  name: string;
  faq_answer?: string;
  details?: string;
  tags?: Array<{ value: string }>;
}
