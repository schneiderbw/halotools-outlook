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
  phonenumber?: string;
  mobile_number?: string;
  jobtitle?: string;
  tags?: Array<{ value: string }>;
}

export interface HaloClient {
  id: number;
  name: string;
  inactive?: boolean;
  client_email_address_domain?: string;
  accountmanager_name?: string;
  accountmanager_id?: number;
  /** Default site Halo associates new users with. Required to create a contact. */
  main_site_id?: number;
  main_site_name?: string;
  tags?: Array<{ value: string }>;
}

export interface CreateContactPayload {
  name: string;
  emailaddress: string;
  client_id?: number;
  phonenumber?: string;
  site_id?: number;
}

export interface HaloTicketType {
  id: number;
  name: string;
  /** Surface where this type is usable: "tickets" / "opps" / "projects" (plural). */
  use: string;
  inactive?: boolean;
  /** Whether agents are allowed to pick this type when creating. False = hidden from the agent picker. */
  agentscanselect?: boolean;
  enduserscanselect?: boolean;
  anonymouscanselect?: boolean;
  /** Visible at all in any picker. Some types are flagged invisible without being inactive. */
  visible?: boolean;
}

export interface HaloStatus {
  id: number;
  name: string;
  /** Status category code (0 = regular, 1 = order, 2 = item, 3 = special). NOT a label string. */
  type?: number;
  /** Hex colour Halo assigns to the status — what we render on the pill. */
  colour?: string;
  /** SLA behaviour: "removehold" | "hold" | "none". Useful as a heuristic for closed-ness. */
  slaaction?: string;
  inactive?: boolean;
}

export interface HaloAgent {
  id: number;
  name: string;
  email?: string;
  inactive?: boolean;
}

export interface HaloPriority {
  /** Halo's GUID identifier. NOT the value you compare to ticket.priority_id. */
  id: string;
  name: string;
  /** Numeric priority ID — THIS is what ticket.priority_id references. */
  priorityid: number;
  colour?: string;
  inactive?: boolean;
  /**
   * SLA scoping (Halo's response uses `slaid`, no underscore). A priority
   * without slaid is global; otherwise it only applies to tickets on that SLA.
   */
  slaid?: number;
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
  // Halo's REST shape for the assigned agent is inconsistent across versions:
  // some tenants return agent_id/agent_name, others agentname or assignedagent_*,
  // and the nested includedetails response uses `agent: { id, name }`.
  agent_id?: number;
  agent_name?: string;
  agentname?: string;
  assignedagent_id?: number;
  assignedagent_name?: string;
  agent?: { id?: number; name?: string };
  priority_id?: number;
  priorityname?: string;
  sla_id?: number;
  tickettype_id?: number;
  category_1?: string;
  dateoccurred?: string;
  dateopened?: string;
  /** ISO datetime. Halo's actual field name is `targetdate` (no underscore). */
  targetdate?: string;
  /** Some Halo versions expose a hard deadline separately. Empty/zero-date when unset. */
  deadlinedate?: string;
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
  /** RFC 5322 Message-ID of the source email — Halo threads on this natively. */
  internetmessageid?: string;
  /** Parent's Message-ID from the In-Reply-To header. */
  inreplyto?: string;
  /** Space-separated ancestor Message-IDs from the References header. */
  references?: string;
  /** Halo customer ("user") this action is on — sets the action to be from-customer. */
  user_id?: number;
  /** Some Halo versions require this explicit field instead of (or alongside) user_id. */
  actionby_user_id?: number;
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
  /** Email source fields — when present, Halo creates the initial action as an email
   *  and stamps internetmessageid on it, enabling native RFC-based threading. */
  emailfrom?: string;
  emailfromname?: string;
  emailsubject?: string;
  internetmessageid?: string;
  inreplyto?: string;
  references?: string;
}

/** Partial update payload for an existing ticket. Halo accepts mutated fields only. */
export interface UpdateTicketPayload {
  id: number;
  status_id?: number;
  agent_id?: number;
  priority_id?: number;
  customfields?: Array<{ name: string; value: string | number | boolean }>;
  /** ISO datetime for the ticket target / due date. Halo expects this exact field name on writes. */
  targetdate?: string;
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
