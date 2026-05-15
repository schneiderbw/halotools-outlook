// Minimal shapes for the HaloPSA endpoints we use.
//
// COPIED from ../../../src/types/halo.ts at extension scaffold time. Keep
// these in sync manually — if upstream changes a shape, mirror it here.
// We don't import from the Outlook side because the extension is built as
// a standalone package with its own dependency tree.

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
  tags?: Array<{ value: string }>;
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
  priority_id?: number;
  priorityname?: string;
  dateoccurred?: string;
  dateopened?: string;
}
