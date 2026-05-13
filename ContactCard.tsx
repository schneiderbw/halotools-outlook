import {
  Text,
  makeStyles,
  tokens,
  Avatar,
  Badge,
} from "@fluentui/react-components";
import type { HaloUser, HaloClient } from "../types/halo";
import { domainOf, type EmailContext } from "../lib/office";
import { SearchPicker, type PickerItem } from "./SearchPicker";
import { searchClients, searchUsers } from "../lib/halo-api";

const useStyles = makeStyles({
  root: {
    display: "flex",
    gap: "12px",
    alignItems: "flex-start",
  },
  details: {
    display: "flex",
    flexDirection: "column",
    gap: "2px",
    flex: 1,
    minWidth: 0,
  },
  name: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase300,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  email: {
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  client: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorBrandForeground1,
    marginTop: "4px",
    display: "flex",
    alignItems: "center",
    gap: "8px",
    flexWrap: "wrap",
  },
  badges: {
    display: "flex",
    gap: "4px",
    marginTop: "4px",
    flexWrap: "wrap",
    alignItems: "center",
  },
  inlineEdit: {
    marginLeft: "auto",
  },
});

interface Props {
  email: EmailContext;
  contact?: HaloUser;
  client?: HaloClient;
  onContactChange: (contact: HaloUser | undefined) => void;
  onClientChange: (client: HaloClient | undefined) => void;
}

export function ContactCard({
  email,
  contact,
  client,
  onContactChange,
  onClientChange,
}: Props) {
  const styles = useStyles();
  const displayName = contact?.name || email.senderName || email.senderEmail;
  const domain = domainOf(email.senderEmail);

  return (
    <div className={styles.root}>
      <Avatar name={displayName} color="colorful" size={36} />
      <div className={styles.details}>
        <Text className={styles.name}>{displayName}</Text>
        <Text className={styles.email}>{email.senderEmail}</Text>
        {client && (
          <div className={styles.client}>
            <Text>{client.name}</Text>
          </div>
        )}
        <div className={styles.badges}>
          {contact ? (
            <Badge appearance="filled" color="success" size="small">
              Contact matched
            </Badge>
          ) : client ? (
            <Badge appearance="filled" color="warning" size="small">
              Domain match only
            </Badge>
          ) : (
            <Badge appearance="filled" color="danger" size="small">
              No match
            </Badge>
          )}

          <SearchPicker<HaloUser>
            triggerLabel={contact ? "Change contact" : "Find contact"}
            title="Find contact in HaloPSA"
            initialQuery={email.senderEmail}
            onSearch={async (q) => {
              const users = await searchUsers(q);
              return users.map<PickerItem<HaloUser>>((u) => ({
                key: u.id,
                primary: u.name,
                secondary: [u.emailaddress, u.client_name].filter(Boolean).join(" · "),
                value: u,
              }));
            }}
            onPick={(u) => {
              onContactChange(u);
              // Picking a contact also implies their client
              if (u.client_id && u.client_name) {
                onClientChange({ id: u.client_id, name: u.client_name });
              }
            }}
          />

          <SearchPicker<HaloClient>
            triggerLabel={client ? "Change client" : "Find client"}
            title="Find client in HaloPSA"
            initialQuery={domain}
            onSearch={async (q) => {
              const clients = await searchClients(q);
              return clients.map<PickerItem<HaloClient>>((c) => ({
                key: c.id,
                primary: c.name,
                value: c,
              }));
            }}
            onPick={(c) => {
              onClientChange(c);
              // Clear contact if it doesn't belong to the new client
              if (contact && contact.client_id && contact.client_id !== c.id) {
                onContactChange(undefined);
              }
            }}
          />
        </div>
      </div>
    </div>
  );
}
