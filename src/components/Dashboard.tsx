import { useEffect, useState, useCallback } from "react";
import {
  Text,
  Spinner,
  makeStyles,
  tokens,
  Divider,
  Button,
  MessageBar,
  MessageBarBody,
  Menu,
  MenuTrigger,
  MenuList,
  MenuItem,
  MenuPopover,
} from "@fluentui/react-components";
import { MoreHorizontal24Regular, ArrowClockwise24Regular } from "@fluentui/react-icons";
import { ContactCard } from "./ContactCard";
import { TicketList } from "./TicketList";
import { LogActions } from "./LogActions";
import { SettingsScreen } from "./SettingsScreen";
import {
  findUserByEmail,
  findClientByDomain,
  listOpenTicketsForClient,
  findTicketsForEmail,
} from "../lib/halo-api";
import { signOut } from "../lib/auth";
import { clearConfig } from "../lib/config";
import { domainOf, type EmailContext } from "../lib/office";
import type { HaloUser, HaloClient, HaloTicket } from "../types/halo";

interface Props {
  email: EmailContext;
  onSignedOut: () => void;
}

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    height: "100%",
    overflowY: "auto",
  },
  header: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    padding: "8px 12px",
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  body: {
    display: "flex",
    flexDirection: "column",
    gap: "12px",
    padding: "12px",
    flex: 1,
  },
  loading: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    flex: 1,
    gap: "8px",
  },
  brand: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase300,
  },
});

export function Dashboard({ email, onSignedOut }: Props) {
  const styles = useStyles();
  const [contact, setContact] = useState<HaloUser | undefined>();
  const [client, setClient] = useState<HaloClient | undefined>();
  const [openTickets, setOpenTickets] = useState<HaloTicket[]>([]);
  const [threadTickets, setThreadTickets] = useState<HaloTicket[]>([]);
  const [error, setError] = useState<string | undefined>();
  const [loadingResolve, setLoadingResolve] = useState(true);
  const [loadingTickets, setLoadingTickets] = useState(false);
  const [settingsOpen, setSettingsOpen] = useState(false);

  // Auto-resolve sender → contact + client whenever the open email changes.
  // Refresh button bumps `refreshTick` to trigger a re-run.
  const [refreshTick, setRefreshTick] = useState(0);

  useEffect(() => {
    let cancelled = false;
    setLoadingResolve(true);
    setError(undefined);
    setContact(undefined);
    setClient(undefined);
    setOpenTickets([]);
    setThreadTickets([]);

    (async () => {
      try {
        const matchedContact = await findUserByEmail(email.senderEmail);
        if (cancelled) return;

        let matchedClient: HaloClient | undefined;
        if (matchedContact?.client_id) {
          matchedClient = {
            id: matchedContact.client_id,
            name: matchedContact.client_name ?? "",
          };
        } else {
          const domain = domainOf(email.senderEmail);
          if (domain) matchedClient = await findClientByDomain(domain);
        }
        if (cancelled) return;

        setContact(matchedContact);
        setClient(matchedClient);

        const threadIds = [
          email.internetMessageId,
          email.inReplyTo,
          ...email.references,
        ].filter((id): id is string => !!id);
        if (threadIds.length > 0) {
          findTicketsForEmail(threadIds)
            .then((t) => !cancelled && setThreadTickets(t))
            .catch(() => {
              /* swallow — non-fatal */
            });
        }
      } catch (e) {
        if (!cancelled) setError((e as Error).message);
      } finally {
        if (!cancelled) setLoadingResolve(false);
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [email.senderEmail, email.internetMessageId, refreshTick]);

  // Refetch open tickets whenever the active client changes (auto or override)
  useEffect(() => {
    if (!client) {
      setOpenTickets([]);
      return;
    }
    let cancelled = false;
    setLoadingTickets(true);
    listOpenTicketsForClient(client.id)
      .then((t) => !cancelled && setOpenTickets(t))
      .catch((e) => !cancelled && setError((e as Error).message))
      .finally(() => !cancelled && setLoadingTickets(false));
    return () => {
      cancelled = true;
    };
  }, [client?.id]);

  const handleTicketUpdated = useCallback(
    async (updated: HaloTicket) => {
      // Optimistic local update for instant feedback
      setOpenTickets((prev) =>
        prev.map((t) => (t.id === updated.id ? { ...t, ...updated } : t)),
      );
      setThreadTickets((prev) =>
        prev.map((t) => (t.id === updated.id ? { ...t, ...updated } : t)),
      );
      // Then refetch open tickets so status changes (e.g., closed) remove tickets from the list
      if (client) {
        try {
          const fresh = await listOpenTicketsForClient(client.id);
          setOpenTickets(fresh);
        } catch {
          /* keep optimistic state */
        }
      }
    },
    [client],
  );

  const handleSignOut = useCallback(async () => {
    await signOut();
    onSignedOut();
  }, [onSignedOut]);

  const handleReconfigure = useCallback(async () => {
    await clearConfig();
    onSignedOut();
  }, [onSignedOut]);

  if (settingsOpen) {
    return (
      <div className={styles.root}>
        <SettingsScreen
          onClose={() => setSettingsOpen(false)}
          onSignOut={handleSignOut}
          onReconfigure={handleReconfigure}
        />
      </div>
    );
  }

  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <Text className={styles.brand}>HaloPSA</Text>
        <div style={{ display: "flex", gap: 4 }}>
          <Button
            appearance="subtle"
            size="small"
            icon={<ArrowClockwise24Regular />}
            onClick={() => setRefreshTick((n) => n + 1)}
            aria-label="Refresh"
            disabled={loadingResolve}
          />
          <Menu>
            <MenuTrigger disableButtonEnhancement>
              <Button
                appearance="subtle"
                size="small"
                icon={<MoreHorizontal24Regular />}
                aria-label="Menu"
              />
            </MenuTrigger>
            <MenuPopover>
              <MenuList>
                <MenuItem onClick={() => setSettingsOpen(true)}>Settings</MenuItem>
                <MenuItem onClick={handleSignOut}>Sign out</MenuItem>
                <MenuItem onClick={handleReconfigure}>Switch HaloPSA tenant</MenuItem>
              </MenuList>
            </MenuPopover>
          </Menu>
        </div>
      </div>

      {loadingResolve && (
        <div className={styles.loading}>
          <Spinner size="small" />
          <Text size={200}>Looking up {email.senderEmail}…</Text>
        </div>
      )}

      {!loadingResolve && error && (
        <div className={styles.body}>
          <MessageBar intent="error">
            <MessageBarBody>{error}</MessageBarBody>
          </MessageBar>
        </div>
      )}

      {!loadingResolve && !error && (
        <div className={styles.body}>
          <ContactCard
            email={email}
            contact={contact}
            client={client}
            onContactChange={setContact}
            onClientChange={setClient}
          />
          <Divider />

          {threadTickets.length > 0 && (
            <TicketList
              label="This conversation"
              tickets={threadTickets}
              onTicketUpdated={handleTicketUpdated}
            />
          )}

          {loadingTickets ? (
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <Spinner size="extra-tiny" /> <Text size={200}>Loading tickets…</Text>
            </div>
          ) : (
            <TicketList
              label={threadTickets.length > 0 ? "Other open tickets" : "Open tickets"}
              tickets={openTickets.filter(
                (t) => !threadTickets.find((tt) => tt.id === t.id),
              )}
              onTicketUpdated={handleTicketUpdated}
            />
          )}

          <Divider />
          <LogActions
            email={email}
            client={client}
            contact={contact}
            candidateTickets={[
              ...threadTickets,
              ...openTickets.filter((t) => !threadTickets.find((tt) => tt.id === t.id)),
            ]}
            preferAppend={threadTickets.length > 0}
          />
        </div>
      )}
    </div>
  );
}
