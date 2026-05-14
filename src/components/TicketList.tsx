import { useEffect, useState } from "react";
import {
  Text,
  makeStyles,
  tokens,
  Badge,
  Menu,
  MenuTrigger,
  MenuList,
  MenuItem,
  MenuPopover,
  MenuButton,
  Spinner,
  MenuDivider,
} from "@fluentui/react-components";
import {
  Open16Regular,
  MoreVertical16Regular,
  CheckmarkCircle16Regular,
  PersonAdd16Regular,
  ArrowUpRight16Regular,
} from "@fluentui/react-icons";
import type { HaloTicket, HaloStatus, HaloAgent } from "../types/halo";
import { getConfig } from "../lib/config";
import { listStatuses, getCurrentAgent, updateTicket } from "../lib/halo-api";
import { getCurrentUserEmail } from "../lib/office";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    gap: "6px",
  },
  label: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground2,
    textTransform: "uppercase",
    letterSpacing: "0.04em",
  },
  empty: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    fontStyle: "italic",
  },
  ticket: {
    display: "flex",
    alignItems: "center",
    gap: "4px",
    padding: "8px",
    borderRadius: tokens.borderRadiusMedium,
    backgroundColor: tokens.colorNeutralBackground2,
    transition: "background-color 80ms",
    ":hover": {
      backgroundColor: tokens.colorNeutralBackground2Hover,
    },
  },
  ticketBody: {
    flex: 1,
    minWidth: 0,
    display: "flex",
    flexDirection: "column",
    gap: "2px",
    cursor: "pointer",
  },
  ticketTitle: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  ticketMeta: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
    display: "flex",
    alignItems: "center",
    gap: "4px",
    flexWrap: "wrap",
  },
  busyOverlay: {
    display: "flex",
    alignItems: "center",
    paddingRight: "4px",
  },
});

interface Props {
  label: string;
  tickets: HaloTicket[];
  onTicketUpdated?: (updated: HaloTicket) => void;
}

export function TicketList({ label, tickets, onTicketUpdated }: Props) {
  const styles = useStyles();
  const cfg = getConfig();
  const haloUrl = cfg?.haloBaseUrl;
  const [statuses, setStatuses] = useState<HaloStatus[]>([]);
  const [currentAgent, setCurrentAgent] = useState<HaloAgent | undefined>();
  const [busyTicketId, setBusyTicketId] = useState<number | undefined>();
  const [actionError, setActionError] = useState<string | undefined>();

  // Lazy-load reference data the first time any quick-action menu opens on this list.
  // For simplicity here we just load on mount.
  useEffect(() => {
    listStatuses().then(setStatuses).catch(() => {});
    const email = getCurrentUserEmail();
    if (email) getCurrentAgent(email).then(setCurrentAgent).catch(() => {});
  }, []);

  const openInHalo = (ticketId: number) => {
    if (!haloUrl) return;
    Office.context.ui.openBrowserWindow(`${haloUrl}/agent?showmenu=false&id=${ticketId}`);
  };

  const apply = async (
    ticket: HaloTicket,
    partial: { status_id?: number; agent_id?: number; priority_id?: number },
  ) => {
    setBusyTicketId(ticket.id);
    setActionError(undefined);
    try {
      const updated = await updateTicket({ id: ticket.id, ...partial });
      onTicketUpdated?.(updated);
    } catch (e) {
      setActionError(`Update failed: ${(e as Error).message}`);
    } finally {
      setBusyTicketId(undefined);
    }
  };

  return (
    <div className={styles.root}>
      <Text className={styles.label}>{label}</Text>
      {actionError && (
        <Text style={{ fontSize: 12, color: tokens.colorPaletteRedForeground1 }}>
          {actionError}
        </Text>
      )}
      {tickets.length === 0 ? (
        <Text className={styles.empty}>None.</Text>
      ) : (
        tickets.map((t) => {
          const isBusy = busyTicketId === t.id;
          return (
            <div key={t.id} className={styles.ticket}>
              <div
                className={styles.ticketBody}
                onClick={() => openInHalo(t.id)}
                role="button"
                tabIndex={0}
                onKeyDown={(e) => e.key === "Enter" && openInHalo(t.id)}
              >
                <Text className={styles.ticketTitle}>
                  #{t.id} · {t.summary}
                </Text>
                <div className={styles.ticketMeta}>
                  {t.statusname && (
                    <Badge appearance="outline" size="small">
                      {t.statusname}
                    </Badge>
                  )}
                  {t.agent_name && <span>· {t.agent_name}</span>}
                </div>
              </div>

              {isBusy ? (
                <div className={styles.busyOverlay}>
                  <Spinner size="extra-tiny" />
                </div>
              ) : (
                <Menu>
                  <MenuTrigger disableButtonEnhancement>
                    <MenuButton
                      appearance="subtle"
                      size="small"
                      icon={<MoreVertical16Regular />}
                      aria-label="Quick actions"
                    />
                  </MenuTrigger>
                  <MenuPopover>
                    <MenuList>
                      <MenuItem icon={<Open16Regular />} onClick={() => openInHalo(t.id)}>
                        Open in HaloPSA
                      </MenuItem>

                      {currentAgent && currentAgent.id !== t.agent_id && (
                        <MenuItem
                          icon={<PersonAdd16Regular />}
                          onClick={() => apply(t, { agent_id: currentAgent.id })}
                        >
                          Assign to me
                        </MenuItem>
                      )}

                      {statuses.length > 0 && (
                        <>
                          <MenuDivider />
                          {statuses
                            .filter((s) => s.id !== t.status_id)
                            .slice(0, 8)
                            .map((s) => (
                              <MenuItem
                                key={s.id}
                                icon={
                                  s.isclosed ? (
                                    <CheckmarkCircle16Regular />
                                  ) : (
                                    <ArrowUpRight16Regular />
                                  )
                                }
                                onClick={() => apply(t, { status_id: s.id })}
                              >
                                {s.name}
                              </MenuItem>
                            ))}
                        </>
                      )}
                    </MenuList>
                  </MenuPopover>
                </Menu>
              )}
            </div>
          );
        })
      )}
    </div>
  );
}
