import { useEffect, useState } from "react";
import {
  Button,
  makeStyles,
  tokens,
  Text,
  Field,
  Combobox,
  Option,
  Switch,
  Divider,
  Spinner,
} from "@fluentui/react-components";
import { ArrowLeft24Regular } from "@fluentui/react-icons";
import { getConfig } from "../lib/config";
import { getDefaults, setDefaults } from "../lib/defaults";
import { listTicketTypes, clearReferenceCache } from "../lib/halo-api";
import type { HaloTicketType } from "../types/halo";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    height: "100%",
    overflowY: "auto",
  },
  header: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    padding: "8px 12px",
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  title: {
    fontWeight: tokens.fontWeightSemibold,
    fontSize: tokens.fontSizeBase300,
  },
  body: {
    display: "flex",
    flexDirection: "column",
    gap: "16px",
    padding: "16px",
    flex: 1,
  },
  sectionLabel: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
    color: tokens.colorNeutralForeground2,
    textTransform: "uppercase",
    letterSpacing: "0.04em",
  },
  meta: {
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
  },
  dangerSection: {
    display: "flex",
    flexDirection: "column",
    gap: "8px",
  },
});

interface Props {
  onClose: () => void;
  onSignOut: () => void;
  onReconfigure: () => void;
}

export function SettingsScreen({ onClose, onSignOut, onReconfigure }: Props) {
  const styles = useStyles();
  const cfg = getConfig();
  const initialDefaults = getDefaults();
  const [ticketTypes, setTicketTypes] = useState<HaloTicketType[]>([]);
  const [defaultTypeId, setDefaultTypeId] = useState<number | undefined>(
    initialDefaults.defaultTicketTypeId,
  );
  const [defaultOutcome, setDefaultOutcome] = useState<string>(
    initialDefaults.defaultAppendOutcome ?? "Email Received",
  );
  const [includeAttach, setIncludeAttach] = useState<boolean>(
    initialDefaults.includeAttachmentsByDefault ?? true,
  );
  const [saving, setSaving] = useState(false);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    listTicketTypes()
      .then(setTicketTypes)
      .catch(() => {
        /* non-fatal — picker just stays empty */
      })
      .finally(() => setLoading(false));
  }, []);

  const save = async () => {
    setSaving(true);
    try {
      await setDefaults({
        defaultTicketTypeId: defaultTypeId,
        defaultAppendOutcome: defaultOutcome,
        includeAttachmentsByDefault: includeAttach,
      });
      onClose();
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className={styles.root}>
      <div className={styles.header}>
        <Button
          appearance="subtle"
          size="small"
          icon={<ArrowLeft24Regular />}
          onClick={onClose}
          aria-label="Back"
        />
        <Text className={styles.title}>Settings</Text>
      </div>

      <div className={styles.body}>
        <div>
          <Text className={styles.sectionLabel}>Connection</Text>
          <Text block className={styles.meta}>
            {cfg?.haloBaseUrl}
          </Text>
          <Text block className={styles.meta}>
            Client ID: {cfg?.clientId}
          </Text>
        </div>

        <Divider />

        <Text className={styles.sectionLabel}>Defaults</Text>

        <Field label="Default ticket type for Create">
          {loading ? (
            <Spinner size="extra-tiny" />
          ) : (
            <Combobox
              placeholder="No default"
              value={ticketTypes.find((t) => t.id === defaultTypeId)?.name ?? ""}
              onOptionSelect={(_, d) =>
                setDefaultTypeId(d.optionValue ? Number(d.optionValue) : undefined)
              }
            >
              <Option value="" text="No default">
                — No default —
              </Option>
              {ticketTypes.map((t) => (
                <Option key={t.id} value={String(t.id)} text={t.name}>
                  {t.name}
                </Option>
              ))}
            </Combobox>
          )}
        </Field>

        <Field
          label="Default action outcome for Append"
          hint="The HaloPSA Action outcome name applied when appending. Must exist in your Halo config."
        >
          <Combobox
            value={defaultOutcome}
            placeholder="Email Received"
            onOptionSelect={(_, d) => setDefaultOutcome(d.optionText ?? "Email Received")}
            freeform
          >
            <Option value="Email Received">Email Received</Option>
            <Option value="Note">Note</Option>
            <Option value="Internal Note">Internal Note</Option>
          </Combobox>
        </Field>

        <Switch
          checked={includeAttach}
          onChange={(_, d) => setIncludeAttach(d.checked)}
          label="Include attachments by default when logging"
        />

        <Divider />

        <Button
          appearance="primary"
          onClick={save}
          disabled={saving}
          icon={saving ? <Spinner size="tiny" /> : undefined}
        >
          {saving ? "Saving…" : "Save"}
        </Button>

        <Button
          appearance="subtle"
          onClick={() => {
            clearReferenceCache();
            listTicketTypes(true).then(setTicketTypes).catch(() => {});
          }}
        >
          Refresh reference data
        </Button>

        <Divider />

        <div className={styles.dangerSection}>
          <Text className={styles.sectionLabel}>Account</Text>
          <Button appearance="secondary" onClick={onSignOut}>
            Sign out
          </Button>
          <Button appearance="secondary" onClick={onReconfigure}>
            Switch HaloPSA tenant
          </Button>
        </div>

        <Divider />

        <Text block className={styles.meta}>
          HaloPSA for Outlook v0.1.0 · tools.iusehalo.com
        </Text>
      </div>
    </div>
  );
}
