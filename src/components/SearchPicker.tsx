import { useEffect, useState, useRef } from "react";
import {
  Dialog,
  DialogTrigger,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
  Button,
  Input,
  Spinner,
  Text,
  makeStyles,
  tokens,
} from "@fluentui/react-components";
import { Search24Regular } from "@fluentui/react-icons";

const useStyles = makeStyles({
  searchRow: {
    marginBottom: "8px",
  },
  results: {
    maxHeight: "260px",
    overflowY: "auto",
    border: `1px solid ${tokens.colorNeutralStroke2}`,
    borderRadius: tokens.borderRadiusMedium,
  },
  row: {
    padding: "8px 10px",
    cursor: "pointer",
    borderBottom: `1px solid ${tokens.colorNeutralStroke3}`,
    ":hover": {
      backgroundColor: tokens.colorNeutralBackground2Hover,
    },
    ":last-child": {
      borderBottom: "none",
    },
  },
  rowSelected: {
    backgroundColor: tokens.colorBrandBackground2,
  },
  primary: {
    fontSize: tokens.fontSizeBase200,
    fontWeight: tokens.fontWeightSemibold,
  },
  secondary: {
    fontSize: tokens.fontSizeBase100,
    color: tokens.colorNeutralForeground3,
  },
  empty: {
    padding: "12px",
    fontSize: tokens.fontSizeBase200,
    color: tokens.colorNeutralForeground3,
    textAlign: "center",
  },
});

export interface PickerItem<T> {
  key: number;
  primary: string;
  secondary?: string;
  value: T;
}

interface Props<T> {
  /** Trigger button rendering — controlled by parent */
  triggerLabel: string;
  /** Optional trigger icon — defaults to search */
  triggerAppearance?: "primary" | "subtle" | "secondary" | "outline" | "transparent";
  /** Title shown in the dialog */
  title: string;
  /** Initial query text */
  initialQuery?: string;
  /** Search function — debounced from inside the picker */
  onSearch: (query: string) => Promise<PickerItem<T>[]>;
  /** Called when user picks an item */
  onPick: (item: T) => void;
}

export function SearchPicker<T>({
  triggerLabel,
  triggerAppearance = "subtle",
  title,
  initialQuery = "",
  onSearch,
  onPick,
}: Props<T>) {
  const styles = useStyles();
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState(initialQuery);
  const [results, setResults] = useState<PickerItem<T>[]>([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | undefined>();
  const [selectedKey, setSelectedKey] = useState<number | undefined>();
  const debounceRef = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);

  // Reset state when opening
  useEffect(() => {
    if (open) {
      setQuery(initialQuery);
      setSelectedKey(undefined);
      setError(undefined);
    }
  }, [open, initialQuery]);

  // Debounced search
  useEffect(() => {
    if (!open) return;
    if (debounceRef.current) clearTimeout(debounceRef.current);
    if (query.trim().length < 2) {
      setResults([]);
      return;
    }
    debounceRef.current = setTimeout(async () => {
      setLoading(true);
      setError(undefined);
      try {
        const r = await onSearch(query.trim());
        setResults(r);
      } catch (e) {
        setError((e as Error).message);
      } finally {
        setLoading(false);
      }
    }, 250);
    return () => {
      if (debounceRef.current) clearTimeout(debounceRef.current);
    };
  }, [query, open, onSearch]);

  const handlePick = () => {
    const item = results.find((r) => r.key === selectedKey);
    if (item) {
      onPick(item.value);
      setOpen(false);
    }
  };

  return (
    <Dialog open={open} onOpenChange={(_, d) => setOpen(d.open)}>
      <DialogTrigger disableButtonEnhancement>
        <Button appearance={triggerAppearance} size="small" icon={<Search24Regular />}>
          {triggerLabel}
        </Button>
      </DialogTrigger>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>{title}</DialogTitle>
          <DialogContent>
            <div className={styles.searchRow}>
              <Input
                value={query}
                placeholder="Type to search…"
                onChange={(_, d) => setQuery(d.value)}
                autoFocus
              />
            </div>

            {error && (
              <Text size={200} style={{ color: tokens.colorPaletteRedForeground1 }}>
                {error}
              </Text>
            )}

            <div className={styles.results}>
              {loading ? (
                <div className={styles.empty}>
                  <Spinner size="extra-tiny" /> Searching…
                </div>
              ) : results.length === 0 ? (
                <div className={styles.empty}>
                  {query.trim().length < 2 ? "Type at least 2 characters." : "No matches."}
                </div>
              ) : (
                results.map((r) => (
                  <div
                    key={r.key}
                    className={
                      styles.row + (r.key === selectedKey ? " " + styles.rowSelected : "")
                    }
                    onClick={() => setSelectedKey(r.key)}
                    onDoubleClick={() => {
                      setSelectedKey(r.key);
                      onPick(r.value);
                      setOpen(false);
                    }}
                  >
                    <Text className={styles.primary}>{r.primary}</Text>
                    {r.secondary && (
                      <>
                        <br />
                        <Text className={styles.secondary}>{r.secondary}</Text>
                      </>
                    )}
                  </div>
                ))
              )}
            </div>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={() => setOpen(false)}>
              Cancel
            </Button>
            <Button appearance="primary" disabled={!selectedKey} onClick={handlePick}>
              Use selection
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
