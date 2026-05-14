import { useEffect, useRef, useState } from "react";

interface State<T> {
  results: T[];
  loading: boolean;
  error?: string;
}

/**
 * Debounced async search hook shared by the compose surface's ticket / KB /
 * log-on-send pickers and any future inline search. Cancels pending searches
 * on query change so older results can't overwrite newer ones.
 *
 * @param query    Live input string. Empty / under-min-length disables search.
 * @param search   Async lookup against Halo (or anything else).
 * @param opts.delayMs        Debounce window in ms. Default 300.
 * @param opts.minLength      Minimum trimmed length before searching. Default 2.
 */
export function useDebouncedSearch<T>(
  query: string,
  search: (q: string) => Promise<T[]>,
  opts: { delayMs?: number; minLength?: number } = {},
): State<T> {
  const { delayMs = 300, minLength = 2 } = opts;
  const [state, setState] = useState<State<T>>({ results: [], loading: false });
  const timerRef = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);
  const reqIdRef = useRef(0);

  useEffect(() => {
    if (timerRef.current) clearTimeout(timerRef.current);
    const trimmed = query.trim();
    if (trimmed.length < minLength) {
      setState({ results: [], loading: false });
      return;
    }
    setState((s) => ({ ...s, loading: true, error: undefined }));
    const myId = ++reqIdRef.current;
    timerRef.current = setTimeout(async () => {
      try {
        const r = await search(trimmed);
        // Stale-response guard.
        if (myId !== reqIdRef.current) return;
        setState({ results: r, loading: false });
      } catch (e) {
        if (myId !== reqIdRef.current) return;
        setState({ results: [], loading: false, error: (e as Error).message });
      }
    }, delayMs);
    return () => {
      if (timerRef.current) clearTimeout(timerRef.current);
    };
    // search is allowed to be inline — callers shouldn't have to memoize it
    // for the hook to behave correctly; we re-evaluate on query change only.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [query, delayMs, minLength]);

  return state;
}
