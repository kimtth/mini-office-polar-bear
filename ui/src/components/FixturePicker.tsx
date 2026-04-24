import { useEffect, useState } from "react";
import { listFixtures, type FixtureEntry } from "../wasm/loader";

interface Props {
  kind: "docx" | "xlsx" | "pptx";
  open: boolean;
  onClose: () => void;
  onPick: (entry: FixtureEntry) => void;
}

/**
 * Modal that lists fixtures on disk via `/api/fixtures`, filtered by kind,
 * and lets the user pick one to open.
 */
export default function FixturePicker({ kind, open, onClose, onPick }: Props) {
  const [items, setItems] = useState<FixtureEntry[] | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (!open) return;
    setError(null);
    setItems(null);
    listFixtures()
      .then((all) => setItems(all.filter((f) => f.kind === kind)))
      .catch((e: Error) => setError(e.message));
  }, [open, kind]);

  if (!open) return null;

  return (
    <div
      className="pbo-modal-backdrop"
      onClick={onClose}
      role="dialog"
      aria-modal="true"
    >
      <div className="pbo-modal" onClick={(e) => e.stopPropagation()}>
        <div className="pbo-modal-header">
          <span>Open sample {kind.toUpperCase()}</span>
          <button className="pbo-close" onClick={onClose} aria-label="Close">
            ×
          </button>
        </div>
        <div className="pbo-modal-body">
          {error && (
            <div style={{ color: "var(--pbo-danger)" }}>Error: {error}</div>
          )}
          {!items && !error && (
            <div style={{ color: "var(--pbo-text-muted)" }}>Loading…</div>
          )}
          {items && items.length === 0 && (
            <div style={{ color: "var(--pbo-text-muted)" }}>
              No fixtures available. Drop .{kind} files into <code>fixtures/{kind}/</code>.
            </div>
          )}
          {items && items.length > 0 && (
            <ul
              style={{
                listStyle: "none",
                padding: 0,
                margin: 0,
                display: "flex",
                flexDirection: "column",
                gap: 4,
              }}
            >
              {items.map((f) => (
                <li key={f.name}>
                  <button
                    className="pbo-btn"
                    style={{ width: "100%", justifyContent: "space-between" }}
                    onClick={() => onPick(f)}
                  >
                    <span>📄 {f.name}</span>
                    <span
                      style={{
                        color: "var(--pbo-text-muted)",
                        fontSize: "0.75rem",
                      }}
                    >
                      {(f.bytes / 1024).toFixed(1)} KB
                    </span>
                  </button>
                </li>
              ))}
            </ul>
          )}
        </div>
        <div className="pbo-modal-footer">
          <button className="pbo-btn" onClick={onClose}>
            Cancel
          </button>
        </div>
      </div>
    </div>
  );
}
