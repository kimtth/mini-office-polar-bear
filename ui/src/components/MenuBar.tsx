import {
  useCallback,
  useEffect,
  useRef,
  useState,
  type ReactNode,
} from "react";

export interface MenuAction {
  label: string;
  onClick?: () => void;
  disabled?: boolean;
  kbd?: string;
  /** render a separator instead of a button */
  separator?: boolean;
}

export interface Menu {
  label: string;
  actions: MenuAction[];
}

interface Props {
  menus: Menu[];
  accent: string;
  appName: string;
  rightSlot?: ReactNode;
}

/**
 * Office-style top menu bar with dropdowns.
 * Click a menu to open it; hover switches to another open menu; click outside closes.
 */
export default function MenuBar({ menus, accent, appName, rightSlot }: Props) {
  const [open, setOpen] = useState<number | null>(null);
  const rootRef = useRef<HTMLDivElement>(null);

  const close = useCallback(() => setOpen(null), []);

  useEffect(() => {
    function onDocClick(e: MouseEvent) {
      if (!rootRef.current?.contains(e.target as Node)) close();
    }
    function onKey(e: KeyboardEvent) {
      if (e.key === "Escape") close();
    }
    document.addEventListener("mousedown", onDocClick);
    document.addEventListener("keydown", onKey);
    return () => {
      document.removeEventListener("mousedown", onDocClick);
      document.removeEventListener("keydown", onKey);
    };
  }, [close]);

  return (
    <div
      ref={rootRef}
      className="pbo-menubar"
      style={{ ["--pbo-accent" as string]: accent } as React.CSSProperties}
    >
      <div className="pbo-brand-dot">
        <span className="pbo-accent" />
        <span>{appName}</span>
      </div>
      {menus.map((m, i) => (
        <div
          key={m.label}
          className="pbo-menu-item"
          aria-expanded={open === i}
          onClick={() => setOpen(open === i ? null : i)}
          onMouseEnter={() => open !== null && setOpen(i)}
        >
          {m.label}
          {open === i && (
            <div className="pbo-menu-pop" onClick={(e) => e.stopPropagation()}>
              {m.actions.map((a, j) =>
                a.separator ? (
                  <div key={j} className="pbo-menu-sep" />
                ) : (
                  <button
                    key={j}
                    disabled={a.disabled}
                    onClick={() => {
                      a.onClick?.();
                      close();
                    }}
                  >
                    <span>{a.label}</span>
                    {a.kbd && <span className="pbo-menu-kbd">{a.kbd}</span>}
                  </button>
                ),
              )}
            </div>
          )}
        </div>
      ))}
      <div style={{ flex: 1 }} />
      {rightSlot}
    </div>
  );
}
