import { useRef } from "react";

/** Invisible <input type="file"> wrapped into an imperative handle. */
export function useHiddenFilePicker(
  accept: string,
  onPick: (file: File) => void,
) {
  const ref = useRef<HTMLInputElement | null>(null);
  const el = (
    <input
      ref={ref}
      type="file"
      accept={accept}
      style={{ display: "none" }}
      onChange={(e) => {
        const f = e.target.files?.[0];
        if (f) onPick(f);
        e.target.value = "";
      }}
    />
  );
  return {
    el,
    open: () => ref.current?.click(),
  };
}
