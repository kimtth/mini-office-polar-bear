import { useEffect, useMemo, useRef, useState } from "react";
import { Link, useLocation } from "react-router-dom";
import logo from "../../../../docs/logo.png";
import {
  documentTypeFor,
  onlyOfficeBaseUrl,
  onlyOfficeDocsApiUrl,
  type AdvancedEditorKind,
} from "./advancedEditor";

type LoadState = "idle" | "loading" | "ready" | "error";

export default function OnlyOfficeEditor() {
  const location = useLocation();
  const editorRef = useRef<OnlyOfficeEditorInstance | null>(null);
  const [loadState, setLoadState] = useState<LoadState>("idle");
  const [error, setError] = useState<string | null>(null);

  const params = useMemo(() => new URLSearchParams(location.search), [location.search]);
  const kind = (params.get("kind") ?? "docx") as AdvancedEditorKind;
  const title = params.get("title") ?? `untitled.${kind}`;
  const file = params.get("file") ?? "";
  const fileUrl = file ? new URL(file, window.location.origin).toString() : "";
  const docsApiUrl = onlyOfficeDocsApiUrl();
  const onlyOfficeUrl = onlyOfficeBaseUrl();

  useEffect(() => {
    if (!fileUrl) {
      setLoadState("error");
      setError("No advanced editor document was provided.");
      return;
    }

    let cancelled = false;
    setLoadState("loading");
    setError(null);

    const createEditor = () => {
      if (cancelled) return;
      if (!window.DocsAPI?.DocEditor) {
        setLoadState("error");
        setError(`OnlyOffice DocsAPI is unavailable. Expected script: ${docsApiUrl}`);
        return;
      }
      editorRef.current?.destroyEditor?.();
      editorRef.current = new window.DocsAPI.DocEditor("pbo-advanced-editor", {
        document: {
          fileType: kind,
          key: `${kind}-${title}`,
          title,
          url: fileUrl,
          permissions: {
            edit: true,
            download: true,
            print: true,
          },
        },
        documentType: documentTypeFor(kind),
        editorConfig: {
          mode: "edit",
          customization: {
            uiTheme: "theme-light",
            compactHeader: false,
            logo: {
              image: logo,
              imageDark: logo,
              url: window.location.origin,
            },
          },
        },
        width: "100%",
        height: "100%",
      });
      setLoadState("ready");
    };

    const existing = document.querySelector<HTMLScriptElement>(`script[data-pbo-onlyoffice="${docsApiUrl}"]`);
    if (window.DocsAPI?.DocEditor) {
      createEditor();
    } else if (existing) {
      const onLoad = () => createEditor();
      const onError = () => {
        if (cancelled) return;
        setLoadState("error");
        setError(`Failed to load OnlyOffice DocsAPI from ${docsApiUrl}.`);
      };
      existing.addEventListener("load", onLoad, { once: true });
      existing.addEventListener("error", onError, { once: true });
    } else {
      const script = document.createElement("script");
      script.src = docsApiUrl;
      script.async = true;
      script.dataset.pboOnlyoffice = docsApiUrl;
      script.onload = createEditor;
      script.onerror = () => {
        if (cancelled) return;
        setLoadState("error");
        setError(`Failed to load OnlyOffice DocsAPI from ${docsApiUrl}.`);
      };
      document.head.appendChild(script);
    }

    return () => {
      cancelled = true;
      editorRef.current?.destroyEditor?.();
      editorRef.current = null;
    };
  }, [docsApiUrl, fileUrl, kind, title]);

  return (
    <div className="pbo-advanced-shell">
      <div className="pbo-advanced-topbar">
        <div className="pbo-advanced-meta">
          <h1>Advanced Editor</h1>
          <p>{title}</p>
        </div>
        <div className="pbo-advanced-actions">
          <Link className="pbo-btn" to="/">Home</Link>
          <Link className="pbo-btn pbo-btn-advanced" to={kind === "docx" ? "/writer" : kind === "xlsx" ? "/calc" : "/slides"}>
            ← Back to Simple Editor
          </Link>
        </div>
      </div>

      <div className="pbo-advanced-banner">
        <strong>Default editor only supports simple editing.</strong> This screen uses the OnlyOffice wrapper for full editing. Configure <code>VITE_ONLYOFFICE_BASE_URL</code> if your Docs assets are not served from <code>{onlyOfficeUrl}</code>.
      </div>

      <div className="pbo-advanced-stage">
        {loadState !== "ready" ? (
          <div className="pbo-advanced-placeholder">
            <div className="pbo-advanced-panel">
              <h2>{loadState === "loading" ? "Loading OnlyOffice" : "OnlyOffice wrapper"}</h2>
              <p>
                {loadState === "loading"
                  ? `Loading DocsAPI from ${docsApiUrl}.`
                  : error ?? "Preparing the advanced editor."}
              </p>
              <p>The current document has already been exported from the simple editor and staged at <code>{fileUrl || "(missing file)"}</code>.</p>
              <ul className="pbo-advanced-list">
                <li>OnlyOffice base URL: <code>{onlyOfficeUrl}</code></li>
                <li>DocsAPI script: <code>{docsApiUrl}</code></li>
                <li>Document type: <code>{kind}</code></li>
              </ul>
            </div>
          </div>
        ) : null}
        <div id="pbo-advanced-editor" className="pbo-advanced-editor" style={{ display: loadState === "ready" ? "block" : "none" }} />
      </div>
    </div>
  );
}