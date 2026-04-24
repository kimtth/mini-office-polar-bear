import { Link } from "react-router-dom";
import logo from "../../../docs/logo.png";
import "../styles/home.css";

interface AppCard {
  to: string;
  title: string;
  tag: string;
  note: string;
}

const APPS: AppCard[] = [
  {
    to: "/writer",
    title: "Writer",
    tag: ".docx",
    note: "Fast paragraph and style edits in the simple editor.",
  },
  {
    to: "/slides",
    title: "Slides",
    tag: ".pptx",
    note: "Quick slide updates, text changes, and layout tweaks.",
  },
  {
    to: "/calc",
    title: "Calc",
    tag: ".xlsx",
    note: "Basic spreadsheet edits without leaving the browser.",
  },
];

export default function Home() {
  return (
    <div className="pbo-home">
      <header>
        <div className="pbo-logo">
          <span>PolarBearOffice</span>
        </div>
        <nav className="pbo-top-links">
          <a
            href="https://github.com"
            target="_blank"
            rel="noreferrer"
            aria-label="GitHub"
          >
            GitHub
          </a>
          <a href="/api/health" target="_blank" rel="noreferrer">
            API
          </a>
        </nav>
      </header>

      <section className="pbo-hero">
        <h1>
          An <span className="pbo-grad">open‑source Office</span><br/> in your browser.
        </h1>
        <p>
          PolarBearOffice (experimental — mini office application) opens, edits, and saves Word, Excel, and PowerPoint
          documents using a fast Rust engine compiled to WebAssembly. Very minimal implementation for the default
          editor (experimental). OnlyOffice wrapper for advanced editing.
        </p>
        <div className="pbo-badges">
          <span className="pbo-badge">v0.1 · Preview</span>
          <span className="pbo-badge">Rust + WebAssembly</span>
          <span className="pbo-badge">DOCX · XLSX · PPTX</span>
        </div>
      </section>

      <section className="pbo-logo-band">
        <img className="pbo-hero-logo" src={logo} alt="PolarBearOffice" />
      </section>

      <section className="pbo-apps">
        {APPS.map((a) => (
          <Link key={a.to} to={a.to} className="pbo-app-card">
            <span className="pbo-app-card-head">
              <span className="pbo-app-title">{a.title}</span>
              <span className="pbo-app-tag">{a.tag}</span>
            </span>
            <span className="pbo-app-note">{a.note}</span>
          </Link>
        ))}
      </section>

      <footer>
        © {new Date().getFullYear()} PolarBearOffice
      </footer>
    </div>
  );
}
