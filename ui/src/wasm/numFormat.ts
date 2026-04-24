// Excel number-format engine.
//
// Implements a useful subset of ECMA-376 §18.8.30 format codes:
//   General, 0, 0.00, #,##0, #,##0.00, 0%, 0.00%, 0.00E+00,
//   date/time tokens (yyyy/yy, mmmm/mmm/mm, dd/d, h/hh, m/mm (minute),
//   s/ss, AM/PM, [h], [m], [s]), positive;negative;zero;text sections,
//   literal text in quotes, escape with backslash, color modifiers [Red]
//   (stripped but not applied), text placeholder @, locale-insensitive
//   thousand separator.
//
// The output is always a plain string; coloring/background from the
// format (`[Red]`) is intentionally not applied — callers can detect the
// negative-section marker via `formatCell` returning a `{ text, color? }`
// helper if needed in the future.

export interface FormatResult {
  text: string;
  color?: string | null;
}

const WEEKDAYS_FULL = [
  "Sunday",
  "Monday",
  "Tuesday",
  "Wednesday",
  "Thursday",
  "Friday",
  "Saturday",
];
const WEEKDAYS_SHORT = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
const MONTHS_FULL = [
  "January",
  "February",
  "March",
  "April",
  "May",
  "June",
  "July",
  "August",
  "September",
  "October",
  "November",
  "December",
];
const MONTHS_SHORT = [
  "Jan",
  "Feb",
  "Mar",
  "Apr",
  "May",
  "Jun",
  "Jul",
  "Aug",
  "Sep",
  "Oct",
  "Nov",
  "Dec",
];

/** Excel epoch is 1900-01-01, serial 1 (with the famous 1900 leap-year bug). */
export function excelSerialToDate(serial: number): Date {
  // Excel treats 1900 as a leap year; days before 1900-03-01 are off by 1.
  const adjusted = serial < 60 ? serial + 1 : serial;
  const epoch = Date.UTC(1899, 11, 30); // 1899-12-30 UTC
  return new Date(epoch + adjusted * 86400000);
}

/** Format a Date/Number/Bool/Text cell according to an Excel format code. */
export function formatCellValue(
  value: unknown,
  code: string | null | undefined,
): FormatResult {
  if (value == null) return { text: "" };
  const fmt = code && code !== "General" ? code.trim() : "General";

  // Plain text values ignore numeric formats but honor `@` text section.
  if (typeof value === "string") {
    // Try date if the format is a date code and the string is ISO-8601.
    if (isDateFormat(fmt)) {
      const d = parseIsoDate(value);
      if (d) return applyFormat(d, true, fmt);
    }
    return applyFormat(value, false, fmt);
  }
  if (typeof value === "boolean") {
    return { text: value ? "TRUE" : "FALSE" };
  }
  if (typeof value === "number") {
    // Dates are detected heuristically: format has date tokens AND value > 0.
    const asDate = isDateFormat(fmt) && value >= 0 && value < 2958466
      ? excelSerialToDate(value)
      : null;
    return asDate ? applyFormat(asDate, true, fmt) : applyFormat(value, false, fmt);
  }
  return { text: String(value) };
}

// ── Detection helpers ────────────────────────────────

function isDateFormat(fmt: string): boolean {
  // Skip escaped literals.
  const stripped = fmt.replace(/\[[^\]]*\]/g, "").replace(/"[^"]*"/g, "");
  return /y|m|d|h|s/i.test(stripped) && !/^@|^0|^#/.test(stripped);
}

function parseIsoDate(s: string): Date | null {
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

// ── Section dispatch ────────────────────────────────

function applyFormat(
  value: Date | number | string,
  isDate: boolean,
  fmt: string,
): FormatResult {
  const sections = splitSections(fmt);
  // sections: [positive, negative, zero, text]; if only one provided,
  // it applies to all numeric kinds.
  let section = sections[0];
  let color: string | null | undefined = undefined;
  if (typeof value === "number") {
    if (value < 0 && sections.length >= 2) {
      section = sections[1];
    } else if (value === 0 && sections.length >= 3) {
      section = sections[2];
    }
  } else if (typeof value === "string" && sections.length >= 4) {
    section = sections[3];
  }
  // Strip leading [Color] token and expose it.
  const colorMatch = /^\s*\[(Red|Blue|Green|Magenta|Yellow|Cyan|Black|White)\]/i.exec(section);
  if (colorMatch) {
    color = colorMatch[1].toLowerCase();
    section = section.slice(colorMatch[0].length);
  }
  const text = renderSection(value, isDate, section);
  return color ? { text, color } : { text };
}

function splitSections(fmt: string): string[] {
  // Split on `;` outside quoted strings & bracket groups.
  const out: string[] = [];
  let buf = "";
  let inQuote = false;
  let inBracket = 0;
  for (let i = 0; i < fmt.length; i++) {
    const ch = fmt[i];
    if (inQuote) {
      if (ch === '"') inQuote = false;
      buf += ch;
      continue;
    }
    if (ch === '"') { inQuote = true; buf += ch; continue; }
    if (ch === "[") { inBracket++; buf += ch; continue; }
    if (ch === "]") { inBracket = Math.max(0, inBracket - 1); buf += ch; continue; }
    if (ch === "\\" && i + 1 < fmt.length) {
      buf += ch + fmt[i + 1];
      i += 1;
      continue;
    }
    if (ch === ";" && inBracket === 0) {
      out.push(buf);
      buf = "";
      continue;
    }
    buf += ch;
  }
  out.push(buf);
  return out;
}

// ── Section renderer ────────────────────────────────

function renderSection(value: Date | number | string, isDate: boolean, section: string): string {
  if (section === "General") return generalNumber(value);
  if (section.trim() === "") return "";

  // Text-only section: just passthrough (or @ substitution).
  if (typeof value === "string" && !isDate) {
    return renderText(section, value);
  }

  if (value instanceof Date || isDate) {
    const d = value instanceof Date ? value : new Date(String(value));
    return renderDateTime(section, d);
  }

  if (typeof value === "number") {
    return renderNumber(section, value);
  }
  return String(value);
}

function generalNumber(value: Date | number | string): string {
  if (value instanceof Date) return value.toLocaleDateString();
  if (typeof value === "number") {
    if (Number.isInteger(value)) return String(value);
    // Trim trailing zeros, stay within 15 sig digits.
    const s = value.toPrecision(15);
    return Number(s).toString();
  }
  return String(value);
}

// ── Text sections ────────────────────────────────

function renderText(section: string, text: string): string {
  let out = "";
  let i = 0;
  while (i < section.length) {
    const ch = section[i];
    if (ch === '"') {
      const end = section.indexOf('"', i + 1);
      if (end === -1) break;
      out += section.slice(i + 1, end);
      i = end + 1;
      continue;
    }
    if (ch === "\\" && i + 1 < section.length) {
      out += section[i + 1];
      i += 2;
      continue;
    }
    if (ch === "@") {
      out += text;
      i += 1;
      continue;
    }
    if (ch === "*" || ch === "_") {
      // Fill character / width placeholder — render as space for UI simplicity.
      i += 2;
      continue;
    }
    out += ch;
    i += 1;
  }
  return out;
}

// ── Number sections ────────────────────────────────

interface NumberSpec {
  prefix: string;
  suffix: string;
  intDigits: string;   // e.g. "#,##0"
  fracDigits: string;  // e.g. "00"
  thousands: boolean;
  percent: boolean;
  scientific: boolean;
  expSign: string;
  expDigits: string;
}

function parseNumberSection(section: string): NumberSpec {
  const spec: NumberSpec = {
    prefix: "",
    suffix: "",
    intDigits: "",
    fracDigits: "",
    thousands: false,
    percent: false,
    scientific: false,
    expSign: "",
    expDigits: "",
  };
  let i = 0;
  // prefix: up to first digit placeholder
  const isDigit = (c: string) => c === "0" || c === "#" || c === "?";
  while (i < section.length && !isDigit(section[i])) {
    const ch = section[i];
    if (ch === '"') {
      const end = section.indexOf('"', i + 1);
      if (end === -1) break;
      spec.prefix += section.slice(i + 1, end);
      i = end + 1;
      continue;
    }
    if (ch === "\\" && i + 1 < section.length) {
      spec.prefix += section[i + 1];
      i += 2;
      continue;
    }
    spec.prefix += ch;
    i += 1;
  }
  // body: int, frac, scientific
  let phase: "int" | "frac" | "exp" = "int";
  for (; i < section.length; i++) {
    const ch = section[i];
    if (phase === "int" && isDigit(ch)) {
      spec.intDigits += ch;
      continue;
    }
    if (phase === "int" && ch === ",") {
      spec.thousands = true;
      continue;
    }
    if (ch === ".") {
      phase = "frac";
      continue;
    }
    if (phase === "frac" && isDigit(ch)) {
      spec.fracDigits += ch;
      continue;
    }
    if (ch === "E" || ch === "e") {
      spec.scientific = true;
      if (i + 1 < section.length && (section[i + 1] === "+" || section[i + 1] === "-")) {
        spec.expSign = section[i + 1];
        i += 1;
      }
      phase = "exp";
      continue;
    }
    if (phase === "exp" && isDigit(ch)) {
      spec.expDigits += ch;
      continue;
    }
    if (ch === "%") {
      spec.percent = true;
      continue;
    }
    // suffix
    if (ch === '"') {
      const end = section.indexOf('"', i + 1);
      if (end === -1) break;
      spec.suffix += section.slice(i + 1, end);
      i = end;
      continue;
    }
    if (ch === "\\" && i + 1 < section.length) {
      spec.suffix += section[i + 1];
      i += 1;
      continue;
    }
    spec.suffix += ch;
  }
  return spec;
}

function renderNumber(section: string, raw: number): string {
  const spec = parseNumberSection(section);
  let value = Math.abs(raw);
  if (spec.percent) value *= 100;

  let body: string;
  if (spec.scientific) {
    const digits = spec.fracDigits.length;
    const expDigits = spec.expDigits.length || 2;
    const formatted = value.toExponential(digits);
    const match = /^(-?)(\d+(?:\.\d+)?)e([+-]?)(\d+)$/i.exec(formatted);
    if (match) {
      const [, , mant, sign, exp] = match;
      const expStr = exp.padStart(expDigits, "0");
      body = `${mant}E${spec.expSign || sign || "+"}${expStr}`;
    } else {
      body = formatted;
    }
  } else {
    const fracLen = spec.fracDigits.length;
    const fixed = value.toFixed(fracLen);
    const [intPart, fracPart = ""] = fixed.split(".");
    const intOut = spec.thousands
      ? intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ",")
      : intPart;
    body = fracLen > 0 ? `${intOut}.${fracPart}` : intOut;
  }

  const signed = raw < 0 ? `-${body}` : body;
  return `${spec.prefix}${signed}${spec.suffix}`;
}

// ── Date sections ────────────────────────────────

function renderDateTime(section: string, d: Date): string {
  let out = "";
  let i = 0;
  // Heuristic: look ahead to classify ambiguous `m` — if preceded by `h`
  // or followed by `:ss` it's minute, else month.
  const hasAmPm = /AM\/PM|A\/P|am\/pm/i.test(section);
  while (i < section.length) {
    const ch = section[i];
    if (ch === '"') {
      const end = section.indexOf('"', i + 1);
      if (end === -1) break;
      out += section.slice(i + 1, end);
      i = end + 1;
      continue;
    }
    if (ch === "\\" && i + 1 < section.length) {
      out += section[i + 1];
      i += 2;
      continue;
    }
    if (ch === "[") {
      const end = section.indexOf("]", i + 1);
      if (end === -1) break;
      const token = section.slice(i + 1, end);
      // Elapsed time: [h], [m], [s]
      if (/^h+$/i.test(token)) {
        out += String(Math.floor(d.getTime() / 3600000)).padStart(token.length, "0");
      } else if (/^m+$/i.test(token)) {
        out += String(Math.floor(d.getTime() / 60000)).padStart(token.length, "0");
      } else if (/^s+$/i.test(token)) {
        out += String(Math.floor(d.getTime() / 1000)).padStart(token.length, "0");
      }
      // Else: color token — already stripped upstream. Ignore.
      i = end + 1;
      continue;
    }
    // Date/time tokens (longest first)
    const rest = section.slice(i);
    if (/^yyyy/i.test(rest)) { out += d.getFullYear(); i += 4; continue; }
    if (/^yy/i.test(rest)) { out += String(d.getFullYear() % 100).padStart(2, "0"); i += 2; continue; }
    if (/^mmmm/i.test(rest)) { out += MONTHS_FULL[d.getMonth()]; i += 4; continue; }
    if (/^mmm/i.test(rest)) { out += MONTHS_SHORT[d.getMonth()]; i += 3; continue; }
    if (/^dddd/i.test(rest)) { out += WEEKDAYS_FULL[d.getDay()]; i += 4; continue; }
    if (/^ddd/i.test(rest)) { out += WEEKDAYS_SHORT[d.getDay()]; i += 3; continue; }
    if (/^dd/i.test(rest)) { out += String(d.getDate()).padStart(2, "0"); i += 2; continue; }
    if (/^d/i.test(rest)) { out += d.getDate(); i += 1; continue; }
    if (/^hh/i.test(rest)) {
      const h = hasAmPm ? ((d.getHours() + 11) % 12) + 1 : d.getHours();
      out += String(h).padStart(2, "0");
      i += 2; continue;
    }
    if (/^h/i.test(rest)) {
      const h = hasAmPm ? ((d.getHours() + 11) % 12) + 1 : d.getHours();
      out += h;
      i += 1; continue;
    }
    if (/^ss/i.test(rest)) { out += String(d.getSeconds()).padStart(2, "0"); i += 2; continue; }
    if (/^s/i.test(rest)) { out += d.getSeconds(); i += 1; continue; }
    // Ambiguous m: minute if near h/ss, else month
    if (/^mm/i.test(rest)) {
      const prev = out.slice(-3).toLowerCase();
      const next = rest.slice(2, 5).toLowerCase();
      const isMinute = /h|:$|:0/.test(prev) || /:ss/i.test(next);
      out += isMinute
        ? String(d.getMinutes()).padStart(2, "0")
        : String(d.getMonth() + 1).padStart(2, "0");
      i += 2; continue;
    }
    if (/^m/i.test(rest)) {
      const prev = out.slice(-3).toLowerCase();
      const isMinute = /h|:$|:0/.test(prev);
      out += isMinute ? d.getMinutes() : d.getMonth() + 1;
      i += 1; continue;
    }
    if (/^AM\/PM/i.test(rest)) {
      out += d.getHours() >= 12 ? "PM" : "AM";
      i += 5; continue;
    }
    if (/^A\/P/i.test(rest)) {
      out += d.getHours() >= 12 ? "P" : "A";
      i += 3; continue;
    }
    out += ch;
    i += 1;
  }
  return out;
}
