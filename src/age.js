import * as XLSX from "xlsx-js-style";

export const DAY = 864e5;
const now = new Date();
export const TODAY = Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()) / DAY;

export const FORMATS = [
  ["days", "In Days"],
  ["months", "In Months (2 dec)"],
  ["years", "In Years (2 dec)"],
  ["ymd", "Y M D"],
];

// [label, order of parts, separator]  — first one is the default
export const DATE_FORMATS = [
  ["DD-MM-YYYY", "dmy", "-"],
  ["DD/MM/YYYY", "dmy", "/"],
  ["DD.MM.YYYY", "dmy", "."],
  ["MM-DD-YYYY", "mdy", "-"],
  ["MM/DD/YYYY", "mdy", "/"],
  ["YYYY-MM-DD", "ymd", "-"],
  ["YYYY/MM/DD", "ymd", "/"],
  ["DD-Mon-YYYY (15-Oct-2005)", "dmon", "-"],
];
const MON = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

// day number -> text in the chosen date format (used for column names / labels)
export const fmtD = (d, f = DATE_FORMATS[0]) => {
  const x = new Date(d * DAY), p = (n) => String(n).padStart(2, "0");
  const D = p(x.getUTCDate()), M = p(x.getUTCMonth() + 1), Y = x.getUTCFullYear(), s = f[2];
  if (f[1] === "mdy") return [M, D, Y].join(s);
  if (f[1] === "ymd") return [Y, M, D].join(s);
  if (f[1] === "dmon") return [D, MON[x.getUTCMonth()], Y].join(s);
  return [D, M, Y].join(s);
};

// Excel serial number OR text date -> day number, else null.
// Understood automatically per cell: YYYY-MM-DD, 15-Oct-2005, and DD/MM or MM/DD text.
// DD/MM vs MM/DD follows `order` strictly (default DD-MM-YYYY); with auto=true a part > 12 decides per cell.
export function toDay(v, order = "dmy", auto = false) {
  if (v === "" || v == null) return null;
  let y, m, d;
  if (typeof v === "number") {
    const p = XLSX.SSF.parse_date_code(v);
    if (!p) return null;
    ({ y, m, d } = p);
  } else {
    const t = String(v).trim();
    let s = t.match(/^(\d{1,2})[-/. ]([A-Za-z]{3,9})[-/., ]+(\d{4})$/);
    if (s) {
      m = MON.findIndex((n) => s[2].toLowerCase().startsWith(n.toLowerCase())) + 1;
      d = +s[1]; y = +s[3];
    } else if ((s = t.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/))) {
      y = +s[1]; m = +s[2]; d = +s[3];
    } else if ((s = t.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/))) {
      const a = +s[1], b = +s[2];
      y = +s[3];
      const mdy = !auto ? order === "mdy" : a > 12 && b <= 12 ? false : b > 12 && a <= 12 ? true : order === "mdy";
      if (mdy) { m = a; d = b; } else { d = a; m = b; }
    } else return null;
  }
  if (!(y >= 1900 && y <= 2100 && m >= 1 && m <= 12 && d >= 1 && d <= 31)) return null;
  const t = Date.UTC(y, m - 1, d);
  return new Date(t).getUTCMonth() === m - 1 ? t / DAY : null; // rejects 31-02 etc.
}

// Look at every value of a column and work out its date format(s).
export function detectColumn(vals) {
  let excel = 0, mon = 0, ymdN = 0, dmy = 0, mdy = 0, amb = 0, bad = 0;
  const sepDM = {}, sepY = {}, samples = [];
  const nb = (v) => { bad++; if (samples.length < 3) samples.push(String(v).slice(0, 28)); };
  const top = (o) => Object.keys(o).sort((x, y) => o[y] - o[x])[0];
  for (const v of vals) {
    if (v === "" || v == null) continue;
    if (typeof v === "number") { toDay(v) != null ? excel++ : nb(v); continue; }
    const t = String(v).trim();
    let s;
    if (/^\d{1,2}[-/. ][A-Za-z]{3,9}[-/., ]+\d{4}$/.test(t)) toDay(t) != null ? mon++ : nb(v);
    else if ((s = t.match(/^(\d{4})([-/.])\d{1,2}[-/.]\d{1,2}$/))) { toDay(t) != null ? ymdN++ : nb(v); sepY[s[2]] = (sepY[s[2]] || 0) + 1; }
    else if ((s = t.match(/^(\d{1,2})([-/.])(\d{1,2})[-/.]\d{4}$/))) {
      const a = +s[1], b = +s[3];
      sepDM[s[2]] = (sepDM[s[2]] || 0) + 1;
      if (toDay(t, "dmy") == null && toDay(t, "mdy") == null) nb(v);
      else if (a > 12 && b <= 12) dmy++; else if (b > 12 && a <= 12) mdy++; else if (a <= 12 && b <= 12) amb++; else nb(v);
    } else nb(v);
  }
  const order = mdy > dmy ? "mdy" : "dmy";
  const parts = [];
  if (excel) parts.push("Excel date");
  if (dmy + mdy + amb) {
    const sp = top(sepDM) || "-";
    let l = (order === "mdy" ? "MM" + sp + "DD" : "DD" + sp + "MM") + sp + "YYYY";
    if (dmy && mdy) l += " + MM/DD mixed (resolved per cell)";
    else if (!dmy && !mdy) l += " (day/month unclear, assumed)";
    parts.push(l);
  }
  if (ymdN) { const sp = top(sepY) || "-"; parts.push("YYYY" + sp + "MM" + sp + "DD"); }
  if (mon) parts.push("DD-Mon-YYYY");
  return { order, bad, samples, good: excel + mon + ymdN + dmy + mdy + amb, label: parts.join(" + ") || "no dates found" };
}

function ymd(a, b) {
  const A = new Date(a * DAY), B = new Date(b * DAY);
  let y = B.getUTCFullYear() - A.getUTCFullYear();
  let m = B.getUTCMonth() - A.getUTCMonth();
  let d = B.getUTCDate() - A.getUTCDate();
  if (d < 0) { m--; d += new Date(Date.UTC(B.getUTCFullYear(), B.getUTCMonth(), 0)).getUTCDate(); }
  if (m < 0) { y--; m += 12; }
  return `${y} Y ${m} M ${d} D`;
}

// age between two day numbers in the chosen format; "" when a date is missing or end < start
export function age(a, b, f) {
  if (a == null || b == null || b < a) return "";
  const n = b - a;
  if (f === "days") return n;
  if (f === "months") return +(n / (365.25 / 12)).toFixed(2);
  if (f === "years") return +(n / 365.25).toFixed(2);
  return ymd(a, b);
}

// Find the header row (1-based) in the first 30 rows: the first row that is filled
// across most columns with unique, non-date, non-number text (skips titles / blank rows).
export function detectHeaderRow(rows) {
  const top = rows.slice(0, 30);
  const vals = (r) => r.map((v) => (typeof v === "string" ? v.trim() : v)).filter((v) => v !== "" && v != null);
  const maxF = Math.max(0, ...top.map((r) => vals(r).length));
  if (maxF < 2) return 1;
  for (let i = 0; i < top.length; i++) {
    const v = vals(top[i]);
    if (v.length < Math.max(2, Math.ceil(maxF * 0.6))) continue;
    const text = v.every((x) => typeof x === "string" && toDay(x) == null && isNaN(Number(x)));
    if (text && new Set(v.map((x) => x.toLowerCase())).size === v.length) return i + 1;
  }
  return 1;
}
