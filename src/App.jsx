import { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx-js-style";
import { TODAY, FORMATS, DATE_FORMATS, fmtD as fmtDate, toDay, detectColumn, detectHeaderRow, age, DAY } from "./age.js";

const Field = ({ label, children }) => (
  <div><label>{label}</label>{children}</div>
);
const Card = ({ title, desc, children }) => (
  <div className="card"><h2>{title}</h2><p>{desc}</p>{children}</div>
);
const ColSel = ({ headers, value, info, onChange }) => (
  <>
    <select value={value} onChange={(e) => onChange(e.target.value)}>
      <option value="">— select column —</option>
      {headers.map((o) => <option key={o.i} value={o.i}>{o.label}{o.isDate ? "  📅" : ""}</option>)}
    </select>
    {info && (
      <div className={"note" + (info.bad > 0 || info.good === 0 ? " red" : "")}>
        Detected in file: {info.label}
        {info.bad > 0 && <> · <b>{info.bad} unreadable</b></>}
        {info.bad > 0 && info.samples.length > 0 && <> (e.g. {info.samples.map((s) => `"${s}"`).join(", ")})</>}
        {info.good === 0 && <> — please check that this is a date column</>}
      </div>
    )}
  </>
);
const fmtSize = (n) => (n > 1048576 ? (n / 1048576).toFixed(1) + " MB" : Math.max(1, Math.round(n / 1024)) + " KB");

const EMPTY = {
  openCol: "", openDate: "2026-04-10", openName: "",
  fromCol: "", toCol: "", disName: "DISPOSAL AGE",
  closeCol: "", closeName: "",
};

export default function App() {
  const [buf, setBuf] = useState(null);
  const [fname, setFname] = useState("");
  const [sheets, setSheets] = useState([]);
  const [sheet, setSheet] = useState("");
  const [hrManual, setHrManual] = useState(""); // "" = auto-detect header row
  const [fmt, setFmt] = useState("days");
  const [dfi, setDfi] = useState(0); // index into DATE_FORMATS (0 = DD-MM-YYYY default); "auto" = detect from file (optional)
  const DF = dfi === "auto" ? DATE_FORMATS[0] : DATE_FORMATS[dfi];
  const fmtD = (d) => fmtDate(d, DF);
  const [c, setC] = useState(EMPTY);
  const [res, setRes] = useState(null);
  const [msg, setMsg] = useState(null);
  const [fsize, setFsize] = useState(0);
  const [drag, setDrag] = useState(false);
  const [busy, setBusy] = useState(false);
  const inp = useRef(null);
  const dragN = useRef(0);

  const set = (k, v) => { setC((o) => ({ ...o, [k]: v })); setRes(null); };

  const rows = useMemo(() => {
    if (!buf || !sheet) return [];
    const wb = XLSX.read(buf, { type: "array" });
    return XLSX.utils.sheet_to_json(wb.Sheets[sheet], { header: 1, defval: "" });
  }, [buf, sheet]);
  const autoHr = useMemo(() => detectHeaderRow(rows), [rows]);
  const hr = hrManual === "" ? autoHr : hrManual;
  const cache = useMemo(() => ({}), [rows, hr]);
  const info = (col) => (col === "" ? null : (cache[col] ||= detectColumn(rows.slice(hr).map((r) => r[+col]))));
  const headers = useMemo(() => (rows[hr - 1] || []).map((v, i) => {
    const d = info(i);
    return {
      i, isDate: d.good > 0 && d.good >= d.bad,
      label: `${XLSX.utils.encode_col(i)}: ${String(v).trim() || "(blank)"}`,
    };
  }), [rows, hr]);
  const ord = (col) => (dfi === "auto" ? info(col)?.order || "dmy" : DF[1]);

  // pre-select the usual columns by header name (user can change any of them)
  useEffect(() => {
    if (!headers.length) return;
    const find = (res) => {
      for (const re of res) {
        const h = headers.find((o) => re.test(o.label.split(": ").slice(1).join(": ")));
        if (h) return String(h.i);
      }
      return "";
    };
    const reg = find([/date.*regist/i, /regist/i, /institut|filing|filed/i, /date of reg|\breg\b/i]);
    const dis = find([/date.*(decision|dispos|decid)/i, /decision|dispos|decid/i, /date.*\bdis\b/i]);
    setC((o) => ({ ...o, openCol: reg, closeCol: reg, fromCol: reg, toCol: dis }));
    setRes(null);
  }, [headers]);

  // drag & drop anywhere on the page
  useEffect(() => {
    const has = (e) => e.dataTransfer && Array.from(e.dataTransfer.types || []).includes("Files");
    const enter = (e) => { if (has(e)) { e.preventDefault(); dragN.current++; setDrag(true); } };
    const over = (e) => { if (has(e)) e.preventDefault(); };
    const leave = (e) => { if (has(e) && --dragN.current <= 0) { dragN.current = 0; setDrag(false); } };
    const drop = (e) => {
      if (!has(e)) return;
      e.preventDefault(); dragN.current = 0; setDrag(false);
      onFile(e.dataTransfer.files[0], e.dataTransfer.files.length);
    };
    window.addEventListener("dragenter", enter); window.addEventListener("dragover", over);
    window.addEventListener("dragleave", leave); window.addEventListener("drop", drop);
    return () => {
      window.removeEventListener("dragenter", enter); window.removeEventListener("dragover", over);
      window.removeEventListener("dragleave", leave); window.removeEventListener("drop", drop);
    };
  }, []);

  function clearFile() {
    setBuf(null); setFname(""); setSheets([]); setSheet(""); setFsize(0);
    setHrManual(""); setRes(null); setMsg(null); setC(EMPTY);
    if (inp.current) inp.current.value = "";
  }

  async function onFile(f, count = 1) {
    if (!f) return;
    if (!/\.(xlsx|xlsm|xls|csv)$/i.test(f.name))
      return setMsg({ t: "er", s: `"${f.name}" is not an Excel/CSV file. Please upload .xlsx, .xlsm, .xls or .csv.` });
    if (f.size > 30 * 1048576)
      return setMsg({ t: "er", s: `"${f.name}" is ${fmtSize(f.size)} — files above 30 MB are too large for the browser.` });
    setBusy(true);
    try {
      const b = await f.arrayBuffer();
      const wb = XLSX.read(b, { type: "array" });
      setBuf(b); setFname(f.name); setFsize(f.size); setSheets(wb.SheetNames); setSheet(wb.SheetNames[0]);
      setRes(null); setHrManual("");
      setMsg(count > 1 ? { t: "er", s: `${count} files dropped — only the first one ("${f.name}") is used.` } : null);
      setC((o) => ({ ...o, openCol: "", fromCol: "", toCol: "", closeCol: "" }));
    } catch {
      setMsg({ t: "er", s: "Could not read this file. Upload a valid .xlsx / .xls / .csv." });
    }
    setBusy(false);
  }

  const openDef = "AGE ON DT " + (c.openDate ? fmtD(Date.parse(c.openDate + "T00:00:00Z") / DAY) : "");
  const closeDef = `CLOSING AGE (TODAY ${fmtD(TODAY)})`;

  function run() {
    setMsg(null);
    const wb = XLSX.read(buf, { type: "array", cellNF: true, cellStyles: true });
    const ws = wb.Sheets[sheet];
    const rng = XLSX.utils.decode_range(ws["!ref"]);
    const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" }).slice(hr);
    const g = (r, i) => toDay(r[+i], ord(i), dfi === "auto");
    const cols = [];

    if (c.openCol !== "") {
      const od = Date.parse(c.openDate + "T00:00:00Z") / DAY;
      cols.push({ name: c.openName || openDef, v: data.map((r) => age(g(r, c.openCol), od, fmt)) });
    }
    if (c.fromCol !== "" && c.toCol !== "") {
      cols.push({ name: c.disName || "DISPOSAL AGE", v: data.map((r) => age(g(r, c.fromCol), g(r, c.toCol), fmt)) });
    } else if (c.fromCol !== "" || c.toCol !== "") {
      return setMsg({ t: "er", s: "Disposal age needs BOTH a From column and a To column." });
    }
    if (c.closeCol !== "") {
      cols.push({
        name: c.closeName || closeDef,
        v: data.map((r) => age(g(r, c.closeCol), TODAY, fmt)),
      });
    }
    if (!cols.length) return setMsg({ t: "er", s: "Select at least one age option." });

    const hRow = rng.s.r + hr - 1;
    cols.forEach((col, j) => {
      const cc = rng.e.c + 1 + j;
      ws[XLSX.utils.encode_cell({ r: hRow, c: cc })] = {
        t: "s", v: col.name,
        s: { font: { bold: true }, alignment: { wrapText: true, vertical: "center", horizontal: "center" } },
      };
      (ws["!cols"] ||= [])[cc] = { wch: 16 }; // narrow column so the header wraps onto lines
      col.v.forEach((v, k) => {
        if (v === "") return;
        ws[XLSX.utils.encode_cell({ r: hRow + 1 + k, c: cc })] =
          typeof v === "number" ? { t: "n", v, z: fmt === "days" ? "0" : "0.00" } : { t: "s", v };
      });
      col.filled = col.v.filter((v) => v !== "").length;
    });
    ws["!ref"] = XLSX.utils.encode_range({ s: rng.s, e: { r: rng.e.r, c: rng.e.c + cols.length } });
    const out = XLSX.write(wb, { type: "array", bookType: "xlsx" });
    setRes({ out, cols, total: data.length, name: fname.replace(/\.[^.]+$/, "") + "_AGE.xlsx" });
  }

  function download() {
    const blob = new Blob([res.out], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = res.name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  }

  return (
    <div className="wrap">
      {drag && <div className="overlay">📥 Drop your Excel file to upload</div>}
      <h1>Excel Age Calculator</h1>
      <p className="sub">
        Upload Excel → choose columns → get Opening, Disposal and Closing age → download as Excel.
        Dates like 15-10-2005 (DD-MM-YYYY) are supported.
      </p>

      <Card title="1. Upload Excel file" desc="Your file is processed in your browser only. Nothing is uploaded to a server.">
        <div className={"drop" + (drag ? " over" : "") + (buf ? " has" : "")} role="button" tabIndex={0}
          onClick={() => inp.current.click()}
          onKeyDown={(e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); inp.current.click(); } }}>
          {busy ? <span>⏳ Reading file…</span>
            : buf ? <span>✅ <b>{fname}</b><br /><small>Click, or drop another file anywhere on the page, to replace</small></span>
            : <span><span className="big">📥</span><br />Drag &amp; drop your Excel file here<br /><small>or <b>click to browse</b> · .xlsx .xlsm .xls .csv · processed in your browser</small></span>}
        </div>
        {buf && (
          <div className="chips">
            <span className="chip">📦 {fmtSize(fsize)}</span>
            <span className="chip">📑 {sheets.length} sheet{sheets.length > 1 ? "s" : ""}</span>
            <span className="chip">↕ {Math.max(0, rows.length - hr)} data rows</span>
            <span className="chip">↔ {headers.length} columns</span>
            <button className="chip rm" onClick={clearFile}>✕ Remove file</button>
          </div>
        )}
        <input ref={inp} type="file" accept=".xlsx,.xlsm,.xls,.csv" hidden onChange={(e) => { onFile(e.target.files[0]); e.target.value = ""; }} />
        {buf && (
          <div className="grid" style={{ marginTop: 12 }}>
            <Field label="Sheet">
              <select value={sheet} onChange={(e) => { setSheet(e.target.value); setRes(null); }}>
                {sheets.map((s) => <option key={s}>{s}</option>)}
              </select>
            </Field>
            <Field label="Header row number (optional)">
              <input type="number" min="1" placeholder={`Auto-detected: row ${autoHr}`} value={hrManual}
                onChange={(e) => { setHrManual(e.target.value === "" ? "" : Math.max(1, +e.target.value || 1)); setRes(null); }} />
              <div className="note">
                {hrManual === "" ? `Header row auto-detected: row ${autoHr}. Type a number only to override.`
                                 : `Using row ${hr} (manual). Clear the box to go back to auto-detect.`}
              </div>
            </Field>
          </div>
        )}
      </Card>

      {buf && (
        <Card title="Date format (optional)"
          desc="Default is DD-MM-YYYY. Real Excel dates are always read correctly. If your file has other formats (DD/MM/YYYY, MM/DD/YYYY, YYYY-MM-DD, mixed), pick that format or choose Auto-detect. The note under each date column shows what was found in the file.">
          <div className="grid">
            <Field label="Date format of text dates">
              <select value={dfi} onChange={(e) => { setDfi(e.target.value === "auto" ? "auto" : +e.target.value); setRes(null); }}>
                {DATE_FORMATS.map((f, i) => <option key={f[0]} value={i}>{f[0]}{i === 0 ? " (default)" : ""}</option>)}
                <option value="auto">Auto-detect from file (optional)</option>
              </select>
            </Field>
          </div>
        </Card>
      )}

      {buf && (
        <>
          <Card title="2. Age format" desc="Applied to all ages calculated below.">
            <div className="seg">
              {FORMATS.map(([k, l]) => (
                <button key={k} className={fmt === k ? "on" : ""} onClick={() => { setFmt(k); setRes(null); }}>{l}</button>
              ))}
            </div>
          </Card>

          <Card title="3. Opening age — age on a specific date"
            desc="Age = TO DATE − date in selected column (e.g. Date of Registration). The To date is selected manually.">
            <div className="grid">
              <Field label="Date column"><ColSel headers={headers} value={c.openCol} info={info(c.openCol)} onChange={(v) => set("openCol", v)} /></Field>
              <Field label="To date (select manually)"><input type="date" value={c.openDate} onChange={(e) => set("openDate", e.target.value)} /></Field>
              <Field label="New column name"><input type="text" placeholder={openDef} value={c.openName} onChange={(e) => set("openName", e.target.value)} /></Field>
            </div>
          </Card>

          <Card title="4. Disposal age — from date to date"
            desc="Age = To date − From date. Rows with an empty date stay blank.">
            <div className="grid">
              <Field label="From date column (e.g. Date of Institution)"><ColSel headers={headers} value={c.fromCol} info={info(c.fromCol)} onChange={(v) => set("fromCol", v)} /></Field>
              <Field label="To date column (e.g. Date of Disposal)"><ColSel headers={headers} value={c.toCol} info={info(c.toCol)} onChange={(v) => set("toCol", v)} /></Field>
              <Field label="New column name"><input type="text" value={c.disName} onChange={(e) => set("disName", e.target.value)} /></Field>
            </div>
          </Card>

          <Card title="5. Closing age — age till today"
            desc={`Age = TO DATE − date in selected column (e.g. Date of Registration). The To date is always TODAY, picked automatically: ${fmtD(TODAY)}.`}>
            <div className="grid">
              <Field label="Date column"><ColSel headers={headers} value={c.closeCol} info={info(c.closeCol)} onChange={(v) => set("closeCol", v)} /></Field>
              <Field label="To date (automatic = today)"><input type="date" value={new Date(TODAY * DAY).toISOString().slice(0, 10)} disabled /></Field>
              <Field label="New column name"><input type="text" placeholder={closeDef} value={c.closeName} onChange={(e) => set("closeName", e.target.value)} /></Field>
            </div>
          </Card>

          <div className="row">
            <button className="btn" onClick={run}>Process</button>
            {res && <button className="btn ok" onClick={download}>⬇ Download Excel ({res.name})</button>}
          </div>
          {msg && <div className={`msg ${msg.t}`}>{msg.s}</div>}

          {res && (
            <Card title="Preview" desc={`${res.total} rows processed. Original columns are kept; new age columns are added at the end.`}>
              <div className="tw">
                <table>
                  <thead><tr>{res.cols.map((x) => <th key={x.name}>{x.name} ({x.filled}/{res.total})</th>)}</tr></thead>
                  <tbody>
                    {Array.from({ length: Math.min(8, res.total) }, (_, k) => (
                      <tr key={k}>{res.cols.map((x) => <td key={x.name}>{String(x.v[k])}</td>)}</tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </Card>
          )}
        </>
      )}
    </div>
  );
}
