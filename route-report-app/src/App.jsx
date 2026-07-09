import { useState, useEffect, useRef, useMemo } from "react";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";

const DEFAULT_TEMPLATE = {
  reportTitle: "Last Visit Outlet Wise Quantity Sales",
  territoryLabel: "Territory",
  routeLabel: "Route",
  totalLabel: "Total sale value",
  dateLabel: "Last visit date",
  colNo: "NO",
  colOutlet: "Outlet Name",
  colProduct: "Product Name",
  colQty: "Quantity",
  accent: "#7A2E33",
  bandingEnabled: true,
  bandingColor: "#F3E3E1",
};

const BANDING_COLORS = [
  { name: "Blush Cherry", value: "#F3E3E1" },
  { name: "Warm Ivory", value: "#F4EFE6" },
  { name: "Sage Whisper", value: "#E7EFE5" },
  { name: "Powder Sky", value: "#E6EEF5" },
  { name: "Soft Lilac", value: "#EFE9F5" },
  { name: "Pale Sand", value: "#F5EFE0" },
  { name: "Ash Gray", value: "#ECEAE7" },
];

const ACCENTS = [
  { name: "Cherry Alloy", value: "#7A2E33" },
  { name: "Stone Charcoal", value: "#4A4743" },
  { name: "Deep Navy", value: "#2E4A7A" },
  { name: "Industrial Forest", value: "#2E7A4E" },
  { name: "Legacy Bronze", value: "#7A5E2E" },
  { name: "Blush Rose", value: "#C97B84" },
  { name: "Dusty Terracotta", value: "#C08552" },
  { name: "Sage Mist", value: "#8FA98C" },
  { name: "Powder Blue", value: "#7FA6C4" },
  { name: "Warm Sand", value: "#D9B98A" },
  { name: "Soft Lilac", value: "#A995C9" },
];

const VISIT_BLOCKS = [
  { key: "1st", label: "PREVIOUS 1ST VISIT", start: 31 },
  { key: "2nd", label: "PREVIOUS 2ND VISIT", start: 24 },
  { key: "3rd", label: "PREVIOUS 3RD VISIT", start: 17 },
  { key: "4th", label: "PREVIOUS 4TH VISIT", start: 10 },
];
const OFF = { qty: 0, invoice: 1, retQty: 2, retVal: 3, freeQty: 4, freeVal: 5, net: 6 };
const PAGE_SIZE = 4;

function num(v) {
  const n = typeof v === "number" ? v : parseFloat(v);
  return isNaN(n) ? 0 : n;
}

function parseMasterReport(rows) {
  let headerRow = -1;
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i] || [];
    if (String(r[0] || "").trim().toUpperCase() === "NO" && String(r[4] || "").trim().toUpperCase() === "TERRITORY") {
      headerRow = i;
      break;
    }
  }
  if (headerRow === -1) throw new Error("Couldn't find the header row (expecting a 'NO' / 'TERRITORY' row). Is this a Route Wise Item Wise Outlet Wise Sales Report export?");

  const dataStart = headerRow + 2;
  const filled = { no: null, rsm: null, asm: null, ase: null, territory: null, routeNumber: null, routeName: null, outletId: null, outletName: null };
  const out = [];

  for (let i = dataStart; i < rows.length; i++) {
    const r = rows[i] || [];
    const product = r[9];
    if (r.every((c) => c === null || c === undefined || c === "")) continue;
    if (!product) continue;

    if (r[0] !== null && r[0] !== undefined && r[0] !== "") filled.no = r[0];
    if (r[4]) filled.territory = String(r[4]).trim();
    if (r[6]) filled.routeName = String(r[6]).trim();
    if (r[8]) filled.outletName = String(r[8]).trim();
    if (r[7]) filled.outletId = r[7];
    if (r[1]) filled.rsm = r[1];
    if (r[2]) filled.asm = r[2];
    if (r[3]) filled.ase = r[3];

    if (!filled.territory || !filled.outletName) continue;

    const visits = {};
    VISIT_BLOCKS.forEach((b) => {
      visits[b.key] = { qty: num(r[b.start + OFF.qty]), net: num(r[b.start + OFF.net]) };
    });

    out.push({
      territory: filled.territory,
      routeName: filled.routeName || "",
      outletId: filled.outletId,
      outletName: filled.outletName,
      product: String(product).trim(),
      visits,
    });
  }
  return out;
}

function resolveLastVisit(row, strategy) {
  const order = strategy === "strict" ? ["1st"] : ["1st", "2nd", "3rd", "4th"];
  for (const key of order) {
    const v = row.visits[key];
    if (v && v.qty > 0) return { slot: key, qty: v.qty, net: v.net };
  }
  return null;
}

function buildReport(rows, territory, routeName, strategy) {
  const filtered = rows.filter((r) => r.territory === territory && (r.routeName || "") === (routeName || ""));
  const outletOrder = [];
  const outletMap = new Map();
  let totalSale = 0;
  let matchedRows = 0;

  filtered.forEach((row) => {
    const resolved = resolveLastVisit(row, strategy);
    if (!resolved) return;
    matchedRows++;
    totalSale += resolved.net;
    if (!outletMap.has(row.outletName)) {
      outletMap.set(row.outletName, []);
      outletOrder.push(row.outletName);
    }
    outletMap.get(row.outletName).push({ product: row.product, qty: resolved.qty, slot: resolved.slot });
  });

  const outlets = outletOrder.map((name) => ({ name, items: outletMap.get(name) }));
  return { outlets, totalSale, matchedRows, consideredRows: filtered.length };
}

// ---------- shared visual tokens ----------
const T = {
  bg: "#FCF9F4",
  surface: "#FFFFFF",
  surfaceLow: "#F6F3EE",
  surfaceContainer: "#F0EDE9",
  border: "#DAC1C0",
  text: "#1C1C19",
  textMuted: "#544242",
  headFont: "'Manrope', sans-serif",
  bodyFont: "'Work Sans', sans-serif",
  capsFont: "'IBM Plex Sans', sans-serif",
};

function Icon({ name, size = 20, style }) {
  return (
    <span className="material-symbols-outlined" style={{ fontSize: size, lineHeight: 1, ...style }}>
      {name}
    </span>
  );
}

function Label({ children }) {
  return (
    <span
      className="block uppercase mb-1.5"
      style={{ fontFamily: T.capsFont, fontSize: 11, letterSpacing: "0.1em", fontWeight: 600, color: T.textMuted }}
    >
      {children}
    </span>
  );
}

function TextField({ value, onChange, accent }) {
  return (
    <input
      className="w-full rounded-md px-3 py-2 text-sm outline-none transition-all"
      style={{ background: T.surface, border: `1px solid ${T.border}`, color: T.text, fontFamily: T.bodyFont }}
      value={value}
      onChange={(e) => onChange(e.target.value)}
      onFocus={(e) => { e.target.style.boxShadow = `0 0 0 2px ${accent}55`; e.target.style.borderColor = accent; }}
      onBlur={(e) => { e.target.style.boxShadow = "none"; e.target.style.borderColor = T.border; }}
    />
  );
}

function Select({ value, onChange, children }) {
  return (
    <select
      className="w-full rounded-md px-2.5 py-2 text-sm outline-none"
      style={{ background: T.surface, border: `1px solid ${T.border}`, color: T.text, fontFamily: T.bodyFont }}
      value={value}
      onChange={onChange}
    >
      {children}
    </select>
  );
}

const MONTH_NAMES = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const WEEKDAY_LETTERS = ["S", "M", "T", "W", "T", "F", "S"];

function toISODate(d) {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function formatDisplayDate(iso) {
  if (!iso) return "";
  const [y, m, d] = iso.split("-").map(Number);
  const dt = new Date(y, m - 1, d);
  return dt.toLocaleDateString(undefined, { day: "numeric", month: "short", year: "numeric" });
}

function DatePicker({ value, onChange, accent }) {
  const [open, setOpen] = useState(false);
  const [viewDate, setViewDate] = useState(() => {
    if (value) {
      const [y, m] = value.split("-").map(Number);
      return new Date(y, m - 1, 1);
    }
    return new Date();
  });
  const wrapRef = useRef(null);

  useEffect(() => {
    function handleClick(e) {
      if (wrapRef.current && !wrapRef.current.contains(e.target)) setOpen(false);
    }
    document.addEventListener("mousedown", handleClick);
    return () => document.removeEventListener("mousedown", handleClick);
  }, []);

  const year = viewDate.getFullYear();
  const month = viewDate.getMonth();
  const firstWeekday = new Date(year, month, 1).getDay();
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const cells = [];
  for (let i = 0; i < firstWeekday; i++) cells.push(null);
  for (let d = 1; d <= daysInMonth; d++) cells.push(d);

  const today = new Date();
  const todayISO = toISODate(today);

  return (
    <div ref={wrapRef} className="relative">
      <button
        type="button"
        onClick={() => setOpen((o) => !o)}
        className="w-full flex items-center justify-between rounded-md px-3 py-2 text-sm text-left outline-none"
        style={{ background: T.surface, border: `1px solid ${T.border}`, color: value ? T.text : T.textMuted, fontFamily: T.bodyFont }}
      >
        <span>{value ? formatDisplayDate(value) : "Select a date"}</span>
        <Icon name="calendar_month" size={17} style={{ color: accent }} />
      </button>

      {open && (
        <div
          className="absolute z-20 mt-2 rounded-lg p-3"
          style={{ background: T.surface, border: `1px solid ${T.border}`, boxShadow: "0px 8px 20px rgba(74,71,67,0.15)", width: 268 }}
        >
          <div className="flex items-center justify-between mb-2">
            <button type="button" onClick={() => setViewDate(new Date(year, month - 1, 1))} className="w-7 h-7 rounded-md flex items-center justify-center" style={{ color: T.textMuted }}>
              <Icon name="chevron_left" size={18} />
            </button>
            <span style={{ fontFamily: T.headFont, fontWeight: 700, fontSize: 13, color: T.text }}>{MONTH_NAMES[month]} {year}</span>
            <button type="button" onClick={() => setViewDate(new Date(year, month + 1, 1))} className="w-7 h-7 rounded-md flex items-center justify-center" style={{ color: T.textMuted }}>
              <Icon name="chevron_right" size={18} />
            </button>
          </div>

          <div className="grid grid-cols-7 gap-y-1 mb-1">
            {WEEKDAY_LETTERS.map((w, i) => (
              <div key={i} className="text-center" style={{ fontFamily: T.capsFont, fontSize: 10, color: T.textMuted }}>{w}</div>
            ))}
          </div>

          <div className="grid grid-cols-7 gap-y-1">
            {cells.map((d, i) => {
              if (d === null) return <div key={i} />;
              const iso = toISODate(new Date(year, month, d));
              const isSelected = iso === value;
              const isToday = iso === todayISO;
              return (
                <button
                  type="button"
                  key={i}
                  onClick={() => { onChange(iso); setOpen(false); }}
                  className="w-8 h-8 rounded-full text-sm mx-auto flex items-center justify-center transition-colors"
                  style={{
                    background: isSelected ? accent : "transparent",
                    color: isSelected ? "#fff" : T.text,
                    fontWeight: isSelected || isToday ? 700 : 400,
                    border: isToday && !isSelected ? `1px solid ${accent}` : "none",
                    fontFamily: T.bodyFont,
                  }}
                >
                  {d}
                </button>
              );
            })}
          </div>

          <div className="flex justify-between mt-3 pt-2" style={{ borderTop: `1px solid ${T.border}` }}>
            <button type="button" onClick={() => { onChange(""); setOpen(false); }} className="text-xs" style={{ color: T.textMuted, fontFamily: T.bodyFont }}>Clear</button>
            <button type="button" onClick={() => { const iso = toISODate(today); onChange(iso); setViewDate(today); setOpen(false); }} className="text-xs font-semibold" style={{ color: accent, fontFamily: T.bodyFont }}>Today</button>
          </div>
        </div>
      )}
    </div>
  );
}

function TemplateEditor({ template, setTemplate, onReset, previewData }) {
  const [saved, setSaved] = useState(false);
  const field = (key) => ({ value: template[key], onChange: (v) => setTemplate({ ...template, [key]: v }), accent: template.accent });

  function handleSave() {
    setSaved(true);
    setTimeout(() => setSaved(false), 1500);
  }

  return (
    <div className="grid grid-cols-1 lg:grid-cols-12 gap-5">
      <section className="lg:col-span-5 flex flex-col gap-5">
        <div className="rounded-lg p-6" style={{ background: T.surface, border: `1px solid ${T.border}`, boxShadow: "0px 4px 12px rgba(74,71,67,0.05)" }}>
          <div className="pb-4 mb-6" style={{ borderBottom: `1px solid ${T.border}` }}>
            <h2 style={{ fontFamily: T.headFont, fontSize: 18, fontWeight: 700, color: template.accent }}>Report Configuration</h2>
            <p className="text-sm mt-1" style={{ color: T.textMuted, fontFamily: T.bodyFont }}>Customize your template's layout and metadata labels.</p>
          </div>

          <div className="space-y-5">
            <div>
              <Label>Report title</Label>
              <TextField {...field("reportTitle")} />
            </div>
            <div className="grid grid-cols-2 gap-4">
              <div><Label>Territory label</Label><TextField {...field("territoryLabel")} /></div>
              <div><Label>Route label</Label><TextField {...field("routeLabel")} /></div>
            </div>
            <div className="grid grid-cols-2 gap-4">
              <div><Label>Total sale label</Label><TextField {...field("totalLabel")} /></div>
              <div><Label>Last visit label</Label><TextField {...field("dateLabel")} /></div>
            </div>
          </div>

          <div className="pt-5">
            <Label>Table column headers</Label>
            <div className="grid grid-cols-2 gap-4">
              <TextField {...field("colNo")} />
              <TextField {...field("colOutlet")} />
              <TextField {...field("colProduct")} />
              <TextField {...field("colQty")} />
            </div>
          </div>

          <div className="pt-5 mt-5" style={{ borderTop: `1px solid ${T.border}` }}>
            <Label>Brand accent color</Label>
            <div className="flex items-center gap-3 flex-wrap">
              {ACCENTS.map((a) => (
                <button
                  key={a.value}
                  onClick={() => setTemplate({ ...template, accent: a.value })}
                  title={a.name}
                  className="w-8 h-8 rounded-full cursor-pointer transition-transform hover:scale-110"
                  style={{
                    backgroundColor: a.value,
                    border: `1px solid ${T.border}`,
                    outline: template.accent === a.value ? `2px solid ${T.text}` : "none",
                    outlineOffset: "2px",
                  }}
                />
              ))}
            </div>
          </div>

          <div className="pt-5 mt-5" style={{ borderTop: `1px solid ${T.border}` }}>
            <div className="flex items-center justify-between mb-2">
              <Label>Outlet row banding</Label>
              <button
                type="button"
                onClick={() => setTemplate({ ...template, bandingEnabled: !template.bandingEnabled })}
                className="relative w-10 h-5.5 rounded-full transition-colors shrink-0"
                style={{ background: template.bandingEnabled ? template.accent : T.border, height: 22, width: 40 }}
                title={template.bandingEnabled ? "Banding on" : "Banding off"}
              >
                <span
                  className="absolute top-0.5 rounded-full bg-white transition-transform"
                  style={{ width: 18, height: 18, left: 2, transform: template.bandingEnabled ? "translateX(18px)" : "translateX(0)" }}
                />
              </button>
            </div>
            <p className="text-xs mb-3" style={{ color: T.textMuted }}>Shades every other outlet's rows so its products are easier to scan. Applies to the table, the Excel export, and the PDF export.</p>
            {template.bandingEnabled && (
              <div className="flex items-center gap-3 flex-wrap">
                {BANDING_COLORS.map((c) => (
                  <button
                    key={c.value}
                    onClick={() => setTemplate({ ...template, bandingColor: c.value })}
                    title={c.name}
                    className="w-8 h-8 rounded-full cursor-pointer transition-transform hover:scale-110"
                    style={{
                      backgroundColor: c.value,
                      border: `1px solid ${T.border}`,
                      outline: template.bandingColor === c.value ? `2px solid ${T.text}` : "none",
                      outlineOffset: "2px",
                    }}
                  />
                ))}
              </div>
            )}
          </div>

          <div className="pt-6 flex items-center justify-between">
            <button onClick={onReset} className="flex items-center gap-1.5 text-sm transition-colors" style={{ color: T.textMuted, fontFamily: T.bodyFont }}>
              <Icon name="restart_alt" size={18} /> Reset to defaults
            </button>
            <button
              onClick={handleSave}
              className="px-6 py-2.5 rounded-lg font-bold text-sm transition-all active:scale-[0.98]"
              style={{ background: template.accent, color: "#fff", fontFamily: T.bodyFont, boxShadow: "0 2px 6px rgba(0,0,0,0.15)" }}
            >
              {saved ? "Saved ✓" : "Save Template"}
            </button>
          </div>
        </div>
      </section>

      <section className="lg:col-span-7">
        <div className="sticky top-5">
          <Label>Live preview (real-time)</Label>
          <div className="rounded-lg overflow-hidden flex flex-col min-h-[560px]" style={{ background: T.surface, border: "1px solid #D9D1C7", boxShadow: "0px 12px 24px rgba(74,71,67,0.12)" }}>
            <div className="p-6" style={{ background: template.accent, color: "#fff" }}>
              <div className="flex justify-between items-start">
                <div>
                  <h3 style={{ fontFamily: T.headFont, fontSize: 22, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.02em" }}>{template.reportTitle}</h3>
                  <p className="text-sm opacity-80 mt-1" style={{ fontFamily: T.bodyFont }}>Outlet-wise breakdown for the most recent visit</p>
                </div>
                <div className="text-right">
                  <p style={{ fontFamily: T.capsFont, fontSize: 10, letterSpacing: "0.1em", opacity: 0.75 }}>{template.dateLabel.toUpperCase()}</p>
                  <p className="font-bold" style={{ fontFamily: T.bodyFont }}>{previewData.lastVisitDate || "—"}</p>
                </div>
              </div>
            </div>

            <div className="p-6 flex-1">
              <div className="grid grid-cols-2 gap-6 pb-6 mb-6" style={{ borderBottom: `1px solid ${T.border}` }}>
                <div className="space-y-4">
                  <div>
                    <p style={{ fontFamily: T.capsFont, fontSize: 10, letterSpacing: "0.08em", color: T.textMuted }}>{template.territoryLabel.toUpperCase()}</p>
                    <p className="font-semibold" style={{ fontFamily: T.bodyFont }}>{previewData.territory || "—"}</p>
                  </div>
                  <div>
                    <p style={{ fontFamily: T.capsFont, fontSize: 10, letterSpacing: "0.08em", color: T.textMuted }}>{template.routeLabel.toUpperCase()}</p>
                    <p className="font-semibold" style={{ fontFamily: T.bodyFont }}>{previewData.routeName || "—"}</p>
                  </div>
                </div>
                <div className="space-y-4">
                  <div>
                    <p style={{ fontFamily: T.capsFont, fontSize: 10, letterSpacing: "0.08em", color: T.textMuted }}>{template.totalLabel.toUpperCase()}</p>
                    <p className="font-bold" style={{ fontFamily: T.headFont, fontSize: 18, color: template.accent }}>{previewData.totalSale}</p>
                  </div>
                </div>
              </div>

              <div className="space-y-2">
                <h4 style={{ fontFamily: T.capsFont, fontSize: 11, letterSpacing: "0.1em", color: T.textMuted }}>OUTLET DATA SUMMARY</h4>
                <table className="w-full border-collapse text-sm" style={{ fontFamily: T.bodyFont }}>
                  <thead>
                    <tr style={{ background: T.surfaceContainer, borderTop: `2px solid ${template.accent}`, textAlign: "left" }}>
                      <th className="p-2.5" style={{ fontFamily: T.capsFont, fontSize: 10 }}>{template.colNo}</th>
                      <th className="p-2.5" style={{ fontFamily: T.capsFont, fontSize: 10 }}>{template.colOutlet}</th>
                      <th className="p-2.5" style={{ fontFamily: T.capsFont, fontSize: 10 }}>{template.colProduct}</th>
                      <th className="p-2.5 text-right" style={{ fontFamily: T.capsFont, fontSize: 10 }}>{template.colQty}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {previewData.sampleRows.map((row, i) => (
                      <tr key={i} style={{ borderBottom: `1px solid ${T.border}` }}>
                        <td className="p-2.5">{row.no}</td>
                        <td className="p-2.5">{row.outlet}</td>
                        <td className="p-2.5">{row.product}</td>
                        <td className="p-2.5 text-right">{row.qty}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
          <p className="text-xs mt-3" style={{ color: T.textMuted, fontFamily: T.bodyFont }}>Saved automatically, and used every time you generate a report on the Dashboard.</p>
        </div>
      </section>
    </div>
  );
}

export default function App() {
  const [screen, setScreen] = useState("dashboard");
  const [template, setTemplate] = useState(DEFAULT_TEMPLATE);
  const [templateLoaded, setTemplateLoaded] = useState(false);

  const [uploads, setUploads] = useState([]);
  const [activeId, setActiveId] = useState(null);
  const [page, setPage] = useState(0);

  const fileInputRef = useRef(null);
  const uploadCounter = useRef(0);

  useEffect(() => {
    try {
      const saved = localStorage.getItem("template-config");
      if (saved) setTemplate({ ...DEFAULT_TEMPLATE, ...JSON.parse(saved) });
    } catch (e) {}
    finally { setTemplateLoaded(true); }
  }, []);

  useEffect(() => {
    if (!templateLoaded) return;
    try { localStorage.setItem("template-config", JSON.stringify(template)); } catch (e) {}
  }, [template, templateLoaded]);

  const activeUpload = uploads.find((u) => u.id === activeId) || null;

  useEffect(() => { setPage(0); }, [activeId, activeUpload?.territory, activeUpload?.routeName, activeUpload?.strategy]);

  function territoriesOf(upload) {
    if (!upload || !upload.rows) return [];
    return [...new Set(upload.rows.map((r) => r.territory))].filter(Boolean).sort();
  }

  function routesOf(upload, territory) {
    if (!upload || !upload.rows || !territory) return [];
    return [...new Set(upload.rows.filter((r) => r.territory === territory).map((r) => r.routeName))].filter((x) => x !== undefined).sort();
  }

  const territories = useMemo(() => territoriesOf(activeUpload), [activeUpload]);
  const routesForTerritory = useMemo(() => routesOf(activeUpload, activeUpload?.territory), [activeUpload]);

  const report = useMemo(() => {
    if (!activeUpload || !activeUpload.rows || !activeUpload.territory) return null;
    return buildReport(activeUpload.rows, activeUpload.territory, activeUpload.routeName, activeUpload.strategy);
  }, [activeUpload]);

  const pageCount = report ? Math.max(1, Math.ceil(report.outlets.length / PAGE_SIZE)) : 1;
  const pagedOutlets = report ? report.outlets.slice(page * PAGE_SIZE, page * PAGE_SIZE + PAGE_SIZE) : [];

  function updateUpload(id, patch) {
    setUploads((prev) => prev.map((u) => (u.id === id ? { ...u, ...patch } : u)));
  }

  function removeUpload(id) {
    setUploads((prev) => prev.filter((u) => u.id !== id));
    if (activeId === id) {
      setActiveId((prevActive) => {
        const remaining = uploads.filter((u) => u.id !== id);
        return remaining.length ? remaining[0].id : null;
      });
    }
  }

  function handleFiles(fileList) {
    const files = Array.from(fileList || []);
    files.forEach((file) => {
      uploadCounter.current += 1;
      const id = `upload-${Date.now()}-${uploadCounter.current}`;
      const draft = { id, fileName: file.name, rows: null, parseError: null, territory: "", routeName: "", lastVisitDate: "", strategy: "fallback" };
      setUploads((prev) => [...prev, draft]);
      setActiveId((prev) => prev || id);

      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const wb = XLSX.read(e.target.result, { type: "array" });
          const sheet = wb.Sheets[wb.SheetNames[0]];
          const arr = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true });
          const parsed = parseMasterReport(arr);
          if (parsed.length === 0) throw new Error("No outlet/product rows found in this file.");
          const firstTerritory = [...new Set(parsed.map((r) => r.territory))].sort()[0] || "";
          const firstRoute = [...new Set(parsed.filter((r) => r.territory === firstTerritory).map((r) => r.routeName))].sort()[0] || "";
          updateUpload(id, { rows: parsed, parseError: null, territory: firstTerritory, routeName: firstRoute });
        } catch (err) {
          updateUpload(id, { rows: null, parseError: err.message || "Couldn't read this file." });
        }
      };
      reader.onerror = () => updateUpload(id, { rows: null, parseError: "The file could not be read." });
      reader.readAsArrayBuffer(file);
    });
  }

  function resetActiveFilters() {
    if (!activeUpload) return;
    const firstTerritory = territoriesOf(activeUpload)[0] || "";
    const firstRoute = routesOf(activeUpload, firstTerritory)[0] || "";
    updateUpload(activeUpload.id, { territory: firstTerritory, routeName: firstRoute, lastVisitDate: "", strategy: "fallback" });
  }

  async function exportExcelFor(upload, rep) {
    if (!rep) return;
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Report");
    ws.columns = [{ width: 6 }, { width: 30 }, { width: 44 }, { width: 12 }];

    const accentARGB = "FF" + template.accent.replace("#", "").toUpperCase();
    const thin = { style: "thin", color: { argb: "FFCBBFAE" } };
    const borderAll = { top: thin, left: thin, bottom: thin, right: thin };

    ws.mergeCells("A1:D1");
    const titleCell = ws.getCell("A1");
    titleCell.value = template.reportTitle;
    titleCell.font = { bold: true, size: 14, color: { argb: "FFFFFFFF" } };
    titleCell.alignment = { horizontal: "center", vertical: "middle" };
    titleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: accentARGB } };
    ws.getRow(1).height = 26;

    const metaRows = [
      [template.territoryLabel, upload.territory],
      [template.routeLabel, upload.routeName],
      [template.totalLabel, Math.round(rep.totalSale * 100) / 100],
      [template.dateLabel, upload.lastVisitDate || ""],
    ];
    metaRows.forEach(([label, value], i) => {
      const rowNum = i + 2;
      const labelCell = ws.getCell(`A${rowNum}`);
      labelCell.value = label;
      labelCell.font = { bold: true };
      ws.mergeCells(`B${rowNum}:D${rowNum}`);
      ws.getCell(`B${rowNum}`).value = value;
    });

    const headerRowNum = 7;
    const headers = [template.colNo, template.colOutlet, template.colProduct, template.colQty];
    headers.forEach((h, i) => {
      const cell = ws.getCell(headerRowNum, i + 1);
      cell.value = h;
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: accentARGB } };
      cell.alignment = { vertical: "middle", horizontal: i === 3 ? "right" : "left" };
      cell.border = borderAll;
    });

    const bandARGB = "FF" + template.bandingColor.replace("#", "").toUpperCase();
    let r = headerRowNum;
    rep.outlets.forEach((outlet, idx) => {
      const startRow = r + 1;
      const banded = template.bandingEnabled && idx % 2 === 1;
      outlet.items.forEach((item, i) => {
        r += 1;
        const row = ws.getRow(r);
        row.getCell(1).value = i === 0 ? idx + 1 : null;
        row.getCell(2).value = i === 0 ? outlet.name : null;
        row.getCell(3).value = item.product;
        row.getCell(4).value = item.qty;
        [1, 2, 3, 4].forEach((c) => {
          row.getCell(c).border = borderAll;
          if (banded) row.getCell(c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: bandARGB } };
        });
        row.getCell(4).alignment = { horizontal: "right" };
        row.getCell(3).alignment = { vertical: "middle" };
      });
      const endRow = r;
      if (outlet.items.length > 1) {
        ws.mergeCells(startRow, 1, endRow, 1);
        ws.mergeCells(startRow, 2, endRow, 2);
      }
      ws.getCell(startRow, 1).alignment = { vertical: "middle", horizontal: "center" };
      ws.getCell(startRow, 2).alignment = { vertical: "middle" };
    });

    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: "application/octet-stream" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    const safeRoute = (upload.routeName || upload.territory).replace(/[^a-z0-9]+/gi, "_").slice(0, 40);
    a.href = url;
    a.download = `${upload.territory}_${safeRoute}_LastVisit.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
  }

  function exportPdfFor(upload, rep) {
    if (!rep) return;
    const win = window.open("", "_blank");
    const rowsHtml = rep.outlets.map((outlet, idx) => {
      const banded = template.bandingEnabled && idx % 2 === 1;
      const bg = banded ? ` style="background:${template.bandingColor}"` : "";
      return outlet.items.map((item, i) => `
      <tr${bg}>
        <td>${i === 0 ? idx + 1 : ""}</td>
        <td>${i === 0 ? outlet.name : ""}</td>
        <td>${item.product}</td>
        <td style="text-align:right">${item.qty}</td>
      </tr>`).join("");
    }).join("");

    win.document.write(`
      <html><head><title>${upload.territory} - ${upload.routeName}</title>
      <style>
        body{font-family:${T.bodyFont};color:${T.text};padding:32px;}
        h1{font-family:${T.headFont};text-align:center;color:white;background:${template.accent};padding:14px;border-radius:6px;font-size:18px;margin-bottom:20px;}
        table{width:100%;border-collapse:collapse;margin-top:16px;font-size:12px;}
        th{text-align:left;border-bottom:2px solid ${template.accent};color:${template.accent};padding:6px 8px;}
        td{padding:5px 8px;border-bottom:1px solid #eee;}
        .meta{display:grid;grid-template-columns:160px 1fr;row-gap:6px;font-size:13px;}
        .meta b{color:${T.textMuted};}
        @media print{ body{padding:0;} }
      </style></head><body>
      <h1>${template.reportTitle}</h1>
      <div class="meta">
        <b>${template.territoryLabel}</b><span>${upload.territory}</span>
        <b>${template.routeLabel}</b><span>${upload.routeName}</span>
        <b>${template.totalLabel}</b><span>${rep.totalSale.toLocaleString(undefined, { minimumFractionDigits: 2 })}</span>
        <b>${template.dateLabel}</b><span>${upload.lastVisitDate || "________________"}</span>
      </div>
      <table><thead><tr><th>${template.colNo}</th><th>${template.colOutlet}</th><th>${template.colProduct}</th><th style="text-align:right">${template.colQty}</th></tr></thead>
      <tbody>${rowsHtml}</tbody></table>
      </body></html>`);
    win.document.close();
    win.focus();
    setTimeout(() => win.print(), 350);
  }

  function exportExcel() { if (activeUpload && report) exportExcelFor(activeUpload, report); }
  function exportPdf() { if (activeUpload && report) exportPdfFor(activeUpload, report); }

  async function exportAllExcel() {
    for (const upload of uploads) {
      if (!upload.rows || !upload.territory) continue;
      const rep = buildReport(upload.rows, upload.territory, upload.routeName, upload.strategy);
      if (rep.outlets.length === 0) continue;
      await exportExcelFor(upload, rep);
      await new Promise((res) => setTimeout(res, 250));
    }
  }

  const previewData = report
    ? {
        territory: activeUpload.territory,
        routeName: activeUpload.routeName,
        lastVisitDate: activeUpload.lastVisitDate,
        totalSale: report.totalSale.toLocaleString(undefined, { minimumFractionDigits: 2 }),
        sampleRows: report.outlets.slice(0, 2).flatMap((o, oi) =>
          o.items.slice(0, oi === 0 ? 2 : 1).map((it, i) => ({ no: i === 0 ? oi + 1 : "", outlet: i === 0 ? o.name : "", product: it.product, qty: it.qty }))
        ),
      }
    : { territory: "Nikaweratiya", routeName: "Anamaduwa to Galkulama", lastVisitDate: "", totalSale: "110,409.00", sampleRows: [{ no: 1, outlet: "Janaka Karawala Kade", product: "Black Chicken 90g", qty: 15 }, { no: "", outlet: "", product: "Sago Seed 100g", qty: 20 }] };

  const canExport = Boolean(report && report.outlets.length > 0);
  const exportableCount = uploads.filter((u) => u.rows && u.territory && buildReport(u.rows, u.territory, u.routeName, u.strategy).outlets.length > 0).length;

  return (
    <div className="min-h-screen" style={{ background: T.bg, fontFamily: T.bodyFont, color: T.text }}>
      <header style={{ background: T.surface, borderBottom: `1px solid ${T.border}`, boxShadow: "0 1px 2px rgba(0,0,0,0.03)" }}>
        <div className="max-w-[1280px] mx-auto px-6 h-16 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <img src="/logo (1).png" alt="Lakmee" className="h-9 w-auto object-contain" />
            <h1 style={{ fontFamily: T.headFont, fontSize: 18, fontWeight: 700, color: template.accent }}>Route Sales Report Builder</h1>
          </div>
          <nav className="rounded-full p-1 flex items-center gap-1" style={{ background: T.surfaceLow, boxShadow: "inset 0 1px 2px rgba(0,0,0,0.04)" }}>
            <button
              onClick={() => setScreen("dashboard")}
              className="px-4 py-1.5 rounded-full text-sm font-semibold transition-colors"
              style={{ fontFamily: T.capsFont, fontSize: 12, letterSpacing: "0.06em", background: screen === "dashboard" ? template.accent : "transparent", color: screen === "dashboard" ? "#fff" : T.textMuted }}
            >
              DASHBOARD
            </button>
            <button
              onClick={() => setScreen("template")}
              className="px-4 py-1.5 rounded-full text-sm font-semibold transition-colors"
              style={{ fontFamily: T.capsFont, fontSize: 12, letterSpacing: "0.06em", background: screen === "template" ? template.accent : "transparent", color: screen === "template" ? "#fff" : T.textMuted }}
            >
              TEMPLATE EDITOR
            </button>
          </nav>
          <div className="flex items-center gap-4">
            <button onClick={exportExcel} disabled={!canExport} title="Download Excel" style={{ color: template.accent, opacity: canExport ? 1 : 0.35, cursor: canExport ? "pointer" : "default" }}>
              <Icon name="download" />
            </button>
            <button onClick={exportPdf} disabled={!canExport} title="Download PDF" style={{ color: template.accent, opacity: canExport ? 1 : 0.35, cursor: canExport ? "pointer" : "default" }}>
              <Icon name="picture_as_pdf" />
            </button>
            <div className="w-8 h-8 rounded-full flex items-center justify-center text-white text-xs font-bold" style={{ background: template.accent, border: `1px solid ${T.border}` }}>RT</div>
          </div>
        </div>
      </header>

      <main className="max-w-[1280px] mx-auto px-6 py-8">
        {screen === "template" ? (
          <TemplateEditor template={template} setTemplate={setTemplate} onReset={() => setTemplate(DEFAULT_TEMPLATE)} previewData={previewData} />
        ) : (
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-5">
            <div className="lg:col-span-4 flex flex-col gap-5">
              <div className="rounded-lg p-5" style={{ background: T.surface, border: `1px solid ${T.border}` }}>
                <Label>Step 1: Upload master report(s)</Label>
                <div
                  onClick={() => fileInputRef.current?.click()}
                  onDrop={(e) => { e.preventDefault(); if (e.dataTransfer.files.length) handleFiles(e.dataTransfer.files); }}
                  onDragOver={(e) => e.preventDefault()}
                  className="rounded-lg p-6 text-center cursor-pointer transition-colors"
                  style={{ border: `2px dashed ${T.border}` }}
                >
                  <Icon name="cloud_upload" size={30} style={{ color: T.textMuted }} />
                  <p className="text-sm font-semibold mt-2">Drop files here or click</p>
                  <p className="text-xs mt-1" style={{ color: T.textMuted }}>You can select multiple Route Wise Item Wise Outlet Wise Sales Reports (.xlsx) at once</p>
                  <input ref={fileInputRef} type="file" accept=".xlsx,.xls" multiple className="hidden" onChange={(e) => { if (e.target.files.length) handleFiles(e.target.files); e.target.value = ""; }} />
                </div>
                {uploads.length > 0 && (
                  <p className="mt-3 text-sm flex items-center gap-1.5" style={{ color: "#2E7A4E" }}><Icon name="check_circle" size={16} />{uploads.length} file{uploads.length === 1 ? "" : "s"} uploaded</p>
                )}
              </div>

              {uploads.length > 0 && (
                <div className="flex flex-col gap-4">
                  {uploads.map((u) => {
                    const isActive = u.id === activeId;
                    const uTerritories = territoriesOf(u);
                    const uRoutes = routesOf(u, u.territory);
                    return (
                      <div
                        key={u.id}
                        onClick={() => setActiveId(u.id)}
                        className="rounded-lg p-4 cursor-pointer transition-colors"
                        style={{ background: T.surface, border: `2px solid ${isActive ? template.accent : T.border}` }}
                      >
                        <div className="flex items-center justify-between mb-2">
                          <div className="flex items-center gap-1.5 min-w-0">
                            <Icon name="description" size={16} style={{ color: isActive ? template.accent : T.textMuted, flexShrink: 0 }} />
                            <span className="text-sm font-semibold truncate" title={u.fileName}>{u.fileName}</span>
                          </div>
                          <button
                            onClick={(e) => { e.stopPropagation(); removeUpload(u.id); }}
                            className="shrink-0 w-6 h-6 rounded-full flex items-center justify-center hover:bg-black/5"
                            style={{ color: T.textMuted }}
                            title="Remove"
                          >
                            <Icon name="close" size={15} />
                          </button>
                        </div>

                        {u.parseError && (
                          <p className="text-xs flex items-start gap-1.5 mb-2" style={{ color: "#ba1a1a" }}><Icon name="error" size={14} />{u.parseError}</p>
                        )}

                        {u.rows && (
                          <div className="space-y-2.5" onClick={(e) => e.stopPropagation()}>
                            <div>
                              <span className="block text-[11px] mb-1" style={{ color: T.textMuted }}>Territory</span>
                              <Select value={u.territory} onChange={(e) => updateUpload(u.id, { territory: e.target.value, routeName: routesOf(u, e.target.value)[0] || "" })}>
                                {uTerritories.map((t) => <option key={t} value={t}>{t}</option>)}
                              </Select>
                            </div>
                            <div>
                              <span className="block text-[11px] mb-1" style={{ color: T.textMuted }}>Route (auto-extracted)</span>
                              <Select value={u.routeName} onChange={(e) => updateUpload(u.id, { routeName: e.target.value })}>
                                {uRoutes.map((r) => <option key={r} value={r}>{r || "(blank route name)"}</option>)}
                              </Select>
                            </div>
                            <div>
                              <span className="block text-[11px] mb-1" style={{ color: T.textMuted }}>Last visit date</span>
                              <DatePicker value={u.lastVisitDate} onChange={(v) => updateUpload(u.id, { lastVisitDate: v })} accent={template.accent} />
                            </div>
                            <div>
                              <span className="block text-[11px] mb-1" style={{ color: T.textMuted }}>Which visit counts as "last"?</span>
                              <Select value={u.strategy} onChange={(e) => updateUpload(u.id, { strategy: e.target.value })}>
                                <option value="fallback">Most recent visit with data</option>
                                <option value="strict">Only PREVIOUS 1ST VISIT column</option>
                              </Select>
                            </div>
                          </div>
                        )}
                      </div>
                    );
                  })}

                  <button
                    onClick={resetActiveFilters}
                    disabled={!activeUpload}
                    className="w-full py-2.5 rounded-lg text-sm font-semibold flex items-center justify-center gap-2 transition-colors disabled:opacity-40"
                    style={{ background: T.surfaceContainer, color: T.text, fontFamily: T.bodyFont }}
                  >
                    <Icon name="restart_alt" size={16} /> Reset Filters (active file)
                  </button>

                  {uploads.length > 1 && (
                    <button
                      onClick={exportAllExcel}
                      disabled={exportableCount === 0}
                      className="w-full py-2.5 rounded-lg text-sm font-semibold flex items-center justify-center gap-2 transition-colors disabled:opacity-40"
                      style={{ background: template.accent, color: "#fff", fontFamily: T.bodyFont }}
                    >
                      <Icon name="folder_zip" size={16} /> Export All as Excel ({exportableCount})
                    </button>
                  )}
                </div>
              )}
            </div>

            <div className="lg:col-span-8">
              {!activeUpload ? (
                <div className="h-full flex flex-col items-center justify-center text-sm py-24 rounded-lg" style={{ border: `2px dashed ${T.border}`, color: T.textMuted }}>
                  <Icon name="table_chart" size={28} style={{ color: T.textMuted }} />
                  <p className="mt-2">Upload one or more reports to see the outlet-wise breakdown here</p>
                </div>
              ) : (
                <div className="rounded-lg overflow-hidden" style={{ background: T.surface, border: `1px solid ${T.border}` }}>
                  <div className="p-6 flex items-center justify-between" style={{ background: template.accent }}>
                    <div>
                      <h2 style={{ fontFamily: T.headFont, fontSize: 22, fontWeight: 700, color: "#fff" }}>{template.reportTitle}</h2>
                      <p className="text-sm opacity-85 mt-1" style={{ color: "#fff" }}>{activeUpload.territory}{activeUpload.routeName ? ` — ${activeUpload.routeName}` : ""}</p>
                    </div>
                    <div className="flex gap-2">
                      <button onClick={exportExcel} disabled={!canExport} className="flex items-center gap-1.5 disabled:opacity-40 text-xs font-bold px-3.5 py-2 rounded-lg" style={{ background: "#fff", color: template.accent }}>
                        <Icon name="grid_on" size={15} /> Download Excel
                      </button>
                      <button onClick={exportPdf} disabled={!canExport} className="flex items-center gap-1.5 disabled:opacity-40 text-xs font-bold px-3.5 py-2 rounded-lg" style={{ background: "rgba(0,0,0,0.2)", color: "#fff" }}>
                        <Icon name="picture_as_pdf" size={15} /> Download PDF
                      </button>
                    </div>
                  </div>

                  {report && (
                    <div className="p-5 grid grid-cols-3 gap-4" style={{ borderBottom: `1px solid ${T.border}` }}>
                      <div className="rounded-lg p-4" style={{ border: `1px solid ${T.border}` }}>
                        <p style={{ fontFamily: T.capsFont, fontSize: 10, letterSpacing: "0.08em", color: T.textMuted }}>{template.totalLabel.toUpperCase()}</p>
                        <p style={{ fontFamily: T.headFont, fontSize: 22, fontWeight: 700, color: template.accent }} className="mt-1">{report.totalSale.toLocaleString(undefined, { minimumFractionDigits: 2 })}</p>
                      </div>
                      <div className="rounded-lg p-4" style={{ border: `1px solid ${T.border}` }}>
                        <p style={{ fontFamily: T.capsFont, fontSize: 10, letterSpacing: "0.08em", color: T.textMuted }}>OUTLETS</p>
                        <p style={{ fontFamily: T.headFont, fontSize: 22, fontWeight: 700 }} className="mt-1">{report.outlets.length}</p>
                      </div>
                      <div className="rounded-lg p-4" style={{ border: `1px solid ${T.border}` }}>
                        <p style={{ fontFamily: T.capsFont, fontSize: 10, letterSpacing: "0.08em", color: T.textMuted }}>ROWS MATCHED</p>
                        <p style={{ fontFamily: T.headFont, fontSize: 22, fontWeight: 700 }} className="mt-1">{report.matchedRows} / {report.consideredRows}</p>
                      </div>
                    </div>
                  )}

                  {report && report.outlets.length === 0 && (
                    <p className="px-5 py-10 text-sm text-center" style={{ color: T.textMuted }}>No rows had quantity in the selected "last visit" column for this route. Try the other strategy, or double-check the route.</p>
                  )}

                  {report && report.outlets.length > 0 && (
                    <>
                      <table className="w-full text-sm">
                        <thead style={{ background: T.surfaceContainer, borderTop: `2px solid ${template.accent}` }}>
                          <tr>
                            <th className="text-left px-5 py-2.5 w-12" style={{ fontFamily: T.capsFont, fontSize: 10, letterSpacing: "0.08em", color: T.textMuted }}>{template.colNo}</th>
                            <th className="text-left px-3 py-2.5" style={{ fontFamily: T.capsFont, fontSize: 10, letterSpacing: "0.08em", color: T.textMuted }}>{template.colOutlet}</th>
                            <th className="text-left px-3 py-2.5" style={{ fontFamily: T.capsFont, fontSize: 10, letterSpacing: "0.08em", color: T.textMuted }}>{template.colProduct}</th>
                            <th className="text-right px-5 py-2.5" style={{ fontFamily: T.capsFont, fontSize: 10, letterSpacing: "0.08em", color: T.textMuted }}>{template.colQty}</th>
                          </tr>
                        </thead>
                        <tbody>
                          {pagedOutlets.map((outlet, idx) => {
                            const absoluteIdx = page * PAGE_SIZE + idx;
                            const banded = template.bandingEnabled && absoluteIdx % 2 === 1;
                            return outlet.items.map((item, i) => (
                              <tr key={`${idx}-${i}`} style={{ borderBottom: `1px solid ${T.border}`, background: banded ? template.bandingColor : "transparent" }}>
                                <td className="px-5 py-2 align-top" style={{ color: T.textMuted }}>{i === 0 ? String(absoluteIdx + 1).padStart(2, "0") : ""}</td>
                                <td className="px-3 py-2 align-top font-semibold">{i === 0 ? outlet.name : ""}</td>
                                <td className="px-3 py-2">{item.product}</td>
                                <td className="px-5 py-2 text-right tabular-nums">{item.qty}</td>
                              </tr>
                            ));
                          })}
                        </tbody>
                      </table>
                      <div className="flex items-center justify-between px-5 py-4 text-sm" style={{ color: T.textMuted, borderTop: `1px solid ${T.border}` }}>
                        <span>Showing {page * PAGE_SIZE + 1}-{Math.min((page + 1) * PAGE_SIZE, report.outlets.length)} of {report.outlets.length} outlets</span>
                        <div className="flex gap-2">
                          <button onClick={() => setPage((p) => Math.max(0, p - 1))} disabled={page === 0} className="w-8 h-8 rounded-md flex items-center justify-center disabled:opacity-30" style={{ border: `1px solid ${T.border}` }}>
                            <Icon name="chevron_left" size={18} />
                          </button>
                          <button onClick={() => setPage((p) => Math.min(pageCount - 1, p + 1))} disabled={page >= pageCount - 1} className="w-8 h-8 rounded-md flex items-center justify-center disabled:opacity-30" style={{ border: `1px solid ${T.border}` }}>
                            <Icon name="chevron_right" size={18} />
                          </button>
                        </div>
                      </div>
                    </>
                  )}
                </div>
              )}
            </div>
          </div>
        )}
      </main>
    </div>
  );
}
