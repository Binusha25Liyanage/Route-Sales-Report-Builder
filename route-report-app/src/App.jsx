import { useState, useEffect, useRef, useMemo } from "react";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import { Upload, FileSpreadsheet, Settings2, LayoutDashboard, Download, FileDown, ChevronRight, CircleAlert, CircleCheck, RotateCcw } from "lucide-react";

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
  titleFont: "Georgia, 'Times New Roman', serif",
  bodyFont: "'Segoe UI', Calibri, Arial, sans-serif",
  accent: "#8a5a2c",
};

const ACCENTS = [
  { name: "Ledger brown", value: "#8a5a2c" },
  { name: "Ink navy", value: "#233246" },
  { name: "Forest", value: "#2f5233" },
  { name: "Maroon", value: "#6b2b3a" },
  { name: "Slate", value: "#475467" },
];

const VISIT_BLOCKS = [
  { key: "1st", label: "PREVIOUS 1ST VISIT", start: 31 },
  { key: "2nd", label: "PREVIOUS 2ND VISIT", start: 24 },
  { key: "3rd", label: "PREVIOUS 3RD VISIT", start: 17 },
  { key: "4th", label: "PREVIOUS 4TH VISIT", start: 10 },
];
// column offsets within a 7-col visit block
const OFF = { qty: 0, invoice: 1, retQty: 2, retVal: 3, freeQty: 4, freeVal: 5, net: 6 };

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

  const dataStart = headerRow + 2; // skip header + sub-header (QUANTITY, INVOICE VALUE...) row
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
      visits[b.key] = {
        qty: num(r[b.start + OFF.qty]),
        net: num(r[b.start + OFF.net]),
      };
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

function TemplateEditor({ template, setTemplate, onReset }) {
  const field = (key, label, placeholder) => (
    <label className="block mb-4">
      <span className="block text-xs font-semibold tracking-wide uppercase text-stone-500 mb-1.5">{label}</span>
      <input
        className="w-full rounded-md border border-stone-300 px-3 py-2 text-sm text-stone-800 focus:outline-none focus:ring-2 focus:ring-offset-1"
        style={{ "--tw-ring-color": template.accent }}
        value={template[key]}
        placeholder={placeholder}
        onChange={(e) => setTemplate({ ...template, [key]: e.target.value })}
      />
    </label>
  );

  return (
    <div className="grid md:grid-cols-2 gap-10">
      <div>
        <h3 className="text-sm font-semibold tracking-wide uppercase text-stone-500 mb-4">Header block text</h3>
        {field("reportTitle", "Report title")}
        {field("territoryLabel", "Territory row label")}
        {field("routeLabel", "Route row label")}
        {field("totalLabel", "Total sale value label")}
        {field("dateLabel", "Last visit date label")}

        <h3 className="text-sm font-semibold tracking-wide uppercase text-stone-500 mb-4 mt-8">Table column headers</h3>
        <div className="grid grid-cols-2 gap-x-4">
          {field("colNo", "Column 1")}
          {field("colOutlet", "Column 2")}
          {field("colProduct", "Column 3")}
          {field("colQty", "Column 4")}
        </div>

        <h3 className="text-sm font-semibold tracking-wide uppercase text-stone-500 mb-4 mt-4">Accent color</h3>
        <div className="flex gap-2 mb-6">
          {ACCENTS.map((a) => (
            <button
              key={a.value}
              onClick={() => setTemplate({ ...template, accent: a.value })}
              title={a.name}
              className="w-8 h-8 rounded-full border-2 transition-transform hover:scale-110"
              style={{ backgroundColor: a.value, borderColor: template.accent === a.value ? "#1c1917" : "transparent" }}
            />
          ))}
        </div>

        <button onClick={onReset} className="flex items-center gap-1.5 text-sm text-stone-500 hover:text-stone-800">
          <RotateCcw size={14} /> Reset to defaults
        </button>
      </div>

      <div>
        <h3 className="text-sm font-semibold tracking-wide uppercase text-stone-500 mb-4">Live preview</h3>
        <div className="border border-stone-300 rounded-lg overflow-hidden shadow-sm bg-white">
          <div className="px-5 py-4" style={{ backgroundColor: template.accent }}>
            <p className="text-white text-center font-semibold" style={{ fontFamily: template.titleFont }}>{template.reportTitle}</p>
          </div>
          <div className="px-5 py-4 text-sm" style={{ fontFamily: template.bodyFont }}>
            <div className="grid grid-cols-[140px_1fr] gap-y-1.5">
              <span className="font-bold text-stone-700">{template.territoryLabel}</span><span className="text-stone-600">Nikaweratiya</span>
              <span className="font-bold text-stone-700">{template.routeLabel}</span><span className="text-stone-600">Anamaduwa to Galkulama</span>
              <span className="font-bold text-stone-700">{template.totalLabel}</span><span className="text-stone-600">110,409.00</span>
              <span className="font-bold text-stone-700">{template.dateLabel}</span><span className="text-stone-600 italic">(fill in after download)</span>
            </div>
            <table className="w-full mt-4 text-xs border-t border-stone-200">
              <thead>
                <tr style={{ color: template.accent }} className="border-b-2" >
                  <th className="text-left py-1.5 font-bold" style={{ borderColor: template.accent }}>{template.colNo}</th>
                  <th className="text-left py-1.5 font-bold">{template.colOutlet}</th>
                  <th className="text-left py-1.5 font-bold">{template.colProduct}</th>
                  <th className="text-left py-1.5 font-bold">{template.colQty}</th>
                </tr>
              </thead>
              <tbody className="text-stone-600">
                <tr className="border-b border-stone-100"><td className="py-1">1</td><td className="py-1">JANAKA KARAWALA KADE</td><td className="py-1">Black Chicken 90g</td><td className="py-1">15</td></tr>
                <tr className="border-b border-stone-100"><td className="py-1"></td><td className="py-1"></td><td className="py-1">Sago Seed 100g</td><td className="py-1">20</td></tr>
              </tbody>
            </table>
          </div>
        </div>
        <p className="text-xs text-stone-400 mt-3">This layout is saved automatically and used every time you generate a report on the Dashboard tab.</p>
      </div>
    </div>
  );
}

export default function App() {
  const [screen, setScreen] = useState("dashboard");
  const [template, setTemplate] = useState(DEFAULT_TEMPLATE);
  const [templateLoaded, setTemplateLoaded] = useState(false);

  const [fileName, setFileName] = useState(null);
  const [rows, setRows] = useState(null);
  const [parseError, setParseError] = useState(null);

  const [territory, setTerritory] = useState("");
  const [routeName, setRouteName] = useState("");
  const [lastVisitDate, setLastVisitDate] = useState("");
  const [strategy, setStrategy] = useState("fallback");

  const fileInputRef = useRef(null);

  useEffect(() => {
    try {
      const saved = localStorage.getItem("template-config");
      if (saved) setTemplate({ ...DEFAULT_TEMPLATE, ...JSON.parse(saved) });
    } catch (e) {
      // no saved config yet
    } finally {
      setTemplateLoaded(true);
    }
  }, []);

  useEffect(() => {
    if (!templateLoaded) return;
    try { localStorage.setItem("template-config", JSON.stringify(template)); } catch (e) {}
  }, [template, templateLoaded]);

  const territories = useMemo(() => {
    if (!rows) return [];
    return [...new Set(rows.map((r) => r.territory))].filter(Boolean).sort();
  }, [rows]);

  const routesForTerritory = useMemo(() => {
    if (!rows || !territory) return [];
    return [...new Set(rows.filter((r) => r.territory === territory).map((r) => r.routeName))].filter((x) => x !== undefined).sort();
  }, [rows, territory]);

  const report = useMemo(() => {
    if (!rows || !territory) return null;
    return buildReport(rows, territory, routeName, strategy);
  }, [rows, territory, routeName, strategy]);

  function handleFile(file) {
    setParseError(null);
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const arr = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true });
        const parsed = parseMasterReport(arr);
        if (parsed.length === 0) throw new Error("No outlet/product rows found in this file.");
        setRows(parsed);
        const firstTerritory = [...new Set(parsed.map((r) => r.territory))].sort()[0];
        setTerritory(firstTerritory || "");
        const firstRoute = [...new Set(parsed.filter((r) => r.territory === firstTerritory).map((r) => r.routeName))].sort()[0];
        setRouteName(firstRoute || "");
      } catch (err) {
        setParseError(err.message || "Couldn't read this file.");
        setRows(null);
      }
    };
    reader.readAsArrayBuffer(file);
  }

  async function exportExcel() {
    if (!report) return;
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
      [template.territoryLabel, territory],
      [template.routeLabel, routeName],
      [template.totalLabel, Math.round(report.totalSale * 100) / 100],
      [template.dateLabel, lastVisitDate || ""],
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

    let r = headerRowNum;
    report.outlets.forEach((outlet, idx) => {
      const startRow = r + 1;
      outlet.items.forEach((item, i) => {
        r += 1;
        const row = ws.getRow(r);
        row.getCell(1).value = i === 0 ? idx + 1 : null;
        row.getCell(2).value = i === 0 ? outlet.name : null;
        row.getCell(3).value = item.product;
        row.getCell(4).value = item.qty;
        [1, 2, 3, 4].forEach((c) => { row.getCell(c).border = borderAll; });
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
    const safeRoute = (routeName || territory).replace(/[^a-z0-9]+/gi, "_").slice(0, 40);
    a.href = url;
    a.download = `${territory}_${safeRoute}_LastVisit.xlsx`;
    a.click();
    URL.revokeObjectURL(url);
  }

  function exportPdf() {
    if (!report) return;
    const win = window.open("", "_blank");
    const rowsHtml = report.outlets.map((outlet, idx) => outlet.items.map((item, i) => `
      <tr>
        <td>${i === 0 ? idx + 1 : ""}</td>
        <td>${i === 0 ? outlet.name : ""}</td>
        <td>${item.product}</td>
        <td style="text-align:right">${item.qty}</td>
      </tr>`).join("")).join("");

    win.document.write(`
      <html><head><title>${territory} - ${routeName}</title>
      <style>
        body{font-family:${template.bodyFont};color:#292524;padding:32px;}
        h1{font-family:${template.titleFont};text-align:center;color:white;background:${template.accent};padding:14px;border-radius:6px;font-size:18px;margin-bottom:20px;}
        table{width:100%;border-collapse:collapse;margin-top:16px;font-size:12px;}
        th{text-align:left;border-bottom:2px solid ${template.accent};color:${template.accent};padding:6px 8px;}
        td{padding:5px 8px;border-bottom:1px solid #eee;}
        .meta{display:grid;grid-template-columns:160px 1fr;row-gap:6px;font-size:13px;}
        .meta b{color:#44403c;}
        @media print{ body{padding:0;} }
      </style></head><body>
      <h1>${template.reportTitle}</h1>
      <div class="meta">
        <b>${template.territoryLabel}</b><span>${territory}</span>
        <b>${template.routeLabel}</b><span>${routeName}</span>
        <b>${template.totalLabel}</b><span>${report.totalSale.toLocaleString(undefined, { minimumFractionDigits: 2 })}</span>
        <b>${template.dateLabel}</b><span>${lastVisitDate || "________________"}</span>
      </div>
      <table><thead><tr><th>${template.colNo}</th><th>${template.colOutlet}</th><th>${template.colProduct}</th><th style="text-align:right">${template.colQty}</th></tr></thead>
      <tbody>${rowsHtml}</tbody></table>
      </body></html>`);
    win.document.close();
    win.focus();
    setTimeout(() => win.print(), 350);
  }

  return (
    <div className="min-h-screen bg-stone-100" style={{ fontFamily: template.bodyFont }}>
      <div className="border-b border-stone-300 bg-white">
        <div className="max-w-6xl mx-auto px-6 py-4 flex items-center justify-between">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded flex items-center justify-center text-white" style={{ backgroundColor: template.accent }}>
              <FileSpreadsheet size={17} />
            </div>
            <span className="font-semibold text-stone-800" style={{ fontFamily: template.titleFont }}>Route Sales Report Builder</span>
          </div>
          <nav className="flex gap-1 bg-stone-100 rounded-lg p-1">
            <button
              onClick={() => setScreen("dashboard")}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-md text-sm font-medium transition-colors ${screen === "dashboard" ? "bg-white shadow-sm text-stone-900" : "text-stone-500 hover:text-stone-700"}`}
            >
              <LayoutDashboard size={15} /> Dashboard
            </button>
            <button
              onClick={() => setScreen("template")}
              className={`flex items-center gap-1.5 px-3 py-1.5 rounded-md text-sm font-medium transition-colors ${screen === "template" ? "bg-white shadow-sm text-stone-900" : "text-stone-500 hover:text-stone-700"}`}
            >
              <Settings2 size={15} /> Template Editor
            </button>
          </nav>
        </div>
      </div>

      <div className="max-w-6xl mx-auto px-6 py-8">
        {screen === "template" ? (
          <TemplateEditor template={template} setTemplate={setTemplate} onReset={() => setTemplate(DEFAULT_TEMPLATE)} />
        ) : (
          <div className="grid lg:grid-cols-[340px_1fr] gap-8">
            <div>
              <div className="bg-white border border-stone-300 rounded-lg p-5">
                <h2 className="text-sm font-semibold tracking-wide uppercase text-stone-500 mb-3">1. Upload master report</h2>
                <div
                  onClick={() => fileInputRef.current?.click()}
                  onDrop={(e) => { e.preventDefault(); if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]); }}
                  onDragOver={(e) => e.preventDefault()}
                  className="border-2 border-dashed border-stone-300 rounded-lg p-6 text-center cursor-pointer hover:border-stone-400 transition-colors"
                >
                  <Upload size={22} className="mx-auto mb-2 text-stone-400" />
                  <p className="text-sm text-stone-600">{fileName || "Drop the Route Wise Item Wise Outlet Wise Sales Report .xlsx here, or click to browse"}</p>
                  <input ref={fileInputRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={(e) => e.target.files[0] && handleFile(e.target.files[0])} />
                </div>
                {parseError && (
                  <p className="mt-3 text-sm text-red-600 flex items-start gap-1.5"><CircleAlert size={16} className="shrink-0 mt-0.5" />{parseError}</p>
                )}
                {rows && !parseError && (
                  <p className="mt-3 text-sm text-emerald-700 flex items-center gap-1.5"><CircleCheck size={16} />{rows.length} outlet/product rows found across {territories.length} territor{territories.length === 1 ? "y" : "ies"}</p>
                )}
              </div>

              {rows && (
                <div className="bg-white border border-stone-300 rounded-lg p-5 mt-5">
                  <h2 className="text-sm font-semibold tracking-wide uppercase text-stone-500 mb-3">2. Choose territory &amp; route</h2>
                  <label className="block mb-3">
                    <span className="block text-xs text-stone-500 mb-1">Territory</span>
                    <select className="w-full border border-stone-300 rounded-md px-2.5 py-2 text-sm" value={territory} onChange={(e) => { setTerritory(e.target.value); setRouteName(""); }}>
                      {territories.map((t) => <option key={t} value={t}>{t}</option>)}
                    </select>
                  </label>
                  <label className="block mb-3">
                    <span className="block text-xs text-stone-500 mb-1">Route (auto-extracted from the file)</span>
                    <select className="w-full border border-stone-300 rounded-md px-2.5 py-2 text-sm" value={routeName} onChange={(e) => setRouteName(e.target.value)}>
                      {routesForTerritory.map((r) => <option key={r} value={r}>{r || "(blank route name)"}</option>)}
                    </select>
                  </label>
                  <label className="block mb-3">
                    <span className="block text-xs text-stone-500 mb-1">Last visit date</span>
                    <input type="date" className="w-full border border-stone-300 rounded-md px-2.5 py-2 text-sm" value={lastVisitDate} onChange={(e) => setLastVisitDate(e.target.value)} />
                  </label>
                  <label className="block">
                    <span className="block text-xs text-stone-500 mb-1">Which visit counts as "last"?</span>
                    <select className="w-full border border-stone-300 rounded-md px-2.5 py-2 text-sm" value={strategy} onChange={(e) => setStrategy(e.target.value)}>
                      <option value="fallback">Most recent visit with data (1st → 4th)</option>
                      <option value="strict">Only PREVIOUS 1ST VISIT column</option>
                    </select>
                  </label>
                </div>
              )}
            </div>

            <div>
              {!rows ? (
                <div className="h-full flex items-center justify-center text-stone-400 text-sm border-2 border-dashed border-stone-200 rounded-lg py-24">
                  Upload a report to see the outlet-wise breakdown here
                </div>
              ) : (
                <div className="bg-white border border-stone-300 rounded-lg overflow-hidden">
                  <div className="px-5 py-4 flex items-center justify-between" style={{ backgroundColor: template.accent }}>
                    <div>
                      <p className="text-white font-semibold" style={{ fontFamily: template.titleFont }}>{template.reportTitle}</p>
                      <p className="text-white/80 text-xs mt-0.5">{territory}{routeName ? ` — ${routeName}` : ""}</p>
                    </div>
                    <div className="flex gap-2">
                      <button onClick={exportExcel} disabled={!report || report.outlets.length === 0} className="flex items-center gap-1.5 bg-white/95 hover:bg-white disabled:opacity-40 text-stone-800 text-xs font-medium px-3 py-1.5 rounded-md">
                        <Download size={13} /> Excel
                      </button>
                      <button onClick={exportPdf} disabled={!report || report.outlets.length === 0} className="flex items-center gap-1.5 bg-white/95 hover:bg-white disabled:opacity-40 text-stone-800 text-xs font-medium px-3 py-1.5 rounded-md">
                        <FileDown size={13} /> PDF
                      </button>
                    </div>
                  </div>

                  {report && (
                    <div className="px-5 py-3 border-b border-stone-100 grid grid-cols-3 gap-4 text-sm bg-stone-50">
                      <div><span className="text-stone-400 text-xs block">{template.totalLabel}</span><span className="font-semibold text-stone-800">{report.totalSale.toLocaleString(undefined, { minimumFractionDigits: 2 })}</span></div>
                      <div><span className="text-stone-400 text-xs block">Outlets with last-visit sales</span><span className="font-semibold text-stone-800">{report.outlets.length}</span></div>
                      <div><span className="text-stone-400 text-xs block">Rows considered / matched</span><span className="font-semibold text-stone-800">{report.consideredRows} / {report.matchedRows}</span></div>
                    </div>
                  )}

                  {report && report.outlets.length === 0 && (
                    <p className="px-5 py-8 text-sm text-stone-500 text-center">No rows had quantity in the selected "last visit" column for this route. Try the fallback strategy, or double-check the route.</p>
                  )}

                  {report && report.outlets.length > 0 && (
                    <div className="max-h-[560px] overflow-y-auto">
                      <table className="w-full text-sm">
                        <thead className="sticky top-0 bg-white border-b border-stone-200">
                          <tr style={{ color: template.accent }}>
                            <th className="text-left px-5 py-2 font-semibold text-xs w-10">{template.colNo}</th>
                            <th className="text-left px-2 py-2 font-semibold text-xs">{template.colOutlet}</th>
                            <th className="text-left px-2 py-2 font-semibold text-xs">{template.colProduct}</th>
                            <th className="text-right px-5 py-2 font-semibold text-xs">{template.colQty}</th>
                          </tr>
                        </thead>
                        <tbody>
                          {report.outlets.map((outlet, idx) => outlet.items.map((item, i) => (
                            <tr key={`${idx}-${i}`} className="border-b border-stone-50 hover:bg-stone-50">
                              <td className="px-5 py-1.5 text-stone-400">{i === 0 ? idx + 1 : ""}</td>
                              <td className="px-2 py-1.5 font-medium text-stone-700">{i === 0 ? outlet.name : ""}</td>
                              <td className="px-2 py-1.5 text-stone-600">{item.product}</td>
                              <td className="px-5 py-1.5 text-right text-stone-700 tabular-nums">{item.qty}</td>
                            </tr>
                          )))}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>
              )}
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
