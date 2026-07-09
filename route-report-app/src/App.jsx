import { useEffect, useMemo, useRef, useState } from "react";
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
};

const ACCENT_SWATCHES = [
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

const PAGE_SIZE = 4;

const THEME = {
  bg: "#FCF9F4",
  surface: "#FFFFFF",
  surfaceLow: "#F6F3EE",
  surfaceAlt: "#F0EDE9",
  border: "#DAC1C0",
  text: "#1C1C19",
  textMuted: "#544242",
  headingFont: '"Manrope", sans-serif',
  bodyFont: '"Work Sans", sans-serif',
  labelFont: '"IBM Plex Sans", sans-serif',
};

function isBlank(value) {
  return value === null || value === undefined || String(value).trim() === "";
}

function text(value) {
  return isBlank(value) ? "" : String(value).trim();
}

function number(value) {
  if (isBlank(value)) return 0;
  const parsed = Number(String(value).replace(/,/g, ""));
  return Number.isFinite(parsed) ? parsed : 0;
}

function formatMoney(value) {
  return Number(value || 0).toLocaleString(undefined, {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function sanitizeFilePart(value) {
  const cleaned = String(value || "")
    .replace(/[<>:"/\\|?*\x00-\x1F]+/g, "_")
    .replace(/\s+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_+|_+$/g, "");
  return cleaned || "Report";
}

function routeKey(routeNumber, routeName) {
  return `${text(routeNumber)}|||${text(routeName)}`;
}

function routeLabel(routeNumber, routeName) {
  const numberPart = text(routeNumber);
  const namePart = text(routeName);
  if (numberPart && namePart) return `${numberPart} - ${namePart}`;
  return namePart || numberPart || "(blank route)";
}

function Icon({ name, size = 20, className = "", style = {} }) {
  return (
    <span className={`material-symbols-outlined ${className}`} style={{ fontSize: size, lineHeight: 1, ...style }}>
      {name}
    </span>
  );
}

function FieldLabel({ children }) {
  return (
    <span
      className="mb-1.5 block uppercase"
      style={{
        fontFamily: THEME.labelFont,
        fontSize: 11,
        fontWeight: 600,
        letterSpacing: "0.12em",
        color: THEME.textMuted,
      }}
    >
      {children}
    </span>
  );
}

function TextInput({ value, onChange, accent, placeholder }) {
  return (
    <input
      value={value}
      placeholder={placeholder}
      onChange={(event) => onChange(event.target.value)}
      className="w-full rounded-lg px-3 py-2.5 text-sm outline-none transition"
      style={{
        background: THEME.surface,
        border: `1px solid ${THEME.border}`,
        color: THEME.text,
        fontFamily: THEME.bodyFont,
      }}
      onFocus={(event) => {
        event.currentTarget.style.boxShadow = `0 0 0 3px ${accent}22`;
        event.currentTarget.style.borderColor = accent;
      }}
      onBlur={(event) => {
        event.currentTarget.style.boxShadow = "none";
        event.currentTarget.style.borderColor = THEME.border;
      }}
    />
  );
}

function SelectInput({ value, onChange, children }) {
  return (
    <select
      value={value}
      onChange={onChange}
      className="w-full rounded-lg px-3 py-2.5 text-sm outline-none"
      style={{
        background: THEME.surface,
        border: `1px solid ${THEME.border}`,
        color: THEME.text,
        fontFamily: THEME.bodyFont,
      }}
    >
      {children}
    </select>
  );
}

function parseWorkbook(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array", raw: true, cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true });

  const headerRowIndex = rows.findIndex((row) => {
    const cells = row || [];
    return text(cells[0]).toUpperCase() === "NO" && text(cells[4]).toUpperCase() === "TERRITORY";
  });

  if (headerRowIndex === -1) {
    throw new Error("Could not find the expected NO / TERRITORY header row.");
  }

  const dataStart = headerRowIndex + 2;
  const carry = {
    no: "",
    rsm: "",
    asm: "",
    ase: "",
    territory: "",
    routeNumber: "",
    routeName: "",
    outletId: "",
    outletName: "",
  };

  const records = [];

  for (let rowIndex = dataStart; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex] || [];

    if (row.every((cell) => isBlank(cell))) {
      continue;
    }

    if (!isBlank(row[0])) carry.no = text(row[0]);
    if (!isBlank(row[1])) carry.rsm = text(row[1]);
    if (!isBlank(row[2])) carry.asm = text(row[2]);
    if (!isBlank(row[3])) carry.ase = text(row[3]);
    if (!isBlank(row[4])) carry.territory = text(row[4]);
    if (!isBlank(row[5])) carry.routeNumber = text(row[5]);
    if (!isBlank(row[6])) carry.routeName = text(row[6]);
    if (!isBlank(row[7])) carry.outletId = text(row[7]);
    if (!isBlank(row[8])) carry.outletName = text(row[8]);

    const product = text(row[9]);
    if (!product) continue;

    const visits = {};
    for (const block of VISIT_BLOCKS) {
      visits[block.key] = {
        qty: number(row[block.start]),
        net: number(row[block.start + 6]),
      };
    }

    records.push({
      ...carry,
      product,
      routeKey: routeKey(carry.routeNumber, carry.routeName),
      routeDisplay: routeLabel(carry.routeNumber, carry.routeName),
      visits,
      sourceRow: rowIndex + 1,
    });
  }

  if (!records.length) {
    throw new Error("No product rows were found after the header.");
  }

  return records;
}

function resolveLastVisit(row, strategy) {
  const visitOrder = strategy === "strict" ? ["1st"] : ["1st", "2nd", "3rd", "4th"];

  for (const visitKey of visitOrder) {
    const block = row.visits[visitKey];
    if (block && block.qty > 0) {
      return { visitKey, qty: block.qty, net: block.net };
    }
  }

  return null;
}

function buildReport(records, territory, selectedRouteKey, strategy) {
  const routeRows = records.filter((row) => row.territory === territory && row.routeKey === selectedRouteKey);
  const outlets = [];
  const outletMap = new Map();
  let matchedRows = 0;
  let totalSale = 0;

  routeRows.forEach((row) => {
    const resolved = resolveLastVisit(row, strategy);
    if (!resolved) return;

    matchedRows += 1;
    totalSale += resolved.net;

    const outletKey = row.outletId || row.outletName || `${row.sourceRow}`;
    if (!outletMap.has(outletKey)) {
      const group = {
        key: outletKey,
        no: row.no || String(outlets.length + 1),
        outletId: row.outletId,
        outletName: row.outletName,
        items: [],
      };
      outletMap.set(outletKey, group);
      outlets.push(group);
    }

    outletMap.get(outletKey).items.push({
      product: row.product,
      qty: resolved.qty,
      net: resolved.net,
      visitKey: resolved.visitKey,
    });
  });

  return {
    routeRows,
    outlets,
    totalSale,
    matchedRows,
    consideredRows: routeRows.length,
  };
}

function getSampleRows(report) {
  if (report && report.outlets.length > 0) {
    const firstOutlet = report.outlets[0];
    const firstItems = firstOutlet.items.slice(0, 2);

    if (firstItems.length === 2) {
      return [
        { no: firstOutlet.no, outlet: firstOutlet.outletName, product: firstItems[0].product, qty: firstItems[0].qty, rowSpan: 2 },
        { no: "", outlet: "", product: firstItems[1].product, qty: firstItems[1].qty, rowSpan: 0 },
      ];
    }

    return [
      { no: firstOutlet.no, outlet: firstOutlet.outletName, product: firstItems[0]?.product || "Sample product", qty: firstItems[0]?.qty ?? 0, rowSpan: 1 },
    ];
  }

  return [
    { no: "1", outlet: "Janaka Karawala Kade", product: "Black Chicken 90g", qty: 15, rowSpan: 2 },
    { no: "", outlet: "", product: "Sago Seed 100g", qty: 20, rowSpan: 0 },
  ];
}

function TopBarButton({ icon, title, onClick, disabled, accent }) {
  return (
    <button
      type="button"
      title={title}
      onClick={onClick}
      disabled={disabled}
      className="inline-flex h-10 w-10 items-center justify-center rounded-full border transition"
      style={{
        borderColor: THEME.border,
        color: accent,
        background: THEME.surface,
        opacity: disabled ? 0.35 : 1,
        cursor: disabled ? "not-allowed" : "pointer",
      }}
    >
      <Icon name={icon} size={20} />
    </button>
  );
}

function TemplateEditor({ template, setTemplate, report, lastVisitDate, onResetToDefaults, onSaveTemplate }) {
  const previewRows = getSampleRows(report);

  const updateTemplate = (key) => (value) => {
    setTemplate((current) => ({ ...current, [key]: value }));
  };

  return (
    <div className="grid grid-cols-1 gap-6 xl:grid-cols-[5fr_7fr]">
      <section className="flex flex-col gap-6">
        <div className="rounded-2xl p-6 shadow-sm" style={{ background: THEME.surface, border: `1px solid ${THEME.border}` }}>
          <div className="mb-6 border-b pb-4" style={{ borderColor: THEME.border }}>
            <h2 style={{ fontFamily: THEME.headingFont, fontSize: 20, fontWeight: 700, color: template.accent }}>
              Report Configuration
            </h2>
            <p className="mt-1 text-sm" style={{ color: THEME.textMuted }}>
              Edit the labels, accent color, and exported report title.
            </p>
          </div>

          <div className="space-y-5">
            <div>
              <FieldLabel>Report title</FieldLabel>
              <TextInput value={template.reportTitle} onChange={updateTemplate("reportTitle")} accent={template.accent} />
            </div>

            <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
              <div>
                <FieldLabel>Territory label</FieldLabel>
                <TextInput value={template.territoryLabel} onChange={updateTemplate("territoryLabel")} accent={template.accent} />
              </div>
              <div>
                <FieldLabel>Route label</FieldLabel>
                <TextInput value={template.routeLabel} onChange={updateTemplate("routeLabel")} accent={template.accent} />
              </div>
            </div>

            <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
              <div>
                <FieldLabel>Total sale label</FieldLabel>
                <TextInput value={template.totalLabel} onChange={updateTemplate("totalLabel")} accent={template.accent} />
              </div>
              <div>
                <FieldLabel>Last visit label</FieldLabel>
                <TextInput value={template.dateLabel} onChange={updateTemplate("dateLabel")} accent={template.accent} />
              </div>
            </div>

            <div>
              <FieldLabel>Table column headers</FieldLabel>
              <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
                <TextInput value={template.colNo} onChange={updateTemplate("colNo")} accent={template.accent} />
                <TextInput value={template.colOutlet} onChange={updateTemplate("colOutlet")} accent={template.accent} />
                <TextInput value={template.colProduct} onChange={updateTemplate("colProduct")} accent={template.accent} />
                <TextInput value={template.colQty} onChange={updateTemplate("colQty")} accent={template.accent} />
              </div>
            </div>

            <div>
              <FieldLabel>Brand accent color</FieldLabel>
              <div className="flex flex-wrap gap-3">
                {ACCENT_SWATCHES.map((swatch) => {
                  const selected = template.accent === swatch.value;
                  return (
                    <button
                      key={swatch.value}
                      type="button"
                      title={swatch.name}
                      onClick={() => updateTemplate("accent")(swatch.value)}
                      className="h-9 w-9 rounded-full border transition hover:scale-105"
                      style={{
                        background: swatch.value,
                        borderColor: THEME.border,
                        boxShadow: selected ? `0 0 0 2px ${THEME.text}` : "none",
                      }}
                    />
                  );
                })}
              </div>
            </div>

            <div className="flex items-center justify-between gap-4 border-t pt-2" style={{ borderColor: THEME.border }}>
              <button
                type="button"
                onClick={onResetToDefaults}
                className="inline-flex items-center gap-2 text-sm transition hover:opacity-80"
                style={{ color: THEME.textMuted, fontFamily: THEME.bodyFont }}
              >
                <Icon name="restart_alt" size={18} />
                Reset to defaults
              </button>

              <button
                type="button"
                onClick={onSaveTemplate}
                className="rounded-lg px-5 py-2.5 text-sm font-semibold text-white shadow-sm transition active:scale-[0.99]"
                style={{ background: template.accent, fontFamily: THEME.bodyFont }}
              >
                Saved ✓
              </button>
            </div>
          </div>
        </div>
      </section>

      <section>
        <div className="sticky top-6">
          <div className="mb-2 flex items-center justify-between">
            <FieldLabel>Live preview (real-time)</FieldLabel>
          </div>

          <div className="overflow-hidden rounded-2xl shadow-lg" style={{ background: THEME.surface, border: `1px solid ${THEME.border}` }}>
            <div className="px-6 py-5 text-white" style={{ background: template.accent }}>
              <div className="flex flex-col gap-4 md:flex-row md:items-start md:justify-between">
                <div>
                  <h3
                    className="uppercase"
                    style={{
                      fontFamily: THEME.headingFont,
                      fontSize: 22,
                      fontWeight: 700,
                      letterSpacing: "0.02em",
                    }}
                  >
                    {template.reportTitle}
                  </h3>
                  <p className="mt-1 text-sm text-white/80">Template-driven report preview</p>
                </div>
                <div className="text-left md:text-right">
                  <p
                    className="uppercase"
                    style={{
                      fontFamily: THEME.labelFont,
                      fontSize: 10,
                      fontWeight: 600,
                      letterSpacing: "0.14em",
                      color: "rgba(255,255,255,0.75)",
                    }}
                  >
                    {template.dateLabel}
                  </p>
                  <p className="mt-1 text-sm font-semibold">{lastVisitDate || "—"}</p>
                </div>
              </div>
            </div>

            <div className="p-6">
              <div className="grid grid-cols-1 gap-5 border-b pb-5 md:grid-cols-2" style={{ borderColor: THEME.border }}>
                <div className="space-y-4">
                  <div>
                    <p className="uppercase" style={{ fontFamily: THEME.labelFont, fontSize: 10, fontWeight: 600, letterSpacing: "0.14em", color: THEME.textMuted }}>
                      {template.territoryLabel}
                    </p>
                    <p className="mt-1 text-sm font-semibold" style={{ color: THEME.text }}>
                      {report?.routeRows[0]?.territory || "—"}
                    </p>
                  </div>
                  <div>
                    <p className="uppercase" style={{ fontFamily: THEME.labelFont, fontSize: 10, fontWeight: 600, letterSpacing: "0.14em", color: THEME.textMuted }}>
                      {template.routeLabel}
                    </p>
                    <p className="mt-1 text-sm font-semibold" style={{ color: THEME.text }}>
                      {report?.routeRows[0]?.routeDisplay || "—"}
                    </p>
                  </div>
                </div>

                <div className="flex flex-col justify-end md:items-end">
                  <p className="uppercase" style={{ fontFamily: THEME.labelFont, fontSize: 10, fontWeight: 600, letterSpacing: "0.14em", color: THEME.textMuted }}>
                    {template.totalLabel}
                  </p>
                  <p className="mt-1 text-2xl font-bold" style={{ fontFamily: THEME.headingFont, color: template.accent }}>
                    {report ? formatMoney(report.totalSale) : formatMoney(110409)}
                  </p>
                </div>
              </div>

              <div className="mt-5 overflow-hidden rounded-xl border" style={{ borderColor: THEME.border }}>
                <table className="w-full border-collapse text-sm" style={{ fontFamily: THEME.bodyFont }}>
                  <thead>
                    <tr style={{ background: THEME.surfaceAlt, borderTop: `2px solid ${template.accent}` }}>
                      <th className="px-3 py-3 text-left" style={{ fontFamily: THEME.labelFont, fontSize: 10, fontWeight: 600, letterSpacing: "0.12em", color: THEME.textMuted }}>
                        {template.colNo}
                      </th>
                      <th className="px-3 py-3 text-left" style={{ fontFamily: THEME.labelFont, fontSize: 10, fontWeight: 600, letterSpacing: "0.12em", color: THEME.textMuted }}>
                        {template.colOutlet}
                      </th>
                      <th className="px-3 py-3 text-left" style={{ fontFamily: THEME.labelFont, fontSize: 10, fontWeight: 600, letterSpacing: "0.12em", color: THEME.textMuted }}>
                        {template.colProduct}
                      </th>
                      <th className="px-3 py-3 text-right" style={{ fontFamily: THEME.labelFont, fontSize: 10, fontWeight: 600, letterSpacing: "0.12em", color: THEME.textMuted }}>
                        {template.colQty}
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    {previewRows.map((row, index) => (
                      <tr key={`${row.no}-${row.product}-${index}`} style={{ borderTop: `1px solid ${THEME.border}` }}>
                        {row.rowSpan > 0 ? (
                          <>
                            <td className="px-3 py-3 align-top" rowSpan={row.rowSpan}>
                              {row.no}
                            </td>
                            <td className="px-3 py-3 align-top font-semibold" rowSpan={row.rowSpan}>
                              {row.outlet}
                            </td>
                          </>
                        ) : null}
                        <td className="px-3 py-3">{row.product}</td>
                        <td className="px-3 py-3 text-right tabular-nums">{row.qty}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        </div>
      </section>
    </div>
  );
}

function Dashboard({
  template,
  records,
  parseError,
  fileName,
  territory,
  setTerritory,
  routeKeyValue,
  setRouteKeyValue,
  lastVisitDate,
  setLastVisitDate,
  strategy,
  setStrategy,
  onFileSelect,
  onDropFile,
  onResetFilters,
  exportExcel,
  exportPdf,
  report,
  territoryOptions,
  routeOptions,
  page,
  setPage,
  canExport,
}) {
  const fileInputRef = useRef(null);
  const currentPage = Math.min(page, Math.max(0, Math.ceil((report?.outlets.length || 0) / PAGE_SIZE) - 1));
  const pageCount = Math.max(1, Math.ceil((report?.outlets.length || 0) / PAGE_SIZE));
  const pageStart = currentPage * PAGE_SIZE;
  const pageOutlets = report ? report.outlets.slice(pageStart, pageStart + PAGE_SIZE) : [];
  const showingStart = report && report.outlets.length > 0 ? pageStart + 1 : 0;
  const showingEnd = report ? Math.min(pageStart + PAGE_SIZE, report.outlets.length) : 0;

  useEffect(() => {
    if (page !== currentPage) setPage(currentPage);
  }, [currentPage, page, setPage]);

  return (
    <div className="grid grid-cols-1 gap-6 xl:grid-cols-[340px_minmax(0,1fr)]">
      <aside className="flex flex-col gap-6 xl:w-[340px] xl:flex-none">
        <div className="rounded-2xl p-5 shadow-sm" style={{ background: THEME.surface, border: `1px solid ${THEME.border}` }}>
          <FieldLabel>Step 1: Upload master report</FieldLabel>

          <div
            className="cursor-pointer rounded-2xl border-2 border-dashed p-6 text-center transition hover:bg-[#faf7f2]"
            style={{ borderColor: THEME.border }}
            onClick={() => fileInputRef.current?.click()}
            onDragOver={(event) => event.preventDefault()}
            onDrop={(event) => {
              event.preventDefault();
              const file = event.dataTransfer.files?.[0];
              if (file) onDropFile(file);
            }}
          >
            <Icon name="cloud_upload" size={34} style={{ color: template.accent }} />
            <p className="mt-3 text-sm font-semibold" style={{ color: THEME.text }}>
              Drop file here or click
            </p>
            <p className="mt-1 text-xs" style={{ color: THEME.textMuted }}>
              Route Wise Item Wise Outlet Wise Sales Report (.xlsx)
            </p>
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls"
              className="hidden"
              onChange={(event) => {
                const file = event.target.files?.[0];
                if (file) onFileSelect(file);
                event.target.value = "";
              }}
            />
          </div>

          {fileName ? (
            <p className="mt-3 text-sm" style={{ color: THEME.textMuted }}>
              Selected file: <span className="font-semibold" style={{ color: THEME.text }}>{fileName}</span>
            </p>
          ) : null}

          {parseError ? (
            <div className="mt-3 flex items-start gap-2 rounded-xl border px-3 py-2.5 text-sm" style={{ borderColor: "#e1a1a1", color: "#9d2626", background: "#fff6f6" }}>
              <Icon name="error" size={18} />
              <span>{parseError}</span>
            </div>
          ) : null}

          {records ? (
            <div className="mt-3 flex items-start gap-2 rounded-xl border px-3 py-2.5 text-sm" style={{ borderColor: "#9cc89f", color: "#2E7A4E", background: "#f4fbf4" }}>
              <Icon name="check_circle" size={18} />
              <span>
                Loaded {records.length} rows across {territoryOptions.length} territories
              </span>
            </div>
          ) : null}
        </div>

        {records ? (
          <div className="rounded-2xl p-5 shadow-sm" style={{ background: THEME.surface, border: `1px solid ${THEME.border}` }}>
            <FieldLabel>Step 2: Choose territory &amp; route</FieldLabel>

            <div className="mt-4 space-y-4">
              <div>
                <FieldLabel>Territory</FieldLabel>
                <SelectInput value={territory} onChange={(event) => setTerritory(event.target.value)}>
                  {territoryOptions.map((value) => (
                    <option key={value} value={value}>
                      {value}
                    </option>
                  ))}
                </SelectInput>
              </div>

              <div>
                <FieldLabel>Route</FieldLabel>
                <SelectInput value={routeKeyValue} onChange={(event) => setRouteKeyValue(event.target.value)}>
                  {routeOptions.map((option) => (
                    <option key={option.key} value={option.key}>
                      {option.label}
                    </option>
                  ))}
                </SelectInput>
                <p className="mt-1 text-xs" style={{ color: THEME.textMuted }}>
                  auto-extracted from file
                </p>
              </div>

              <div>
                <FieldLabel>Last visit date</FieldLabel>
                <input
                  type="date"
                  value={lastVisitDate}
                  onChange={(event) => setLastVisitDate(event.target.value)}
                  className="w-full rounded-lg px-3 py-2.5 text-sm outline-none"
                  style={{ background: THEME.surface, border: `1px solid ${THEME.border}`, color: THEME.text, fontFamily: THEME.bodyFont }}
                />
              </div>

              <div>
                <FieldLabel>Which visit counts as last?</FieldLabel>
                <SelectInput value={strategy} onChange={(event) => setStrategy(event.target.value)}>
                  <option value="fallback">Most recent visit with data</option>
                  <option value="strict">Only PREVIOUS 1ST VISIT column</option>
                </SelectInput>
              </div>
            </div>

            <button
              type="button"
              onClick={onResetFilters}
              className="mt-5 inline-flex w-full items-center justify-center gap-2 rounded-xl px-4 py-2.5 text-sm font-semibold transition"
              style={{ background: THEME.surfaceAlt, color: THEME.text }}
            >
              <Icon name="restart_alt" size={18} />
              Reset Filters
            </button>
          </div>
        ) : null}
      </aside>

      <main className="min-w-0">
        {!records ? (
          <div className="flex min-h-[560px] items-center justify-center rounded-2xl border-2 border-dashed px-6 text-center shadow-sm" style={{ borderColor: THEME.border, background: THEME.surface, color: THEME.textMuted }}>
            <div>
              <Icon name="table_chart" size={42} style={{ color: template.accent }} />
              <p className="mt-3 text-sm font-semibold" style={{ color: THEME.text }}>
                Upload a report to see the outlet-wise breakdown here
              </p>
            </div>
          </div>
        ) : (
          <div className="overflow-hidden rounded-2xl shadow-sm" style={{ background: THEME.surface, border: `1px solid ${THEME.border}` }}>
            <div className="flex flex-col gap-4 bg-white px-6 py-5 md:flex-row md:items-center md:justify-between" style={{ background: template.accent }}>
              <div className="text-white">
                <h2 style={{ fontFamily: THEME.headingFont, fontSize: 24, fontWeight: 700 }}>{template.reportTitle}</h2>
                <p className="mt-1 text-sm text-white/85">
                  {territory} {selectedRouteLabel ? ` - ${selectedRouteLabel}` : ""}
                </p>
              </div>

              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={exportExcel}
                  disabled={!canExport}
                  className="inline-flex items-center gap-2 rounded-lg px-4 py-2 text-xs font-semibold transition"
                  style={{ background: THEME.surface, color: template.accent, opacity: canExport ? 1 : 0.35 }}
                >
                  <Icon name="download" size={16} />
                  Download Excel
                </button>
                <button
                  type="button"
                  onClick={exportPdf}
                  disabled={!canExport}
                  className="inline-flex items-center gap-2 rounded-lg px-4 py-2 text-xs font-semibold transition"
                  style={{ background: "rgba(0,0,0,0.22)", color: "#fff", opacity: canExport ? 1 : 0.35 }}
                >
                  <Icon name="picture_as_pdf" size={16} />
                  Download PDF
                </button>
              </div>
            </div>

            <div className="grid grid-cols-1 gap-4 border-b p-6 md:grid-cols-3" style={{ borderColor: THEME.border }}>
              <div className="rounded-xl border p-4" style={{ borderColor: THEME.border, background: THEME.surface }}>
                <p className="uppercase" style={{ fontFamily: THEME.labelFont, fontSize: 10, fontWeight: 600, letterSpacing: "0.14em", color: THEME.textMuted }}>
                  {template.totalLabel}
                </p>
                <p className="mt-2 text-2xl font-bold" style={{ fontFamily: THEME.headingFont, color: template.accent }}>
                  {formatMoney(report?.totalSale || 0)}
                </p>
              </div>

              <div className="rounded-xl border p-4" style={{ borderColor: THEME.border, background: THEME.surface }}>
                <p className="uppercase" style={{ fontFamily: THEME.labelFont, fontSize: 10, fontWeight: 600, letterSpacing: "0.14em", color: THEME.textMuted }}>
                  Outlets
                </p>
                <p className="mt-2 text-2xl font-bold" style={{ fontFamily: THEME.headingFont }}>
                  {report?.outlets.length || 0}
                </p>
              </div>

              <div className="rounded-xl border p-4" style={{ borderColor: THEME.border, background: THEME.surface }}>
                <p className="uppercase" style={{ fontFamily: THEME.labelFont, fontSize: 10, fontWeight: 600, letterSpacing: "0.14em", color: THEME.textMuted }}>
                  Rows matched
                </p>
                <p className="mt-2 text-2xl font-bold" style={{ fontFamily: THEME.headingFont }}>
                  {report?.matchedRows || 0} / {report?.consideredRows || 0}
                </p>
              </div>
            </div>

            {report && report.outlets.length === 0 ? (
              <div className="px-6 py-16 text-center">
                <p className="text-sm" style={{ color: THEME.textMuted }}>
                  No results were found for this territory and route with the selected last-visit strategy.
                </p>
                <p className="mt-2 text-sm" style={{ color: THEME.textMuted }}>
                  Try the other strategy or double check the route.
                </p>
              </div>
            ) : null}

            {report && report.outlets.length > 0 ? (
              <>
                <div className="overflow-auto">
                  <table className="w-full border-collapse text-sm" style={{ fontFamily: THEME.bodyFont }}>
                    <thead className="sticky top-0 z-10">
                      <tr style={{ background: THEME.surfaceAlt, borderTop: `2px solid ${template.accent}` }}>
                        <th className="px-6 py-3 text-left" style={{ fontFamily: THEME.labelFont, fontSize: 10, fontWeight: 600, letterSpacing: "0.14em", color: THEME.textMuted }}>
                          {template.colNo}
                        </th>
                        <th className="px-4 py-3 text-left" style={{ fontFamily: THEME.labelFont, fontSize: 10, fontWeight: 600, letterSpacing: "0.14em", color: THEME.textMuted }}>
                          {template.colOutlet}
                        </th>
                        <th className="px-4 py-3 text-left" style={{ fontFamily: THEME.labelFont, fontSize: 10, fontWeight: 600, letterSpacing: "0.14em", color: THEME.textMuted }}>
                          {template.colProduct}
                        </th>
                        <th className="px-6 py-3 text-right" style={{ fontFamily: THEME.labelFont, fontSize: 10, fontWeight: 600, letterSpacing: "0.14em", color: THEME.textMuted }}>
                          {template.colQty}
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {pageOutlets.map((outlet) => {
                        const rowSpan = outlet.items.length;
                        return outlet.items.map((item, itemIndex) => (
                          <tr key={`${outlet.key}-${item.product}-${itemIndex}`} style={{ borderTop: `1px solid ${THEME.border}` }}>
                            {itemIndex === 0 ? (
                              <>
                                <td className="px-6 py-3 align-top" rowSpan={rowSpan}>
                                  {outlet.no}
                                </td>
                                <td className="px-4 py-3 align-top font-semibold" rowSpan={rowSpan}>
                                  {outlet.outletName}
                                </td>
                              </>
                            ) : null}
                            <td className="px-4 py-3">{item.product}</td>
                            <td className="px-6 py-3 text-right tabular-nums">{item.qty}</td>
                          </tr>
                        ));
                      })}
                    </tbody>
                  </table>
                </div>

                <div className="flex items-center justify-between gap-4 border-t px-6 py-4 text-sm" style={{ borderColor: THEME.border, color: THEME.textMuted }}>
                  <span>
                    Showing {showingStart}-{showingEnd} of {report.outlets.length} outlets
                  </span>
                  <div className="flex items-center gap-2">
                    <button
                      type="button"
                      onClick={() => setPage((current) => Math.max(0, current - 1))}
                      disabled={currentPage === 0}
                      className="inline-flex h-9 w-9 items-center justify-center rounded-lg border transition disabled:cursor-not-allowed disabled:opacity-35"
                      style={{ borderColor: THEME.border, background: THEME.surface }}
                    >
                      <Icon name="chevron_left" size={18} />
                    </button>
                    <button
                      type="button"
                      onClick={() => setPage((current) => Math.min(pageCount - 1, current + 1))}
                      disabled={currentPage >= pageCount - 1}
                      className="inline-flex h-9 w-9 items-center justify-center rounded-lg border transition disabled:cursor-not-allowed disabled:opacity-35"
                      style={{ borderColor: THEME.border, background: THEME.surface }}
                    >
                      <Icon name="chevron_right" size={18} />
                    </button>
                  </div>
                </div>
              </>
            ) : null}
          </div>
        )}
      </main>
    </div>
  );
}

function App() {
  const [screen, setScreen] = useState("dashboard");
  const [template, setTemplate] = useState(DEFAULT_TEMPLATE);
  const [templateLoaded, setTemplateLoaded] = useState(false);
  const [records, setRecords] = useState(null);
  const [fileName, setFileName] = useState("");
  const [parseError, setParseError] = useState("");
  const [territory, setTerritory] = useState("");
  const [routeKeyValue, setRouteKeyValue] = useState("");
  const [lastVisitDate, setLastVisitDate] = useState("");
  const [strategy, setStrategy] = useState("fallback");
  const [page, setPage] = useState(0);
  const [saveFlash, setSaveFlash] = useState(false);
  const saveFlashRef = useRef(null);

  useEffect(() => {
    try {
      const saved = localStorage.getItem("template-config");
      if (saved) {
        setTemplate({ ...DEFAULT_TEMPLATE, ...JSON.parse(saved) });
      }
    } catch {
      setTemplate(DEFAULT_TEMPLATE);
    } finally {
      setTemplateLoaded(true);
    }
  }, []);

  useEffect(() => {
    if (!templateLoaded) return;
    try {
      localStorage.setItem("template-config", JSON.stringify(template));
    } catch {
      // Ignore storage errors in private or blocked sessions.
    }
  }, [template, templateLoaded]);

  const territoryOptions = useMemo(() => {
    if (!records) return [];
    return Array.from(new Set(records.map((row) => row.territory).filter(Boolean))).sort((left, right) => left.localeCompare(right));
  }, [records]);

  const routeOptions = useMemo(() => {
    if (!records || !territory) return [];

    const seen = new Map();
    records
      .filter((row) => row.territory === territory)
      .forEach((row) => {
        if (!seen.has(row.routeKey)) {
          seen.set(row.routeKey, row.routeDisplay);
        }
      });

    return Array.from(seen.entries())
      .map(([key, label]) => ({ key, label }))
      .sort((left, right) => left.label.localeCompare(right.label));
  }, [records, territory]);

  useEffect(() => {
    if (!territoryOptions.length) return;
    if (!territory || !territoryOptions.includes(territory)) {
      setTerritory(territoryOptions[0]);
    }
  }, [territory, territoryOptions]);

  useEffect(() => {
    if (!routeOptions.length) {
      if (routeKeyValue) setRouteKeyValue("");
      return;
    }

    if (!routeOptions.some((route) => route.key === routeKeyValue)) {
      setRouteKeyValue(routeOptions[0].key);
    }
  }, [routeOptions, routeKeyValue]);

  useEffect(() => {
    setPage(0);
  }, [territory, routeKeyValue, strategy]);

  const report = useMemo(() => {
    if (!records || !territory || !routeKeyValue) return null;
    return buildReport(records, territory, routeKeyValue, strategy);
  }, [records, territory, routeKeyValue, strategy]);

  const canExport = Boolean(report && report.outlets.length > 0);

  function loadFile(file) {
    setParseError("");
    setFileName(file.name);

    const reader = new FileReader();
    reader.onerror = () => {
      setRecords(null);
      setParseError("The file could not be read.");
    };
    reader.onload = (event) => {
      try {
        const parsed = parseWorkbook(event.target.result);
        setRecords(parsed);
        setParseError("");
        setLastVisitDate("");
      } catch (error) {
        setRecords(null);
        setParseError(error instanceof Error ? error.message : "Could not parse this report.");
      }
    };
    reader.readAsArrayBuffer(file);
  }

  async function exportExcel() {
    if (!canExport || !report) return;

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Report");
    const accentArgb = `FF${template.accent.replace("#", "").toUpperCase()}`;
    const thin = { style: "thin", color: { argb: "FFD8C1C0" } };
    const border = { top: thin, left: thin, bottom: thin, right: thin };

    sheet.columns = [{ width: 8 }, { width: 32 }, { width: 42 }, { width: 14 }];

    sheet.mergeCells("A1:D1");
    const titleCell = sheet.getCell("A1");
    titleCell.value = template.reportTitle;
    titleCell.font = { bold: true, size: 14, color: { argb: "FFFFFFFF" } };
    titleCell.alignment = { horizontal: "center", vertical: "middle" };
    titleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: accentArgb } };
    sheet.getRow(1).height = 24;

    const metadata = [
      [template.territoryLabel, territory],
      [template.routeLabel, routeOptions.find((route) => route.key === routeKeyValue)?.label || ""],
      [template.totalLabel, formatMoney(report.totalSale)],
      [template.dateLabel, lastVisitDate || ""],
    ];

    metadata.forEach(([label, value], index) => {
      const rowNumber = index + 2;
      sheet.getCell(`A${rowNumber}`).value = label;
      sheet.getCell(`A${rowNumber}`).font = { bold: true };
      sheet.getCell(`A${rowNumber}`).border = border;
      sheet.mergeCells(`B${rowNumber}:D${rowNumber}`);
      const valueCell = sheet.getCell(`B${rowNumber}`);
      valueCell.value = value;
      valueCell.border = border;
    });

    const headerRowNumber = 7;
    [template.colNo, template.colOutlet, template.colProduct, template.colQty].forEach((header, index) => {
      const cell = sheet.getCell(headerRowNumber, index + 1);
      cell.value = header;
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: accentArgb } };
      cell.border = border;
      cell.alignment = { vertical: "middle", horizontal: index === 3 ? "right" : "left" };
    });

    let currentRow = headerRowNumber;
    report.outlets.forEach((outlet) => {
      const startRow = currentRow + 1;
      outlet.items.forEach((item, itemIndex) => {
        currentRow += 1;
        const excelRow = sheet.getRow(currentRow);
        excelRow.getCell(1).value = itemIndex === 0 ? outlet.no : null;
        excelRow.getCell(2).value = itemIndex === 0 ? outlet.outletName : null;
        excelRow.getCell(3).value = item.product;
        excelRow.getCell(4).value = item.qty;

        for (let column = 1; column <= 4; column += 1) {
          excelRow.getCell(column).border = border;
          excelRow.getCell(column).alignment = {
            vertical: "middle",
            horizontal: column === 4 ? "right" : "left",
          };
        }
      });

      if (outlet.items.length > 1) {
        sheet.mergeCells(startRow, 1, currentRow, 1);
        sheet.mergeCells(startRow, 2, currentRow, 2);
      }
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    const routePart = sanitizeFilePart(routeOptions.find((route) => route.key === routeKeyValue)?.label || routeKeyValue);
    link.href = url;
    link.download = `${sanitizeFilePart(territory)}_${routePart}_LastVisit.xlsx`;
    link.click();
    URL.revokeObjectURL(url);
  }

  function exportPdf() {
    if (!canExport || !report) return;

    const routeLabelValue = routeOptions.find((route) => route.key === routeKeyValue)?.label || routeKeyValue;
    const rowsHtml = report.outlets
      .map((outlet) =>
        outlet.items
          .map((item, itemIndex) => {
            const cells = [];
            if (itemIndex === 0) {
              cells.push(`<td rowspan="${outlet.items.length}">${escapeHtml(outlet.no)}</td>`);
              cells.push(`<td rowspan="${outlet.items.length}">${escapeHtml(outlet.outletName)}</td>`);
            }
            cells.push(`<td>${escapeHtml(item.product)}</td>`);
            cells.push(`<td class="num">${escapeHtml(item.qty)}</td>`);
            return `<tr>${cells.join("")}</tr>`;
          })
          .join("")
      )
      .join("");

    const printWindow = window.open("", "_blank");
    if (!printWindow) return;

    printWindow.document.write(`
      <!doctype html>
      <html>
        <head>
          <title>${escapeHtml(template.reportTitle)}</title>
          <style>
            :root {
              color-scheme: light;
              --accent: ${template.accent};
              --bg: ${THEME.bg};
              --border: ${THEME.border};
              --text: ${THEME.text};
              --muted: ${THEME.textMuted};
            }
            * { box-sizing: border-box; }
            body {
              margin: 0;
              padding: 24px;
              background: var(--bg);
              color: var(--text);
              font-family: ${THEME.bodyFont};
            }
            .sheet {
              max-width: 1100px;
              margin: 0 auto;
              background: #fff;
              border: 1px solid var(--border);
              border-radius: 18px;
              overflow: hidden;
            }
            .band {
              background: var(--accent);
              color: #fff;
              padding: 24px 28px;
              display: flex;
              justify-content: space-between;
              gap: 16px;
              align-items: flex-start;
            }
            .band h1 {
              margin: 0;
              font-family: ${THEME.headingFont};
              font-size: 22px;
              font-weight: 700;
              letter-spacing: 0.02em;
              text-transform: uppercase;
            }
            .band .sub {
              margin-top: 6px;
              font-size: 13px;
              opacity: 0.85;
            }
            .meta {
              display: grid;
              grid-template-columns: 1fr 1fr;
              gap: 20px;
              padding: 22px 28px;
              border-bottom: 1px solid var(--border);
            }
            .meta-grid {
              display: grid;
              grid-template-columns: 170px 1fr;
              gap: 10px 16px;
            }
            .meta-label {
              font-family: ${THEME.labelFont};
              font-size: 10px;
              letter-spacing: 0.14em;
              text-transform: uppercase;
              color: var(--muted);
              font-weight: 600;
            }
            .meta-value.total {
              color: var(--accent);
              font-size: 24px;
              font-family: ${THEME.headingFont};
              font-weight: 700;
            }
            .table-wrap {
              padding: 0 28px 28px;
            }
            table {
              width: 100%;
              border-collapse: collapse;
              margin-top: 20px;
            }
            thead th {
              text-align: left;
              background: #f0ede9;
              color: var(--muted);
              font-size: 10px;
              letter-spacing: 0.14em;
              text-transform: uppercase;
              font-family: ${THEME.labelFont};
              padding: 12px 14px;
              border: 1px solid var(--border);
              border-top: 2px solid var(--accent);
            }
            tbody td {
              padding: 10px 14px;
              border: 1px solid var(--border);
              vertical-align: top;
            }
            tbody td.num {
              text-align: right;
              font-variant-numeric: tabular-nums;
            }
            @media print {
              body { padding: 0; }
              .sheet { border: none; border-radius: 0; }
            }
          </style>
        </head>
        <body>
          <div class="sheet">
            <div class="band">
              <div>
                <h1>${escapeHtml(template.reportTitle)}</h1>
                <div class="sub">${escapeHtml(territory)}${routeLabelValue ? ` - ${escapeHtml(routeLabelValue)}` : ""}</div>
              </div>
              <div style="text-align:right; font-size: 13px;">
                <div style="font-family: ${THEME.labelFont}; text-transform: uppercase; letter-spacing: 0.14em; font-size: 10px; opacity: 0.8;">${escapeHtml(template.dateLabel)}</div>
                <div style="margin-top: 6px; font-weight: 600;">${escapeHtml(lastVisitDate || "—")}</div>
              </div>
            </div>
            <div class="meta">
              <div class="meta-grid">
                <div class="meta-label">${escapeHtml(template.territoryLabel)}</div>
                <div>${escapeHtml(territory)}</div>
                <div class="meta-label">${escapeHtml(template.routeLabel)}</div>
                <div>${escapeHtml(routeLabelValue)}</div>
              </div>
              <div style="display:flex; flex-direction:column; align-items:flex-end; justify-content:flex-end;">
                <div class="meta-label">${escapeHtml(template.totalLabel)}</div>
                <div class="meta-value total">${escapeHtml(formatMoney(report.totalSale))}</div>
              </div>
            </div>
            <div class="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>${escapeHtml(template.colNo)}</th>
                    <th>${escapeHtml(template.colOutlet)}</th>
                    <th>${escapeHtml(template.colProduct)}</th>
                    <th style="text-align:right">${escapeHtml(template.colQty)}</th>
                  </tr>
                </thead>
                <tbody>${rowsHtml}</tbody>
              </table>
            </div>
          </div>
        </body>
      </html>
    `);
    printWindow.document.close();
    printWindow.focus();
    setTimeout(() => printWindow.print(), 350);
  }

  function saveTemplate() {
    if (saveFlashRef.current) window.clearTimeout(saveFlashRef.current);
    setSaveFlash(true);
    saveFlashRef.current = window.setTimeout(() => setSaveFlash(false), 1500);
  }

  function resetTemplate() {
    setTemplate(DEFAULT_TEMPLATE);
    if (saveFlashRef.current) window.clearTimeout(saveFlashRef.current);
    setSaveFlash(false);
  }

  function resetFilters() {
    if (!territoryOptions.length) return;
    const firstTerritory = territoryOptions[0];
    const firstRoute =
      records
        ?.filter((row) => row.territory === firstTerritory)
        .map((row) => ({ key: row.routeKey, label: row.routeDisplay }))
        .find((route, index, all) => all.findIndex((candidate) => candidate.key === route.key) === index)?.key || "";

    setTerritory(firstTerritory);
    setRouteKeyValue(firstRoute);
    setLastVisitDate("");
    setStrategy("fallback");
  }

  const selectedRouteLabel = routeOptions.find((route) => route.key === routeKeyValue)?.label || "";

  return (
    <div className="min-h-screen" style={{ background: THEME.bg, color: THEME.text, fontFamily: THEME.bodyFont }}>
      <header className="sticky top-0 z-30 border-b" style={{ background: THEME.surface, borderColor: THEME.border }}>
        <div className="mx-auto grid max-w-[1540px] grid-cols-[1fr_auto_1fr] items-center gap-4 px-5 py-3 md:px-8">
          <div className="flex min-w-0 items-center gap-3">
            <img src="/logo.png" alt="Lakmee" className="h-9 w-auto object-contain" />
            <div className="min-w-0">
              <div className="truncate text-[18px] font-bold" style={{ fontFamily: THEME.headingFont, color: template.accent }}>
                Route Sales Report Builder
              </div>
            </div>
          </div>

          <div className="rounded-full p-1" style={{ background: THEME.surfaceLow, border: `1px solid ${THEME.border}` }}>
            <div className="flex items-center gap-1">
              <button
                type="button"
                onClick={() => setScreen("dashboard")}
                className="rounded-full px-4 py-2 text-xs font-semibold uppercase tracking-[0.14em] transition"
                style={{
                  background: screen === "dashboard" ? template.accent : "transparent",
                  color: screen === "dashboard" ? "#fff" : THEME.textMuted,
                  fontFamily: THEME.labelFont,
                }}
              >
                Dashboard
              </button>
              <button
                type="button"
                onClick={() => setScreen("template")}
                className="rounded-full px-4 py-2 text-xs font-semibold uppercase tracking-[0.14em] transition"
                style={{
                  background: screen === "template" ? template.accent : "transparent",
                  color: screen === "template" ? "#fff" : THEME.textMuted,
                  fontFamily: THEME.labelFont,
                }}
              >
                Template Editor
              </button>
            </div>
          </div>

          <div className="flex items-center justify-end gap-3">
            <TopBarButton icon="download" title="Download Excel" onClick={exportExcel} disabled={!canExport} accent={template.accent} />
            <TopBarButton icon="picture_as_pdf" title="Download PDF" onClick={exportPdf} disabled={!canExport} accent={template.accent} />
            <div className="flex h-9 w-9 items-center justify-center rounded-full text-sm font-semibold text-white" style={{ background: template.accent, fontFamily: THEME.labelFont }}>
              RT
            </div>
          </div>
        </div>
      </header>

      <main className="mx-auto max-w-[1540px] px-5 py-6 md:px-8 md:py-8">
        {screen === "dashboard" ? (
          <Dashboard
            template={template}
            records={records}
            parseError={parseError}
            fileName={fileName}
            territory={territory}
            setTerritory={setTerritory}
            routeKeyValue={routeKeyValue}
            setRouteKeyValue={setRouteKeyValue}
            lastVisitDate={lastVisitDate}
            setLastVisitDate={setLastVisitDate}
            strategy={strategy}
            setStrategy={setStrategy}
            onFileSelect={loadFile}
            onDropFile={loadFile}
            onResetFilters={resetFilters}
            exportExcel={exportExcel}
            exportPdf={exportPdf}
            report={report}
            territoryOptions={territoryOptions}
            routeOptions={routeOptions}
            page={page}
            setPage={setPage}
            canExport={canExport}
          />
        ) : (
          <TemplateEditor
            template={template}
            setTemplate={setTemplate}
            report={report}
            lastVisitDate={lastVisitDate}
            onResetToDefaults={resetTemplate}
            onSaveTemplate={saveTemplate}
            saveFlash={saveFlash}
          />
        )}
      </main>
    </div>
  );
}

export default App;