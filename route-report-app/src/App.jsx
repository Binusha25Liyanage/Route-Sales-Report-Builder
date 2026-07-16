import { useState, useEffect, useRef, useMemo, Fragment } from "react";
import * as XLSX from "xlsx";
import ExcelJS from "exceljs";
import { jsPDF } from "jspdf";
import autoTable from "jspdf-autotable";
import { NOTO_SANS_SINHALA_BASE64 } from "./notoSansSinhalaFont.js";

const SALES_DEFAULT_COLUMNS = [
  { id: "c1", label: "NO", field: "no" },
  { id: "c2", label: "Outlet Name", field: "outlet" },
  { id: "c3", label: "Product Name", field: "product" },
  { id: "c4", label: "Quantity", field: "qty" },
];

const DEFAULT_TEMPLATE = {
  reportTitle: "Last Visit Outlet Wise Quantity Sales",
  territoryLabel: "Territory",
  routeLabel: "Route",
  totalLabel: "Total sale value",
  dateLabel: "Last visit date",
  columns: SALES_DEFAULT_COLUMNS,
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

const ATTENDANCE_WEEK_SIZE = 7;

const ATTENDANCE_DEFAULT_COLUMNS = [
  { id: "a1", label: "Employee ID", field: "empId" },
  { id: "a2", label: "First Name", field: "name" },
  { id: "a3", label: "Department", field: "dept" },
];

const ATTENDANCE_DEFAULT_TEMPLATE = {
  reportTitle: "Employee Daily Attendance",
  columns: ATTENDANCE_DEFAULT_COLUMNS,
  checkInLabel: "Check in",
  checkOutLabel: "Check out",
  accent: "#7A2E33",
  bandingEnabled: true,
  bandingColor: "#F3E3E1",
};

// ---------- field options for the column-mapping dropdowns ----------
const SALES_FIELD_OPTIONS = [
  { value: "no", label: "Row number (auto)" },
  { value: "outlet", label: "Outlet name (from data)" },
  { value: "product", label: "Product name (from data)" },
  { value: "qty", label: "Quantity (from data)" },
  { value: "", label: "Custom / blank column" },
];
const SALES_MERGE_FIELDS = new Set(["no", "outlet"]);

const ATTENDANCE_FIELD_OPTIONS = [
  { value: "empId", label: "Employee ID (from data)" },
  { value: "name", label: "Employee name (from data)" },
  { value: "dept", label: "Department (from data)" },
  { value: "", label: "Custom / blank column" },
];

const DAILY_TX_DEFAULT_COLUMNS = [
  { id: "d1", label: "Employee ID", field: "empId" },
  { id: "d2", label: "First Name", field: "name" },
  { id: "d3", label: "Department", field: "dept" },
];

const DAILY_TX_DEFAULT_TEMPLATE = {
  reportTitle: "Transaction",
  columns: DAILY_TX_DEFAULT_COLUMNS,
  timeLabel: "Time",
  stateLabel: "Punch State",
  absentPlaceholder: "-",
  accent: "#7A2E33",
  bandingEnabled: true,
  bandingColor: "#F3E3E1",
};

const DAILY_TX_FIELD_OPTIONS = [
  { value: "empId", label: "Employee ID (from data)" },
  { value: "name", label: "Employee name (from data)" },
  { value: "dept", label: "Department (from data)" },
  { value: "", label: "Custom / blank column" },
];

function hexToRgb(hex) {
  const clean = String(hex || "#7A2E33").replace("#", "");
  const r = parseInt(clean.substring(0, 2), 16);
  const g = parseInt(clean.substring(2, 4), 16);
  const b = parseInt(clean.substring(4, 6), 16);
  return [r, g, b];
}

function createUnicodePdf(options) {
  const doc = new jsPDF(options);
  doc.addFileToVFS("NotoSansSinhala.ttf", NOTO_SANS_SINHALA_BASE64);
  doc.addFont("NotoSansSinhala.ttf", "NotoSansSinhala", "normal");
  doc.addFont("NotoSansSinhala.ttf", "NotoSansSinhala", "bold");
  doc.setFont("NotoSansSinhala", "normal");
  return doc;
}

function makeColumnId() {
  return `col-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
}

function normalizeHeaderText(v) {
  return String(v || "").trim().toLowerCase();
}

function autoMapField(label, kind) {
  const t = normalizeHeaderText(label);
  if (kind === "sales") {
    if (["no", "sl", "s.no", "sl no", "#", "serial", "serial no"].includes(t)) return "no";
    if (t.includes("outlet") || t.includes("shop") || t.includes("customer")) return "outlet";
    if (t.includes("product") || t.includes("item") || t.includes("description")) return "product";
    if (t.includes("qty") || t.includes("quantity")) return "qty";
    return "";
  }
  if (kind === "attendance") {
    if (t.includes("employee id") || t.includes("emp id") || t === "id" || t.includes("employee no")) return "empId";
    if (t.includes("name")) return "name";
    if (t.includes("dept") || t.includes("department")) return "dept";
    return "";
  }
  return "";
}

function parseImportedColumns(workbook, kind) {
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const arr = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true });

    if (kind === "sales") {
      for (const row of arr) {
        const cells = (row || []).map((c) => normalizeHeaderText(c));
        const hasQty = cells.some((c) => c.includes("qty") || c.includes("quantity"));
        const hasOutletOrProduct = cells.some((c) => c.includes("outlet") || c.includes("product") || c.includes("item"));
        if (hasQty && hasOutletOrProduct) {
          const detected = (row || []).filter((c) => c !== null && c !== undefined && String(c).trim() !== "");
          if (detected.length >= 2) {
            return detected.map((label) => ({ id: makeColumnId(), label: String(label).trim(), field: autoMapField(label, "sales") }));
          }
        }
      }
    } else if (kind === "attendance") {
      for (const row of arr) {
        const cells = (row || []).map((c) => normalizeHeaderText(c));
        const checkInIdx = cells.findIndex((c) => c.includes("check in") || c.includes("check-in") || c === "checkin");
        const hasEmployeeHints = cells.some((c) => c.includes("employee") || c.includes("department") || c.includes("first name"));
        if (checkInIdx > 0 && hasEmployeeHints) {
          const leading = (row || []).slice(0, checkInIdx).filter((c) => c !== null && c !== undefined && String(c).trim() !== "");
          if (leading.length >= 1) {
            return leading.map((label) => ({ id: makeColumnId(), label: String(label).trim(), field: autoMapField(label, "attendance") }));
          }
        }
      }
    } else if (kind === "dailytx") {
      for (const row of arr) {
        const cells = (row || []).map((c) => normalizeHeaderText(c));
        const punchIdx = cells.findIndex((c) => c.includes("time"));
        const hasEmployeeHints = cells.some((c) => c.includes("employee") || c.includes("department") || c.includes("first name"));
        if (punchIdx > 0 && hasEmployeeHints) {
          const leading = (row || []).slice(0, punchIdx).filter((c) => c !== null && c !== undefined && String(c).trim() !== "");
          if (leading.length >= 1) {
            return leading.map((label) => ({ id: makeColumnId(), label: String(label).trim(), field: autoMapField(label, "attendance") }));
          }
        }
      }
    }
  }
  throw new Error(
    kind === "sales"
      ? "Couldn't detect a header row with outlet/product and quantity columns in this file."
      : kind === "dailytx"
      ? "Couldn't detect a header row with Employee ID/Department columns followed by Time/Punch State in this file."
      : "Couldn't detect a header row with Employee ID/Department columns followed by Check in/Check out in this file."
  );
}

function parseAttendanceWorkbook(workbook) {
  let headerInfo = null;
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const arr = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true });
    for (let i = 0; i < arr.length; i++) {
      const r = arr[i] || [];
      const norm = r.map((c) => String(c || "").trim().toLowerCase());
      if (norm.includes("employee id") && norm.includes("date") && norm.includes("time")) {
        headerInfo = { arr, headerRow: i, cols: {
          empId: norm.indexOf("employee id"),
          name: norm.indexOf("first name"),
          dept: norm.indexOf("department"),
          date: norm.indexOf("date"),
          time: norm.indexOf("time"),
        } };
        break;
      }
    }
    if (headerInfo) break;
  }
  if (!headerInfo) throw new Error("Couldn't find a sheet with Employee ID / Date / Time columns. Is this an attendance transaction export?");

  const { arr, headerRow, cols } = headerInfo;
  const employees = new Map();
  const byEmpDate = new Map();
  const dateSet = new Set();

  for (let i = headerRow + 1; i < arr.length; i++) {
    const r = arr[i] || [];
    const empId = r[cols.empId];
    const date = r[cols.date];
    const time = r[cols.time];
    if (empId === null || empId === undefined || empId === "" || !date || !time) continue;

    const empIdStr = String(empId).trim();
    const dateStr = String(date).trim();
    const timeStr = String(time).trim();

    if (!employees.has(empIdStr)) {
      employees.set(empIdStr, { empId: empIdStr, name: String(r[cols.name] || "").trim(), dept: String(r[cols.dept] || "").trim() });
    }
    dateSet.add(dateStr);

    const key = `${empIdStr}|${dateStr}`;
    if (!byEmpDate.has(key)) byEmpDate.set(key, []);
    byEmpDate.get(key).push(timeStr);
  }

  if (employees.size === 0) throw new Error("No attendance transaction rows were found.");

  byEmpDate.forEach((times) => times.sort());

  const dates = [...dateSet].sort();
  const employeeList = [...employees.values()].sort((a, b) => Number(a.empId) - Number(b.empId) || a.empId.localeCompare(b.empId));

  const weeks = [];
  for (let i = 0; i < dates.length; i += ATTENDANCE_WEEK_SIZE) {
    weeks.push(dates.slice(i, i + ATTENDANCE_WEEK_SIZE));
  }

  return { employees: employeeList, byEmpDate, dates, weeks };
}

function formatAttDate(iso) {
  const [y, m, d] = iso.split("-").map(Number);
  return new Date(y, m - 1, d).toLocaleDateString(undefined, { day: "numeric", month: "short" });
}

// ---------- Daily Transaction report ----------
function normHeader(v) {
  return String(v || "").trim().toLowerCase();
}

function parseDailyTransactionWorkbook(workbook) {
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const arr = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true });

    let headerRowIdx = -1;
    let colMap = null;

    for (let i = 0; i < arr.length; i++) {
      const cells = (arr[i] || []).map(normHeader);
      const hasEmpId = cells.some((c) => c.includes("employee id") || c === "emp id" || c === "id");
      const hasPunch = cells.some((c) => c.includes("punch state") || c === "state");
      if (hasEmpId && hasPunch) {
        headerRowIdx = i;
        colMap = {};
        cells.forEach((c, ci) => {
          if (c.includes("employee id") || c === "emp id") colMap.empId = ci;
          else if (c === "id" && colMap.empId === undefined) colMap.empId = ci;
          else if (c.includes("first name")) colMap.name = ci;
          else if (c === "name" && colMap.name === undefined) colMap.name = ci;
          else if (c.includes("department") || c === "dept") colMap.dept = ci;
          else if (c === "date") colMap.date = ci;
          else if (c === "time") colMap.time = ci;
          else if (c.includes("punch state") || c === "state") colMap.state = ci;
        });
        break;
      }
    }
    if (headerRowIdx === -1 || colMap.empId === undefined) continue;

    let reportDate = null;
    if (colMap.date === undefined) {
      for (let i = 0; i < headerRowIdx; i++) {
        const row = arr[i] || [];
        for (const cell of row) {
          const s = String(cell || "").trim();
          const m = s.match(/^date\s*:?\s*(.+)$/i);
          if (m && m[1] && m[1].trim()) { reportDate = m[1].trim(); break; }
        }
        if (reportDate) break;
      }
    }

    const employees = new Map();
    const punches = new Map();

    for (let i = headerRowIdx + 1; i < arr.length; i++) {
      const row = arr[i] || [];
      if (row.every((c) => c === null || c === undefined || String(c).trim() === "")) continue;
      const empId = String(row[colMap.empId] ?? "").trim();
      if (!empId) continue;
      const name = colMap.name !== undefined ? String(row[colMap.name] ?? "").trim() : "";
      const dept = colMap.dept !== undefined ? String(row[colMap.dept] ?? "").trim() : "";
      const time = colMap.time !== undefined ? String(row[colMap.time] ?? "").trim() : "";
      const state = colMap.state !== undefined ? String(row[colMap.state] ?? "").trim() : "";
      if (colMap.date !== undefined && !reportDate) {
        const d = String(row[colMap.date] ?? "").trim();
        if (d) reportDate = d;
      }

      if (!employees.has(empId)) employees.set(empId, { empId, name, dept });
      else {
        const existing = employees.get(empId);
        if (!existing.name && name) existing.name = name;
        if (!existing.dept && dept) existing.dept = dept;
      }

      if (time || state) {
        if (!punches.has(empId)) punches.set(empId, []);
        const list = punches.get(empId);
        if (!list.some((p) => p.time === time && p.state === state)) list.push({ time, state });
      }
    }

    if (employees.size === 0) continue;
    punches.forEach((list) => list.sort((a, b) => a.time.localeCompare(b.time)));

    return {
      date: reportDate,
      employees: [...employees.values()].sort((a, b) => Number(a.empId) - Number(b.empId) || a.empId.localeCompare(b.empId)),
      punches,
    };
  }
  throw new Error("Couldn't find a header row with Employee ID and Punch State columns in this file.");
}

function mergeRoster(existingRoster, newEmployees) {
  const map = new Map(existingRoster.map((e) => [e.empId, e]));
  newEmployees.forEach((e) => {
    if (map.has(e.empId)) {
      const cur = map.get(e.empId);
      map.set(e.empId, { empId: e.empId, name: e.name || cur.name, dept: e.dept || cur.dept });
    } else {
      map.set(e.empId, e);
    }
  });
  return [...map.values()].sort((a, b) => Number(a.empId) - Number(b.empId) || a.empId.localeCompare(b.empId));
}

function buildDailyTxRows(roster, parsed, absentPlaceholder) {
  const source = roster && roster.length > 0 ? roster : parsed.employees;
  const rows = [];
  source.forEach((emp) => {
    const recs = parsed.punches.get(emp.empId) || [];
    const empInfo = parsed.employees.find((e) => e.empId === emp.empId) || emp;
    if (recs.length === 0) {
      rows.push({ empId: empInfo.empId, name: empInfo.name || emp.name, dept: empInfo.dept || emp.dept, time: absentPlaceholder, state: absentPlaceholder, absent: true, groupKey: emp.empId });
    } else {
      recs.forEach((r) => rows.push({ empId: empInfo.empId, name: empInfo.name || emp.name, dept: empInfo.dept || emp.dept, time: r.time, state: r.state, absent: false, groupKey: emp.empId }));
    }
  });
  return rows;
}

function dtCellValue(field, emp) {
  switch (field) {
    case "empId": return emp.empId;
    case "name": return emp.name;
    case "dept": return emp.dept;
    default: return "";
  }
}

function weekRangeLabel(weekDates) {
  if (!weekDates.length) return "";
  const first = formatAttDate(weekDates[0]);
  const last = formatAttDate(weekDates[weekDates.length - 1]);
  return weekDates.length > 1 ? `${first} – ${last}` : first;
}

function buildAttendanceRows(parsed, weekDates) {
  return parsed.employees.map((emp) => ({
    ...emp,
    cells: weekDates.map((d) => {
      const times = parsed.byEmpDate.get(`${emp.empId}|${d}`) || [];
      if (times.length === 0) return { checkIn: "", checkOut: "" };
      if (times.length === 1) return { checkIn: times[0], checkOut: "" };
      return { checkIn: times[0], checkOut: times[times.length - 1] };
    }),
  }));
}

function attCellValue(field, emp) {
  switch (field) {
    case "empId": return emp.empId;
    case "name": return emp.name;
    case "dept": return emp.dept;
    default: return "";
  }
}

function escapeAttHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

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

function salesCellValue(field, { outletIdx, item, isFirst, outlet }) {
  switch (field) {
    case "no": return isFirst ? outletIdx + 1 : "";
    case "outlet": return isFirst ? outlet.name : "";
    case "product": return item.product;
    case "qty": return item.qty;
    default: return "";
  }
}

function getTerritoryRouteCombos(rows) {
  const seen = new Set();
  const combos = [];
  rows.forEach((r) => {
    const key = `${r.territory}|||${r.routeName || ""}`;
    if (!seen.has(key)) {
      seen.add(key);
      combos.push({ territory: r.territory, routeName: r.routeName || "" });
    }
  });
  return combos;
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

function ColumnListEditor({ columns, setColumns, fieldOptions, accent, importKind, importLabel }) {
  const fileRef = useRef(null);
  const [importError, setImportError] = useState("");

  function updateCol(id, patch) {
    setColumns(columns.map((c) => (c.id === id ? { ...c, ...patch } : c)));
  }
  function removeCol(id) {
    if (columns.length <= 1) return;
    setColumns(columns.filter((c) => c.id !== id));
  }
  function addCol() {
    setColumns([...columns, { id: makeColumnId(), label: "New Column", field: "" }]);
  }
  function move(id, dir) {
    const idx = columns.findIndex((c) => c.id === id);
    const swapIdx = idx + dir;
    if (swapIdx < 0 || swapIdx >= columns.length) return;
    const next = [...columns];
    [next[idx], next[swapIdx]] = [next[swapIdx], next[idx]];
    setColumns(next);
  }

  function handleImportFile(file) {
    setImportError("");
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, { type: "array" });
        const detected = parseImportedColumns(wb, importKind);
        setColumns(detected);
      } catch (err) {
        setImportError(err.message || "Couldn't read that file's column layout.");
      }
    };
    reader.onerror = () => setImportError("The file could not be read.");
    reader.readAsArrayBuffer(file);
  }

  return (
    <div>
      <div className="flex items-center justify-between mb-2">
        <Label>{importLabel || "Columns"}</Label>
        <button
          type="button"
          onClick={() => fileRef.current?.click()}
          className="flex items-center gap-1 text-xs font-semibold"
          style={{ color: accent, fontFamily: T.bodyFont }}
        >
          <Icon name="upload_file" size={14} /> Import from format file
        </button>
        <input ref={fileRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={(e) => { if (e.target.files[0]) handleImportFile(e.target.files[0]); e.target.value = ""; }} />
      </div>
      {importError && (
        <p className="text-xs flex items-start gap-1.5 mb-2" style={{ color: "#ba1a1a" }}><Icon name="error" size={14} />{importError}</p>
      )}

      <div className="space-y-2">
        {columns.map((c, i) => (
          <div key={c.id} className="flex items-center gap-1.5">
            <div className="flex flex-col">
              <button type="button" onClick={() => move(c.id, -1)} disabled={i === 0} className="disabled:opacity-25" style={{ color: T.textMuted, lineHeight: 0.7 }}>
                <Icon name="expand_less" size={16} />
              </button>
              <button type="button" onClick={() => move(c.id, 1)} disabled={i === columns.length - 1} className="disabled:opacity-25" style={{ color: T.textMuted, lineHeight: 0.7 }}>
                <Icon name="expand_more" size={16} />
              </button>
            </div>
            <input
              value={c.label}
              onChange={(e) => updateCol(c.id, { label: e.target.value })}
              className="flex-1 min-w-0 rounded-md px-2.5 py-1.5 text-sm outline-none"
              style={{ background: T.surface, border: `1px solid ${T.border}`, color: T.text, fontFamily: T.bodyFont }}
            />
            <select
              value={c.field || ""}
              onChange={(e) => updateCol(c.id, { field: e.target.value })}
              className="w-40 shrink-0 rounded-md px-2 py-1.5 text-xs outline-none"
              style={{ background: T.surface, border: `1px solid ${T.border}`, color: T.textMuted, fontFamily: T.bodyFont }}
            >
              {fieldOptions.map((opt) => <option key={opt.value} value={opt.value}>{opt.label}</option>)}
            </select>
            <button type="button" onClick={() => removeCol(c.id)} disabled={columns.length <= 1} className="shrink-0 w-7 h-7 rounded-md flex items-center justify-center disabled:opacity-25" style={{ color: "#ba1a1a" }}>
              <Icon name="delete" size={16} />
            </button>
          </div>
        ))}
      </div>

      <button
        type="button"
        onClick={addCol}
        className="mt-2 flex items-center gap-1.5 text-xs font-semibold"
        style={{ color: accent, fontFamily: T.bodyFont }}
      >
        <Icon name="add" size={15} /> Add column
      </button>
      <p className="text-xs mt-2" style={{ color: T.textMuted }}>Columns mapped to "Custom / blank" always appear empty in the output for you to fill in by hand. Use the arrows to reorder.</p>
    </div>
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

function TemplateEditor({ template, setTemplate, onReset, previewData, attTemplate, setAttTemplate, onResetAtt, dtTemplate, setDtTemplate, onResetDt, templateMode, setTemplateMode }) {
  const [saved, setSaved] = useState(false);
  const isAtt = templateMode === "attendance";
  const isDt = templateMode === "dailytx";
  const current = isDt ? dtTemplate : isAtt ? attTemplate : template;
  const setCurrent = isDt ? setDtTemplate : isAtt ? setAttTemplate : setTemplate;
  const field = (key) => ({ value: current[key], onChange: (v) => setCurrent({ ...current, [key]: v }), accent: current.accent });

  function handleSave() {
    setSaved(true);
    setTimeout(() => setSaved(false), 1500);
  }

  const sampleWeekLabel1 = "01 Jul";
  const sampleWeekLabel2 = "02 Jul";

  return (
    <div className="grid grid-cols-1 lg:grid-cols-12 gap-5">
      <section className="lg:col-span-5 flex flex-col gap-5">
        <div className="rounded-lg p-1 flex items-center gap-1 w-fit flex-wrap" style={{ background: T.surfaceLow, border: `1px solid ${T.border}` }}>
          <button
            onClick={() => setTemplateMode("sales")}
            className="px-4 py-1.5 rounded-md text-xs font-semibold transition-colors"
            style={{ fontFamily: T.capsFont, letterSpacing: "0.06em", background: !isAtt && !isDt ? template.accent : "transparent", color: !isAtt && !isDt ? "#fff" : T.textMuted }}
          >
            SALES REPORT
          </button>
          <button
            onClick={() => setTemplateMode("attendance")}
            className="px-4 py-1.5 rounded-md text-xs font-semibold transition-colors"
            style={{ fontFamily: T.capsFont, letterSpacing: "0.06em", background: isAtt ? attTemplate.accent : "transparent", color: isAtt ? "#fff" : T.textMuted }}
          >
            ATTENDANCE
          </button>
          <button
            onClick={() => setTemplateMode("dailytx")}
            className="px-4 py-1.5 rounded-md text-xs font-semibold transition-colors"
            style={{ fontFamily: T.capsFont, letterSpacing: "0.06em", background: isDt ? dtTemplate.accent : "transparent", color: isDt ? "#fff" : T.textMuted }}
          >
            DAILY TRANSACTION
          </button>
        </div>

        <div className="rounded-lg p-6" style={{ background: T.surface, border: `1px solid ${T.border}`, boxShadow: "0px 4px 12px rgba(74,71,67,0.05)" }}>
          <div className="pb-4 mb-6" style={{ borderBottom: `1px solid ${T.border}` }}>
            <h2 style={{ fontFamily: T.headFont, fontSize: 18, fontWeight: 700, color: current.accent }}>Report Configuration</h2>
            <p className="text-sm mt-1" style={{ color: T.textMuted, fontFamily: T.bodyFont }}>Customize your template's layout and metadata labels.</p>
          </div>

          {!isAtt ? (
            <>
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
                <ColumnListEditor
                  columns={template.columns}
                  setColumns={(cols) => setTemplate({ ...template, columns: cols })}
                  fieldOptions={SALES_FIELD_OPTIONS}
                  accent={template.accent}
                  importKind="sales"
                  importLabel="Table columns"
                />
              </div>
            </>
          ) : (
            <>
              <div className="space-y-5">
                <div>
                  <Label>Report title</Label>
                  <TextField {...field("reportTitle")} />
                </div>
                <div className="grid grid-cols-2 gap-4">
                  <div><Label>Check-in label</Label><TextField {...field("checkInLabel")} /></div>
                  <div><Label>Check-out label</Label><TextField {...field("checkOutLabel")} /></div>
                </div>
              </div>

              <div className="pt-5">
                <ColumnListEditor
                  columns={attTemplate.columns}
                  setColumns={(cols) => setAttTemplate({ ...attTemplate, columns: cols })}
                  fieldOptions={ATTENDANCE_FIELD_OPTIONS}
                  accent={attTemplate.accent}
                  importKind="attendance"
                  importLabel="Employee columns (before the daily Check in / Check out)"
                />
              </div>
            </>
          )}

          <div className="pt-5 mt-5" style={{ borderTop: `1px solid ${T.border}` }}>
            <Label>Brand accent color</Label>
            <div className="flex items-center gap-3 flex-wrap">
              {ACCENTS.map((a) => (
                <button
                  key={a.value}
                  onClick={() => setCurrent({ ...current, accent: a.value })}
                  title={a.name}
                  className="w-8 h-8 rounded-full cursor-pointer transition-transform hover:scale-110"
                  style={{
                    backgroundColor: a.value,
                    border: `1px solid ${T.border}`,
                    outline: current.accent === a.value ? `2px solid ${T.text}` : "none",
                    outlineOffset: "2px",
                  }}
                />
              ))}
            </div>
          </div>

          <div className="pt-5 mt-5" style={{ borderTop: `1px solid ${T.border}` }}>
            <div className="flex items-center justify-between mb-2">
              <Label>{isAtt ? "Employee row banding" : "Outlet row banding"}</Label>
              <button
                type="button"
                onClick={() => setCurrent({ ...current, bandingEnabled: !current.bandingEnabled })}
                className="relative w-10 h-5.5 rounded-full transition-colors shrink-0"
                style={{ background: current.bandingEnabled ? current.accent : T.border, height: 22, width: 40 }}
                title={current.bandingEnabled ? "Banding on" : "Banding off"}
              >
                <span
                  className="absolute top-0.5 rounded-full bg-white transition-transform"
                  style={{ width: 18, height: 18, left: 2, transform: current.bandingEnabled ? "translateX(18px)" : "translateX(0)" }}
                />
              </button>
            </div>
            <p className="text-xs mb-3" style={{ color: T.textMuted }}>
              {isAtt ? "Shades every other employee's row so it's easier to scan across the week. Applies to the table, the Excel export, and the PDF export." : "Shades every other outlet's rows so its products are easier to scan. Applies to the table, the Excel export, and the PDF export."}
            </p>
            {current.bandingEnabled && (
              <div className="flex items-center gap-3 flex-wrap">
                {BANDING_COLORS.map((c) => (
                  <button
                    key={c.value}
                    onClick={() => setCurrent({ ...current, bandingColor: c.value })}
                    title={c.name}
                    className="w-8 h-8 rounded-full cursor-pointer transition-transform hover:scale-110"
                    style={{
                      backgroundColor: c.value,
                      border: `1px solid ${T.border}`,
                      outline: current.bandingColor === c.value ? `2px solid ${T.text}` : "none",
                      outlineOffset: "2px",
                    }}
                  />
                ))}
              </div>
            )}
          </div>

          <div className="pt-6 flex items-center justify-between">
            <button onClick={isAtt ? onResetAtt : onReset} className="flex items-center gap-1.5 text-sm transition-colors" style={{ color: T.textMuted, fontFamily: T.bodyFont }}>
              <Icon name="restart_alt" size={18} /> Reset to defaults
            </button>
            <button
              onClick={handleSave}
              className="px-6 py-2.5 rounded-lg font-bold text-sm transition-all active:scale-[0.98]"
              style={{ background: current.accent, color: "#fff", fontFamily: T.bodyFont, boxShadow: "0 2px 6px rgba(0,0,0,0.15)" }}
            >
              {saved ? "Saved ✓" : "Save Template"}
            </button>
          </div>
        </div>
      </section>

      <section className="lg:col-span-7">
        <div className="sticky top-5">
          <Label>Live preview (real-time)</Label>
          {!isAtt ? (
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
                        {template.columns.map((col) => (
                          <th key={col.id} className={`p-2.5 ${col.field === "qty" ? "text-right" : ""}`} style={{ fontFamily: T.capsFont, fontSize: 10 }}>{col.label}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {previewData.sampleRows.map((row, i) => (
                        <tr key={i} style={{ borderBottom: `1px solid ${T.border}` }}>
                          {template.columns.map((col) => (
                            <td key={col.id} className={`p-2.5 ${col.field === "qty" ? "text-right" : ""}`}>{col.field ? row[col.field] : ""}</td>
                          ))}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>
          ) : (
            <div className="rounded-lg overflow-hidden flex flex-col min-h-[560px]" style={{ background: T.surface, border: "1px solid #D9D1C7", boxShadow: "0px 12px 24px rgba(74,71,67,0.12)" }}>
              <div className="p-6" style={{ background: attTemplate.accent, color: "#fff" }}>
                <h3 style={{ fontFamily: T.headFont, fontSize: 22, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.02em" }}>{attTemplate.reportTitle}</h3>
                <p className="text-sm opacity-80 mt-1" style={{ fontFamily: T.bodyFont }}>Week 1 of 4 — {sampleWeekLabel1} – {sampleWeekLabel2}</p>
              </div>

              <div className="p-6 flex-1">
                <table className="w-full border-collapse text-sm" style={{ fontFamily: T.bodyFont }}>
                  <thead>
                    <tr style={{ background: T.surfaceContainer, borderTop: `2px solid ${attTemplate.accent}`, textAlign: "left" }}>
                      {attTemplate.columns.map((col) => (
                        <th key={col.id} className="p-2.5" style={{ fontFamily: T.capsFont, fontSize: 10 }}>{col.label}</th>
                      ))}
                      <th className="p-2.5 text-center" colSpan={2} style={{ fontFamily: T.capsFont, fontSize: 10 }}>{sampleWeekLabel1}</th>
                    </tr>
                    <tr style={{ background: T.surfaceContainer, textAlign: "left" }}>
                      {attTemplate.columns.map((col) => <th key={col.id} className="p-2.5"></th>)}
                      <th className="p-2.5 text-center" style={{ fontFamily: T.capsFont, fontSize: 9, color: T.textMuted }}>{attTemplate.checkInLabel}</th>
                      <th className="p-2.5 text-center" style={{ fontFamily: T.capsFont, fontSize: 9, color: T.textMuted }}>{attTemplate.checkOutLabel}</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr style={{ borderBottom: `1px solid ${T.border}` }}>
                      {attTemplate.columns.map((col) => (
                        <td key={col.id} className="p-2.5">{col.field === "empId" ? "1750" : col.field === "name" ? "Binusha" : col.field === "dept" ? "ICT" : ""}</td>
                      ))}
                      <td className="p-2.5 text-center">08:02</td>
                      <td className="p-2.5 text-center">17:15</td>
                    </tr>
                    <tr style={{ borderBottom: `1px solid ${T.border}`, background: attTemplate.bandingEnabled ? attTemplate.bandingColor : "transparent" }}>
                      {attTemplate.columns.map((col) => (
                        <td key={col.id} className="p-2.5">{col.field === "empId" ? "1751" : col.field === "name" ? "Manohari" : col.field === "dept" ? "SA / Finance" : ""}</td>
                      ))}
                      <td className="p-2.5 text-center">07:58</td>
                      <td className="p-2.5 text-center">17:02</td>
                    </tr>
                  </tbody>
                </table>
              </div>
            </div>
          )}
          <p className="text-xs mt-3" style={{ color: T.textMuted, fontFamily: T.bodyFont }}>Saved automatically, and used every time you generate a report.</p>
        </div>
      </section>
    </div>
  );
}

function AttendanceScreen({
  attTemplate,
  attUploads,
  attActiveId,
  setAttActiveId,
  activeAttUpload,
  attWeekIndex,
  setAttWeekIndex,
  activeAttWeekDates,
  activeAttRows,
  attFileInputRef,
  handleAttFiles,
  removeAttUpload,
  exportAttendanceExcelFor,
  exportAttendancePdfFor,
  exportAllAttendanceWeeksExcel,
  exportAllAttendanceWeeksPdf,
}) {
  const weeks = activeAttUpload && activeAttUpload.parsed ? activeAttUpload.parsed.weeks : [];
  const canExportActive = Boolean(activeAttUpload && activeAttUpload.parsed && activeAttRows.length > 0);

  return (
    <div className="grid grid-cols-1 lg:grid-cols-12 gap-5">
      <div className="lg:col-span-4 flex flex-col gap-5">
        <div className="rounded-lg p-5" style={{ background: T.surface, border: `1px solid ${T.border}` }}>
          <Label>Step 1: Upload attendance report(s)</Label>
          <div
            onClick={() => attFileInputRef.current?.click()}
            onDrop={(e) => { e.preventDefault(); if (e.dataTransfer.files.length) handleAttFiles(e.dataTransfer.files); }}
            onDragOver={(e) => e.preventDefault()}
            className="rounded-lg p-6 text-center cursor-pointer transition-colors"
            style={{ border: `2px dashed ${T.border}` }}
          >
            <Icon name="cloud_upload" size={30} style={{ color: T.textMuted }} />
            <p className="text-sm font-semibold mt-2">Drop files here or click</p>
            <p className="text-xs mt-1" style={{ color: T.textMuted }}>Employee attendance / punch-log export (.xlsx). Weeks are split automatically from whatever date range the file covers.</p>
            <input ref={attFileInputRef} type="file" accept=".xlsx,.xls" multiple className="hidden" onChange={(e) => { if (e.target.files.length) handleAttFiles(e.target.files); e.target.value = ""; }} />
          </div>
          {attUploads.length > 0 && (
            <p className="mt-3 text-sm flex items-center gap-1.5" style={{ color: "#2E7A4E" }}><Icon name="check_circle" size={16} />{attUploads.length} file{attUploads.length === 1 ? "" : "s"} uploaded</p>
          )}
        </div>

        {attUploads.length > 0 && (
          <div className="flex flex-col gap-4">
            {attUploads.map((u) => {
              const isActive = u.id === attActiveId;
              return (
                <div
                  key={u.id}
                  onClick={() => setAttActiveId(u.id)}
                  className="rounded-lg p-4 cursor-pointer transition-colors"
                  style={{ background: T.surface, border: `2px solid ${isActive ? attTemplate.accent : T.border}` }}
                >
                  <div className="flex items-center justify-between mb-2">
                    <div className="flex items-center gap-1.5 min-w-0">
                      <Icon name="badge" size={16} style={{ color: isActive ? attTemplate.accent : T.textMuted, flexShrink: 0 }} />
                      <span className="text-sm font-semibold truncate" title={u.fileName}>{u.fileName}</span>
                    </div>
                    <button
                      onClick={(e) => { e.stopPropagation(); removeAttUpload(u.id); }}
                      className="shrink-0 w-6 h-6 rounded-full flex items-center justify-center hover:bg-black/5"
                      style={{ color: T.textMuted }}
                      title="Remove"
                    >
                      <Icon name="close" size={15} />
                    </button>
                  </div>

                  {u.parseError && (
                    <p className="text-xs flex items-start gap-1.5" style={{ color: "#ba1a1a" }}><Icon name="error" size={14} />{u.parseError}</p>
                  )}

                  {u.parsed && (
                    <p className="text-xs" style={{ color: T.textMuted }}>
                      {u.parsed.employees.length} employees · {u.parsed.dates.length} days · {u.parsed.weeks.length} week sheet{u.parsed.weeks.length === 1 ? "" : "s"}
                    </p>
                  )}

                  {isActive && u.parsed && u.parsed.weeks.length > 0 && (
                    <div className="mt-3 pt-3 flex flex-wrap gap-1.5" style={{ borderTop: `1px solid ${T.border}` }} onClick={(e) => e.stopPropagation()}>
                      {u.parsed.weeks.map((weekDates, i) => (
                        <button
                          key={i}
                          onClick={() => setAttWeekIndex(i)}
                          className="px-2.5 py-1 rounded-md text-xs font-semibold transition-colors"
                          style={{
                            background: attWeekIndex === i ? attTemplate.accent : T.surfaceContainer,
                            color: attWeekIndex === i ? "#fff" : T.textMuted,
                          }}
                          title={weekRangeLabel(weekDates)}
                        >
                          Week {i + 1}
                        </button>
                      ))}
                    </div>
                  )}
                </div>
              );
            })}

            {activeAttUpload && activeAttUpload.parsed && activeAttUpload.parsed.weeks.length > 1 && (
              <>
                <button
                  onClick={() => exportAllAttendanceWeeksExcel(activeAttUpload)}
                  className="w-full py-2.5 rounded-lg text-sm font-semibold flex items-center justify-center gap-2 transition-colors"
                  style={{ background: attTemplate.accent, color: "#fff", fontFamily: T.bodyFont }}
                  title="Choose a folder once — all week files save there automatically"
                >
                  <Icon name="folder_zip" size={16} /> Export All Weeks as Excel ({activeAttUpload.parsed.weeks.length})
                </button>
                <button
                  onClick={() => exportAllAttendanceWeeksPdf(activeAttUpload)}
                  className="w-full py-2.5 rounded-lg text-sm font-semibold flex items-center justify-center gap-2 transition-colors"
                  style={{ background: T.surface, color: attTemplate.accent, border: `1.5px solid ${attTemplate.accent}`, fontFamily: T.bodyFont }}
                  title="Choose a folder once — a PDF for each week saves there automatically"
                >
                  <Icon name="picture_as_pdf" size={16} /> Export All Weeks as PDF ({activeAttUpload.parsed.weeks.length})
                </button>
                <p className="text-xs text-center -mt-2" style={{ color: T.textMuted }}>Excel: pick a destination folder once — every week saves there automatically. PDF: pick a destination folder once — a PDF for each week saves there automatically.</p>
              </>
            )}
          </div>
        )}
      </div>

      <div className="lg:col-span-8">
        {!activeAttUpload || !activeAttUpload.parsed ? (
          <div className="h-full flex flex-col items-center justify-center text-sm py-24 rounded-lg" style={{ border: `2px dashed ${T.border}`, color: T.textMuted }}>
            <Icon name="badge" size={28} style={{ color: T.textMuted }} />
            <p className="mt-2">Upload an attendance report to see the weekly breakdown here</p>
          </div>
        ) : (
          <div className="rounded-lg overflow-hidden" style={{ background: T.surface, border: `1px solid ${T.border}` }}>
            <div className="p-6 flex items-center justify-between flex-wrap gap-3" style={{ background: attTemplate.accent }}>
              <div>
                <h2 style={{ fontFamily: T.headFont, fontSize: 22, fontWeight: 700, color: "#fff" }}>{attTemplate.reportTitle}</h2>
                <p className="text-sm opacity-85 mt-1" style={{ color: "#fff" }}>Week {attWeekIndex + 1} of {weeks.length} — {weekRangeLabel(activeAttWeekDates)}</p>
              </div>
              <div className="flex gap-2">
                <button onClick={() => exportAttendanceExcelFor(activeAttUpload, activeAttWeekDates, attWeekIndex)} disabled={!canExportActive} className="flex items-center gap-1.5 disabled:opacity-40 text-xs font-bold px-3.5 py-2 rounded-lg" style={{ background: "#fff", color: attTemplate.accent }}>
                  <Icon name="grid_on" size={15} /> Download Excel
                </button>
                <button onClick={() => exportAttendancePdfFor(activeAttUpload, activeAttWeekDates)} disabled={!canExportActive} className="flex items-center gap-1.5 disabled:opacity-40 text-xs font-bold px-3.5 py-2 rounded-lg" style={{ background: "rgba(0,0,0,0.2)", color: "#fff" }}>
                  <Icon name="picture_as_pdf" size={15} /> Download PDF
                </button>
              </div>
            </div>

            <div className="p-5 grid grid-cols-3 gap-4" style={{ borderBottom: `1px solid ${T.border}` }}>
              <div className="rounded-lg p-4" style={{ border: `1px solid ${T.border}` }}>
                <p style={{ fontFamily: T.capsFont, fontSize: 10, letterSpacing: "0.08em", color: T.textMuted }}>EMPLOYEES</p>
                <p style={{ fontFamily: T.headFont, fontSize: 22, fontWeight: 700, color: attTemplate.accent }} className="mt-1">{activeAttUpload.parsed.employees.length}</p>
              </div>
              <div className="rounded-lg p-4" style={{ border: `1px solid ${T.border}` }}>
                <p style={{ fontFamily: T.capsFont, fontSize: 10, letterSpacing: "0.08em", color: T.textMuted }}>DAYS THIS WEEK</p>
                <p style={{ fontFamily: T.headFont, fontSize: 22, fontWeight: 700 }} className="mt-1">{activeAttWeekDates.length}</p>
              </div>
              <div className="rounded-lg p-4" style={{ border: `1px solid ${T.border}` }}>
                <p style={{ fontFamily: T.capsFont, fontSize: 10, letterSpacing: "0.08em", color: T.textMuted }}>TOTAL WEEKS</p>
                <p style={{ fontFamily: T.headFont, fontSize: 22, fontWeight: 700 }} className="mt-1">{weeks.length}</p>
              </div>
            </div>

            <div className="overflow-x-auto">
              <table className="text-sm" style={{ minWidth: 620 + activeAttWeekDates.length * 160 }}>
                <thead>
                  <tr>
                    {attTemplate.columns.map((col) => (
                      <th key={col.id} rowSpan={2} className="text-left px-4 py-2.5 align-bottom" style={{ background: T.surfaceContainer, borderTop: `2px solid ${attTemplate.accent}`, fontFamily: T.capsFont, fontSize: 10, letterSpacing: "0.08em", color: T.textMuted }}>{col.label}</th>
                    ))}
                    {activeAttWeekDates.map((d) => (
                      <th key={d} colSpan={2} className="text-center px-2 py-1.5" style={{ background: T.surfaceContainer, borderTop: `2px solid ${attTemplate.accent}`, borderLeft: `1px solid ${T.border}`, fontFamily: T.capsFont, fontSize: 10, letterSpacing: "0.06em", color: T.text }}>{formatAttDate(d)}</th>
                    ))}
                  </tr>
                  <tr>
                    {activeAttWeekDates.map((d) => (
                      <Fragment key={d}>
                        <th className="text-center px-2 py-1.5" style={{ background: T.surfaceContainer, borderLeft: `1px solid ${T.border}`, fontFamily: T.capsFont, fontSize: 9, letterSpacing: "0.06em", color: T.textMuted }}>{attTemplate.checkInLabel}</th>
                        <th className="text-center px-2 py-1.5" style={{ background: T.surfaceContainer, fontFamily: T.capsFont, fontSize: 9, letterSpacing: "0.06em", color: T.textMuted }}>{attTemplate.checkOutLabel}</th>
                      </Fragment>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {activeAttRows.map((emp, idx) => {
                    const banded = attTemplate.bandingEnabled && idx % 2 === 1;
                    return (
                      <tr key={emp.empId} style={{ borderBottom: `1px solid ${T.border}`, background: banded ? attTemplate.bandingColor : "transparent" }}>
                        {attTemplate.columns.map((col, ci) => (
                          <td key={col.id} className={`px-4 py-2 ${ci === 1 ? "font-semibold" : ""}`} style={ci === 0 ? { color: T.textMuted } : undefined}>
                            {col.field === "empId" ? emp.empId : col.field === "name" ? emp.name : col.field === "dept" ? emp.dept : ""}
                          </td>
                        ))}
                        {emp.cells.map((c, i) => (
                          <Fragment key={i}>
                            <td className="px-2 py-2 text-center tabular-nums" style={{ borderLeft: `1px solid ${T.border}` }}>{c.checkIn || "—"}</td>
                            <td className="px-2 py-2 text-center tabular-nums">{c.checkOut || "—"}</td>
                          </Fragment>
                        ))}
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

function DailyTransactionScreen({
  dtTemplate,
  dtUploads,
  dtActiveId,
  setDtActiveId,
  activeDtUpload,
  activeDtRows,
  dtRoster,
  dtFileInputRef,
  handleDtFiles,
  removeDtUpload,
  exportDtExcelFor,
  exportDtPdfFor,
  exportAllDtExcel,
  exportAllDtPdf,
}) {
  const canExportActive = Boolean(activeDtUpload && activeDtUpload.parsed && activeDtRows.length > 0);
  const presentCount = activeDtRows.reduce((set, r) => { if (!r.absent) set.add(r.groupKey); return set; }, new Set()).size;
  const absentCount = dtRoster.length - presentCount;

  return (
    <div className="grid grid-cols-1 lg:grid-cols-12 gap-5">
      <div className="lg:col-span-4 flex flex-col gap-5">
        <div className="rounded-lg p-5" style={{ background: T.surface, border: `1px solid ${T.border}` }}>
          <Label>Step 1: Upload transaction report(s)</Label>
          <div
            onClick={() => dtFileInputRef.current?.click()}
            onDrop={(e) => { e.preventDefault(); if (e.dataTransfer.files.length) handleDtFiles(e.dataTransfer.files); }}
            onDragOver={(e) => e.preventDefault()}
            className="rounded-lg p-6 text-center cursor-pointer transition-colors"
            style={{ border: `2px dashed ${T.border}` }}
          >
            <Icon name="cloud_upload" size={30} style={{ color: T.textMuted }} />
            <p className="text-sm font-semibold mt-2">Drop files here or click</p>
            <p className="text-xs mt-1" style={{ color: T.textMuted }}>Daily punch/transaction export (.xlsx). Any of the file layouts you use are accepted — one file per day.</p>
            <input ref={dtFileInputRef} type="file" accept=".xlsx,.xls" multiple className="hidden" onChange={(e) => { if (e.target.files.length) handleDtFiles(e.target.files); e.target.value = ""; }} />
          </div>
          {dtUploads.length > 0 && (
            <p className="mt-3 text-sm flex items-center gap-1.5" style={{ color: "#2E7A4E" }}><Icon name="check_circle" size={16} />{dtUploads.length} file{dtUploads.length === 1 ? "" : "s"} uploaded</p>
          )}
          {dtRoster.length > 0 && (
            <p className="mt-2 text-xs flex items-center gap-1.5" style={{ color: T.textMuted }}><Icon name="groups" size={14} />{dtRoster.length} employees remembered in your staff list</p>
          )}
        </div>

        {dtUploads.length > 0 && (
          <div className="flex flex-col gap-4">
            {dtUploads.map((u) => {
              const isActive = u.id === dtActiveId;
              return (
                <div
                  key={u.id}
                  onClick={() => setDtActiveId(u.id)}
                  className="rounded-lg p-4 cursor-pointer transition-colors"
                  style={{ background: T.surface, border: `2px solid ${isActive ? dtTemplate.accent : T.border}` }}
                >
                  <div className="flex items-center justify-between mb-2">
                    <div className="flex items-center gap-1.5 min-w-0">
                      <Icon name="fingerprint" size={16} style={{ color: isActive ? dtTemplate.accent : T.textMuted, flexShrink: 0 }} />
                      <span className="text-sm font-semibold truncate" title={u.fileName}>{u.fileName}</span>
                    </div>
                    <button
                      onClick={(e) => { e.stopPropagation(); removeDtUpload(u.id); }}
                      className="shrink-0 w-6 h-6 rounded-full flex items-center justify-center hover:bg-black/5"
                      style={{ color: T.textMuted }}
                      title="Remove"
                    >
                      <Icon name="close" size={15} />
                    </button>
                  </div>

                  {u.parseError && (
                    <p className="text-xs flex items-start gap-1.5" style={{ color: "#ba1a1a" }}><Icon name="error" size={14} />{u.parseError}</p>
                  )}

                  {u.parsed && (
                    <p className="text-xs" style={{ color: T.textMuted }}>
                      {u.parsed.date ? `Date: ${u.parsed.date} · ` : ""}{u.parsed.employees.length} employees with punches
                    </p>
                  )}
                </div>
              );
            })}

            {dtUploads.filter((u) => u.parsed).length > 1 && (
              <>
                <button
                  onClick={exportAllDtExcel}
                  className="w-full py-2.5 rounded-lg text-sm font-semibold flex items-center justify-center gap-2 transition-colors"
                  style={{ background: dtTemplate.accent, color: "#fff", fontFamily: T.bodyFont }}
                  title="Choose a folder once — every day's file saves there automatically"
                >
                  <Icon name="folder_zip" size={16} /> Export All Days as Excel ({dtUploads.filter((u) => u.parsed).length})
                </button>
                <button
                  onClick={exportAllDtPdf}
                  className="w-full py-2.5 rounded-lg text-sm font-semibold flex items-center justify-center gap-2 transition-colors"
                  style={{ background: T.surface, color: dtTemplate.accent, border: `1.5px solid ${dtTemplate.accent}`, fontFamily: T.bodyFont }}
                  title="Choose a folder once — a PDF for each day saves there automatically"
                >
                  <Icon name="picture_as_pdf" size={16} /> Export All Days as PDF ({dtUploads.filter((u) => u.parsed).length})
                </button>
                <p className="text-xs text-center -mt-2" style={{ color: T.textMuted }}>Excel: pick a destination folder once — every day saves there automatically. PDF: pick a destination folder once — a PDF for each day saves there automatically.</p>
              </>
            )}
          </div>
        )}
      </div>

      <div className="lg:col-span-8">
        {!activeDtUpload || !activeDtUpload.parsed ? (
          <div className="h-full flex flex-col items-center justify-center text-sm py-24 rounded-lg" style={{ border: `2px dashed ${T.border}`, color: T.textMuted }}>
            <Icon name="fingerprint" size={28} style={{ color: T.textMuted }} />
            <p className="mt-2">Upload a transaction report to see the daily punch sheet here</p>
          </div>
        ) : (
          <div className="rounded-lg overflow-hidden" style={{ background: T.surface, border: `1px solid ${T.border}` }}>
            <div className="p-6 flex items-center justify-between flex-wrap gap-3" style={{ background: dtTemplate.accent }}>
              <div>
                <h2 style={{ fontFamily: T.headFont, fontSize: 22, fontWeight: 700, color: "#fff" }}>{dtTemplate.reportTitle}</h2>
                <p className="text-sm opacity-85 mt-1" style={{ color: "#fff" }}>Date: {activeDtUpload.parsed.date || "unknown"}</p>
              </div>
              <div className="flex gap-2">
                <button onClick={() => exportDtExcelFor(activeDtUpload)} disabled={!canExportActive} className="flex items-center gap-1.5 disabled:opacity-40 text-xs font-bold px-3.5 py-2 rounded-lg" style={{ background: "#fff", color: dtTemplate.accent }}>
                  <Icon name="grid_on" size={15} /> Download Excel
                </button>
                <button onClick={() => exportDtPdfFor(activeDtUpload)} disabled={!canExportActive} className="flex items-center gap-1.5 disabled:opacity-40 text-xs font-bold px-3.5 py-2 rounded-lg" style={{ background: "rgba(0,0,0,0.2)", color: "#fff" }}>
                  <Icon name="picture_as_pdf" size={15} /> Download PDF
                </button>
              </div>
            </div>

            <div className="p-5 grid grid-cols-3 gap-4" style={{ borderBottom: `1px solid ${T.border}` }}>
              <div className="rounded-lg p-4" style={{ border: `1px solid ${T.border}` }}>
                <p style={{ fontFamily: T.capsFont, fontSize: 10, letterSpacing: "0.08em", color: T.textMuted }}>STAFF LIST</p>
                <p style={{ fontFamily: T.headFont, fontSize: 22, fontWeight: 700 }} className="mt-1">{dtRoster.length}</p>
              </div>
              <div className="rounded-lg p-4" style={{ border: `1px solid ${T.border}` }}>
                <p style={{ fontFamily: T.capsFont, fontSize: 10, letterSpacing: "0.08em", color: T.textMuted }}>PRESENT TODAY</p>
                <p style={{ fontFamily: T.headFont, fontSize: 22, fontWeight: 700, color: dtTemplate.accent }} className="mt-1">{presentCount}</p>
              </div>
              <div className="rounded-lg p-4" style={{ border: `1px solid ${T.border}` }}>
                <p style={{ fontFamily: T.capsFont, fontSize: 10, letterSpacing: "0.08em", color: T.textMuted }}>ABSENT ({dtTemplate.absentPlaceholder})</p>
                <p style={{ fontFamily: T.headFont, fontSize: 22, fontWeight: 700 }} className="mt-1">{Math.max(0, absentCount)}</p>
              </div>
            </div>

            <div className="overflow-x-auto max-h-[560px] overflow-y-auto">
              <table className="w-full text-sm">
                <thead className="sticky top-0 z-10">
                  <tr>
                    {dtTemplate.columns.map((col) => (
                      <th key={col.id} className="text-left px-4 py-2.5" style={{ background: T.surfaceContainer, borderTop: `2px solid ${dtTemplate.accent}`, fontFamily: T.capsFont, fontSize: 10, letterSpacing: "0.08em", color: T.textMuted }}>{col.label}</th>
                    ))}
                    <th className="text-left px-4 py-2.5" style={{ background: T.surfaceContainer, borderTop: `2px solid ${dtTemplate.accent}`, fontFamily: T.capsFont, fontSize: 10, letterSpacing: "0.08em", color: T.textMuted }}>{dtTemplate.timeLabel}</th>
                    <th className="text-left px-4 py-2.5" style={{ background: T.surfaceContainer, borderTop: `2px solid ${dtTemplate.accent}`, fontFamily: T.capsFont, fontSize: 10, letterSpacing: "0.08em", color: T.textMuted }}>{dtTemplate.stateLabel}</th>
                  </tr>
                </thead>
                <tbody>
                  {(() => {
                    let groupToggle = false;
                    let lastGroupKey = null;
                    return activeDtRows.map((row, idx) => {
                      if (row.groupKey !== lastGroupKey) { groupToggle = !groupToggle; lastGroupKey = row.groupKey; }
                      const banded = dtTemplate.bandingEnabled && groupToggle;
                      return (
                        <tr key={idx} style={{ borderBottom: `1px solid ${T.border}`, background: banded ? dtTemplate.bandingColor : "transparent" }}>
                          {dtTemplate.columns.map((col, ci) => (
                            <td key={col.id} className={`px-4 py-2 ${ci === 1 ? "font-semibold" : ""}`} style={ci === 0 ? { color: T.textMuted } : undefined}>
                              {dtCellValue(col.field, row)}
                            </td>
                          ))}
                          <td className="px-4 py-2 tabular-nums" style={row.absent ? { color: T.textMuted } : undefined}>{row.time}</td>
                          <td className="px-4 py-2" style={row.absent ? { color: T.textMuted } : undefined}>{row.state}</td>
                        </tr>
                      );
                    });
                  })()}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

export default function App() {
  const [screen, setScreen] = useState("dashboard");
  const [templateMode, setTemplateMode] = useState("sales");
  const [template, setTemplate] = useState(DEFAULT_TEMPLATE);
  const [templateLoaded, setTemplateLoaded] = useState(false);
  const [attTemplate, setAttTemplate] = useState(ATTENDANCE_DEFAULT_TEMPLATE);
  const [attTemplateLoaded, setAttTemplateLoaded] = useState(false);
  const [dtTemplate, setDtTemplate] = useState(DAILY_TX_DEFAULT_TEMPLATE);
  const [dtTemplateLoaded, setDtTemplateLoaded] = useState(false);

  const [uploads, setUploads] = useState([]);
  const [activeId, setActiveId] = useState(null);
  const [page, setPage] = useState(0);

  const [attUploads, setAttUploads] = useState([]);
  const [attActiveId, setAttActiveId] = useState(null);
  const [attWeekIndex, setAttWeekIndex] = useState(0);

  const fileInputRef = useRef(null);
  const uploadCounter = useRef(0);
  const attFileInputRef = useRef(null);
  const attUploadCounter = useRef(0);

  useEffect(() => {
    try {
      const saved = localStorage.getItem("attendance-template-config");
      if (saved) {
        const parsed = JSON.parse(saved);
        if (!Array.isArray(parsed.columns) || parsed.columns.length === 0) delete parsed.columns;
        setAttTemplate({ ...ATTENDANCE_DEFAULT_TEMPLATE, ...parsed });
      }
    } catch (e) {}
    finally { setAttTemplateLoaded(true); }
  }, []);

  useEffect(() => {
    if (!attTemplateLoaded) return;
    try { localStorage.setItem("attendance-template-config", JSON.stringify(attTemplate)); } catch (e) {}
  }, [attTemplate, attTemplateLoaded]);

  useEffect(() => {
    try {
      const saved = localStorage.getItem("daily-tx-template-config");
      if (saved) {
        const parsed = JSON.parse(saved);
        if (!Array.isArray(parsed.columns) || parsed.columns.length === 0) delete parsed.columns;
        setDtTemplate({ ...DAILY_TX_DEFAULT_TEMPLATE, ...parsed });
      }
    } catch (e) {}
    finally { setDtTemplateLoaded(true); }
  }, []);

  useEffect(() => {
    if (!dtTemplateLoaded) return;
    try { localStorage.setItem("daily-tx-template-config", JSON.stringify(dtTemplate)); } catch (e) {}
  }, [dtTemplate, dtTemplateLoaded]);

  const activeAttUpload = attUploads.find((u) => u.id === attActiveId) || null;
  const activeAttWeekDates = activeAttUpload && activeAttUpload.parsed ? (activeAttUpload.parsed.weeks[attWeekIndex] || []) : [];
  const activeAttRows = activeAttUpload && activeAttUpload.parsed ? buildAttendanceRows(activeAttUpload.parsed, activeAttWeekDates) : [];

  useEffect(() => { setAttWeekIndex(0); }, [attActiveId]);

  function updateAttUpload(id, patch) {
    setAttUploads((prev) => prev.map((u) => (u.id === id ? { ...u, ...patch } : u)));
  }

  function removeAttUpload(id) {
    setAttUploads((prev) => prev.filter((u) => u.id !== id));
    if (attActiveId === id) {
      setAttActiveId((prevActive) => {
        const remaining = attUploads.filter((u) => u.id !== id);
        return remaining.length ? remaining[0].id : null;
      });
    }
  }

  function handleAttFiles(fileList) {
    const files = Array.from(fileList || []);
    files.forEach((file) => {
      attUploadCounter.current += 1;
      const id = `att-upload-${Date.now()}-${attUploadCounter.current}`;
      const draft = { id, fileName: file.name, parsed: null, parseError: null };
      setAttUploads((prev) => [...prev, draft]);
      setAttActiveId((prev) => prev || id);

      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const wb = XLSX.read(e.target.result, { type: "array" });
          const parsed = parseAttendanceWorkbook(wb);
          updateAttUpload(id, { parsed, parseError: null });
        } catch (err) {
          updateAttUpload(id, { parsed: null, parseError: err.message || "Couldn't read this file." });
        }
      };
      reader.onerror = () => updateAttUpload(id, { parsed: null, parseError: "The file could not be read." });
      reader.readAsArrayBuffer(file);
    });
  }

  const [dtUploads, setDtUploads] = useState([]);
  const [dtActiveId, setDtActiveId] = useState(null);
  const [dtRoster, setDtRoster] = useState([]);
  const [dtRosterLoaded, setDtRosterLoaded] = useState(false);
  const dtFileInputRef = useRef(null);
  const dtUploadCounter = useRef(0);

  useEffect(() => {
    try {
      const saved = localStorage.getItem("daily-tx-roster");
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed)) setDtRoster(parsed);
      }
    } catch (e) {}
    finally { setDtRosterLoaded(true); }
  }, []);

  useEffect(() => {
    if (!dtRosterLoaded) return;
    try { localStorage.setItem("daily-tx-roster", JSON.stringify(dtRoster)); } catch (e) {}
  }, [dtRoster, dtRosterLoaded]);

  const activeDtUpload = dtUploads.find((u) => u.id === dtActiveId) || null;
  const activeDtRows = activeDtUpload && activeDtUpload.parsed ? buildDailyTxRows(dtRoster, activeDtUpload.parsed, dtTemplate.absentPlaceholder) : [];

  function updateDtUpload(id, patch) {
    setDtUploads((prev) => prev.map((u) => (u.id === id ? { ...u, ...patch } : u)));
  }

  function removeDtUpload(id) {
    setDtUploads((prev) => prev.filter((u) => u.id !== id));
    if (dtActiveId === id) {
      setDtActiveId((prevActive) => {
        const remaining = dtUploads.filter((u) => u.id !== id);
        return remaining.length ? remaining[0].id : null;
      });
    }
  }

  function handleDtFiles(fileList) {
    const files = Array.from(fileList || []);
    files.forEach((file) => {
      dtUploadCounter.current += 1;
      const id = `dt-upload-${Date.now()}-${dtUploadCounter.current}`;
      const draft = { id, fileName: file.name, parsed: null, parseError: null };
      setDtUploads((prev) => [...prev, draft]);
      setDtActiveId((prev) => prev || id);

      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const wb = XLSX.read(e.target.result, { type: "array" });
          const parsed = parseDailyTransactionWorkbook(wb);
          updateDtUpload(id, { parsed, parseError: null });
          setDtRoster((prev) => mergeRoster(prev, parsed.employees));
        } catch (err) {
          updateDtUpload(id, { parsed: null, parseError: err.message || "Couldn't read this file." });
        }
      };
      reader.onerror = () => updateDtUpload(id, { parsed: null, parseError: "The file could not be read." });
      reader.readAsArrayBuffer(file);
    });
  }

  useEffect(() => {
    try {
      const saved = localStorage.getItem("template-config");
      if (saved) {
        const parsed = JSON.parse(saved);
        if (!Array.isArray(parsed.columns) || parsed.columns.length === 0) delete parsed.columns;
        setTemplate({ ...DEFAULT_TEMPLATE, ...parsed });
      }
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

  async function pickDirectory() {
    if (!window.showDirectoryPicker) return null;
    try {
      return await window.showDirectoryPicker({ mode: "readwrite" });
    } catch (err) {
      return null; // user cancelled, or not supported/permitted
    }
  }

  async function saveBlob(blob, suggestedName, { pickLocation = true, dirHandle = null } = {}) {
    if (dirHandle) {
      try {
        const fileHandle = await dirHandle.getFileHandle(suggestedName, { create: true });
        const writable = await fileHandle.createWritable();
        await writable.write(blob);
        await writable.close();
        return true;
      } catch (err) {
        // fall through to normal download if writing into the folder fails
      }
    }

    if (pickLocation && window.showSaveFilePicker) {
      try {
        const handle = await window.showSaveFilePicker({
          suggestedName,
          types: [{ description: "Excel Workbook", accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] } }],
        });
        const writable = await handle.createWritable();
        await writable.write(blob);
        await writable.close();
        return true;
      } catch (err) {
        if (err && err.name === "AbortError") return false; // user cancelled the picker
        // any other error: fall through to the default-download fallback below
      }
    }
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = suggestedName;
    a.click();
    URL.revokeObjectURL(url);
    return true;
  }

  async function exportExcelFor(upload, rep, exportAllMode = false, dirHandle = null) {
    if (!rep) return;
    const cols = template.columns;
    const colCount = cols.length;
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Report");
    ws.columns = cols.map((c) => ({ width: c.field === "product" ? 44 : c.field === "outlet" ? 30 : c.field === "no" ? 6 : 20 }));

    const accentARGB = "FF" + template.accent.replace("#", "").toUpperCase();
    const thin = { style: "thin", color: { argb: "FFCBBFAE" } };
    const borderAll = { top: thin, left: thin, bottom: thin, right: thin };

    ws.mergeCells(1, 1, 1, colCount);
    const titleCell = ws.getCell(1, 1);
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
      const labelCell = ws.getCell(rowNum, 1);
      labelCell.value = label;
      labelCell.font = { bold: true };
      if (colCount > 1) ws.mergeCells(rowNum, 2, rowNum, colCount);
      ws.getCell(rowNum, 2).value = value;
    });

    const headerRowNum = 7;
    cols.forEach((col, i) => {
      const cell = ws.getCell(headerRowNum, i + 1);
      cell.value = col.label;
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: accentARGB } };
      cell.alignment = { vertical: "middle", horizontal: col.field === "qty" ? "right" : "left" };
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
        cols.forEach((col, ci) => {
          const isMerge = SALES_MERGE_FIELDS.has(col.field);
          const cell = row.getCell(ci + 1);
          cell.value = (!isMerge || i === 0) ? salesCellValue(col.field, { outletIdx: idx, item, isFirst: true, outlet }) : null;
          cell.border = borderAll;
          cell.alignment = { horizontal: col.field === "qty" ? "right" : "left", vertical: "middle" };
          if (banded) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: bandARGB } };
        });
      });
      const endRow = r;
      if (outlet.items.length > 1) {
        cols.forEach((col, ci) => {
          if (SALES_MERGE_FIELDS.has(col.field)) {
            ws.mergeCells(startRow, ci + 1, endRow, ci + 1);
            ws.getCell(startRow, ci + 1).alignment = { vertical: "middle", horizontal: col.field === "no" ? "center" : "left" };
          }
        });
      }
    });

    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: "application/octet-stream" });
    const safeRoute = (upload.routeName || upload.territory).replace(/[^a-z0-9]+/gi, "_").slice(0, 40);
    const fileName = `${upload.territory}_${safeRoute}_LastVisit.xlsx`;
    return saveBlob(blob, fileName, { pickLocation: exportAllMode !== true, dirHandle });
  }

  function exportPdfFor(upload, rep, exportAllMode = false, dirHandle = null) {
    if (!rep) return;
    const cols = template.columns;
    const [ar, ag, ab] = hexToRgb(template.accent);
    const doc = createUnicodePdf({ unit: "pt", format: "a4" });
    const pageWidth = doc.internal.pageSize.getWidth();
    const margin = 40;

    doc.setFillColor(ar, ag, ab);
    doc.rect(0, 0, pageWidth, 56, "F");
    doc.setTextColor(255, 255, 255);
    doc.setFont("NotoSansSinhala", "bold");
    doc.setFontSize(15);
    doc.text(template.reportTitle, pageWidth / 2, 34, { align: "center" });

    doc.setTextColor(30, 28, 25);
    doc.setFontSize(10);
    let y = 78;
    const metaRows = [
      [template.territoryLabel, upload.territory],
      [template.routeLabel, upload.routeName],
      [template.totalLabel, rep.totalSale.toLocaleString(undefined, { minimumFractionDigits: 2 })],
      [template.dateLabel, upload.lastVisitDate || "—"],
    ];
    metaRows.forEach(([label, value]) => {
      doc.setFont("NotoSansSinhala", "bold");
      doc.text(String(label), margin, y);
      doc.setFont("NotoSansSinhala", "normal");
      doc.text(String(value), margin + 130, y);
      y += 16;
    });

    const head = [cols.map((c) => c.label)];
    const body = [];
    const groupStartRows = [];
    rep.outlets.forEach((outlet, idx) => {
      groupStartRows.push(body.length);
      outlet.items.forEach((item, i) => {
        const row = cols.map((col) => {
          const isMerge = SALES_MERGE_FIELDS.has(col.field);
          return (!isMerge || i === 0) ? String(salesCellValue(col.field, { outletIdx: idx, item, isFirst: true, outlet })) : "";
        });
        body.push(row);
      });
    });

    const bandRgb = template.bandingEnabled ? hexToRgb(template.bandingColor) : null;
    const qtyColIdx = cols.findIndex((c) => c.field === "qty");

    autoTable(doc, {
      startY: y + 8,
      head,
      body,
      margin: { left: margin, right: margin },
      styles: { font: "NotoSansSinhala", fontSize: 9, cellPadding: 6, lineColor: [226, 217, 211], lineWidth: exportAllMode ? 0 : 0.5 },
      headStyles: { fillColor: [ar, ag, ab], textColor: [255, 255, 255], fontStyle: "bold" },
      columnStyles: qtyColIdx >= 0 ? { [qtyColIdx]: { halign: "right" } } : {},
      didParseCell: (data) => {
        if (data.section !== "body") return;
        const outletIdx = rep.outlets.findIndex((o, oi) => {
          const start = groupStartRows[oi];
          const end = start + o.items.length;
          return data.row.index >= start && data.row.index < end;
        });
        if (bandRgb && outletIdx % 2 === 1) data.cell.styles.fillColor = bandRgb;
        if (exportAllMode && groupStartRows.includes(data.row.index) && data.row.index > 0) {
          data.cell.styles.lineWidth = { top: 0.75, right: 0, bottom: 0, left: 0 };
          data.cell.styles.lineColor = [226, 217, 211];
        }
      },
    });

    const blob = doc.output("blob");
    const safeRoute = (upload.routeName || upload.territory).replace(/[^a-z0-9]+/gi, "_").slice(0, 40);
    const fileName = `${upload.territory}_${safeRoute}_LastVisit.pdf`;
    return saveBlob(blob, fileName, { pickLocation: exportAllMode !== true, dirHandle });
  }

  function exportExcel() { if (activeUpload && report) exportExcelFor(activeUpload, report); }
  function exportPdf() { if (activeUpload && report) exportPdfFor(activeUpload, report); }

  async function exportAllExcel() {
    const exportable = uploads.filter((u) => u.rows && u.territory && buildReport(u.rows, u.territory, u.routeName, u.strategy).outlets.length > 0);
    if (exportable.length === 0) return;

    const dirHandle = exportable.length > 1 ? await pickDirectory() : null;

    for (const upload of exportable) {
      const rep = buildReport(upload.rows, upload.territory, upload.routeName, upload.strategy);
      await exportExcelFor(upload, rep, true, dirHandle);
    }
  }

  async function exportAllPdf() {
    const jobs = [];
    uploads.forEach((upload) => {
      if (!upload.rows) return;
      const combos = getTerritoryRouteCombos(upload.rows);
      combos.forEach((combo) => {
        const rep = buildReport(upload.rows, combo.territory, combo.routeName, upload.strategy);
        if (rep.outlets.length === 0) return;
        jobs.push({ upload: { ...upload, territory: combo.territory, routeName: combo.routeName }, rep });
      });
    });
    if (jobs.length === 0) return;

    const dirHandle = jobs.length > 1 ? await pickDirectory() : null;
    for (const job of jobs) {
      await exportPdfFor(job.upload, job.rep, true, dirHandle);
    }
  }

  async function exportAttendanceExcelFor(upload, weekDates, weekIdx, pickLocation = true, dirHandle = null) {
    const rowsData = buildAttendanceRows(upload.parsed, weekDates);
    const leadCols = attTemplate.columns;
    const leadCount = leadCols.length;
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Attendance");
    ws.columns = [...leadCols.map(() => ({ width: 16 })), ...weekDates.flatMap(() => [{ width: 11 }, { width: 11 }])];

    const accentARGB = "FF" + attTemplate.accent.replace("#", "").toUpperCase();
    const thin = { style: "thin", color: { argb: "FFCBBFAE" } };
    const borderAll = { top: thin, left: thin, bottom: thin, right: thin };

    ws.mergeCells(1, 1, 1, leadCount);
    const titleCell = ws.getCell(1, 1);
    titleCell.value = attTemplate.reportTitle;
    titleCell.font = { bold: true, size: 13, color: { argb: "FFFFFFFF" } };
    titleCell.alignment = { horizontal: "center", vertical: "middle" };
    titleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: accentARGB } };
    ws.getRow(1).height = 22;

    weekDates.forEach((d, i) => {
      const startCol = leadCount + 1 + i * 2;
      ws.mergeCells(1, startCol, 1, startCol + 1);
      const cell = ws.getCell(1, startCol);
      cell.value = formatAttDate(d);
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: accentARGB } };
    });

    const headerLabels = [...leadCols.map((c) => c.label), ...weekDates.flatMap(() => [attTemplate.checkInLabel, attTemplate.checkOutLabel])];
    headerLabels.forEach((label, i) => {
      const cell = ws.getCell(2, i + 1);
      cell.value = label;
      cell.font = { bold: true };
      cell.border = borderAll;
      cell.alignment = { horizontal: "center", vertical: "middle" };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF0EDE9" } };
    });

    const bandARGB = "FF" + attTemplate.bandingColor.replace("#", "").toUpperCase();
    rowsData.forEach((emp, idx) => {
      const rowNum = idx + 3;
      const banded = attTemplate.bandingEnabled && idx % 2 === 1;
      const values = [...leadCols.map((c) => attCellValue(c.field, emp)), ...emp.cells.flatMap((c) => [c.checkIn, c.checkOut])];
      values.forEach((v, i) => {
        const cell = ws.getCell(rowNum, i + 1);
        cell.value = v;
        cell.border = borderAll;
        cell.alignment = { horizontal: i < leadCount ? "left" : "center", vertical: "middle" };
        if (banded) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: bandARGB } };
      });
    });

    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: "application/octet-stream" });
    const label = weekRangeLabel(weekDates).replace(/[^a-z0-9]+/gi, "_");
    const fileName = `Attendance_Week${weekIdx + 1}_${label}.xlsx`;
    return saveBlob(blob, fileName, { pickLocation, dirHandle });
  }

  function exportAttendancePdfFor(upload, weekDates, exportAllMode = false, dirHandle = null) {
    const rowsData = buildAttendanceRows(upload.parsed, weekDates);
    const leadCols = attTemplate.columns;
    const [ar, ag, ab] = hexToRgb(attTemplate.accent);
    const doc = createUnicodePdf({ unit: "pt", format: "a4", orientation: "landscape" });
    const pageWidth = doc.internal.pageSize.getWidth();
    const margin = 30;

    doc.setFillColor(ar, ag, ab);
    doc.rect(0, 0, pageWidth, 46, "F");
    doc.setTextColor(255, 255, 255);
    doc.setFont("NotoSansSinhala", "bold");
    doc.setFontSize(13);
    doc.text(`${attTemplate.reportTitle} — ${weekRangeLabel(weekDates)}`, pageWidth / 2, 29, { align: "center" });

    const head1 = [
      ...leadCols.map((c) => ({ content: c.label, rowSpan: 2, styles: { valign: "middle" } })),
      ...weekDates.map((d) => ({ content: formatAttDate(d), colSpan: 2, styles: { halign: "center" } })),
    ];
    const head2 = weekDates.flatMap(() => [attTemplate.checkInLabel, attTemplate.checkOutLabel]);

    const bandRgb = attTemplate.bandingEnabled ? hexToRgb(attTemplate.bandingColor) : null;
    const body = rowsData.map((emp) => [
      ...leadCols.map((c) => String(attCellValue(c.field, emp))),
      ...emp.cells.flatMap((c) => [c.checkIn || "—", c.checkOut || "—"]),
    ]);

    autoTable(doc, {
      startY: 60,
      head: [head1, head2],
      body,
      margin: { left: margin, right: margin },
      styles: { font: "NotoSansSinhala", fontSize: 8, cellPadding: 5, halign: "center", lineColor: [226, 217, 211], lineWidth: 0.5 },
      headStyles: { fillColor: [ar, ag, ab], textColor: [255, 255, 255], fontStyle: "bold" },
      columnStyles: leadCols.reduce((acc, _, i) => ({ ...acc, [i]: { halign: "left" } }), {}),
      didParseCell: (data) => {
        if (data.section === "body" && bandRgb && data.row.index % 2 === 1) {
          data.cell.styles.fillColor = bandRgb;
        }
      },
    });

    const blob = doc.output("blob");
    const label = weekRangeLabel(weekDates).replace(/[^a-z0-9]+/gi, "_");
    const fileName = `Attendance_Week_${label}.pdf`;
    return saveBlob(blob, fileName, { pickLocation: exportAllMode !== true, dirHandle });
  }

  async function exportAllAttendanceWeeksExcel(upload) {
    if (!upload.parsed || upload.parsed.weeks.length === 0) return;
    const dirHandle = upload.parsed.weeks.length > 1 ? await pickDirectory() : null;
    for (let i = 0; i < upload.parsed.weeks.length; i++) {
      await exportAttendanceExcelFor(upload, upload.parsed.weeks[i], i, false, dirHandle);
    }
  }

  async function exportAllAttendanceWeeksPdf(upload) {
    if (!upload.parsed || upload.parsed.weeks.length === 0) return;
    const dirHandle = upload.parsed.weeks.length > 1 ? await pickDirectory() : null;
    for (const weekDates of upload.parsed.weeks) {
      await exportAttendancePdfFor(upload, weekDates, true, dirHandle);
    }
  }

  async function exportDtExcelFor(upload, exportAllMode = false, dirHandle = null) {
    if (!upload.parsed) return;
    const rows = buildDailyTxRows(dtRoster, upload.parsed, dtTemplate.absentPlaceholder);
    const leadCols = dtTemplate.columns;
    const leadCount = leadCols.length;
    const totalCols = leadCount + 2;
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Transaction");
    ws.columns = [...leadCols.map(() => ({ width: 16 })), { width: 12 }, { width: 14 }];

    const accentARGB = "FF" + dtTemplate.accent.replace("#", "").toUpperCase();
    const thin = { style: "thin", color: { argb: "FFCBBFAE" } };
    const borderAll = { top: thin, left: thin, bottom: thin, right: thin };

    ws.mergeCells(1, 1, 1, totalCols);
    const titleCell = ws.getCell(1, 1);
    titleCell.value = dtTemplate.reportTitle;
    titleCell.font = { bold: true, size: 14, color: { argb: "FFFFFFFF" } };
    titleCell.alignment = { horizontal: "center", vertical: "middle" };
    titleCell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: accentARGB } };
    ws.getRow(1).height = 24;

    ws.mergeCells(2, 1, 2, totalCols);
    const dateCell = ws.getCell(2, 1);
    dateCell.value = `Date: ${upload.parsed.date || ""}`;
    dateCell.font = { bold: true };
    dateCell.alignment = { horizontal: "left", vertical: "middle" };

    const headerLabels = [...leadCols.map((c) => c.label), dtTemplate.timeLabel, dtTemplate.stateLabel];
    headerLabels.forEach((label, i) => {
      const cell = ws.getCell(3, i + 1);
      cell.value = label;
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: accentARGB } };
      cell.alignment = { vertical: "middle", horizontal: "left" };
      cell.border = borderAll;
    });

    const bandARGB = "FF" + dtTemplate.bandingColor.replace("#", "").toUpperCase();
    let groupToggle = false;
    let lastGroupKey = null;
    rows.forEach((row, idx) => {
      const rowNum = idx + 4;
      if (row.groupKey !== lastGroupKey) { groupToggle = !groupToggle; lastGroupKey = row.groupKey; }
      const banded = dtTemplate.bandingEnabled && groupToggle;
      const values = [...leadCols.map((c) => dtCellValue(c.field, row)), row.time, row.state];
      values.forEach((v, i) => {
        const cell = ws.getCell(rowNum, i + 1);
        cell.value = v;
        cell.border = borderAll;
        cell.alignment = { horizontal: "left", vertical: "middle" };
        if (banded) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: bandARGB } };
      });
    });

    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: "application/octet-stream" });
    const safeDate = (upload.parsed.date || upload.fileName).replace(/[^a-z0-9]+/gi, "_").slice(0, 40);
    const fileName = `Transaction_${safeDate}.xlsx`;
    return saveBlob(blob, fileName, { pickLocation: exportAllMode !== true, dirHandle });
  }

  function exportDtPdfFor(upload, exportAllMode = false, dirHandle = null) {
    if (!upload.parsed) return;
    const rows = buildDailyTxRows(dtRoster, upload.parsed, dtTemplate.absentPlaceholder);
    const leadCols = dtTemplate.columns;
    const [ar, ag, ab] = hexToRgb(dtTemplate.accent);
    const doc = createUnicodePdf({ unit: "pt", format: "a4" });
    const pageWidth = doc.internal.pageSize.getWidth();
    const margin = 40;

    doc.setFillColor(ar, ag, ab);
    doc.rect(0, 0, pageWidth, 50, "F");
    doc.setTextColor(255, 255, 255);
    doc.setFont("NotoSansSinhala", "bold");
    doc.setFontSize(14);
    doc.text(dtTemplate.reportTitle, pageWidth / 2, 32, { align: "center" });

    doc.setTextColor(30, 28, 25);
    doc.setFont("NotoSansSinhala", "bold");
    doc.setFontSize(10);
    doc.text(`Date: ${upload.parsed.date || ""}`, margin, 68);

    const head = [[...leadCols.map((c) => c.label), dtTemplate.timeLabel, dtTemplate.stateLabel]];
    const bandRgb = dtTemplate.bandingEnabled ? hexToRgb(dtTemplate.bandingColor) : null;
    const groupKeys = rows.map((r) => r.groupKey);
    const body = rows.map((row) => [...leadCols.map((c) => String(dtCellValue(c.field, row))), row.time, row.state]);

    autoTable(doc, {
      startY: 82,
      head,
      body,
      margin: { left: margin, right: margin },
      styles: { font: "NotoSansSinhala", fontSize: 9, cellPadding: 6, lineColor: [204, 204, 204], lineWidth: 0.5 },
      headStyles: { fillColor: [ar, ag, ab], textColor: [255, 255, 255], fontStyle: "bold" },
      didParseCell: (data) => {
        if (data.section !== "body" || !bandRgb) return;
        let groupToggle = false;
        let lastKey = null;
        for (let i = 0; i <= data.row.index; i++) {
          if (groupKeys[i] !== lastKey) { groupToggle = !groupToggle; lastKey = groupKeys[i]; }
        }
        if (groupToggle) data.cell.styles.fillColor = bandRgb;
      },
    });

    const blob = doc.output("blob");
    const safeDate = (upload.parsed.date || upload.fileName).replace(/[^a-z0-9]+/gi, "_").slice(0, 40);
    const fileName = `Transaction_${safeDate}.pdf`;
    return saveBlob(blob, fileName, { pickLocation: exportAllMode !== true, dirHandle });
  }

  async function exportAllDtExcel() {
    const exportable = dtUploads.filter((u) => u.parsed);
    if (exportable.length === 0) return;
    const dirHandle = exportable.length > 1 ? await pickDirectory() : null;
    for (const upload of exportable) {
      await exportDtExcelFor(upload, true, dirHandle);
    }
  }

  async function exportAllDtPdf() {
    const exportable = dtUploads.filter((u) => u.parsed);
    if (exportable.length === 0) return;
    const dirHandle = exportable.length > 1 ? await pickDirectory() : null;
    for (const upload of exportable) {
      await exportDtPdfFor(upload, true, dirHandle);
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
  const pdfCombosCount = uploads.reduce((sum, u) => {
    if (!u.rows) return sum;
    const combos = getTerritoryRouteCombos(u.rows);
    return sum + combos.filter((c) => buildReport(u.rows, c.territory, c.routeName, u.strategy).outlets.length > 0).length;
  }, 0);

  return (
    <div className="min-h-screen" style={{ background: T.bg, fontFamily: T.bodyFont, color: T.text }}>
      <header style={{ background: T.surface, borderBottom: `1px solid ${T.border}`, boxShadow: "0 1px 2px rgba(0,0,0,0.03)" }}>
        <div className="max-w-[1280px] mx-auto px-6 h-16 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <img src="/logo.png" alt="Lakmee" className="h-9 w-auto object-contain" />
            <h1 style={{ fontFamily: T.headFont, fontSize: 18, fontWeight: 700, color: template.accent }}>Route Sales Report Builder</h1>
          </div>
          <nav className="rounded-full p-1 flex items-center gap-1" style={{ background: T.surfaceLow, boxShadow: "inset 0 1px 2px rgba(0,0,0,0.04)" }}>
            <button
              onClick={() => setScreen("dashboard")}
              className="px-4 py-1.5 rounded-full text-sm font-semibold transition-colors"
              style={{ fontFamily: T.capsFont, fontSize: 12, letterSpacing: "0.06em", background: screen === "dashboard" ? template.accent : "transparent", color: screen === "dashboard" ? "#fff" : T.textMuted }}
            >
              SALES REPORT
            </button>
            <button
              onClick={() => setScreen("attendance")}
              className="px-4 py-1.5 rounded-full text-sm font-semibold transition-colors"
              style={{ fontFamily: T.capsFont, fontSize: 12, letterSpacing: "0.06em", background: screen === "attendance" ? attTemplate.accent : "transparent", color: screen === "attendance" ? "#fff" : T.textMuted }}
            >
              ATTENDANCE
            </button>
            <button
              onClick={() => setScreen("dailytx")}
              className="px-4 py-1.5 rounded-full text-sm font-semibold transition-colors"
              style={{ fontFamily: T.capsFont, fontSize: 12, letterSpacing: "0.06em", background: screen === "dailytx" ? dtTemplate.accent : "transparent", color: screen === "dailytx" ? "#fff" : T.textMuted }}
            >
              DAILY TRANSACTION
            </button>
            <button
              onClick={() => setScreen("template")}
              className="px-4 py-1.5 rounded-full text-sm font-semibold transition-colors"
              style={{ fontFamily: T.capsFont, fontSize: 12, letterSpacing: "0.06em", background: screen === "template" ? (templateMode === "attendance" ? attTemplate.accent : templateMode === "dailytx" ? dtTemplate.accent : template.accent) : "transparent", color: screen === "template" ? "#fff" : T.textMuted }}
            >
              TEMPLATE EDITOR
            </button>
          </nav>
          <div className="flex items-center gap-4">
            {screen === "attendance" ? (
              <>
                <button onClick={() => activeAttUpload && exportAttendanceExcelFor(activeAttUpload, activeAttWeekDates, attWeekIndex)} disabled={!activeAttUpload || !activeAttUpload.parsed} title="Download this week's Excel" style={{ color: attTemplate.accent, opacity: activeAttUpload && activeAttUpload.parsed ? 1 : 0.35, cursor: activeAttUpload && activeAttUpload.parsed ? "pointer" : "default" }}>
                  <Icon name="download" />
                </button>
                <button onClick={() => activeAttUpload && exportAttendancePdfFor(activeAttUpload, activeAttWeekDates)} disabled={!activeAttUpload || !activeAttUpload.parsed} title="Download this week's PDF" style={{ color: attTemplate.accent, opacity: activeAttUpload && activeAttUpload.parsed ? 1 : 0.35, cursor: activeAttUpload && activeAttUpload.parsed ? "pointer" : "default" }}>
                  <Icon name="picture_as_pdf" />
                </button>
              </>
            ) : screen === "dailytx" ? (
              <>
                <button onClick={() => activeDtUpload && exportDtExcelFor(activeDtUpload)} disabled={!activeDtUpload || !activeDtUpload.parsed} title="Download today's Excel" style={{ color: dtTemplate.accent, opacity: activeDtUpload && activeDtUpload.parsed ? 1 : 0.35, cursor: activeDtUpload && activeDtUpload.parsed ? "pointer" : "default" }}>
                  <Icon name="download" />
                </button>
                <button onClick={() => activeDtUpload && exportDtPdfFor(activeDtUpload)} disabled={!activeDtUpload || !activeDtUpload.parsed} title="Download today's PDF" style={{ color: dtTemplate.accent, opacity: activeDtUpload && activeDtUpload.parsed ? 1 : 0.35, cursor: activeDtUpload && activeDtUpload.parsed ? "pointer" : "default" }}>
                  <Icon name="picture_as_pdf" />
                </button>
              </>
            ) : (
              <>
                <button onClick={exportExcel} disabled={!canExport} title="Download Excel" style={{ color: template.accent, opacity: canExport ? 1 : 0.35, cursor: canExport ? "pointer" : "default" }}>
                  <Icon name="download" />
                </button>
                <button onClick={exportPdf} disabled={!canExport} title="Download PDF" style={{ color: template.accent, opacity: canExport ? 1 : 0.35, cursor: canExport ? "pointer" : "default" }}>
                  <Icon name="picture_as_pdf" />
                </button>
              </>
            )}
            <div className="w-8 h-8 rounded-full flex items-center justify-center text-white text-xs font-bold" style={{ background: screen === "attendance" ? attTemplate.accent : screen === "dailytx" ? dtTemplate.accent : template.accent, border: `1px solid ${T.border}` }}>RT</div>
          </div>
        </div>
      </header>

      <main className="max-w-[1280px] mx-auto px-6 py-8">
        {screen === "template" ? (
          <TemplateEditor
            template={template}
            setTemplate={setTemplate}
            onReset={() => setTemplate(DEFAULT_TEMPLATE)}
            previewData={previewData}
            attTemplate={attTemplate}
            setAttTemplate={setAttTemplate}
            onResetAtt={() => setAttTemplate(ATTENDANCE_DEFAULT_TEMPLATE)}
            dtTemplate={dtTemplate}
            setDtTemplate={setDtTemplate}
            onResetDt={() => setDtTemplate(DAILY_TX_DEFAULT_TEMPLATE)}
            templateMode={templateMode}
            setTemplateMode={setTemplateMode}
          />
        ) : screen === "attendance" ? (
          <AttendanceScreen
            attTemplate={attTemplate}
            attUploads={attUploads}
            attActiveId={attActiveId}
            setAttActiveId={setAttActiveId}
            activeAttUpload={activeAttUpload}
            attWeekIndex={attWeekIndex}
            setAttWeekIndex={setAttWeekIndex}
            activeAttWeekDates={activeAttWeekDates}
            activeAttRows={activeAttRows}
            attFileInputRef={attFileInputRef}
            handleAttFiles={handleAttFiles}
            removeAttUpload={removeAttUpload}
            exportAttendanceExcelFor={exportAttendanceExcelFor}
            exportAttendancePdfFor={exportAttendancePdfFor}
            exportAllAttendanceWeeksExcel={exportAllAttendanceWeeksExcel}
            exportAllAttendanceWeeksPdf={exportAllAttendanceWeeksPdf}
          />
        ) : screen === "dailytx" ? (
          <DailyTransactionScreen
            dtTemplate={dtTemplate}
            dtUploads={dtUploads}
            dtActiveId={dtActiveId}
            setDtActiveId={setDtActiveId}
            activeDtUpload={activeDtUpload}
            activeDtRows={activeDtRows}
            dtRoster={dtRoster}
            dtFileInputRef={dtFileInputRef}
            handleDtFiles={handleDtFiles}
            removeDtUpload={removeDtUpload}
            exportDtExcelFor={exportDtExcelFor}
            exportDtPdfFor={exportDtPdfFor}
            exportAllDtExcel={exportAllDtExcel}
            exportAllDtPdf={exportAllDtPdf}
          />
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
                      title="Choose a folder once — all files save there automatically"
                      className="w-full py-2.5 rounded-lg text-sm font-semibold flex items-center justify-center gap-2 transition-colors disabled:opacity-40"
                      style={{ background: template.accent, color: "#fff", fontFamily: T.bodyFont }}
                    >
                      <Icon name="folder_zip" size={16} /> Export All as Excel ({exportableCount})
                    </button>
                  )}
                  {uploads.length > 1 && (
                    <button
                      onClick={exportAllPdf}
                      disabled={pdfCombosCount === 0}
                      title="Choose a folder once — a PDF for each territory/route saves there automatically"
                      className="w-full py-2.5 rounded-lg text-sm font-semibold flex items-center justify-center gap-2 transition-colors disabled:opacity-40"
                      style={{ background: T.surface, color: template.accent, border: `1.5px solid ${template.accent}`, fontFamily: T.bodyFont }}
                    >
                      <Icon name="picture_as_pdf" size={16} /> Export All as PDF ({pdfCombosCount})
                    </button>
                  )}
                  {uploads.length > 1 && (
                    <p className="text-xs text-center -mt-2" style={{ color: T.textMuted }}>Excel: pick a destination folder once — every file saves there automatically. PDF: pick a destination folder once — a clean PDF for each territory/route saves there automatically.</p>
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
                            {template.columns.map((col, ci) => (
                              <th
                                key={col.id}
                                className={`px-3 py-2.5 ${ci === 0 ? "pl-5 w-12" : ""} ${col.field === "qty" ? "text-right pr-5" : "text-left"}`}
                                style={{ fontFamily: T.capsFont, fontSize: 10, letterSpacing: "0.08em", color: T.textMuted }}
                              >
                                {col.label}
                              </th>
                            ))}
                          </tr>
                        </thead>
                        <tbody>
                          {pagedOutlets.map((outlet, idx) => {
                            const absoluteIdx = page * PAGE_SIZE + idx;
                            const banded = template.bandingEnabled && absoluteIdx % 2 === 1;
                            return outlet.items.map((item, i) => (
                              <tr key={`${idx}-${i}`} style={{ borderBottom: `1px solid ${T.border}`, background: banded ? template.bandingColor : "transparent" }}>
                                {template.columns.map((col, ci) => {
                                  const isMerge = SALES_MERGE_FIELDS.has(col.field);
                                  const value = salesCellValue(col.field, { outletIdx: absoluteIdx, item, isFirst: i === 0, outlet });
                                  return (
                                    <td
                                      key={col.id}
                                      className={`px-3 py-2 ${ci === 0 ? "pl-5" : ""} ${col.field === "qty" ? "text-right pr-5 tabular-nums" : ""} ${isMerge ? "align-top" : ""} ${col.field === "outlet" ? "font-semibold" : ""}`}
                                      style={col.field === "no" ? { color: T.textMuted } : undefined}
                                    >
                                      {value}
                                    </td>
                                  );
                                })}
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
