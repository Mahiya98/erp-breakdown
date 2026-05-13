// ====== CONFIG ======
const SPREADSHEET_ID = "17isMrQuxVMbFjsL8sIiB6iwm3xRTr-4gELPxZmPeOTQ";
const GID = "482789258"; // ⬅ Breakdown sheet tab

const URLS = [
  `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq?tqx=out:csv&gid=${GID}`,
  `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/export?format=csv&gid=${GID}`,
];

let allData = [];
let filteredData = [];

// ====== HELPERS ======
const $ = (id) => document.getElementById(id);
const setText = (id, txt) => { const el = $(id); if (el) el.textContent = txt; };
const setStatus = (t) => setText("status", t);

const num = (v) => {
  if (v === null || v === undefined || v === "") return 0;
  const n = parseFloat(String(v).replace(/[%,\s]/g, ""));
  return isNaN(n) ? 0 : n;
};
const fmt = (n) => Math.round(n).toLocaleString();
const normalize = (s) => String(s || "").toLowerCase().replace(/\(.*?\)/g, "").replace(/[%\s_\-\/\.]/g, "").trim();

function getCol(row, ...names) {
  const keys = Object.keys(row);
  for (const n of names) {
    const target = normalize(n);
    for (const k of keys) if (normalize(k) === target) return row[k];
  }
  for (const n of names) {
    const target = normalize(n);
    if (target.length < 4) continue;
    for (const k of keys) if (normalize(k).includes(target)) return row[k];
  }
  return "";
}

// Parse "1 Dec", "12/1/2025", "2025-12-01" all into a Date
function parseDate(s) {
  if (!s) return null;
  const t = String(s).trim();
  let m;

  // m/d/yyyy or m/d/yy
  if ((m = t.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/))) {
    const y = m[3].length === 2 ? 2000 + +m[3] : +m[3];
    return new Date(y, +m[1] - 1, +m[2]);
  }
  // yyyy-mm-dd
  if ((m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/))) return new Date(+m[1], +m[2] - 1, +m[3]);

  // "1 Dec" / "31 Dec" — assume current year (2026 from operating context)
  const monthMap = { jan:0, feb:1, mar:2, apr:3, may:4, jun:5, jul:6, aug:7, sep:8, oct:9, nov:10, dec:11 };
  if ((m = t.match(/^(\d{1,2})\s+([a-z]{3,})/i))) {
    const mo = monthMap[m[2].toLowerCase().slice(0,3)];
    if (mo !== undefined) {
      const yr = new Date().getFullYear();
      return new Date(yr, mo, +m[1]);
    }
  }

  const d = new Date(t);
  return isNaN(d) ? null : d;
}

// ====== CSV PARSER ======
function parseCsv(str) {
  const rows = [];
  let row = [], field = "", inQuotes = false;
  for (let i = 0; i < str.length; i++) {
    const c = str[i];
    if (inQuotes) {
      if (c === '"' && str[i + 1] === '"') { field += '"'; i++; }
      else if (c === '"') inQuotes = false;
      else field += c;
    } else {
      if (c === '"') inQuotes = true;
      else if (c === ",") { row.push(field); field = ""; }
      else if (c === "\n") { row.push(field); rows.push(row); row = []; field = ""; }
      else if (c !== "\r") field += c;
    }
  }
  if (field.length || row.length) { row.push(field); rows.push(row); }

  // Find header row — looks for 'Date' AND 'Downtime' or 'Breakdown'
  let headerIdx = rows.findIndex(r =>
    r.some(c => /^date$/i.test((c||"").trim())) &&
    r.some(c => /downtime|breakdown/i.test((c||"").trim()))
  );
  if (headerIdx < 0) {
    headerIdx = rows.findIndex(r => r.some(c => /^date$/i.test((c||"").trim())));
  }
  if (headerIdx < 0) headerIdx = 0;

  const headers = rows[headerIdx].map(h => (h||"").trim());
  return rows.slice(headerIdx + 1)
    .filter(r => r.some(c => c && String(c).trim() !== ""))
    .map(r => Object.fromEntries(headers.map((h, i) => [h, (r[i] ?? "").toString().trim()])));
}

// ====== FETCH ======
async function loadData() {
  setStatus("Loading…");
  let lastErr;
  for (const url of URLS) {
    try {
      const res = await fetch(`${url}${url.includes("?") ? "&" : "?"}_=${Date.now()}`,
        { cache: "no-store", redirect: "follow" });
      if (!res.ok) throw new Error("HTTP " + res.status);
      const txt = await res.text();
      if (txt.trim().startsWith("<")) throw new Error("Got HTML — sheet not public");
      const parsed = parseCsv(txt);
      if (!parsed.length) throw new Error("Empty CSV");

      // Sort by date ascending
      parsed.sort((a, b) => {
        const da = parseDate(getCol(a, "Date"));
        const db = parseDate(getCol(b, "Date"));
        if (!da && !db) return 0;
        if (!da) return 1;
        if (!db) return -1;
        return da - db;
      });

      allData = parsed;
      console.log(`✅ ${parsed.length} rows. Headers:`, Object.keys(parsed[0]));
      setStatus(`✅ ${parsed.length} rows loaded`);
      initFilters();
      applyFilters();
      return;
    } catch (e) {
      console.warn("❌", url, "→", e.message);
      lastErr = e;
    }
  }
  setStatus("❌ " + (lastErr?.message || "Failed"));
  alert("Cannot load data. Make sure sheet is shared 'Anyone with the link → Viewer'.");
}

// ====== FILTERS ======
function uniqueValues(...names) {
  const set = new Set();
  allData.forEach(r => {
    const v = getCol(r, ...names);
    if (v && String(v).trim() !== "") set.add(String(v).trim());
  });
  return [...set].sort();
}

function fillSelect(id, values) {
  const sel = $(id); if (!sel) return;
  const cur = sel.value;
  sel.innerHTML = '<option value="">Select All</option>' +
    values.map(v => `<option value="${v}">${v}</option>`).join("");
  sel.value = cur;
}

function initFilters() {
  const sbuVals = uniqueValues("SBU");
  const sbuField = $("sbuField");
  if (sbuField) sbuField.style.display = sbuVals.length === 0 ? "none" : "";
  fillSelect("filterSBU", sbuVals);

  // Shop Floor → use "Shop Floor" if exists, else fall back to "Product Criteria"
  fillSelect("filterShop",
    uniqueValues("Shop Floor").length ? uniqueValues("Shop Floor") : uniqueValues("Product Criteria"));
  fillSelect("filterMill", uniqueValues("Mill Name", "Machine"));
}

function applyFilters() {
  const sbu = $("filterSBU")?.value || "";
  const shop = $("filterShop")?.value || "";
  const mill = $("filterMill")?.value || "";
  const from = parseDate($("fromDate")?.value);
  const to = parseDate($("toDate")?.value);
  if (to) to.setHours(23, 59, 59, 999);

  filteredData = allData.filter(r => {
    if (sbu && getCol(r, "SBU") !== sbu) return false;
    const shopVal = getCol(r, "Shop Floor") || getCol(r, "Product Criteria");
    if (shop && shopVal !== shop) return false;
    if (mill && getCol(r, "Mill Name", "Machine") !== mill) return false;
    const d = parseDate(getCol(r, "Date"));
    if (from && (!d || d < from)) return false;
    if (to && (!d || d > to)) return false;
    return true;
  });

  renderCards();
  renderTable();
}

// ====== RENDER CARDS ======
function renderCards() {
  let totalDowntime = 0;
  const typeCount = {};
  const reasonCount = {};
  const typeSet = new Set();

  filteredData.forEach(r => {
    const dt = num(getCol(r, "Downtime"));
    totalDowntime += dt;

    const type = getCol(r, "Breakdown Type") || "—";
    const reason = getCol(r, "Reason") || "—";
    if (type !== "—") typeSet.add(type);

    typeCount[type] = (typeCount[type] || 0) + dt;
    reasonCount[reason] = (reasonCount[reason] || 0) + dt;
  });

  // Top by total downtime
  const topType = Object.entries(typeCount).filter(([k]) => k !== "—").sort((a, b) => b[1] - a[1])[0];
  const topReason = Object.entries(reasonCount).filter(([k]) => k !== "—").sort((a, b) => b[1] - a[1])[0];

  setText("cCount", filteredData.length.toLocaleString());
  setText("cDowntime", fmt(totalDowntime));
  setText("cAvg", filteredData.length ? fmt(totalDowntime / filteredData.length) : "0");
  setText("cTypes", typeSet.size.toLocaleString());
  setText("cTopType", topType ? `${topType[0]} (${fmt(topType[1])} min)` : "—");
  setText("cTopReason", topReason ? `${topReason[0]} (${fmt(topReason[1])} min)` : "—");
}

// ====== RENDER TABLE ======
function renderTable() {
  const table = $("dataTable"); if (!table) return;
  const tbody = table.querySelector("tbody"); if (!tbody) return;

  if (!filteredData.length) {
    tbody.innerHTML = `<tr><td colspan="8" style="padding:30px;color:#9ca3af;text-align:center;">No data for selected filters</td></tr>`;
    return;
  }

  tbody.innerHTML = filteredData.map(r => `
    <tr>
      <td>${getCol(r, "Date") || ""}</td>
      <td>${getCol(r, "Product Criteria") || ""}</td>
      <td>${getCol(r, "Shift") || ""}</td>
      <td>${getCol(r, "Mill Name", "Machine") || ""}</td>
      <td>${getCol(r, "Breakdown Type") || ""}</td>
      <td>${getCol(r, "Breakdown Name", "Breakdwon Name") || ""}</td>
      <td style="text-align:left;">${getCol(r, "Reason") || ""}</td>
      <td><b>${num(getCol(r, "Downtime")) ? fmt(num(getCol(r, "Downtime"))) : ""}</b></td>
    </tr>
  `).join("");
}

// ====== EXCEL EXPORT ======
function exportExcel() {
  if (!filteredData.length) { alert("No data to export!"); return; }

  const headers = ["Date", "Product Criteria", "Shift", "Mill Name",
    "Breakdown Type", "Breakdown Name", "Reason", "Downtime (min)"];

  const rows = filteredData.map(r => [
    getCol(r, "Date"),
    getCol(r, "Product Criteria"),
    getCol(r, "Shift"),
    getCol(r, "Mill Name", "Machine"),
    getCol(r, "Breakdown Type"),
    getCol(r, "Breakdown Name", "Breakdwon Name"),
    getCol(r, "Reason"),
    getCol(r, "Downtime"),
  ]);

  const html = `
    <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">
    <head><meta charset="UTF-8"></head>
    <body><table border="1">
      <tr style="background:#c2410c;color:#fff;font-weight:bold">${headers.map(h=>`<th>${h}</th>`).join("")}</tr>
      ${rows.map(r => `<tr>${r.map(c => `<td>${c ?? ""}</td>`).join("")}</tr>`).join("")}
    </table></body></html>`;

  const blob = new Blob([html], { type: "application/vnd.ms-excel" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = `Breakdown_Report_${new Date().toISOString().slice(0,10)}.xls`;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
}

// ====== EVENTS ======
document.addEventListener("DOMContentLoaded", () => {
  ["filterSBU","filterShop","filterMill","fromDate","toDate"]
    .forEach(id => { const el = $(id); if (el) el.addEventListener("change", applyFilters); });

  $("resetBtn")?.addEventListener("click", () => {
    ["filterSBU","filterShop","filterMill","fromDate","toDate"]
      .forEach(id => { const el = $(id); if (el) el.value = ""; });
    applyFilters();
  });
  $("excelBtn")?.addEventListener("click", exportExcel);

  setInterval(loadData, 5 * 60 * 1000); // auto-refresh every 5 min
  loadData();
});
