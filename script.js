// ==== CONFIG: your published CSVs ====
const MAST_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vROnxP7pLqg7Tfg30SNF0NPvjPUDdszRqLMWGZ5HAP3xpEom02mmzGxF50sa_iAtvt7HWbkuyCqajYr/pub?gid=1561831764&single=true&output=csv"; // MAST_TMBV-VAVE
const DASH_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vROnxP7pLqg7Tfg30SNF0NPvjPUDdszRqLMWGZ5HAP3xpEom02mmzGxF50sa_iAtvt7HWbkuyCqajYr/pub?gid=595572358&single=true&output=csv"; // Dash Board

// Exact material order (rows) — normalized to single spaces
const MATERIALS = [
  "ASTM A479 SS304L / SS316L",
  "ASTM A479 SS304 / SS316",
  "ASTM A479 SS 904L/N08904", // fixed spacing/newline issue
  "Alloy 20",
  "INCOLLOY 825",
  "Hast - C",
  "Monel 400",
  "ASTM A276 XM19",
  "INCONEL 625",
  "ASTM A182 F51",
  "ASTM A479 SS316 SH",
  "ASTM A 182 Gr F53",
  "ASTM A 182 Gr F55",
  "ASTM A479 SS410",
  "ASTM A564 SS630 (17-4 ph) (1150)",
  "INCONEL 718"
];

// Columns we must show
const COLS = [
  "Ball to Stem",
  "Sealing",
  "Operator",
  "Min of all",
  "Verification",
  "FOS to Valve Torque",
  "FOS to Actuator Torque"
];

// Globals
let mast = [];
let dash = [];
let trimMap = {};
let valveTorque = null;
let actuatorTorque = null;
let lastRows = [];

// ---------------- Helpers ----------------
function papaFetch(url) {
  return new Promise((resolve, reject) => {
    Papa.parse(url, {
      download: true,
      dynamicTyping: true,
      skipEmptyLines: "greedy",
      complete: (res) => resolve(res.data),
      error: reject
    });
  });
}
function textEquals(a, b) {
  if (a == null || b == null) return false;
  return String(a).trim() === String(b).trim();
}
// Normalize labels so "ASTM A479\n SS 904L/N08904" matches "ASTM A479 SS 904L/N08904"
function normLabel(s) {
  return String(s ?? "").replace(/\s+/g, " ").trim().toUpperCase();
}
function findAllInRow(row, labels) {
  const indices = {};
  labels.forEach(l => indices[l] = -1);
  for (let c = 0; c < row.length; c++) {
    const val = row[c];
    if (val == null) continue;
    labels.forEach(l => {
      if (indices[l] === -1 && textEquals(val, l)) indices[l] = c;
    });
  }
  return indices;
}
// Parse (size,class)->trim table from Dash
function buildTrimMap(dashData) {
  const map = {};
  for (let r = 0; r < dashData.length; r++) {
    for (let c = 0; c < dashData[r].length - 1; c++) {
      const cell = dashData[r][c];
      const next = dashData[r][c + 1];
      if (typeof cell === "number" && typeof next === "number") {
        const code = cell;
        const cls = Math.round(code % 1000);
        const size = Math.round((code - cls) / 1000);
        const trim = next;
        if (size > 0 && (cls === 150 || cls === 300 || cls === 600)) {
          if (!map[cls]) map[cls] = {};
          map[cls][size] = trim;
        }
      }
    }
  }
  return map;
}
function readTorque(dashData, label) {
  for (let r = 0; r < dashData.length; r++) {
    for (let c = 0; c < dashData[r].length; c++) {
      if (textEquals(dashData[r][c], label)) {
        for (let cc = c + 1; cc < dashData[r].length; cc++) {
          const v = dashData[r][cc];
          if (typeof v === "number") return v;
        }
      }
    }
  }
  return null;
}
function getBlockStarts(mastData) {
  const row0 = mastData[0] || [];
  const blocks = findAllInRow(row0, ["Ball to Stem", "Circular", "Operator Side"]);
  const entries = Object.entries(blocks).filter(([_, idx]) => idx >= 0);
  entries.sort((a, b) => a[1] - b[1]);
  const result = [];
  for (let i = 0; i < entries.length; i++) {
    const [name, start] = entries[i];
    const end = i < entries.length - 1 ? entries[i + 1][1] - 1 : (mastData[0].length - 1);
    result.push({ name, start, end });
  }
  return result;
}
function getMaterialColsForBlock(mastData, block) {
  const row1 = mastData[1] || [];
  const map = {};
  for (let c = block.start; c <= block.end; c++) {
    const label = row1[c];
    if (label && typeof label === "string") {
      map[normLabel(label)] = c;
    }
  }
  return map;
}
function findTrimRow(mastData, trimValue) {
  const needle = "Trim " + String(Math.round(trimValue));
  for (let r = 0; r < mastData.length; r++) {
    for (let c = 0; c < Math.min(8, mastData[r].length); c++) {
      const v = mastData[r][c];
      if (typeof v === "string" && v.trim() === needle) return r;
    }
  }
  return -1;
}
function numberOrBlank(v) {
  const n = Number(v);
  return Number.isFinite(n) ? n : "";
}
function min3(a, b, c) {
  const arr = [a, b, c].filter(x => typeof x === "number" && Number.isFinite(x));
  return arr.length ? Math.min(...arr) : "";
}
function formatBool(val) { return val ? "True" : "False"; }
function setLoading(isLoading) { $("#loading").toggleClass("hidden", !isLoading); }
function setEmpty(isEmpty) { $("#empty").toggleClass("hidden", !isEmpty); }
// 2 decimals for normal numbers
function fmt2(v) {
  return (typeof v === "number" && Number.isFinite(v)) ? v.toFixed(2) : "—";
}
// For KPI torques: no decimals if the value is an integer, otherwise 2 decimals
function fmtSmart(v) {
  if (!(typeof v === "number" && Number.isFinite(v))) return "—";
  const r = Math.round(v);
  return Math.abs(v - r) < 1e-9 ? String(r) : v.toFixed(2);
}

// ---------------- Rendering ----------------
function renderTable(rows) {
  const thead = $("#data-table thead");
  const tbody = $("#data-table tbody");
  thead.empty();
  tbody.empty();

  let h = "<tr><th>Material</th>";
  COLS.forEach(c => { h += `<th>${c}</th>`; });
  h += "</tr>";
  thead.append(h);

  rows.forEach(r => {
    let tr = `<tr><td>${r.material}</td>`;
    COLS.forEach(cn => {
      const v = r[cn];
      let cell =
        (typeof v === "number" && Number.isFinite(v)) ? v.toFixed(2) :
        (v === true || v === false) ? `<span class="badge ${v ? 'ok' : 'fail'}">${formatBool(v)}</span>` :
        (v ?? "");
      tr += `<td>${cell}</td>`;
    });
    tr += "</tr>";
    tbody.append(tr);
  });

  setEmpty(rows.length === 0);
}

function renderKPIs({ cls, size, trim, vt, at }) {
  $("#kpi-class").text(cls ?? "—");
  $("#kpi-size").text(size ?? "—");
  $("#kpi-trim").text((trim ?? "") === "" ? "—" : trim);

  // Hide Valve/Actuator Torque until both Size & Class are selected
  if (cls && size) {
    $("#vt-card, #at-card").removeClass("hidden");
    $("#kpi-vt").text(fmtSmart(vt));
    $("#kpi-at").text(fmtSmart(at));
  } else {
    $("#vt-card, #at-card").addClass("hidden");
  }
}

// ---------------- Export ----------------
function exportCSV(rows) {
  if (!rows || !rows.length) return;
  const headers = ["Material", ...COLS];
  const csv = [
    headers.join(","),
    ...rows.map(r => [r.material, ...COLS.map(cn => {
      const v = r[cn];
      if (typeof v === "number" && Number.isFinite(v)) return v.toFixed(2);
      if (v === true || v === false) return formatBool(v);
      return (v ?? "");
    })].join(","))
  ].join("\n");

  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "mast_dashboard.csv";
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

// ---------------- Load & Search ----------------
async function loadAll() {
  try {
    setLoading(true);
    setEmpty(true);

    [mast, dash] = await Promise.all([papaFetch(MAST_URL), papaFetch(DASH_URL)]);
    trimMap = buildTrimMap(dash);
    valveTorque = readTorque(dash, "Valve Torque");
    actuatorTorque = readTorque(dash, "Actuator Torque");

    // Populate dropdowns
    const classes = Object.keys(trimMap).map(Number).sort((a,b)=>a-b);
    const sizes = [...new Set(Object.values(trimMap).flatMap(obj => Object.keys(obj).map(Number)))].sort((a,b)=>a-b);

    $("#classDropdown").empty().append('<option value=""></option>');
    $("#sizeDropdown").empty().append('<option value=""></option>');
    sizes.forEach(s => $("#sizeDropdown").append(`<option value="${s}">${s}</option>`));
    classes.forEach(c => $("#classDropdown").append(`<option value="${c}">${c}</option>`));

    // Select2 with clear placeholders
    $("#classDropdown").select2({ placeholder: "Select Class", allowClear: true, width: "resolve", dropdownAutoWidth: true });
    $("#sizeDropdown").select2({ placeholder: "Select Size", allowClear: true, width: "resolve", dropdownAutoWidth: true });

    // Improve the inner select2 search placeholder
    $(document).on('select2:open', () => {
      document.querySelectorAll('.select2-search__field').forEach(i => i.setAttribute('placeholder','Type to filter…'));
    });

    renderTable([]);
    renderKPIs({ cls: null, size: null, trim: null, vt: valveTorque, at: actuatorTorque });
  } catch (e) {
    console.error(e);
    alert("Failed to load sheets. Check that the CSV links are published and accessible.");
  } finally {
    setLoading(false);
  }
}

function onSearch() {
  // show loader just while computing
  setLoading(true);

  const cls = Number($("#classDropdown").val());
  const size = Number($("#sizeDropdown").val());

  if (!cls || !size) {
    lastRows = [];
    renderTable([]);
    renderKPIs({ cls: $("#classDropdown").val() || "—", size: $("#sizeDropdown").val() || "—", trim: "—", vt: valveTorque, at: actuatorTorque });
    setLoading(false);
    return;
  }

  const trim = trimMap?.[cls]?.[size];
  if (trim == null) {
    lastRows = [];
    renderTable([]);
    renderKPIs({ cls, size, trim: "—", vt: valveTorque, at: actuatorTorque });
    setLoading(false);
    alert(`No Trim mapping found for Size ${size} & Class ${cls}.`);
    return;
  }

  const blocks = getBlockStarts(mast);
  const ballBlk = blocks.find(b => b.name === "Ball to Stem");
  const sealBlk = blocks.find(b => b.name === "Circular");       // "Circular" = Sealing
  const operBlk = blocks.find(b => b.name === "Operator Side");  // Operator

  if (!ballBlk || !sealBlk || !operBlk) {
    setLoading(false);
    alert("Could not locate required blocks (Ball to Stem / Circular / Operator Side) in MAST sheet.");
    return;
  }

  const trimRow = findTrimRow(mast, trim);
  if (trimRow < 0) {
    lastRows = [];
    renderTable([]);
    renderKPIs({ cls, size, trim, vt: valveTorque, at: actuatorTorque });
    setLoading(false);
    alert(`Could not find row for "Trim ${trim}".`);
    return;
  }

  // Build maps using normalized labels
  const ballCols = getMaterialColsForBlock(mast, ballBlk);
  const sealCols = getMaterialColsForBlock(mast, sealBlk);
  const operCols = getMaterialColsForBlock(mast, operBlk);

  const rows = MATERIALS.map(mat => {
    const key = normLabel(mat);
    const vBall = numberOrBlank(mast[trimRow][ballCols[key]]);
    const vSeal = numberOrBlank(mast[trimRow][sealCols[key]]);
    const vOper = numberOrBlank(mast[trimRow][operCols[key]]);
    const vMin  = min3(vBall, vSeal, vOper);

    const fosValve = (typeof valveTorque === "number" && typeof vMin === "number") ? (vMin / valveTorque) : "";
    const fosAct   = (typeof actuatorTorque === "number" && typeof vMin === "number") ? (vMin / actuatorTorque) : "";
    const verification = (typeof valveTorque === "number" && typeof vMin === "number") ? (vMin >= valveTorque) : "";

    return {
      material: mat,
      "Ball to Stem": vBall,
      "Sealing": vSeal,
      "Operator": vOper,
      "Min of all": vMin,
      "Verification": verification,
      "FOS to Valve Torque": fosValve,
      "FOS to Actuator Torque": fosAct
    };
  });

  lastRows = rows;
  renderTable(rows);
  renderKPIs({ cls, size, trim, vt: valveTorque, at: actuatorTorque });
  setLoading(false);
}

// ---------------- Events ----------------
$("#searchBtn").on("click", onSearch);
$("#clearBtn").on("click", () => {
  $("#classDropdown").val(null).trigger("change");
  $("#sizeDropdown").val(null).trigger("change");
  lastRows = [];
  renderTable([]);
  renderKPIs({ cls: null, size: null, trim: null, vt: valveTorque, at: actuatorTorque });
});
$("#exportBtn").on("click", () => exportCSV(lastRows));
// Auto search when either select changes
$("#classDropdown").on("change", onSearch);
$("#sizeDropdown").on("change", onSearch);

// Start
loadAll();
