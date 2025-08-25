// ==== CONFIG: your published CSVs ====
const MAST_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vROnxP7pLqg7Tfg30SNF0NPvjPUDdszRqLMWGZ5HAP3xpEom02mmzGxF50sa_iAtvt7HWbkuyCqajYr/pub?gid=1561831764&single=true&output=csv"; 
const DASH_URL = "https://docs.google.com/spreadsheets/d/e/2PACX-1vROnxP7pLqg7Tfg30SNF0NPvjPUDdszRqLMWGZ5HAP3xpEom02mmzGxF50sa_iAtvt7HWbkuyCqajYr/pub?gid=595572358&single=true&output=csv"; 

// Exact material order (rows)
const MATERIALS = [
  "ASTM A479 SS304L / SS316L",
  "ASTM A479 SS304 / SS316",
  "ASTM A479 SS 904L/N08904",
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
function fmtSmart(v) {
  if (!(typeof v === "number" && Number.isFinite(v))) return "";
  const r = Math.round(v);
  return Math.abs(v - r) < 1e-9 ? String(r) : v.toFixed(2);
}
function safeDivide(n, d) {
  if (typeof n !== "number" || !Number.isFinite(n)) return "";
  if (typeof d !== "number" || !Number.isFinite(d) || d <= 0) return "";
  return n / d;
}

// Generate timestamp for filenames
function getTimestamp() {
  const now = new Date();
  return now.toISOString().slice(0, 19).replace(/[:-]/g, '').replace('T', '_');
}

// Get current configuration info
function getCurrentConfig() {
  const cls = $("#classDropdown").val();
  const size = $("#sizeDropdown").val();
  const vt = $("#input-vt").val();
  const at = $("#input-at").val();
  return { cls, size, vt, at };
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
      let v = r[cn];
      let cell = "";

      if (cn === "Verification") {
        cell = `<span class="badge ${v ? 'ok' : 'fail'}">${formatBool(v)}</span>`;
      } else if (cn === "FOS to Valve Torque") {
        if (typeof v === "number" && v < 2) {
          cell = `<span class="badge fail">${v.toFixed(2)}</span>`;
        } else if (typeof v === "number") {
          cell = `<span class="badge ok">${v.toFixed(2)}</span>`;
        } else cell = v ?? "";
      } else if (cn === "FOS to Actuator Torque") {
        if (typeof v === "number" && v < 1) {
          cell = `<span class="badge fail">${v.toFixed(2)}</span>`;
        } else if (typeof v === "number") {
          cell = `<span class="badge ok">${v.toFixed(2)}</span>`;
        } else cell = v ?? "";
      } else {
        if (typeof v === "number" && Number.isFinite(v)) cell = v.toFixed(2);
        else cell = v ?? "";
      }

      tr += `<td>${cell}</td>`;
    });
    tr += "</tr>";
    tbody.append(tr);
  });

  setEmpty(rows.length === 0);
}

// ---------------- Export Functions ----------------

// Export to XLSX
function exportExcel(rows) {
  if (!rows || !rows.length) {
    alert("No data to export. Please run a search first.");
    return;
  }

  const wb = XLSX.utils.book_new();
  const ws = {};

  // Headers
  const headers = ["Material", ...COLS];
  headers.forEach((header, col) => {
    const cellRef = XLSX.utils.encode_cell({r: 0, c: col});
    ws[cellRef] = {
      v: header,
      t: 's',
      s: {
        font: { bold: true, color: { rgb: "24292f" } },
        fill: { fgColor: { rgb: "f3f4f6" } },
        border: {
          top: { style: "thin", color: { rgb: "d0d7de" } },
          bottom: { style: "thin", color: { rgb: "d0d7de" } },
          left: { style: "thin", color: { rgb: "d0d7de" } },
          right: { style: "thin", color: { rgb: "d0d7de" } }
        },
        alignment: { horizontal: "center", vertical: "center" }
      }
    };
  });

  // Data rows
  rows.forEach((row, rowIndex) => {
    const excelRow = rowIndex + 1;
    
    const materialCellRef = XLSX.utils.encode_cell({r: excelRow, c: 0});
    ws[materialCellRef] = {
      v: row.material,
      t: 's',
      s: {
        font: { bold: true, color: { rgb: "24292f" } },
        fill: { fgColor: { rgb: "f7f7f8" } },
        border: {
          top: { style: "thin", color: { rgb: "d0d7de" } },
          bottom: { style: "thin", color: { rgb: "d0d7de" } },
          left: { style: "thin", color: { rgb: "d0d7de" } },
          right: { style: "thin", color: { rgb: "d0d7de" } }
        },
        alignment: { horizontal: "left", vertical: "center" }
      }
    };

    COLS.forEach((colName, colIndex) => {
      const col = colIndex + 1;
      const cellRef = XLSX.utils.encode_cell({r: excelRow, c: col});
      const value = row[colName];
      
      let cellData = {
        s: {
          border: {
            top: { style: "thin", color: { rgb: "d0d7de" } },
            bottom: { style: "thin", color: { rgb: "d0d7de" } },
            left: { style: "thin", color: { rgb: "d0d7de" } },
            right: { style: "thin", color: { rgb: "d0d7de" } }
          },
          alignment: { horizontal: "center", vertical: "center" },
          fill: { fgColor: { rgb: "ffffff" } }
        }
      };

      if (colName === "Verification") {
        cellData.v = formatBool(value);
        cellData.t = 's';
        if (value) {
          cellData.s.font = { color: { rgb: "1e7e34" }, bold: true };
          cellData.s.fill = { fgColor: { rgb: "e9f7ef" } };
        } else {
          cellData.s.font = { color: { rgb: "b71c1c" }, bold: true };
          cellData.s.fill = { fgColor: { rgb: "fdecea" } };
        }
      } else if (colName === "FOS to Valve Torque") {
        if (typeof value === "number") {
          cellData.v = value;
          cellData.t = 'n';
          cellData.z = "0.00";
          cellData.s.font = { bold: true };
          if (value < 2) {
            cellData.s.font.color = { rgb: "b71c1c" };
            cellData.s.fill = { fgColor: { rgb: "fdecea" } };
          } else {
            cellData.s.font.color = { rgb: "1e7e34" };
            cellData.s.fill = { fgColor: { rgb: "e9f7ef" } };
          }
        } else {
          cellData.v = "";
          cellData.t = 's';
        }
      } else if (colName === "FOS to Actuator Torque") {
        if (typeof value === "number") {
          cellData.v = value;
          cellData.t = 'n';
          cellData.z = "0.00";
          cellData.s.font = { bold: true };
          if (value < 1) {
            cellData.s.font.color = { rgb: "b71c1c" };
            cellData.s.fill = { fgColor: { rgb: "fdecea" } };
          } else {
            cellData.s.font.color = { rgb: "1e7e34" };
            cellData.s.fill = { fgColor: { rgb: "e9f7ef" } };
          }
        } else {
          cellData.v = "";
          cellData.t = 's';
        }
      } else {
        if (typeof value === "number" && Number.isFinite(value)) {
          cellData.v = value;
          cellData.t = 'n';
          cellData.z = "0.00";
        } else {
          cellData.v = value ?? "";
          cellData.t = 's';
        }
      }

      ws[cellRef] = cellData;
    });
  });

  const colWidths = [
    { wch: 35 }, { wch: 12 }, { wch: 12 }, { wch: 12 },
    { wch: 12 }, { wch: 12 }, { wch: 15 }, { wch: 15 }
  ];
  ws['!cols'] = colWidths;

  const range = XLSX.utils.encode_range({
    s: { c: 0, r: 0 },
    e: { c: headers.length - 1, r: rows.length }
  });
  ws['!ref'] = range;

  XLSX.utils.book_append_sheet(wb, ws, "Material Capability Matrix");

  const filename = `MAST_Dashboard_${getTimestamp()}.xlsx`;
  XLSX.writeFile(wb, filename);
}

// Export to CSV
function exportCSV(rows) {
  if (!rows || !rows.length) {
    alert("No data to export. Please run a search first.");
    return;
  }

  const headers = ["Material", ...COLS];
  const csvData = [headers];

  rows.forEach(row => {
    const csvRow = [row.material];
    COLS.forEach(colName => {
      let value = row[colName];
      
      if (colName === "Verification") {
        value = formatBool(value);
      } else if (typeof value === "number" && Number.isFinite(value)) {
        value = value.toFixed(2);
      } else {
        value = value ?? "";
      }
      
      csvRow.push(value);
    });
    csvData.push(csvRow);
  });

  const csv = Papa.unparse(csvData);
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement('a');
  const url = URL.createObjectURL(blob);
  
  link.setAttribute('href', url);
  link.setAttribute('download', `MAST_Dashboard_${getTimestamp()}.csv`);
  link.style.visibility = 'hidden';
  
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
}

// Export to PDF
function exportPDF(rows) {
  if (!rows || !rows.length) {
    alert("No data to export. Please run a search first.");
    return;
  }

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF('l', 'mm', 'a4'); // Landscape orientation
  
  const config = getCurrentConfig();
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  
  // Header
  doc.setFontSize(18);
  doc.setFont(undefined, 'bold');
  doc.text('MASTER TMBV - Material Capability Matrix', 15, 20);
  
  doc.setFontSize(12);
  doc.setFont(undefined, 'normal');
  doc.text('(MAST Values in Nm)', 15, 28);
  
  // Configuration info
  doc.setFontSize(10);
  doc.text(`Configuration: Size ${config.size || 'N/A'}, Class ${config.cls || 'N/A'}`, 15, 36);
  doc.text(`Valve Torque: ${config.vt || 'N/A'} Nm, Actuator Torque: ${config.at || 'N/A'} Nm`, 15, 42);
  doc.text(`Generated: ${new Date().toLocaleString()}`, 15, 48);
  
  // Prepare table data
  const headers = ["Material", ...COLS];
  const tableData = rows.map(row => {
    const rowData = [row.material];
    COLS.forEach(colName => {
      let value = row[colName];
      
      if (colName === "Verification") {
        value = formatBool(value);
      } else if (typeof value === "number" && Number.isFinite(value)) {
        value = value.toFixed(2);
      } else {
        value = value ?? "";
      }
      
      rowData.push(value);
    });
    return rowData;
  });
  
  // Table styling
  doc.autoTable({
    head: [headers],
    body: tableData,
    startY: 55,
    styles: {
      fontSize: 8,
      cellPadding: 2,
      halign: 'center',
      valign: 'middle'
    },
    headStyles: {
      fillColor: [243, 244, 246],
      textColor: [36, 41, 47],
      fontStyle: 'bold',
      halign: 'center'
    },
    columnStyles: {
      0: { 
        halign: 'left', 
        cellWidth: 45,
        fillColor: [247, 247, 248],
        fontStyle: 'bold'
      }
    },
    didParseCell: function(data) {
      const value = data.cell.text[0];
      const colIndex = data.column.index;
      
      // Color coding for verification column
      if (headers[colIndex] === "Verification") {
        if (value === "True") {
          data.cell.styles.fillColor = [233, 247, 239];
          data.cell.styles.textColor = [30, 126, 52];
          data.cell.styles.fontStyle = 'bold';
        } else if (value === "False") {
          data.cell.styles.fillColor = [253, 236, 234];
          data.cell.styles.textColor = [183, 28, 28];
          data.cell.styles.fontStyle = 'bold';
        }
      }
      
      // Color coding for FOS columns
      if (headers[colIndex] === "FOS to Valve Torque" && value !== "") {
        const numValue = parseFloat(value);
        if (!isNaN(numValue)) {
          if (numValue < 2) {
            data.cell.styles.fillColor = [253, 236, 234];
            data.cell.styles.textColor = [183, 28, 28];
            data.cell.styles.fontStyle = 'bold';
          } else {
            data.cell.styles.fillColor = [233, 247, 239];
            data.cell.styles.textColor = [30, 126, 52];
            data.cell.styles.fontStyle = 'bold';
          }
        }
      }
      
      if (headers[colIndex] === "FOS to Actuator Torque" && value !== "") {
        const numValue = parseFloat(value);
        if (!isNaN(numValue)) {
          if (numValue < 1) {
            data.cell.styles.fillColor = [253, 236, 234];
            data.cell.styles.textColor = [183, 28, 28];
            data.cell.styles.fontStyle = 'bold';
          } else {
            data.cell.styles.fillColor = [233, 247, 239];
            data.cell.styles.textColor = [30, 126, 52];
            data.cell.styles.fontStyle = 'bold';
          }
        }
      }
    },
    margin: { top: 15, right: 15, bottom: 15, left: 15 },
    tableWidth: 'auto',
    theme: 'grid'
  });
  
  // Footer
  const finalY = doc.lastAutoTable.finalY || 55;
  if (finalY < pageHeight - 30) {
    doc.setFontSize(8);
    doc.setTextColor(128, 128, 128);
    doc.text('Generated by MASTER TMBV Dashboard', 15, pageHeight - 10);
  }
  
  // Save PDF
  const filename = `MAST_Dashboard_${getTimestamp()}.pdf`;
  doc.save(filename);
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

    const classes = Object.keys(trimMap).map(Number).sort((a,b)=>a-b);
    const sizes = [...new Set(Object.values(trimMap).flatMap(obj => Object.keys(obj).map(Number)))].sort((a,b)=>a-b);

    $("#classDropdown").empty().append('<option value=""></option>');
    $("#sizeDropdown").empty().append('<option value=""></option>');
    sizes.forEach(s => $("#sizeDropdown").append(`<option value="${s}">${s}</option>`));
    classes.forEach(c => $("#classDropdown").append(`<option value="${c}">${c}</option>`));

    $("#classDropdown").select2({ placeholder: "Select Class", allowClear: true, width: "resolve", dropdownAutoWidth: true });
    $("#sizeDropdown").select2({ placeholder: "Select Size", allowClear: true, width: "resolve", dropdownAutoWidth: true });

    $(document).on('select2:open', () => {
      document.querySelectorAll('.select2-search__field').forEach(i => i.setAttribute('placeholder','Type to filter…'));
    });

    $("#input-vt").val(valveTorque ?? "");
    $("#input-at").val(actuatorTorque ?? "");
  } catch (e) {
    console.error(e);
    alert("Failed to load sheets. Check that the CSV links are published and accessible.");
  } finally {
    setLoading(false);
  }
}

function onSearch() {
  setLoading(true);

  const cls = Number($("#classDropdown").val());
  const size = Number($("#sizeDropdown").val());

  const vtInput = Number($("#input-vt").val());
  const atInput = Number($("#input-at").val());
  const vt = Number.isFinite(vtInput) ? vtInput : null;
  const at = Number.isFinite(atInput) ? atInput : null;

  if (!cls || !size) {
    lastRows = [];
    renderTable([]);
    setLoading(false);
    return;
  }

  const trim = trimMap?.[cls]?.[size];
  if (trim == null) {
    lastRows = [];
    renderTable([]);
    setLoading(false);
    alert(`No Trim mapping found for Size ${size} & Class ${cls}.`);
    return;
  }

  const blocks = getBlockStarts(mast);
  const ballBlk = blocks.find(b => b.name === "Ball to Stem");
  const sealBlk = blocks.find(b => b.name === "Circular");
  const operBlk = blocks.find(b => b.name === "Operator Side");

  if (!ballBlk || !sealBlk || !operBlk) {
    setLoading(false);
    alert("Could not locate required blocks in MAST sheet.");
    return;
  }

  const trimRow = findTrimRow(mast, trim);
  if (trimRow < 0) {
    lastRows = [];
    renderTable([]);
    setLoading(false);
    alert(`Could not find row for "Trim ${trim}".`);
    return;
  }

  const ballCols = getMaterialColsForBlock(mast, ballBlk);
  const sealCols = getMaterialColsForBlock(mast, sealBlk);
  const operCols = getMaterialColsForBlock(mast, operBlk);

  const rows = MATERIALS.map(mat => {
    const key = normLabel(mat);
    const vBall = numberOrBlank(mast[trimRow][ballCols[key]]);
    const vSeal = numberOrBlank(mast[trimRow][sealCols[key]]);
    const vOper = numberOrBlank(mast[trimRow][operCols[key]]);
    const vMin  = min3(vBall, vSeal, vOper);

    const fosValve = safeDivide(vMin, vt);
    const fosAct   = safeDivide(vMin, at);

    const verification = (typeof vOper === "number" && vOper === vMin);

    return {
      material: mat,
      "Ball to Stem": vBall,
      "Sealing": vSeal,
      "Operator": vOper,
      "Min of all": vMin,
      "Verification": verification,
      "FOS to Valve Torque": typeof fosValve === "number" ? fosValve : "",
      "FOS to Actuator Torque": typeof fosAct === "number" ? fosAct : ""
    };
  });

  lastRows = rows;
  renderTable(rows);
  setLoading(false);
}

// ---------------- Events ----------------
$(function () {
  loadAll();

  // Dropdown toggle functionality
  $("#exportMenuBtn").on("click", function(e) {
    e.stopPropagation();
    $("#exportMenu").toggleClass("hidden");
  });

  // Close dropdown when clicking outside
  $(document).on("click", function() {
    $("#exportMenu").addClass("hidden");
  });

  // Prevent dropdown from closing when clicking inside
  $("#exportMenu").on("click", function(e) {
    e.stopPropagation();
  });

  // Export button handlers
  $("#exportXlsxBtn").on("click", () => {
    exportExcel(lastRows);
    $("#exportMenu").addClass("hidden");
  });

  $("#exportCsvBtn").on("click", () => {
    exportCSV(lastRows);
    $("#exportMenu").addClass("hidden");
  });

  $("#exportPdfBtn").on("click", () => {
    exportPDF(lastRows);
    $("#exportMenu").addClass("hidden");
  });

  // Other existing handlers
  $("#searchBtn").on("click", onSearch);
  $("#clearBtn").on("click", () => {
    $("#classDropdown").val(null).trigger("change");
    $("#sizeDropdown").val(null).trigger("change");
    $("#input-vt").val(valveTorque ?? "");
    $("#input-at").val(actuatorTorque ?? "");
    lastRows = [];
    renderTable([]);
  });
  
  $("#classDropdown").on("change", onSearch);
  $("#sizeDropdown").on("change", onSearch);
  $("#input-vt, #input-at").on("input", onSearch);
});
