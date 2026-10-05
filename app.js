const yesterdayInput = document.getElementById("yesterdayFile");
const todayInput = document.getElementById("todayFile");
const compareBtn = document.getElementById("compareBtn");

const results = document.getElementById("results");
const statusBox = document.getElementById("status");

let latestResult = null;

yesterdayInput.addEventListener("change", updateFileNames);
todayInput.addEventListener("change", updateFileNames);

compareBtn.addEventListener("click", compareFiles);
document.getElementById("downloadExcel").addEventListener("click", downloadExcel);
document.getElementById("downloadTxt").addEventListener("click", downloadTxt);

function updateFileNames() {
  document.getElementById("yesterdayName").textContent =
    yesterdayInput.files[0]?.name || "No file selected";

  document.getElementById("todayName").textContent =
    todayInput.files[0]?.name || "No file selected";

  compareBtn.disabled =
    !(yesterdayInput.files.length && todayInput.files.length);
}

function setStatus(message, error = false) {
  statusBox.textContent = message;
  statusBox.className = "status" + (error ? " error" : "");
  statusBox.classList.remove("hidden");
}

function isBlank(value) {
  return value === null ||
         value === undefined ||
         String(value).trim() === "";
}

function displayValue(value) {
  return isBlank(value) ? "<BLANK>" : String(value);
}

function valuesDifferent(oldValue, newValue) {
  if (isBlank(oldValue) && isBlank(newValue)) {
    return false;
  }

  return String(oldValue) !== String(newValue);
}

function detectSkuColumn(headers) {
  return ["sku", "SKU", "#SKU"].find(
    column => headers.includes(column)
  ) || null;
}

/*
 * CSV parser supporting:
 * - commas
 * - quoted fields
 * - commas inside quoted fields
 * - escaped quotes
 * - line breaks inside quoted fields
 */
function parseCSV(text) {
  const rows = [];

  let row = [];
  let field = "";
  let quoted = false;

  for (let i = 0; i < text.length; i++) {
    const character = text[i];

    if (quoted) {
      if (character === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i++;
        } else {
          quoted = false;
        }
      } else {
        field += character;
      }

      continue;
    }

    if (character === '"') {
      quoted = true;
    } else if (character === ",") {
      row.push(field);
      field = "";
    } else if (character === "\n") {
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else if (character !== "\r") {
      field += character;
    }
  }

  if (field.length || row.length) {
    row.push(field);
    rows.push(row);
  }

  if (!rows.length) {
    return [];
  }

  const headers = rows[0].map((header, index) => {
    let result = header.trim();

    if (index === 0) {
      result = result.replace(/^\uFEFF/, "");
    }

    return result;
  });

  return rows
    .slice(1)
    .filter(row => row.some(value => value !== ""))
    .map(row => {
      const object = {};

      headers.forEach((header, index) => {
        object[header] = row[index] ?? "";
      });

      return object;
    });
}

async function readCSVFile(file) {
  const text = await file.text();
  return parseCSV(text);
}

async function compareFiles() {
  results.classList.add("hidden");

  setStatus(
    "Reading and comparing files locally in your browser..."
  );

  try {
    const yesterdayRows =
      await readCSVFile(yesterdayInput.files[0]);

    const todayRows =
      await readCSVFile(todayInput.files[0]);

    if (!yesterdayRows.length || !todayRows.length) {
      throw new Error("One of the CSV files is empty.");
    }

    const yesterdayHeaders =
      Object.keys(yesterdayRows[0]);

    const todayHeaders =
      Object.keys(todayRows[0]);

    const sharedHeaders =
      yesterdayHeaders.filter(
        header => todayHeaders.includes(header)
      );

    const skuColumn =
      detectSkuColumn(sharedHeaders);

    if (!skuColumn) {
      throw new Error(
        "SKU column was not found in both CSV files. Expected sku, SKU, or #SKU."
      );
    }

    const yesterdayMap = new Map();
    const todayMap = new Map();

    const yesterdayDuplicateSet = new Set();
    const todayDuplicateSet = new Set();

    for (const row of yesterdayRows) {
      const sku = row[skuColumn];

      if (!isBlank(sku)) {
        if (yesterdayMap.has(sku)) {
          yesterdayDuplicateSet.add(sku);
        } else {
          yesterdayMap.set(sku, row);
        }
      }
    }

    for (const row of todayRows) {
      const sku = row[skuColumn];

      if (!isBlank(sku)) {
        if (todayMap.has(sku)) {
          todayDuplicateSet.add(sku);
        } else {
          todayMap.set(sku, row);
        }
      }
    }

    const yesterdaySkus =
      new Set(yesterdayMap.keys());

    const todaySkus =
      new Set(todayMap.keys());

    const newSkus =
      [...todaySkus]
        .filter(sku => !yesterdaySkus.has(sku))
        .sort();

    const removedSkus =
      [...yesterdaySkus]
        .filter(sku => !todaySkus.has(sku))
        .sort();

    const commonSkus =
      [...yesterdaySkus]
        .filter(sku => todaySkus.has(sku));

    const commonColumns =
      yesterdayHeaders.filter(
        column =>
          todayHeaders.includes(column) &&
          column !== skuColumn
      );

    const changes = [];

    for (const sku of commonSkus) {
      const oldRow = yesterdayMap.get(sku);
      const newRow = todayMap.get(sku);

      for (const column of commonColumns) {
        const oldValue = oldRow[column];
        const newValue = newRow[column];

        if (valuesDifferent(oldValue, newValue)) {
          changes.push({
            SKU: sku,
            Column: column,
            Yesterday: displayValue(oldValue),
            Today: displayValue(newValue)
          });
        }
      }
    }

    const changedSkus =
      new Set(changes.map(change => change.SKU));

    const columnStats = {};

    for (const change of changes) {
      if (!columnStats[change.Column]) {
        columnStats[change.Column] = {
          changes: 0,
          skus: new Set()
        };
      }

      columnStats[change.Column].changes++;
      columnStats[change.Column].skus.add(change.SKU);
    }

    latestResult = {
      yesterdayFile:
        yesterdayInput.files[0].name,

      todayFile:
        todayInput.files[0].name,

      yesterdayRows,
      todayRows,

      yesterdayHeaders,
      todayHeaders,

      skuColumn,

      yesterdayMap,
      todayMap,

      newSkus,
      removedSkus,
      commonSkus,

      changes,
      changedSkus,
      columnStats,

      yesterdayDuplicates:
        [...yesterdayDuplicateSet].sort(),

      todayDuplicates:
        [...todayDuplicateSet].sort()
    };

    renderResults(latestResult);

    statusBox.classList.add("hidden");
    results.classList.remove("hidden");

  } catch (error) {
    setStatus(
      error.message || "Comparison failed.",
      true
    );
  }
}

function renderResults(result) {
  const summary = [
    [
      "Yesterday SKUs",
      result.yesterdayMap.size
    ],
    [
      "Today SKUs",
      result.todayMap.size
    ],
    [
      "New SKUs",
      result.newSkus.length
    ],
    [
      "Removed SKUs",
      result.removedSkus.length
    ],
    [
      "Changed SKUs",
      result.changedSkus.size
    ],
    [
      "Field Changes",
      result.changes.length
    ]
  ];

  document.getElementById("summaryCards").innerHTML =
    summary.map(item => `
      <div class="summary-card">
        <div class="summary-label">
          ${escapeHTML(item[0])}
        </div>

        <div class="summary-value">
          ${item[1].toLocaleString()}
        </div>
      </div>
    `).join("");

  const columnStats =
    Object.entries(result.columnStats)
      .sort(
        (a, b) =>
          b[1].changes - a[1].changes
      );

  document.getElementById("columnTable").innerHTML = `
    <thead>
      <tr>
        <th>Column</th>
        <th>Field Changes</th>
        <th>SKUs Affected</th>
      </tr>
    </thead>

    <tbody>
      ${
        columnStats.length
          ? columnStats.map(([column, stats]) => `
              <tr>
                <td>${escapeHTML(column)}</td>
                <td>${stats.changes.toLocaleString()}</td>
                <td>${stats.skus.size.toLocaleString()}</td>
              </tr>
            `).join("")
          : `
              <tr>
                <td colspan="3">
                  No column changes found.
                </td>
              </tr>
            `
      }
    </tbody>
  `;

  document.getElementById("newSkus").innerHTML =
    result.newSkus.length
      ? result.newSkus
          .map(sku =>
            `<div class="sku">${escapeHTML(sku)}</div>`
          )
          .join("")
      : "None";

  document.getElementById("removedSkus").innerHTML =
    result.removedSkus.length
      ? result.removedSkus
          .map(sku =>
            `<div class="sku">${escapeHTML(sku)}</div>`
          )
          .join("")
      : "None";

  document.getElementById("txtReport").textContent =
    buildTXTReport(result);
}

function buildTXTReport(result) {
  const line = "=".repeat(100);
  const separator = "-".repeat(100);

  const report = [];

  report.push(
    line,
    "PRODUCT CSV - YESTERDAY vs TODAY COMPARISON REPORT",
    line
  );

  report.push(
    `Yesterday file : ${result.yesterdayFile}`,
    `Today file     : ${result.todayFile}`
  );

  report.push(
    "",
    line,
    "1. DATASET SUMMARY",
    line
  );

  report.push(
    `Yesterday rows : ${result.yesterdayRows.length.toLocaleString()}`,
    `Today rows     : ${result.todayRows.length.toLocaleString()}`,
    `Yesterday cols : ${result.yesterdayHeaders.length.toLocaleString()}`,
    `Today cols     : ${result.todayHeaders.length.toLocaleString()}`
  );

  report.push(
    "",
    "Yesterday columns:",
    result.yesterdayHeaders.join(", ")
  );

  report.push(
    "",
    "Today columns:",
    result.todayHeaders.join(", ")
  );

  report.push(
    "",
    line,
    "2. SKU SUMMARY",
    line
  );

  report.push(
    `SKU column            : ${result.skuColumn}`,
    `Yesterday unique SKUs : ${result.yesterdayMap.size.toLocaleString()}`,
    `Today unique SKUs     : ${result.todayMap.size.toLocaleString()}`,
    `Common SKUs           : ${result.commonSkus.length.toLocaleString()}`
  );

  report.push(
    "",
    "Duplicate SKU check:",
    `Yesterday duplicate SKUs : ${result.yesterdayDuplicates.length.toLocaleString()}`,
    `Today duplicate SKUs     : ${result.todayDuplicates.length.toLocaleString()}`
  );

  if (result.yesterdayDuplicates.length) {
    report.push(
      "",
      "WARNING - Duplicate SKUs in yesterday:",
      result.yesterdayDuplicates.join("\n")
    );
  }

  if (result.todayDuplicates.length) {
    report.push(
      "",
      "WARNING - Duplicate SKUs in today:",
      result.todayDuplicates.join("\n")
    );
  }

  report.push(
    "",
    line,
    "3. NEW SKUs",
    line,
    `New SKU count: ${result.newSkus.length.toLocaleString()}`
  );

  report.push(
    result.newSkus.length
      ? result.newSkus.join("\n")
      : "None"
  );

  report.push(
    "",
    line,
    "4. REMOVED SKUs",
    line,
    `Removed SKU count: ${result.removedSkus.length.toLocaleString()}`
  );

  report.push(
    result.removedSkus.length
      ? result.removedSkus.join("\n")
      : "None"
  );

  report.push(
    "",
    line,
    "5. COLUMN-LEVEL CHANGE COUNTS",
    line
  );

  report.push(
    `${"Column".padEnd(40)} ` +
    `${"Field Changes".padStart(15)} ` +
    `${"SKUs Affected".padStart(15)}`
  );

  report.push("-".repeat(75));

  const columnStats =
    Object.entries(result.columnStats)
      .sort(
        (a, b) =>
          b[1].changes - a[1].changes
      );

  if (!columnStats.length) {
    report.push("No column changes found.");
  } else {
    for (const [column, stats] of columnStats) {
      report.push(
        `${String(column).padEnd(40)} ` +
        `${String(stats.changes).padStart(15)} ` +
        `${String(stats.skus.size).padStart(15)}`
      );
    }
  }

  report.push(
    "",
    line,
    "6. DETAILED VALUE CHANGES",
    line
  );

  report.push(
    `SKUs with value changes : ${result.changedSkus.size.toLocaleString()}`,
    `Total field changes     : ${result.changes.length.toLocaleString()}`
  );

  if (!result.changes.length) {
    report.push(
      "No value changes found for existing SKUs."
    );
  } else {
    const grouped = {};

    for (const change of result.changes) {
      if (!grouped[change.SKU]) {
        grouped[change.SKU] = [];
      }

      grouped[change.SKU].push(change);
    }

    for (const sku of Object.keys(grouped).sort()) {
      report.push(
        "",
        separator,
        `SKU: ${sku}`,
        separator
      );

      grouped[sku]
        .sort(
          (a, b) =>
            a.Column.localeCompare(b.Column)
        )
        .forEach(change => {
          report.push(
            `Column    : ${change.Column}`,
            `Yesterday : ${change.Yesterday}`,
            `Today     : ${change.Today}`,
            ""
          );
        });
    }
  }

  report.push(
    "",
    line,
    "7. FINAL SUMMARY",
    line
  );

  report.push(
    `Yesterday total SKUs      : ${result.yesterdayMap.size.toLocaleString()}`,
    `Today total SKUs          : ${result.todayMap.size.toLocaleString()}`,
    `New SKUs                  : ${result.newSkus.length.toLocaleString()}`,
    `Removed SKUs              : ${result.removedSkus.length.toLocaleString()}`,
    `SKUs with value changes   : ${result.changedSkus.size.toLocaleString()}`,
    `Total field-level changes : ${result.changes.length.toLocaleString()}`,
    `Unchanged SKUs            : ${(result.commonSkus.length - result.changedSkus.size).toLocaleString()}`
  );

  report.push(
    "",
    line,
    "END OF REPORT",
    line
  );

  return report.join("\n");
}

function downloadTXT() {
  if (!latestResult) {
    return;
  }

  const content =
    buildTXTReport(latestResult);

  const blob =
    new Blob(
      [content],
      {
        type: "text/plain;charset=utf-8"
      }
    );

  downloadBlob(
    blob,
    "product_comparison_report.txt"
  );
}

function downloadExcel() {
  if (!latestResult) {
    return;
  }

  if (typeof XLSX === "undefined") {
    setStatus(
      "Excel library could not be loaded. Check your internet connection.",
      true
    );

    return;
  }

  const result = latestResult;

  const workbook =
    XLSX.utils.book_new();

  const summary = [
    ["Metric", "Value"],

    ["Yesterday File",
      result.yesterdayFile],

    ["Today File",
      result.todayFile],

    ["SKU Column",
      result.skuColumn],

    ["Yesterday Rows",
      result.yesterdayRows.length],

    ["Today Rows",
      result.todayRows.length],

    ["Yesterday Columns",
      result.yesterdayHeaders.length],

    ["Today Columns",
      result.todayHeaders.length],

    ["Yesterday Unique SKUs",
      result.yesterdayMap.size],

    ["Today Unique SKUs",
      result.todayMap.size],

    ["Common SKUs",
      result.commonSkus.length],

    ["New SKUs",
      result.newSkus.length],

    ["Removed SKUs",
      result.removedSkus.length],

    ["SKUs With Value Changes",
      result.changedSkus.size],

    ["Total Field-Level Changes",
      result.changes.length],

    ["Unchanged SKUs",
      result.commonSkus.length -
      result.changedSkus.size],

    ["Yesterday Duplicate SKUs",
      result.yesterdayDuplicates.length],

    ["Today Duplicate SKUs",
      result.todayDuplicates.length]
  ];

  addWorksheet(
    workbook,
    "Summary",
    summary
  );

  addWorksheet(
    workbook,
    "New SKUs",
    [
      [result.skuColumn],
      ...result.newSkus.map(sku => [sku])
    ]
  );

  addWorksheet(
    workbook,
    "Removed SKUs",
    [
      [result.skuColumn],
      ...result.removedSkus.map(sku => [sku])
    ]
  );

  const columnRows = [
    [
      "Column",
      "Field Changes",
      "SKUs Affected"
    ]
  ];

  Object.entries(result.columnStats)
    .sort(
      (a, b) =>
        b[1].changes - a[1].changes
    )
    .forEach(([column, stats]) => {
      columnRows.push([
        column,
        stats.changes,
        stats.skus.size
      ]);
    });

  addWorksheet(
    workbook,
    "Column Changes",
    columnRows
  );

  addWorksheet(
    workbook,
    "Detailed Changes",
    [
      [
        "SKU",
        "Column",
        "Yesterday",
        "Today"
      ],
      ...result.changes.map(change => [
        change.SKU,
        change.Column,
        change.Yesterday,
        change.Today
      ])
    ]
  );

  addWorksheet(
    workbook,
    "Duplicate SKUs",
    [
      ["File", "SKU"],

      ...result.yesterdayDuplicates.map(
        sku => [
          result.yesterdayFile,
          sku
        ]
      ),

      ...result.todayDuplicates.map(
        sku => [
          result.todayFile,
          sku
        ]
      )
    ]
  );

  const changedSet =
    result.changedSkus;

  addWorksheet(
    workbook,
    "Common SKU Summary",
    [
      ["SKU", "Status"],

      ...result.commonSkus
        .sort()
        .map(sku => [
          sku,
          changedSet.has(sku)
            ? "Changed"
            : "Unchanged"
        ])
    ]
  );

  XLSX.writeFile(
    workbook,
    "product_comparison_report.xlsx"
  );
}

function addWorksheet(
  workbook,
  sheetName,
  data
) {
  const worksheet =
    XLSX.utils.aoa_to_sheet(data);

  if (worksheet["!ref"]) {
    worksheet["!autofilter"] = {
      ref: worksheet["!ref"]
    };

    worksheet["!cols"] =
      calculateColumnWidths(data);
  }

  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    sheetName
  );
}

function calculateColumnWidths(data) {
  const widths = [];

  for (const row of data) {
    row.forEach((value, index) => {
      const length =
        String(value ?? "").length;

      widths[index] =
        Math.min(
          Math.max(
            widths[index] || 10,
            length + 2
          ),
          60
        );
    });
  }

  return widths.map(width => ({
    wch: width
  }));
}

function downloadBlob(blob, filename) {
  const link =
    document.createElement("a");

  const url =
    URL.createObjectURL(blob);

  link.href = url;
  link.download = filename;

  document.body.appendChild(link);
  link.click();
  link.remove();

  setTimeout(
    () => URL.revokeObjectURL(url),
    1000
  );
}

function escapeHTML(value) {
  return String(value)
    .replace(
      /[&<>"']/g,
      character => ({
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#39;"
      })[character]
    );
}
