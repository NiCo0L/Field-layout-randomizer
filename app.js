(() => {
  const treatmentsInput = document.getElementById("treatmentsInput");
  const replicationsEl = document.getElementById("replications");
  const rowsCountEl = document.getElementById("rowsCount");
  const colsCountEl = document.getElementById("colsCount");
  const orderTypeEl = document.getElementById("orderType");
  const startCornerEl = document.getElementById("startCorner");
  const useSeedEl = document.getElementById("useSeed");
  const replicationSelect = document.getElementById("replicationSelect");
  const gridWrap = document.getElementById("gridWrap");
  const generateBtn = document.getElementById("generateBtn");
  const downloadBtn = document.getElementById("downloadBtn");
  const statusEl = document.getElementById("status");

  let generatedLayouts = null;
  let lastConfig = null;

  function mulberry32(seed) {
    let a = seed >>> 0;
    return function rng() {
      a |= 0;
      a = (a + 0x6d2b79f5) | 0;
      let t = Math.imul(a ^ (a >>> 15), 1 | a);
      t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }

  function seedFromString(s) {
    // Simple stable hash -> uint32
    let h = 2166136261 >>> 0;
    for (let i = 0; i < s.length; i++) {
      h ^= s.charCodeAt(i);
      h = Math.imul(h, 16777619) >>> 0;
    }
    return h >>> 0;
  }

  function getRng() {
    const seedText = (useSeedEl.value || "").trim();
    if (!seedText) return Math.random;

    // If user entered a number, use it; otherwise hash string.
    const asNum = Number(seedText);
    const seed = Number.isFinite(asNum) ? (asNum >>> 0) : seedFromString(seedText);
    return mulberry32(seed);
  }

  function normalizeType(typeRaw) {
    const t = (typeRaw || "").toString().trim().toLowerCase();
    if (!t) return "test";
    if (t.startsWith("c") || t === "1" || t.includes("check")) return "check";
    if (t.startsWith("t") || t === "0" || t.includes("test")) return "test";
    // Fallback: if unknown, keep as-is
    return t;
  }

  function parseTreatments() {
    const lines = (treatmentsInput.value || "")
      .split(/\r?\n/)
      .map((l) => l.trim())
      .filter(Boolean);

    if (lines.length === 0) return [];

    const treatments = [];
    for (const line of lines) {
      // Support formats:
      // code,type
      // code type
      // code;type
      const parts = line.split(/[,\t;]+/).map((p) => p.trim()).filter(Boolean);
      if (parts.length === 1) {
        const ws = line.split(/\s+/).map((p) => p.trim()).filter(Boolean);
        if (ws.length >= 2) {
          treatments.push({ code: ws[0], type: normalizeType(ws[1]) });
        } else {
          treatments.push({ code: parts[0], type: "test" });
        }
        continue;
      }

      treatments.push({
        code: parts[0],
        type: normalizeType(parts[1]),
      });
    }

    return treatments;
  }

  function shuffleInPlace(arr, rng) {
    for (let i = arr.length - 1; i > 0; i--) {
      const j = Math.floor(rng() * (i + 1));
      const tmp = arr[i];
      arr[i] = arr[j];
      arr[j] = tmp;
    }
  }

  function computePositions(rows, cols, orderType, startCorner) {
    // Physical coordinates:
    // r: 0..rows-1 top->bottom
    // c: 0..cols-1 left->right
    const isTop = startCorner.startsWith("top");
    const isLeft = startCorner.endsWith("left");

    const rowAsc = Array.from({ length: rows }, (_, i) => i);
    const rowDesc = Array.from({ length: rows }, (_, i) => rows - 1 - i);
    const colAsc = Array.from({ length: cols }, (_, i) => i);
    const colDesc = Array.from({ length: cols }, (_, i) => cols - 1 - i);

    if (orderType === "row-col") {
      const rowOrder = isTop ? rowAsc : rowDesc;
      const colOrder = isLeft ? colAsc : colDesc;
      const positions = [];
      for (const r of rowOrder) {
        for (const c of colOrder) positions.push({ r, c });
      }
      return positions;
    }

    // Column serpentine:
    // Scan columns in left->right or right->left order, and alternate row direction per column.
    const colOrder = isLeft ? colAsc : colDesc;
    const positions = [];
    for (let seqIdx = 0; seqIdx < colOrder.length; seqIdx++) {
      const c = colOrder[seqIdx];
      const traverseTopToBottom = isTop ? seqIdx % 2 === 0 : seqIdx % 2 === 1;
      const rowsForCol = traverseTopToBottom ? rowAsc : rowDesc;
      for (const r of rowsForCol) positions.push({ r, c });
    }
    return positions;
  }

  function renderGrid(repIndex) {
    if (!generatedLayouts) return;

    const rows = lastConfig.rows;
    const cols = lastConfig.cols;
    const grid = generatedLayouts[repIndex];

    // Build table with axis labels
    const table = document.createElement("table");
    table.className = "grid";

    const theadRow = document.createElement("tr");
    const corner = document.createElement("td");
    corner.className = "axis";
    corner.textContent = "";
    theadRow.appendChild(corner);
    for (let c = 0; c < cols; c++) {
      const td = document.createElement("td");
      td.className = "axis";
      td.textContent = c + 1;
      theadRow.appendChild(td);
    }
    table.appendChild(theadRow);

    for (let r = 0; r < rows; r++) {
      const tr = document.createElement("tr");
      const rowHeader = document.createElement("td");
      rowHeader.className = "axis";
      rowHeader.textContent = r + 1;
      tr.appendChild(rowHeader);

      for (let c = 0; c < cols; c++) {
        const td = document.createElement("td");
        const cell = grid[r][c];
        if (!cell) {
          td.className = "cell-blank";
          td.textContent = "";
        } else {
          const cls = cell.type === "check" ? "cell-check" : "cell-test";
          td.className = cls;
          td.textContent = cell.code;
          td.title = `Replication ${repIndex + 1} | Row ${r + 1} | Col ${c + 1} | ${cell.type}`;
        }
        tr.appendChild(td);
      }
      table.appendChild(tr);
    }

    gridWrap.innerHTML = "";
    gridWrap.appendChild(table);
  }

  function setStatus(msg) {
    statusEl.textContent = msg;
  }

  function generate() {
    const treatments = parseTreatments();
    const replications = Math.max(1, Number(replicationsEl.value));
    const rows = Math.max(1, Number(rowsCountEl.value));
    const cols = Math.max(1, Number(colsCountEl.value));
    const orderType = orderTypeEl.value;
    const startCorner = startCornerEl.value;

    if (treatments.length === 0) {
      generatedLayouts = null;
      downloadBtn.disabled = true;
      replicationSelect.innerHTML = "";
      setStatus("Please enter at least one entry (code + type).");
      return;
    }

    const nPlots = rows * cols;
    if (treatments.length > nPlots) {
      generatedLayouts = null;
      downloadBtn.disabled = true;
      replicationSelect.innerHTML = "";
      setStatus(
        `Too many entries for the field: entries=${treatments.length}, rows*cols=${nPlots}. Reduce entries or increase field size.`
      );
      return;
    }

    const rng = getRng();
    const positions = computePositions(rows, cols, orderType, startCorner);
    const scanIndexByRC = new Map();
    for (let i = 0; i < positions.length; i++) {
      scanIndexByRC.set(`${positions[i].r},${positions[i].c}`, i + 1); // 1-based
    }

    const layouts = [];
    for (let rep = 0; rep < replications; rep++) {
      const treatmentsCopy = treatments.map((t) => ({ code: t.code, type: t.type }));
      shuffleInPlace(treatmentsCopy, rng);

      const grid = Array.from({ length: rows }, () => Array.from({ length: cols }, () => null));

      // Fill only the first N (entries) scan positions with randomized treatments.
      // Remaining positions are always blank, so blanks are a "tail" in the scan order.
      for (let i = 0; i < treatmentsCopy.length; i++) {
        const { r, c } = positions[i];
        grid[r][c] = treatmentsCopy[i];
      }
      layouts.push(grid);
    }

    generatedLayouts = layouts;
    lastConfig = { rows, cols, replications, orderType, startCorner, scanIndexByRC };

    // Replication selector
    replicationSelect.disabled = false;
    replicationSelect.innerHTML = "";
    for (let i = 0; i < replications; i++) {
      const opt = document.createElement("option");
      opt.value = String(i);
      opt.textContent = `Replication ${i + 1}`;
      replicationSelect.appendChild(opt);
    }

    downloadBtn.disabled = false;

    const blankCount = nPlots - treatments.length;
    setStatus(
      `Generated ${replications} block(s). Field: ${rows}x${cols} (${nPlots} plots). Entries: ${treatments.length}. Blank plots: ${blankCount}.\nOrder: ${orderType.replace("-", " ")} | Start: ${startCorner}.`
    );

    // Quick validation: show that each replication contains all entries (once per block).
    // Useful if user is reviewing the Excel and wants to confirm replication coverage.
    const uniqueByRep = [];
    for (let rep = 0; rep < replications; rep++) {
      const set = new Set();
      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
          const cell = generatedLayouts[rep][r][c];
          if (cell) set.add(cell.code);
        }
      }
      uniqueByRep.push({ rep: rep + 1, found: set.size });
    }

    // Count-based validation (handles duplicates codes better than Set-size).
    const expectedCounts = new Map();
    for (const t of treatments) expectedCounts.set(t.code, (expectedCounts.get(t.code) || 0) + 1);
    const mismatch = [];
    for (let rep = 0; rep < replications; rep++) {
      const foundCounts = new Map();
      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
          const cell = generatedLayouts[rep][r][c];
          if (!cell) continue;
          foundCounts.set(cell.code, (foundCounts.get(cell.code) || 0) + 1);
        }
      }
      for (const [code, expected] of expectedCounts.entries()) {
        const got = foundCounts.get(code) || 0;
        if (got !== expected) {
          mismatch.push(`Rep ${rep + 1} code ${code}: expected ${expected}, found ${got}`);
          break;
        }
      }
    }
    if (mismatch.length > 0) {
      setStatus((statusEl.textContent || "") + `\n\nWarning: Entry count mismatch detected:\n${mismatch.join("\n")}`);
    }

    renderGrid(0);
  }

  function exportExcel() {
    if (!generatedLayouts || !lastConfig) return;
    const XLSX = window.XLSX;
    if (!XLSX) {
      setStatus("Excel export failed: SheetJS library not loaded.");
      return;
    }

    const { rows, cols, replications, orderType, startCorner, scanIndexByRC } = lastConfig;
    const positionsTotal = rows * cols;
    const treatmentsCount = (() => {
      // Count non-null from first replication
      let k = 0;
      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
          if (generatedLayouts[0][r][c]) k++;
        }
      }
      return k;
    })();

    const aoa = [];
    aoa.push([
      "Replication",
      "Row",
      "Column",
      "X",
      "Y",
      "ScanIndex",
      "EntryCode",
      "EntryType",
      "OrderType",
      "StartCorner",
    ]);

    // X/Y definition for Excel:
    // X = column number (1..cols) from left to right
    // Y = row number (1..rows) from top to bottom
    for (let rep = 0; rep < replications; rep++) {
      for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
          const cell = generatedLayouts[rep][r][c];
          aoa.push([
            rep + 1,
            r + 1,
            c + 1,
            c + 1,
            r + 1,
            scanIndexByRC ? scanIndexByRC.get(`${r},${c}`) || "" : "",
            cell ? cell.code : "",
            cell ? cell.type : "",
            orderType,
            startCorner,
          ]);
        }
      }
    }

    const ws = XLSX.utils.aoa_to_sheet(aoa);

    // Basic column sizing (non-critical)
    ws["!cols"] = [
      { wch: 11 },
      { wch: 6 },
      { wch: 8 },
      { wch: 6 },
      { wch: 6 },
      { wch: 14 },
      { wch: 11 },
      { wch: 12 },
      { wch: 12 },
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "FieldLayout");

    const stamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, "-");
    const fileName = `field_layout_randomizer_${rows}x${cols}_${treatmentsCount}entries_${stamp}.xlsx`;
    XLSX.writeFile(wb, fileName);
  }

  // Wire events
  generateBtn.addEventListener("click", generate);
  downloadBtn.addEventListener("click", exportExcel);
  replicationSelect.addEventListener("change", () => {
    const idx = Number(replicationSelect.value);
    renderGrid(idx);
  });

  // Initial state
  replicationSelect.disabled = true;
  setStatus("Enter entries, then click “Generate layout”.");
})();

