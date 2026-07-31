/**
 * importer.js
 * -----------
 * Handles copying selected worksheets into the active Excel workbook
 * via the Office.js API.
 *
 * Strategy:
 *   1. For each selected (workbook, sheet) pair, read the sheet data with SheetJS.
 *   2. Use Excel.run() to create a new worksheet in the active workbook.
 *   3. Write values, number formats, column widths, and row heights.
 *   4. Report progress after each sheet.
 *
 * Note on fidelity:
 *   The browser Office.js APIs can write cell values, formulas, number formats,
 *   column widths, and row heights. Charts, images, and embedded objects require
 *   the Office.js insertWorksheetsFromBase64 API (Desktop Excel 2021 / M365 only).
 *   We try insertWorksheetsFromBase64 first; if unavailable we fall back to a
 *   cell-by-cell copy which preserves data + formatting.
 */

const Importer = (() => {

  /**
   * Main entry point — import all selected sheets.
   *
   * @param {Array<{ wb: WorkbookEntry, sheet: SheetEntry }>} pairs
   * @param {{ onProgress: (done: number, total: number, label: string) => void }} options
   * @returns {Promise<{ imported: number, errors: string[] }>}
   */
  async function importSheets(pairs, { onProgress } = {}) {
    const total   = pairs.length;
    let   done    = 0;
    const errors  = [];

    // Group pairs by workbook file to avoid reading the same file multiple times
    const byFile = _groupByFile(pairs);

    for (const [file, filePairs] of byFile) {
      // Try the high-fidelity path first (insertWorksheetsFromBase64)
      const base64 = await _toBase64(file);
      const sheetNames = filePairs.map(p => p.sheet.name);

      const highFiResult = await _tryHighFidelityImport(base64, sheetNames, filePairs, onProgress, done, total);

      if (highFiResult.success) {
        done += highFiResult.count;
      } else {
        // Fall back to cell-by-cell copy
        for (const pair of filePairs) {
          onProgress && onProgress(done, total, `Importing "${pair.sheet.name}"…`);
          try {
            await _importSheetCellByCellFromBase64(base64, pair.sheet.name, pair.wb.filename);
          } catch (err) {
            errors.push(`"${pair.sheet.name}" from "${pair.wb.filename}": ${err.message}`);
          }
          done++;
          onProgress && onProgress(done, total, `Imported "${pair.sheet.name}"`);
        }
      }
    }

    return { imported: done - errors.length, errors };
  }

  // ── Private helpers ─────────────────────────────────────────────────────────

  /**
   * Group import pairs by File object (same file instance = same upload).
   * @param {Array<{ wb, sheet }>} pairs
   * @returns {Map<File, Array<{ wb, sheet }>>}
   */
  function _groupByFile(pairs) {
    const map = new Map();
    pairs.forEach(pair => {
      const file = pair.wb.file;
      if (!map.has(file)) map.set(file, []);
      map.get(file).push(pair);
    });
    return map;
  }

  /**
   * Read a File as a base64 string.
   * @param {File} file
   * @returns {Promise<string>}
   */
  function _toBase64(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload  = e => {
        // result is "data:<mime>;base64,<data>" — strip the prefix
        const b64 = e.target.result.split(',')[1];
        resolve(b64);
      };
      reader.onerror = () => reject(new Error(`Cannot read file: ${file.name}`));
      reader.readAsDataURL(file);
    });
  }

  /**
   * Attempt to use insertWorksheetsFromBase64 (high-fidelity — preserves charts,
   * images, tables, conditional formatting, etc.).
   * Available in Excel Desktop 2021+ / Microsoft 365.
   *
   * @returns {Promise<{ success: boolean, count: number }>}
   */
  async function _tryHighFidelityImport(base64, sheetNames, pairs, onProgress, doneSoFar, total) {
    try {
      await Excel.run(async context => {
        // insertWorksheetsFromBase64 is the modern high-fidelity API
        // Check existence defensively
        if (typeof context.workbook.insertWorksheetsFromBase64 !== 'function') {
          throw new Error('API not available');
        }

        const existingSheets = context.workbook.worksheets;
        existingSheets.load('items/name');
        await context.sync();

        const existingNames = existingSheets.items.map(s => s.name);

        for (const pair of pairs) {
          onProgress && onProgress(doneSoFar, total, `Importing "${pair.sheet.name}"…`);

          // Resolve a safe target name
          const safeName = _uniqueName(pair.sheet.name, existingNames);
          existingNames.push(safeName);

          context.workbook.insertWorksheetsFromBase64(base64, {
            sheetNamesToInsert: [pair.sheet.name],
            positionType: Excel.WorksheetPositionType.end,
          });

          // Rename the inserted sheet if needed
          await context.sync();

          const insertedSheets = context.workbook.worksheets;
          insertedSheets.load('items/name');
          await context.sync();

          // Find the sheet that was just inserted (it keeps the original name)
          const inserted = insertedSheets.items.find(s => s.name === pair.sheet.name);
          if (inserted && safeName !== pair.sheet.name) {
            inserted.name = safeName;
            await context.sync();
          }

          doneSoFar++;
          onProgress && onProgress(doneSoFar, total, `Imported "${safeName}"`);
        }
      });

      return { success: true, count: pairs.length };
    } catch (_) {
      // Not available or failed — caller will try fallback
      return { success: false, count: 0 };
    }
  }

  /**
   * Fallback: import a single sheet cell-by-cell using SheetJS data.
   * Preserves: values, formulas, number formats, column widths, row heights.
   *
   * @param {string} base64  - The source workbook as base64
   * @param {string} sheetName
   * @param {string} filename - For naming collision prefix
   */
  async function _importSheetCellByCellFromBase64(base64, sheetName, filename) {
    // Parse source workbook from base64
    const buffer = _base64ToUint8Array(base64);
    const wb = XLSX.read(buffer, {
      type: 'array',
      cellStyles: true,
      cellFormulas: true,
      cellDates: true,
    });

    const ws = wb.Sheets[sheetName];
    if (!ws) throw new Error(`Sheet "${sheetName}" not found in workbook.`);

    const ref = ws['!ref'];
    if (!ref) {
      // Empty sheet — just create a blank sheet
      await _createBlankSheet(sheetName, filename);
      return;
    }

    // Decode the range
    const range    = XLSX.utils.decode_range(ref);
    const numRows  = range.e.r - range.s.r + 1;
    const numCols  = range.e.c - range.s.c + 1;

    // Build a 2D array of values and a parallel 2D array of number formats
    const values  = [];
    const formats = [];

    for (let r = range.s.r; r <= range.e.r; r++) {
      const rowVals = [];
      const rowFmts = [];
      for (let c = range.s.c; c <= range.e.c; c++) {
        const addr = XLSX.utils.encode_cell({ r, c });
        const cell = ws[addr];
        if (!cell) {
          rowVals.push(null);
          rowFmts.push(null);
        } else {
          // Prefer formula if present
          if (cell.f) {
            rowVals.push(`=${cell.f}`);
          } else {
            rowVals.push(cell.v !== undefined ? cell.v : null);
          }
          rowFmts.push(cell.z || null);
        }
      }
      values.push(rowVals);
      formats.push(rowFmts);
    }

    // Extract column widths (!cols) and row heights (!rows)
    const colWidths = (ws['!cols'] || []).map(c => c ? (c.wch || null) : null);
    const rowHeights = (ws['!rows'] || []).map(r => r ? (r.hpx || null) : null);

    // Merged cells
    const merges = ws['!merges'] || [];

    await Excel.run(async context => {
      const workbook  = context.workbook;
      const sheets    = workbook.worksheets;
      sheets.load('items/name');
      await context.sync();

      const existingNames = sheets.items.map(s => s.name);
      const safeName      = _uniqueName(sheetName, existingNames);

      // Add the new sheet
      const newSheet = sheets.add(safeName);
      newSheet.position = sheets.items.length; // append at end

      // Write values (Excel is 1-indexed, row/col are 0-indexed offsets)
      const startCell = XLSX.utils.encode_cell({ r: range.s.r, c: range.s.c });
      const rangeRef  = `${_xlsxToExcelAddr(range.s)}:${_xlsxToExcelAddr(range.e)}`;

      const excelRange = newSheet.getRange(rangeRef);
      excelRange.values = values;

      // Apply number formats where non-null
      formats.forEach((rowFmts, ri) => {
        rowFmts.forEach((fmt, ci) => {
          if (!fmt) return;
          const cell = newSheet.getCell(range.s.r + ri, range.s.c + ci);
          cell.numberFormat = [[fmt]];
        });
      });

      // Apply column widths (SheetJS wch is in characters; Excel width is similar)
      colWidths.forEach((wch, i) => {
        if (wch == null) return;
        const col = newSheet.getColumn(range.s.c + i + 1); // 1-indexed
        col.width = wch * 7; // approx conversion to pixels
      });

      // Apply row heights
      rowHeights.forEach((hpx, i) => {
        if (hpx == null) return;
        const row = newSheet.getRow(range.s.r + i + 1); // 1-indexed
        row.height = hpx * 0.75; // px → pt approx
      });

      // Apply merged cells
      merges.forEach(merge => {
        const mergeRef = `${_xlsxToExcelAddr(merge.s)}:${_xlsxToExcelAddr(merge.e)}`;
        try {
          newSheet.getRange(mergeRef).merge();
        } catch (_) { /* Ignore merge conflicts */ }
      });

      await context.sync();
    });
  }

  /**
   * Create an empty sheet (fallback for empty source sheets).
   */
  async function _createBlankSheet(sheetName, filename) {
    await Excel.run(async context => {
      const sheets = context.workbook.worksheets;
      sheets.load('items/name');
      await context.sync();
      const existingNames = sheets.items.map(s => s.name);
      const safeName = _uniqueName(sheetName, existingNames);
      sheets.add(safeName);
      await context.sync();
    });
  }

  /**
   * Convert SheetJS {r, c} cell address to Excel A1 notation.
   * @param {{ r: number, c: number }} cellAddr
   * @returns {string}
   */
  function _xlsxToExcelAddr({ r, c }) {
    return XLSX.utils.encode_cell({ r, c });
  }

  /**
   * Generate a unique sheet name avoiding conflicts.
   * e.g. "Sales" → "Sales (2)" → "Sales (3)"
   *
   * @param {string} desired
   * @param {string[]} existing
   * @returns {string}
   */
  function _uniqueName(desired, existing) {
    const lower = existing.map(n => n.toLowerCase());
    if (!lower.includes(desired.toLowerCase())) return desired;
    let n = 2;
    while (lower.includes(`${desired} (${n})`.toLowerCase())) n++;
    return `${desired} (${n})`;
  }

  /**
   * Convert a base64 string to a Uint8Array (for SheetJS).
   * @param {string} b64
   * @returns {Uint8Array}
   */
  function _base64ToUint8Array(b64) {
    const binary = atob(b64);
    const bytes  = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    return bytes;
  }

  return { importSheets };
})();
