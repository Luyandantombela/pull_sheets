/**
 * file-reader.js
 * --------------
 * Reads uploaded Excel files using SheetJS (xlsx).
 * Exposes helpers to:
 *   1. Extract sheet names from a File object.
 *   2. Read a specific sheet as a full workbook object for import.
 */

const FileReader_ = (() => {

  /**
   * Read a File and return the list of sheet names.
   * @param {File} file
   * @returns {Promise<string[]>}
   */
  async function getSheetNames(file) {
    const buffer = await _readFileAsArrayBuffer(file);
    // Parse with full cell formulas disabled for speed (we just want sheet names here)
    const wb = XLSX.read(buffer, { type: 'array', bookSheets: true });
    return wb.SheetNames;
  }

  /**
   * Read a specific sheet from a File and return the full SheetJS workbook
   * and the worksheet object for that sheet.
   *
   * @param {File} file
   * @param {string} sheetName
   * @returns {Promise<{ workbook: Object, worksheet: Object }>}
   */
  async function readSheet(file, sheetName) {
    const buffer = await _readFileAsArrayBuffer(file);
    // cellStyles: true preserves formatting metadata
    const workbook = XLSX.read(buffer, {
      type: 'array',
      cellStyles: true,
      cellFormulas: true,
      cellDates: true,
      sheetRows: 0, // 0 = all rows
    });
    const worksheet = workbook.Sheets[sheetName];
    return { workbook, worksheet };
  }

  /**
   * Read a File as ArrayBuffer.
   * @param {File} file
   * @returns {Promise<ArrayBuffer>}
   */
  function _readFileAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload  = e => resolve(e.target.result);
      reader.onerror = () => reject(new Error(`Failed to read file: ${file.name}`));
      reader.readAsArrayBuffer(file);
    });
  }

  return { getSheetNames, readSheet };
})();
