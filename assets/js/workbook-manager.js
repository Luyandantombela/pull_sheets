/**
 * workbook-manager.js
 * -------------------
 * Central state store for uploaded workbooks and their selected sheets.
 * All other modules read from / write to this single source of truth.
 */

const WorkbookManager = (() => {
  /** @type {Map<string, WorkbookEntry>} keyed by a unique workbook id */
  const _workbooks = new Map();

  /** Counter used to generate unique IDs */
  let _idCounter = 0;

  /**
   * @typedef {Object} SheetEntry
   * @property {string}  name       - Sheet name
   * @property {boolean} selected   - Whether the user has checked it
   */

  /**
   * @typedef {Object} WorkbookEntry
   * @property {string}       id        - Unique ID for this upload
   * @property {string}       filename  - Original file name
   * @property {File}         file      - Raw File object
   * @property {SheetEntry[]} sheets    - List of discovered sheets
   * @property {boolean}      expanded  - Whether the card is open
   */

  /**
   * Add a new workbook to the store.
   * @param {File} file
   * @param {string[]} sheetNames - Sheet names read from the file
   * @returns {string} The generated id
   */
  function add(file, sheetNames) {
    const id = `wb_${++_idCounter}`;
    _workbooks.set(id, {
      id,
      filename: file.name,
      file,
      sheets: sheetNames.map(name => ({ name, selected: true })),
      expanded: true,
    });
    return id;
  }

  /**
   * Remove a workbook from the store.
   * @param {string} id
   */
  function remove(id) {
    _workbooks.delete(id);
  }

  /**
   * Toggle a single sheet's selected state.
   * @param {string} wbId
   * @param {string} sheetName
   * @param {boolean} selected
   */
  function setSheetSelected(wbId, sheetName, selected) {
    const wb = _workbooks.get(wbId);
    if (!wb) return;
    const sheet = wb.sheets.find(s => s.name === sheetName);
    if (sheet) sheet.selected = selected;
  }

  /**
   * Select or deselect all sheets in a workbook.
   * @param {string} wbId
   * @param {boolean} selected
   */
  function setAllSelected(wbId, selected) {
    const wb = _workbooks.get(wbId);
    if (!wb) return;
    wb.sheets.forEach(s => (s.selected = selected));
  }

  /**
   * Toggle the expanded/collapsed state of a card.
   * @param {string} wbId
   */
  function toggleExpanded(wbId) {
    const wb = _workbooks.get(wbId);
    if (wb) wb.expanded = !wb.expanded;
  }

  /**
   * Returns a snapshot of all workbooks as an array.
   * @returns {WorkbookEntry[]}
   */
  function getAll() {
    return [..._workbooks.values()];
  }

  /**
   * Returns total number of selected sheets across all workbooks.
   * @returns {number}
   */
  function totalSelectedSheets() {
    let count = 0;
    _workbooks.forEach(wb => {
      count += wb.sheets.filter(s => s.selected).length;
    });
    return count;
  }

  /**
   * Returns true if at least one sheet is selected.
   * @returns {boolean}
   */
  function hasSelection() {
    for (const wb of _workbooks.values()) {
      if (wb.sheets.some(s => s.selected)) return true;
    }
    return false;
  }

  /**
   * Returns all selected (workbook, sheet) pairs ready for import.
   * @returns {{ wb: WorkbookEntry, sheet: SheetEntry }[]}
   */
  function getSelectedPairs() {
    const pairs = [];
    _workbooks.forEach(wb => {
      wb.sheets.forEach(sheet => {
        if (sheet.selected) pairs.push({ wb, sheet });
      });
    });
    return pairs;
  }

  return {
    add,
    remove,
    setSheetSelected,
    setAllSelected,
    toggleExpanded,
    getAll,
    totalSelectedSheets,
    hasSelection,
    getSelectedPairs,
    get size() { return _workbooks.size; },
  };
})();
