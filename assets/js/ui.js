/**
 * ui.js
 * -----
 * All DOM rendering and UI update logic.
 * Reads from WorkbookManager and refreshes the view.
 */

const UI = (() => {

  // ── DOM references (assigned in init) ───────────────────────────────────────
  let _listEl, _wbCountEl, _sheetCountEl, _importBtn,
      _progressWrap, _progressFill, _progressLabel, _progressBar,
      _resultMsg, _emptyState;

  /**
   * Cache DOM refs and wire up no-arg summary refresh.
   * Called once from taskpane.js after DOMContentLoaded.
   */
  function init() {
    _listEl         = document.getElementById('workbookList');
    _wbCountEl      = document.getElementById('wbCount');
    _sheetCountEl   = document.getElementById('sheetCount');
    _importBtn      = document.getElementById('importBtn');
    _progressWrap   = document.getElementById('progressWrap');
    _progressFill   = document.getElementById('progressFill');
    _progressLabel  = document.getElementById('progressLabel');
    _progressBar    = document.getElementById('progressBar');
    _resultMsg      = document.getElementById('resultMsg');
    _emptyState     = document.getElementById('emptyState');
  }

  // ── Card rendering ──────────────────────────────────────────────────────────

  /**
   * Fully re-render the workbook card list from WorkbookManager state.
   * Called after any state change that affects the list.
   */
  function renderCards() {
    const workbooks = WorkbookManager.getAll();

    // Clear existing cards
    _listEl.innerHTML = '';

    if (workbooks.length === 0) {
      _emptyState && (_emptyState.hidden = false);
    } else {
      _emptyState && (_emptyState.hidden = true);
      workbooks.forEach(wb => _listEl.appendChild(_buildCard(wb)));
    }

    _updateSummary();
  }

  /**
   * Build the DOM element for a single workbook card.
   * @param {WorkbookEntry} wb
   * @returns {HTMLElement}
   */
  function _buildCard(wb) {
    const card = document.createElement('div');
    card.className = 'workbook-card';
    card.dataset.wbId = wb.id;

    const selectedCount = wb.sheets.filter(s => s.selected).length;
    const totalCount    = wb.sheets.length;

    card.innerHTML = `
      <!-- Card header -->
      <div class="card-header">
        <span class="card-icon" aria-hidden="true">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
            <rect x="3" y="3" width="18" height="18" rx="2" fill="#107c10" opacity=".15"/>
            <path d="M8 8h8M8 12h8M8 16h5" stroke="#107c10" stroke-width="1.5" stroke-linecap="round"/>
          </svg>
        </span>
        <div class="card-title-wrap">
          <div class="card-filename" title="${_escHtml(wb.filename)}">${_escHtml(wb.filename)}</div>
          <div class="card-meta js-card-meta">${selectedCount} of ${totalCount} sheet${totalCount !== 1 ? 's' : ''} selected</div>
        </div>
        <div class="card-actions">
          <button
            class="toggle-btn ${wb.expanded ? 'expanded' : ''}"
            data-action="toggle"
            data-wb-id="${wb.id}"
            title="${wb.expanded ? 'Collapse' : 'Expand'}"
            aria-expanded="${wb.expanded}"
            aria-label="${wb.expanded ? 'Collapse' : 'Expand'} ${_escHtml(wb.filename)}"
          >
            <svg class="chevron" width="16" height="16" viewBox="0 0 16 16" fill="currentColor">
              <path d="M4.5 6.5l3.5 3.5 3.5-3.5" stroke="currentColor" stroke-width="1.5" fill="none" stroke-linecap="round" stroke-linejoin="round"/>
            </svg>
          </button>
          <button
            class="btn btn--danger btn--sm"
            data-action="remove"
            data-wb-id="${wb.id}"
            title="Remove workbook"
            aria-label="Remove ${_escHtml(wb.filename)}"
          >
            <svg width="14" height="14" viewBox="0 0 16 16" fill="none">
              <path d="M4 4l8 8M12 4l-8 8" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"/>
            </svg>
          </button>
        </div>
      </div>

      <!-- Card body (collapsible) -->
      <div class="card-body" ${wb.expanded ? '' : 'hidden'}>
        <div class="card-toolbar">
          <button class="btn btn--ghost btn--sm" data-action="select-all" data-wb-id="${wb.id}">Select All</button>
          <button class="btn btn--ghost btn--sm" data-action="deselect-all" data-wb-id="${wb.id}">Deselect All</button>
        </div>
        <ul class="sheet-list" role="group" aria-label="Sheets in ${_escHtml(wb.filename)}">
          ${wb.sheets.map(sheet => `
            <li class="sheet-item">
              <label style="display:flex;align-items:center;gap:8px;width:100%;cursor:pointer;">
                <input
                  type="checkbox"
                  class="sheet-check"
                  data-action="toggle-sheet"
                  data-wb-id="${wb.id}"
                  data-sheet="${_escHtml(sheet.name)}"
                  ${sheet.selected ? 'checked' : ''}
                  aria-label="${_escHtml(sheet.name)}"
                />
                <span class="sheet-name">${_escHtml(sheet.name)}</span>
              </label>
            </li>
          `).join('')}
        </ul>
      </div>
    `;

    return card;
  }

  // ── Summary bar ─────────────────────────────────────────────────────────────

  /**
   * Update the summary counts and import button state.
   */
  function _updateSummary() {
    const wbCount     = WorkbookManager.size;
    const sheetCount  = WorkbookManager.totalSelectedSheets();
    const hasAny      = WorkbookManager.hasSelection();

    _wbCountEl.textContent   = `${wbCount} workbook${wbCount !== 1 ? 's' : ''}`;
    _sheetCountEl.textContent = `${sheetCount} sheet${sheetCount !== 1 ? 's' : ''} selected`;
    _importBtn.disabled = !hasAny;
  }

  // ── Progress ─────────────────────────────────────────────────────────────────

  /**
   * Show and update the progress bar.
   * @param {number} done
   * @param {number} total
   * @param {string} label
   */
  function showProgress(done, total, label) {
    _resultMsg.hidden = true;
    _progressWrap.hidden = false;
    _progressLabel.textContent = label || `Importing ${total} sheet${total !== 1 ? 's' : ''}…`;

    const pct = total > 0 ? Math.round((done / total) * 100) : 0;
    _progressFill.style.width = `${pct}%`;
    _progressBar.setAttribute('aria-valuenow', pct);
  }

  /** Hide the progress bar. */
  function hideProgress() {
    _progressWrap.hidden = true;
    _progressFill.style.width = '0%';
    _progressBar.setAttribute('aria-valuenow', 0);
  }

  // ── Result message ───────────────────────────────────────────────────────────

  /**
   * Show a success or error result message.
   * @param {'success'|'error'} type
   * @param {string} message
   */
  function showResult(type, message) {
    _resultMsg.hidden = false;
    _resultMsg.className = `result-msg ${type}`;
    _resultMsg.textContent = message;
  }

  /** Hide the result message. */
  function hideResult() {
    _resultMsg.hidden = true;
  }

  // ── Utility ──────────────────────────────────────────────────────────────────

  /**
   * Escape HTML special characters.
   * @param {string} str
   * @returns {string}
   */
  function _escHtml(str) {
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  return {
    init,
    renderCards,
    showProgress,
    hideProgress,
    showResult,
    hideResult,
  };
})();
