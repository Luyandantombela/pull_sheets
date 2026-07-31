/**
 * taskpane.js
 * -----------
 * Entry point for the Sheet Importer add-in.
 * Initialises Office.js and wires up all UI event listeners.
 */

/* global Office, Excel, WorkbookManager, FileReader_, UI, Importer */

Office.onReady(({ host }) => {
  // Only run inside Excel
  if (host !== Office.HostType.Excel) {
    document.getElementById('app').innerHTML =
      '<p style="padding:16px;color:#a4262c;">This add-in is designed for Microsoft Excel only.</p>';
    return;
  }

  // Initialise UI module
  UI.init();
  UI.renderCards(); // show empty state initially

  // ── Element refs ──────────────────────────────────────────────────────────
  const dropZone  = document.getElementById('dropZone');
  const browseBtn = document.getElementById('browseBtn');
  const fileInput = document.getElementById('fileInput');
  const importBtn = document.getElementById('importBtn');

  // ── Browse button ─────────────────────────────────────────────────────────
  browseBtn.addEventListener('click', e => {
    e.stopPropagation();
    fileInput.click();
  });

  fileInput.addEventListener('change', () => {
    if (fileInput.files.length) {
      _handleFiles([...fileInput.files]);
      fileInput.value = ''; // allow re-uploading the same file
    }
  });

  // ── Drag and drop ─────────────────────────────────────────────────────────
  dropZone.addEventListener('click', e => {
    if (e.target === browseBtn) return; // browseBtn has its own handler
    fileInput.click();
  });

  dropZone.addEventListener('keydown', e => {
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      fileInput.click();
    }
  });

  dropZone.addEventListener('dragover', e => {
    e.preventDefault();
    dropZone.classList.add('drag-over');
  });

  ['dragleave', 'dragend'].forEach(evt =>
    dropZone.addEventListener(evt, () => dropZone.classList.remove('drag-over'))
  );

  dropZone.addEventListener('drop', e => {
    e.preventDefault();
    dropZone.classList.remove('drag-over');

    const excelFiles = [...e.dataTransfer.files].filter(f =>
      f.name.endsWith('.xlsx') || f.name.endsWith('.xlsm')
    );

    if (excelFiles.length === 0) {
      UI.showResult('error', 'Only .xlsx and .xlsm files are supported.');
      return;
    }
    _handleFiles(excelFiles);
  });

  // ── Delegated card events ─────────────────────────────────────────────────
  //  All card interactions bubble up to the workbook list container.
  document.getElementById('workbookList').addEventListener('click', _handleCardClick);
  document.getElementById('workbookList').addEventListener('change', _handleCardChange);

  // ── Import button ─────────────────────────────────────────────────────────
  importBtn.addEventListener('click', _handleImport);

  // ── File processing ───────────────────────────────────────────────────────

  /**
   * Process an array of File objects:
   * read sheet names → add to WorkbookManager → re-render.
   * @param {File[]} files
   */
  async function _handleFiles(files) {
    UI.hideResult();

    // Process each file
    for (const file of files) {
      try {
        const sheetNames = await FileReader_.getSheetNames(file);
        if (sheetNames.length === 0) {
          UI.showResult('error', `"${file.name}" contains no worksheets.`);
          continue;
        }
        WorkbookManager.add(file, sheetNames);
      } catch (err) {
        UI.showResult('error', `Failed to read "${file.name}": ${err.message}`);
      }
    }

    UI.renderCards();
  }

  // ── Card event handlers ───────────────────────────────────────────────────

  /**
   * Handle all click-based card actions via delegation.
   * @param {MouseEvent} e
   */
  function _handleCardClick(e) {
    const btn    = e.target.closest('[data-action]');
    if (!btn) return;

    const action = btn.dataset.action;
    const wbId   = btn.dataset.wbId;

    switch (action) {
      case 'toggle':
        WorkbookManager.toggleExpanded(wbId);
        UI.renderCards();
        break;

      case 'remove':
        WorkbookManager.remove(wbId);
        UI.hideResult();
        UI.renderCards();
        break;

      case 'select-all':
        WorkbookManager.setAllSelected(wbId, true);
        UI.renderCards();
        break;

      case 'deselect-all':
        WorkbookManager.setAllSelected(wbId, false);
        UI.renderCards();
        break;
    }
  }

  /**
   * Handle checkbox change events (individual sheet toggles).
   * @param {Event} e
   */
  function _handleCardChange(e) {
    const checkbox = e.target.closest('[data-action="toggle-sheet"]');
    if (!checkbox) return;

    const { wbId, sheet } = checkbox.dataset;
    WorkbookManager.setSheetSelected(wbId, sheet, checkbox.checked);

    // Update the meta line and summary without full re-render for performance
    const card      = checkbox.closest('.workbook-card');
    const metaEl    = card.querySelector('.js-card-meta');
    const wb        = WorkbookManager.getAll().find(w => w.id === wbId);
    if (wb && metaEl) {
      const selCount  = wb.sheets.filter(s => s.selected).length;
      const total     = wb.sheets.length;
      metaEl.textContent = `${selCount} of ${total} sheet${total !== 1 ? 's' : ''} selected`;
    }

    // Update summary bar
    const wbCount     = WorkbookManager.size;
    const sheetCount  = WorkbookManager.totalSelectedSheets();
    const hasAny      = WorkbookManager.hasSelection();
    document.getElementById('wbCount').textContent    = `${wbCount} workbook${wbCount !== 1 ? 's' : ''}`;
    document.getElementById('sheetCount').textContent = `${sheetCount} sheet${sheetCount !== 1 ? 's' : ''} selected`;
    document.getElementById('importBtn').disabled     = !hasAny;
  }

  // ── Import handler ────────────────────────────────────────────────────────

  /**
   * Run the import flow for all selected sheets.
   */
  async function _handleImport() {
    const pairs = WorkbookManager.getSelectedPairs();
    if (pairs.length === 0) return;

    const total = pairs.length;

    // Disable UI during import
    importBtn.disabled = true;
    UI.hideResult();
    UI.showProgress(0, total, `Importing ${total} sheet${total !== 1 ? 's' : ''}…`);

    try {
      const { imported, errors } = await Importer.importSheets(pairs, {
        onProgress: (done, total, label) => UI.showProgress(done, total, label),
      });

      // Show 100% briefly
      UI.showProgress(total, total, 'Finalising…');

      await _sleep(400); // let the progress bar animate to 100 %
      UI.hideProgress();

      if (errors.length === 0) {
        UI.showResult('success', `✓ Successfully imported ${imported} worksheet${imported !== 1 ? 's' : ''}.`);
      } else {
        const errMsg = errors.length < 3
          ? errors.join('; ')
          : `${errors.length} sheets failed.`;
        UI.showResult(
          imported > 0 ? 'success' : 'error',
          `Imported ${imported} sheet${imported !== 1 ? 's' : ''}. ${errors.length} error${errors.length !== 1 ? 's' : ''}: ${errMsg}`
        );
      }
    } catch (err) {
      UI.hideProgress();
      UI.showResult('error', `Import failed: ${err.message}`);
    } finally {
      importBtn.disabled = !WorkbookManager.hasSelection();
    }
  }

  /** Simple promise-based sleep. */
  function _sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }
});
