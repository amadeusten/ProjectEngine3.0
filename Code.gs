/**
 * @OnlyCurrentDoc
 * ProjectEngine 3.0
 * This script manages multiple production applications with a shared project management system.
 */

// ============================================================================
// MATERIALS SHEET
// ============================================================================

const materialsSheet = {
  NAME: 'Materials',

  getSheet: function() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(this.NAME);
    if (!sheet) {
      console.error(`Sheet '${this.NAME}' not found.`);
    }
    return sheet;
  },

  getData: function(category = 'FABRICATION') {
    const sheet = this.getSheet();
    if (!sheet) return [];
    const range = sheet.getRange('A2:P' + sheet.getLastRow());
    const values = range.getValues();

    const filteredValues = values.filter(row => {
      const name = row[1];
      const primaryCategory = row[4];
      return name && name.toString().trim() !== "" && primaryCategory && primaryCategory.toString().toUpperCase().includes(category);
    });

    if (category.includes('PRINT')) {
      return filteredValues.map(row => {
        const name = row[1].toString().trim();
        const type = row[6] ? row[6].toString().trim().toUpperCase() : 'SHEET';
        const width = parseFloat(row[7]) || 0;
        const length = parseFloat(row[8]) || 0;
        let sheetCost = row[9];
        if (sheetCost && typeof sheetCost === 'string') {
          sheetCost = parseFloat(sheetCost.replace(/[^0-9.-]+/g, "")) || 0;
        } else if (typeof sheetCost !== 'number') {
          sheetCost = 0;
        }
        let costLinFt = 0;
        if (type === 'ROLL' && length > 0) {
          costLinFt = sheetCost / length;
        }
        return [name, type, width, length, sheetCost, costLinFt];
      });
    } else {
      return filteredValues.map(row => {
        const name = row[1].toString().trim();
        let unitCost = row[9];
        if (unitCost && typeof unitCost === 'string') {
          unitCost = parseFloat(unitCost.replace(/[^0-9.-]+/g, "")) || 0;
        } else if (typeof unitCost !== 'number') {
          unitCost = 0;
        }
        return { name: name, unitCost: unitCost };
      });
    }
  }
};

// ============================================================================
// PERSONNEL SHEET
// ============================================================================

const personnelSheet = {
  NAME: 'Personnel',

  getSheet: function() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(this.NAME);
    if (!sheet) {
      console.error(`Sheet '${this.NAME}' not found.`);
    }
    return sheet;
  },

  getData: function() {
    const sheet = this.getSheet();
    if (!sheet) return [];
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    const values = sheet.getRange('A2:K' + lastRow).getValues();
    return values
      .filter(row => row[0] && row[0].toString().trim() !== '')
      .map(row => {
        const firstName  = row[0].toString().trim();
        const lastName   = row[1].toString().trim();
        const hourlyRate = typeof row[2] === 'number' ? row[2] : (parseFloat(row[2]) || 0);
        const dayRate    = typeof row[3] === 'number' ? row[3] : (parseFloat(row[3]) || 0);
        const skills     = row[4]  ? row[4].toString().trim()  : '';
        const email      = row[5]  ? row[5].toString().trim()  : '';
        const phone      = row[6]  ? row[6].toString().trim()  : '';
        const address    = row[7]  ? row[7].toString().trim()  : '';
        const city       = row[8]  ? row[8].toString().trim()  : '';
        const state      = row[9]  ? row[9].toString().trim()  : '';
        const note       = row[10] ? row[10].toString().trim() : '';
        const fullName   = `${firstName} ${lastName}`.trim();
        return [firstName, lastName, hourlyRate, dayRate, skills, email, phone, address, city, state, note, fullName];
      });
  }
};

// ============================================================================
// APPLICATION OBJECTS
// ============================================================================

const fabricationApp = {
  showDialog: function() {
    const htmlOutput = HtmlService.createHtmlOutputFromFile('FabricationIndex').setWidth(850).setHeight(850);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Fabrication Details');
  },

  getMaterials: function() {
    try { return materialsSheet.getData('FABRICATION'); }
    catch (e) { console.error("Error in fabricationApp.getMaterials: " + e.toString()); return []; }
  },

  getPersonnel: function() {
    try { return personnelSheet.getData(); }
    catch (e) { console.error("Error in fabricationApp.getPersonnel: " + e.toString()); return []; }
  },

  openForEdit: function(logId) {
    const formData = projectSheet.getLoggedFormData(logId, 'FabricationLog');
    if (formData) {
      const htmlOutput = HtmlService.createHtmlOutputFromFile('FabricationIndex').setWidth(850).setHeight(850);
      const htmlContent = htmlOutput.getContent();
      const modifiedContent = htmlContent.replace('<script>', `<script>window.editFormData = ${JSON.stringify(formData)};`);
      const modifiedOutput = HtmlService.createHtmlOutput(modifiedContent).setWidth(850).setHeight(850);
      SpreadsheetApp.getUi().showModalDialog(modifiedOutput, 'Edit Fabrication Details');
    } else {
      this.showDialog();
    }
  },

  addToProject: function(fabricationData) {
    try { return projectSheet.addProjectItem(fabricationData, 'FAB', 'FabricationLog'); }
    catch (e) {
      console.error("Error in fabricationApp.addToProject: " + e.toString());
      return { success: false, message: `Error adding to project: ${e.toString()}`, rowNumber: null, logId: null };
    }
  }
};

const apparelApp = {
  showDialog: function() {
    const htmlOutput = HtmlService.createHtmlOutputFromFile('ApparelIndex').setWidth(850).setHeight(850);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Apparel / Screen Printing');
  },

  openForEdit: function(logId) {
    const formData = projectSheet.getLoggedFormData(logId, 'ApparelLog');
    if (formData) {
      const htmlOutput = HtmlService.createHtmlOutputFromFile('ApparelIndex').setWidth(850).setHeight(850);
      const htmlContent = htmlOutput.getContent();
      const modifiedContent = htmlContent.replace('<script>', `<script>window.editFormData = ${JSON.stringify(formData)};`);
      const modifiedOutput = HtmlService.createHtmlOutput(modifiedContent).setWidth(850).setHeight(850);
      SpreadsheetApp.getUi().showModalDialog(modifiedOutput, 'Edit Apparel');
    } else {
      this.showDialog();
    }
  },

  addToProject: function(apparelData) {
    try { return projectSheet.addProjectItem(apparelData, 'APP', 'ApparelLog'); }
    catch (e) {
      console.error("Error in apparelApp.addToProject: " + e.toString());
      return { success: false, message: `Error adding to project: ${e.toString()}`, rowNumber: null, logId: null };
    }
  }
};

const printingApp = {
  showDialog: function() {
    const htmlOutput = HtmlService.createHtmlOutputFromFile('PrintingIndex').setWidth(850).setHeight(850);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'PrintCut Estimate');
  },

  getMaterials: function() {
    try { return materialsSheet.getData('PRINT'); }
    catch (e) { console.error("Error in printingApp.getMaterials: " + e.toString()); return []; }
  },

  openForEdit: function(logId) {
    const formData = projectSheet.getLoggedFormData(logId, 'PrintingLog');
    if (formData) {
      const htmlOutput = HtmlService.createHtmlOutputFromFile('PrintingIndex').setWidth(850).setHeight(850);
      const htmlContent = htmlOutput.getContent();
      const modifiedContent = htmlContent.replace('<script>', `<script>window.editFormData = ${JSON.stringify(formData)};`);
      const modifiedOutput = HtmlService.createHtmlOutput(modifiedContent).setWidth(750).setHeight(850);
      SpreadsheetApp.getUi().showModalDialog(modifiedOutput, 'Edit PrintCut Estimate');
    } else {
      this.showDialog();
    }
  },

  addToProject: function(printingData) {
    try { return projectSheet.addProjectItem(printingData, 'PRT', 'PrintingLog'); }
    catch (e) {
      console.error("Error in printingApp.addToProject: " + e.toString());
      return { success: false, message: `Error adding to project: ${e.toString()}`, rowNumber: null, logId: null };
    }
  }
};

// ============================================================================
// PROJECT SHEET — Core Data Management
// ============================================================================

const projectSheet = {
  getActiveSheet: function() {
    return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  },

  // ── Blueprint #1: LogID-Based Row Lookup ──────────────────────
  /**
   * Scans Column G notes to find the row containing a specific LogID.
   * Returns the 1-based row number, or null if not found.
   */
  findRowByLogId: function(logId) {
    if (!logId) return null;
    const sheet = this.getActiveSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;

    // Get all notes in Column G (column 7)
    const notes = sheet.getRange(2, 7, lastRow - 1, 1).getNotes();
    for (let i = 0; i < notes.length; i++) {
      const note = notes[i][0];
      if (note && note.includes('LogID:')) {
        const match = note.match(/LogID:\s*([^\n\r]+)/);
        if (match && match[1].trim() === logId) {
          return i + 2; // Convert 0-based array index to 1-based row number
        }
      }
    }
    return null;
  },

  // ── Blueprint #2: Shared Edit Cell Styling ────────────────────
  /**
   * Applies consistent styling to the Edit cell (Column G).
   * @param {Range} cell - The Column G cell range
   * @param {string} logId - The LogID to store in the note
   * @param {boolean} isMultiItem - Whether this is a multi-item PRT entry (header row)
   * @param {boolean} isUpdate - Whether this is an update vs. new add
   */
  _styleEditCell: function(cell, logId, isMultiItem, isUpdate) {
    const timestamp = new Date().toLocaleString();
    let noteText = `LogID: ${logId}\n\nTo edit this item:\n1. Select this cell\n2. Go to Production > Edit Selected Item`;
    if (isUpdate) {
      noteText += `\n\nLast updated: ${timestamp}`;
    }
    cell.setValue('Edit');
    cell.setNote(noteText);
    cell.setBackground('#e3f2fd');
    cell.setFontColor('#1976d2');
    cell.setFontWeight('bold');
  },

  /**
   * Styles a child row's Edit cell with a parent LogID reference.
   * Used for multi-item PRT child rows so cleanup can find them by parent.
   * @param {Range} cell - The Column G cell range on the child row
   * @param {string} parentLogId - The parent header row's LogID
   */
  _styleChildEditCell: function(cell, parentLogId) {
    cell.setValue('');
    cell.setNote(`ParentLogID: ${parentLogId}`);
    cell.setBackground(null);
    cell.setFontColor(null);
    cell.setFontWeight(null);
  },

  // ── Log Data Management ───────────────────────────────────────
  logFormData: function(formData, projectRowNumber, logIdPrefix, logSheetName) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      let logSheet = spreadsheet.getSheetByName(logSheetName);
      if (!logSheet) {
        logSheet = spreadsheet.insertSheet(logSheetName);
        logSheet.hideSheet();
        logSheet.getRange(1, 1, 1, 4).setValues([['LogID', 'ProjectRow', 'Timestamp', 'FormData']]);
      }
      const logId = `${logIdPrefix}_${Date.now()}_${projectRowNumber}`;
      const timestamp = new Date();
      const formDataWithRow = { ...formData, originalRowNumber: projectRowNumber, logId: logId };
      const formDataJson = JSON.stringify(formDataWithRow);
      const nextLogRow = logSheet.getLastRow() + 1;
      logSheet.getRange(nextLogRow, 1, 1, 4).setValues([[logId, projectRowNumber, timestamp, formDataJson]]);
      return logId;
    } catch (error) {
      console.error('Error logging form data:', error);
      return null;
    }
  },

  updateLogFormData: function(formData, projectRowNumber, logIdPrefix, logSheetName) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const logSheet = spreadsheet.getSheetByName(logSheetName);
      if (!logSheet) {
        return this.logFormData(formData, projectRowNumber, logIdPrefix, logSheetName);
      }
      const dataRange = logSheet.getDataRange();
      const values = dataRange.getValues();
      let existingRowIndex = -1;
      for (let i = 1; i < values.length; i++) {
        if (values[i][1] === projectRowNumber) {
          existingRowIndex = i + 1;
          break;
        }
      }
      const logId = `${logIdPrefix}_${Date.now()}_${projectRowNumber}`;
      const timestamp = new Date();
      const formDataWithRow = { ...formData, originalRowNumber: projectRowNumber, logId: logId };
      const formDataJson = JSON.stringify(formDataWithRow);
      if (existingRowIndex > 0) {
        logSheet.getRange(existingRowIndex, 1, 1, 4).setValues([[logId, projectRowNumber, timestamp, formDataJson]]);
      } else {
        const nextLogRow = logSheet.getLastRow() + 1;
        logSheet.getRange(nextLogRow, 1, 1, 4).setValues([[logId, projectRowNumber, timestamp, formDataJson]]);
      }
      return logId;
    } catch (error) {
      console.error('Error updating log data:', error);
      return null;
    }
  },

  getLoggedFormData: function(logId, logSheetName) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const logSheet = spreadsheet.getSheetByName(logSheetName);
      if (!logSheet) return null;
      const dataRange = logSheet.getDataRange();
      const values = dataRange.getValues();
      for (let i = 1; i < values.length; i++) {
        if (values[i][0] === logId) {
          return JSON.parse(values[i][3]);
        }
      }
      return null;
    } catch (error) {
      console.error('Error retrieving form data:', error);
      return null;
    }
  },

  createEditInstruction: function(logId) {
    return "Edit";
  },

  // ── ID Generators ─────────────────────────────────────────────
  getNextFabricationId: function() {
    const sheet = this.getActiveSheet();
    const values = sheet.getDataRange().getValues();
    let maxNumber = 0;
    for (let i = 1; i < values.length; i++) {
      const cellValue = values[i][1];
      if (cellValue && typeof cellValue === 'string' && cellValue.startsWith('F')) {
        const number = parseInt(cellValue.substring(1));
        if (!isNaN(number) && number > maxNumber) maxNumber = number;
      }
    }
    return `F${(maxNumber + 1).toString().padStart(2, '0')}`;
  },

  getNextApparelId: function() {
    const sheet = this.getActiveSheet();
    const values = sheet.getDataRange().getValues();
    let maxNumber = 0;
    for (let i = 1; i < values.length; i++) {
      const cellValue = values[i][1];
      if (cellValue && typeof cellValue === 'string' && cellValue.startsWith('AP')) {
        const number = parseInt(cellValue.substring(2));
        if (!isNaN(number) && number > maxNumber) maxNumber = number;
      }
    }
    return `AP${(maxNumber + 1).toString().padStart(2, '0')}`;
  },

  getNextPrintingId: function() {
    const sheet = this.getActiveSheet();
    const values = sheet.getDataRange().getValues();
    let maxNumber = 0;
    for (let i = 1; i < values.length; i++) {
      const cellValue = values[i][1];
      if (cellValue && typeof cellValue === 'string' && cellValue.startsWith('PR')) {
        const number = parseInt(cellValue.substring(2));
        if (!isNaN(number) && number > maxNumber) maxNumber = number;
      }
    }
    return `PR${(maxNumber + 1).toString().padStart(2, '0')}`;
  },

  // ── Blueprint #3: Multi-Item PRT Child Row Cleanup ────────────
  /**
   * Removes all child rows belonging to a parent LogID.
   * Scans Column G notes for `ParentLogID: <parentLogId>` and deletes those rows.
   * Scans from bottom up to avoid index shifting.
   */
  _cleanupChildRows: function(sheet, parentRowNum, parentLogId) {
    const lastRow = sheet.getLastRow();
    if (lastRow <= parentRowNum) return;

    // Scan all rows below the parent for child notes
    const numRowsBelow = lastRow - parentRowNum;
    const notes = sheet.getRange(parentRowNum + 1, 7, numRowsBelow, 1).getNotes();
    const colBValues = sheet.getRange(parentRowNum + 1, 2, numRowsBelow, 1).getValues();

    // Collect row numbers to delete (1-based), bottom-up
    const rowsToDelete = [];
    for (let i = 0; i < notes.length; i++) {
      const note = notes[i][0];
      const colB = colBValues[i][0];
      // Match by ParentLogID note OR by adjacency (blank Column B with no other LogID)
      if (note && note.includes(`ParentLogID: ${parentLogId}`)) {
        rowsToDelete.push(parentRowNum + 1 + i);
      }
    }

    // Delete from bottom up to preserve row indices
    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      sheet.deleteRow(rowsToDelete[i]);
    }
  },

  // ── Update Project Item ───────────────────────────────────────
  updateProjectItem: function(itemData, logIdPrefix, logSheetName) {
    try {
      const sheet = this.getActiveSheet();
      if (!itemData || typeof itemData !== 'object') {
        throw new Error('Invalid item data provided');
      }

      const { description, quantity, dimensions, totalPrice, formData, originalRowNumber } = itemData;

      // Blueprint #1: Try LogID-based lookup first, fall back to stored row number
      let rowNum = null;
      const storedLogId = formData && formData.logId ? formData.logId : null;
      if (storedLogId) {
        rowNum = this.findRowByLogId(storedLogId);
      }
      if (!rowNum) {
        rowNum = parseInt(originalRowNumber);
      }

      if (!rowNum || isNaN(rowNum) || rowNum < 1) {
        throw new Error(`Invalid row number for update: ${originalRowNumber}`);
      }

      const maxRows = sheet.getMaxRows();
      if (rowNum > maxRows) {
        throw new Error(`Row number ${rowNum} exceeds sheet maximum rows ${maxRows}`);
      }

      const logId = this.updateLogFormData(formData, rowNum, logIdPrefix, logSheetName);

      if (logIdPrefix === 'FAB') {
        const existingId = sheet.getRange(rowNum, 2).getValue() || this.getNextFabricationId();
        sheet.getRange(rowNum, 1, 1, 6).setValues([[
          '', existingId, description || '', dimensions || '', '', totalPrice || 0
        ]]);
        sheet.getRange(rowNum, 6).setNumberFormat('$#,##0.00');
        if (logId) this._styleEditCell(sheet.getRange(rowNum, 7), logId, false, true);

      } else if (logIdPrefix === 'APP') {
        const existingId = sheet.getRange(rowNum, 2).getValue() || this.getNextApparelId();
        const unitPrice = quantity && quantity > 0 ? (totalPrice / quantity) : 0;
        sheet.getRange(rowNum, 1, 1, 6).setValues([[
          '', existingId, description || '', quantity || '', unitPrice, totalPrice || 0
        ]]);
        sheet.getRange(rowNum, 5).setNumberFormat('$#,##0.00');
        sheet.getRange(rowNum, 6).setNumberFormat('$#,##0.00');
        if (logId) this._styleEditCell(sheet.getRange(rowNum, 7), logId, false, true);

      } else if (logIdPrefix === 'PRT') {
        // Blueprint #3: Multi-item PRT update with child row cleanup
        const existingId = sheet.getRange(rowNum, 2).getValue() || this.getNextPrintingId();

        // Clean up old child rows using parent LogID note approach
        if (storedLogId) {
          this._cleanupChildRows(sheet, rowNum, storedLogId);
        }

        // Write using multi-item scenarios
        this._writePrtRows(sheet, rowNum, existingId, itemData, logId, true);
      }

      return {
        success: true,
        message: `Item updated in row ${rowNum}`,
        rowNumber: rowNum,
        logId: logId,
        isUpdate: true
      };

    } catch (error) {
      console.error('Error updating project item:', error);
      return { success: false, message: `Error updating item: ${error.message}`, rowNumber: null, logId: null, isUpdate: false };
    }
  },

  // ── Add Project Item ──────────────────────────────────────────
  addProjectItem: function(itemData, logIdPrefix, logSheetName) {
    try {
      let originalRowNumber = null;
      if (itemData.originalRowNumber) {
        originalRowNumber = itemData.originalRowNumber;
      } else if (itemData.formData && itemData.formData.originalRowNumber) {
        originalRowNumber = itemData.formData.originalRowNumber;
      }

      if (originalRowNumber && originalRowNumber > 0) {
        return this.updateProjectItem({ ...itemData, originalRowNumber: originalRowNumber }, logIdPrefix, logSheetName);
      }

      const sheet = this.getActiveSheet();
      if (!itemData || typeof itemData !== 'object') {
        throw new Error('Invalid item data provided');
      }

      const { description, quantity, dimensions, totalPrice, formData } = itemData;
      const lastRow = sheet.getLastRow();
      const nextRow = lastRow + 1;

      const logId = this.logFormData(formData, nextRow, logIdPrefix, logSheetName);

      if (logIdPrefix === 'FAB') {
        const fabricationId = this.getNextFabricationId();
        sheet.getRange(nextRow, 1, 1, 7).setValues([[
          '', fabricationId, description || '', dimensions || '', '', totalPrice || 0, ''
        ]]);
        sheet.getRange(nextRow, 6).setNumberFormat('$#,##0.00');
        if (logId) this._styleEditCell(sheet.getRange(nextRow, 7), logId, false, false);

      } else if (logIdPrefix === 'APP') {
        const apparelId = this.getNextApparelId();
        const unitPrice = quantity && quantity > 0 ? (totalPrice / quantity) : 0;
        sheet.getRange(nextRow, 1, 1, 7).setValues([[
          '', apparelId, description || '', quantity || '', unitPrice, totalPrice || 0, ''
        ]]);
        sheet.getRange(nextRow, 5).setNumberFormat('$#,##0.00');
        sheet.getRange(nextRow, 6).setNumberFormat('$#,##0.00');
        if (logId) this._styleEditCell(sheet.getRange(nextRow, 7), logId, false, false);

      } else if (logIdPrefix === 'PRT') {
        // Blueprint #3: Multi-item PRT writing
        const printingId = this.getNextPrintingId();
        this._writePrtRows(sheet, nextRow, printingId, itemData, logId, false);
      }

      return {
        success: true,
        message: `Item added to row ${nextRow}`,
        rowNumber: nextRow,
        logId: logId,
        isUpdate: false
      };

    } catch (error) {
      console.error('Error adding project item:', error);
      return { success: false, message: `Error adding item: ${error.message}`, rowNumber: null, logId: null, isUpdate: false };
    }
  },

  // ── Blueprint #3: PRT Multi-Item Sheet Writer ─────────────────
  /**
   * Writes PRT rows to the sheet based on three scenarios:
   *   1. Single item: Description + dimensions, Qty, Unit Price, Total — no child rows
   *   2. Multi + Track Quantities ON: Header (description bold, no prices), children (name+dims, Qty, Unit Price, Total)
   *   3. Multi + Track Quantities OFF: Header (description bold, Total Qty, Grand Total), children (name+dims, Qty only)
   *
   * @param {Sheet} sheet - The active project sheet
   * @param {number} headerRow - Row number for the header/only row
   * @param {string} printingId - The PR## ID
   * @param {Object} itemData - Full item data payload
   * @param {string} logId - The LogID for note storage
   * @param {boolean} isUpdate - Whether this is an update
   */
  _writePrtRows: function(sheet, headerRow, printingId, itemData, logId, isUpdate) {
    const { description, totalPrice, formData } = itemData;
    const items = (formData && formData.items) || [];
    const isMultiItem = (itemData.isMultiItem || (formData && formData.isMultiItem)) && items.length > 1;
    const trackQuantities = itemData.trackQuantities || (formData && formData.trackQuantities);
    // itemPricing lives on itemData (top-level payload), not inside formData
    const itemPricing = itemData.itemPricing || [];

    if (!isMultiItem) {
      // ── Scenario 1: Single item ──
      const item = items[0] || {};
      const qty = Number(item.quantity) || 1;
      const dims = (item.width && item.height) ? `${item.width}" × ${item.height}"` : '';
      const descWithDims = dims ? `${description || ''} — ${dims}` : (description || '');
      const unitPrice = qty > 0 ? (totalPrice / qty) : 0;

      sheet.getRange(headerRow, 1, 1, 7).setValues([[
        '', printingId, descWithDims, qty, unitPrice, totalPrice || 0, ''
      ]]);
      sheet.getRange(headerRow, 5).setNumberFormat('$#,##0.00');
      sheet.getRange(headerRow, 6).setNumberFormat('$#,##0.00');
      if (logId) this._styleEditCell(sheet.getRange(headerRow, 7), logId, false, isUpdate);

    } else if (trackQuantities) {
      // ── Scenario 2: Multi + Track Quantities ON ──
      // Header row: Description bold, no prices
      sheet.getRange(headerRow, 1, 1, 7).setValues([[
        '', printingId, description || '', '', '', '', ''
      ]]);
      sheet.getRange(headerRow, 3).setFontWeight('bold');
      if (logId) this._styleEditCell(sheet.getRange(headerRow, 7), logId, true, isUpdate);

      // Insert child rows below header
      const childCount = items.length;
      if (childCount > 0) {
        sheet.insertRowsAfter(headerRow, childCount);
        for (let i = 0; i < childCount; i++) {
          const childRow = headerRow + 1 + i;
          const item = items[i];
          const dims = (item.width && item.height) ? `${item.width}" × ${item.height}"` : '';
          const childDesc = item.itemName ? `${item.itemName} — ${dims}` : dims;
          const qty = Number(item.quantity) || 0;

          // Get per-item pricing from top-level payload
          const pricing = itemPricing[i] || {};
          const childUnitPrice = pricing.unitPrice || 0;
          const childTotal = pricing.total || 0;

          sheet.getRange(childRow, 1, 1, 7).setValues([[
            '', '', `    ${childDesc}`, qty, childUnitPrice, childTotal, ''
          ]]);
          sheet.getRange(childRow, 5).setNumberFormat('$#,##0.00');
          sheet.getRange(childRow, 6).setNumberFormat('$#,##0.00');
          // Ensure child rows are not bold
          sheet.getRange(childRow, 1, 1, 7).setFontWeight('normal');
          // Mark child with parent LogID for safe cleanup
          if (logId) this._styleChildEditCell(sheet.getRange(childRow, 7), logId);
        }
      }

    } else {
      // ── Scenario 3: Multi + Track Quantities OFF ──
      // Header row: Description bold, Total Qty, Grand Total
      const totalQty = items.reduce((sum, it) => sum + (Number(it.quantity) || 0), 0);

      sheet.getRange(headerRow, 1, 1, 7).setValues([[
        '', printingId, description || '', totalQty, '', totalPrice || 0, ''
      ]]);
      sheet.getRange(headerRow, 3).setFontWeight('bold');
      sheet.getRange(headerRow, 6).setNumberFormat('$#,##0.00');
      if (logId) this._styleEditCell(sheet.getRange(headerRow, 7), logId, true, isUpdate);

      // Insert child rows below header — Qty only, no prices
      const childCount = items.length;
      if (childCount > 0) {
        sheet.insertRowsAfter(headerRow, childCount);
        for (let i = 0; i < childCount; i++) {
          const childRow = headerRow + 1 + i;
          const item = items[i];
          const dims = (item.width && item.height) ? `${item.width}" × ${item.height}"` : '';
          const childDesc = item.itemName ? `${item.itemName} — ${dims}` : dims;
          const qty = Number(item.quantity) || 0;

          sheet.getRange(childRow, 1, 1, 7).setValues([[
            '', '', `    ${childDesc}`, qty, '', '', ''
          ]]);
          // Ensure child rows are not bold
          sheet.getRange(childRow, 1, 1, 7).setFontWeight('normal');
          if (logId) this._styleChildEditCell(sheet.getRange(childRow, 7), logId);
        }
      }
    }
  }
};

// ============================================================================
// NICH DOCS FUNCTIONALITY
// ============================================================================

const nichDocs = {

  // ── Apparel Pricing Tiers (mirrors ApparelIndex.html) ─────────
  _apparelPricingTiers: [
    { min: 1, max: 12, prices: [0.00, 3.15, 3.46, 3.67, 3.94] },
    { min: 13, max: 35, prices: [0.00, 2.10, 2.41, 2.62, 2.77, 3.08] },
    { min: 36, max: 75, prices: [0.00, 1.57, 1.89, 2.10, 2.35, 2.56, 2.77, 2.98, 3.19] },
    { min: 76, max: 150, prices: [0.00, 1.38, 1.75, 1.96, 2.18, 2.46, 2.69, 2.92, 3.05] },
    { min: 151, max: 250, prices: [0.00, 1.01, 1.32, 1.42, 1.55, 1.66, 1.78, 1.92, 2.02] },
    { min: 251, max: 500, prices: [0.00, 0.89, 1.08, 1.12, 1.18, 1.29, 1.33, 1.43, 1.54] },
    { min: 501, max: 1000, prices: [0.00, 0.75, 0.87, 0.86, 0.91, 0.96, 1.01, 1.07, 1.12] },
    { min: 1001, max: 2500, prices: [0.00, 0.69, 0.75, 0.77, 0.84, 0.91, 0.96, 1.01, 1.07] },
    { min: 2501, max: 5000, prices: [0.00, 0.52, 0.57, 0.69, 0.68, 0.73, 0.78, 0.84, 0.89] }
  ],

  _getApparelPricingTier: function(quantity) {
    for (let tier of this._apparelPricingTiers) {
      if (quantity >= tier.min && quantity <= tier.max) return tier;
    }
    return this._apparelPricingTiers[this._apparelPricingTiers.length - 1];
  },

  _getUnitPrintPrice: function(colors, quantity) {
    const tier = this._getApparelPricingTier(quantity);
    if (colors >= tier.prices.length) return tier.prices[tier.prices.length - 1];
    return tier.prices[colors] || 0;
  },

  // Location upcharges
  _locationUpcharges: { 'sleeve': 0.25, 'pocket': 0.10, 'tag': 0, 'back-collar': 0 },

  _calculateLocationCost: function(colors, quantity, location) {
    if (colors <= 0) return 0;
    const unitPrintPrice = this._getUnitPrintPrice(colors, quantity);
    const locUpcharge = this._locationUpcharges[location] || 0;
    return (unitPrintPrice + locUpcharge) * quantity;
  },

  // ── Profit & Loss ─────────────────────────────────────────────
  createProfitLoss: function() {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const plSheet = spreadsheet.getSheetByName('PL');
      if (!plSheet) { SpreadsheetApp.getUi().alert('Error: PL sheet not found.'); return; }
      const mainSheet = spreadsheet.getSheetByName('Main');
      if (!mainSheet) { SpreadsheetApp.getUi().alert('Error: Main sheet not found.'); return; }

      const revenueData = this.collectRevenueFromMain(mainSheet);
      const costData = this.collectCostData();

      this.writeRevenueToPL(plSheet, revenueData);
      this.writeCostsToPL(plSheet, costData);

      SpreadsheetApp.getUi().alert('Profit & Loss data updated successfully!');
    } catch (error) {
      console.error('Error creating Profit & Loss:', error);
      SpreadsheetApp.getUi().alert('Error creating Profit & Loss: ' + error.message);
    }
  },

  // ── Blueprint #5: Revenue Guard ───────────────────────────────
  collectRevenueFromMain: function(mainSheet) {
    const revenueData = { printing: 0, fabrication: 0, apparel: 0 };
    const lastRow = mainSheet.getLastRow();
    if (lastRow < 2) return revenueData;

    const data = mainSheet.getRange(2, 2, lastRow - 1, 5).getValues();
    data.forEach(row => {
      const id = row[0];
      const totalPrice = parseFloat(row[4]) || 0;
      // Guard: skip blank or non-string IDs (e.g. child rows with blank Column B)
      if (!id || typeof id !== 'string') return;
      if (id.startsWith('PR')) revenueData.printing += totalPrice;
      else if (id.startsWith('F')) revenueData.fabrication += totalPrice;
      else if (id.startsWith('AP')) revenueData.apparel += totalPrice;
    });
    return revenueData;
  },

  writeRevenueToPL: function(plSheet, revenueData) {
    plSheet.getRange(8, 4).setValue(revenueData.printing);
    plSheet.getRange(9, 4).setValue(revenueData.fabrication);
    plSheet.getRange(10, 4).setValue(revenueData.apparel);
    plSheet.getRange(8, 4, 3, 1).setNumberFormat('$#,##0.00');
  },

  writeCostsToPL: function(plSheet, costData) {
    plSheet.getRange(3, 7).setValue(costData.printing.materials);
    plSheet.getRange(4, 7).setValue(costData.printing.ink);
    plSheet.getRange(5, 7).setValue(costData.printing.equipment);
    plSheet.getRange(6, 7).setValue(costData.printing.operator);
    plSheet.getRange(7, 7).setValue(costData.printing.design);
    plSheet.getRange(11, 7).setValue(costData.fabrication.materials);
    plSheet.getRange(12, 7).setValue(costData.fabrication.labor);
    plSheet.getRange(13, 7).setValue(costData.fabrication.components);
    plSheet.getRange(17, 7).setValue(costData.apparel.garments);
    plSheet.getRange(18, 7).setValue(costData.apparel.printing);
    plSheet.getRange(19, 7).setValue(costData.apparel.labor);
    [plSheet.getRange(3,7,5,1), plSheet.getRange(11,7,3,1), plSheet.getRange(17,7,3,1)].forEach(r => r.setNumberFormat('$#,##0.00'));
  },

  collectCostData: function() {
    const costData = {
      printing: { materials: 0, ink: 0, equipment: 0, design: 0, operator: 0 },
      fabrication: { materials: 0, labor: 0, components: 0 },
      apparel: { garments: 0, printing: 0, labor: 0 }
    };
    const printingCosts = this.getPrintingCosts();
    costData.printing.materials = printingCosts.materialCost + printingCosts.laminationCost;
    costData.printing.ink = printingCosts.inkCost;
    costData.printing.equipment = printingCosts.cuttingCost + printingCosts.equipmentCost;
    costData.printing.design = printingCosts.designCost;
    costData.printing.operator = printingCosts.operatorCost;

    const fabricationCosts = this.getFabricationCosts();
    costData.fabrication.materials = fabricationCosts.materialTotal;
    costData.fabrication.labor = fabricationCosts.personnelTotal;
    costData.fabrication.components = fabricationCosts.componentTotal;

    const apparelCosts = this.getApparelCosts();
    costData.apparel.garments = apparelCosts.garmentTotal;
    costData.apparel.printing = apparelCosts.totalPrintCosts + apparelCosts.screenSetupCosts;
    costData.apparel.labor = apparelCosts.additionalOptionsCosts;
    return costData;
  },

  // ── Printing Costs ────────────────────────────────────────────
  getPrintingCosts: function() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = spreadsheet.getSheetByName('PrintingLog');
    const costs = { materialCost:0, laminationCost:0, inkCost:0, cuttingCost:0, equipmentCost:0, designCost:0, operatorCost:0 };
    if (!logSheet) return costs;
    const values = logSheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      const formDataJson = values[i][3];
      if (!formDataJson) continue;
      try {
        const formData = JSON.parse(formDataJson);
        const calc = this.calculatePrintingCosts(formData);
        costs.materialCost += calc.materialCost;
        costs.laminationCost += calc.laminationCost;
        costs.inkCost += calc.inkCost;
        costs.cuttingCost += calc.cuttingCost;
        costs.equipmentCost += calc.equipmentCost;
        costs.designCost += calc.designCost;
        costs.operatorCost += calc.operatorCost;
      } catch (error) { console.error('Error parsing PrintingLog data:', error); }
    }
    return costs;
  },

  calculatePrintingCosts: function(formData) {
    const costs = { materialCost:0, laminationCost:0, inkCost:0, cuttingCost:0, equipmentCost:0, designCost:0, operatorCost:0 };
    const materialName = formData.materialName || '';
    if (!materialName) return costs;

    const material = this.getMaterialByName(materialName);
    if (!material) return costs;

    // Build items list: use formData.items[] if available, fall back to top-level fields
    let items = [];
    if (formData.items && Array.isArray(formData.items) && formData.items.length > 0) {
      items = formData.items.filter(it => Number(it.quantity) > 0 && Number(it.width) > 0 && Number(it.height) > 0);
    } else if (Number(formData.quantity) > 0 && Number(formData.width) > 0 && Number(formData.height) > 0) {
      items = [{ quantity: formData.quantity, width: formData.width, height: formData.height }];
    }
    if (items.length === 0) return costs;

    const bleed = 0.25, spacing = 0.25;
    let totalLinearFeet = 0, totalMatCost = 0;
    let totalSqFt = 0, totalPerimeter = 0;

    // Iterate over items (mirrors PrintingIndex.html calculateEstimate)
    for (const item of items) {
      const qty = Number(item.quantity);
      const aw = Number(item.width), ah = Number(item.height);
      const artW = aw + bleed, artH = ah + bleed;

      if (material.type === 'ROLL') {
        const rw = material.width;
        if (artW > rw && artH > rw) continue;
        const cP = Math.floor((rw + spacing) / (artW + spacing));
        const cL = Math.floor((rw + spacing) / (artH + spacing));
        const cols = Math.max(cP, cL, 1);
        const layoutH = cL > cP ? artW : artH;
        const rows = Math.ceil(qty / cols);
        const linIn = (rows * layoutH) + ((rows - 1) * spacing);
        totalLinearFeet += linIn / 12;
      } else if (material.type === 'SHEET') {
        const sw = material.width, sh = material.height, sc = material.costSheet;
        const fp = (artW <= sw && artH <= sh);
        const fl = (artW <= sh && artH <= sw);
        if (!fp && !fl) continue;
        const p1 = fp ? Math.floor((sw+spacing)/(artW+spacing)) * Math.floor((sh+spacing)/(artH+spacing)) : 0;
        const p2 = fl ? Math.floor((sw+spacing)/(artH+spacing)) * Math.floor((sh+spacing)/(artW+spacing)) : 0;
        const pps = Math.max(p1, p2, 1);
        const sheets = Math.ceil(qty / pps);
        totalMatCost += Math.max(sheets * sc, sc * 0.5);
      }

      let sqft = (artW * artH / 144) * qty;
      if (formData.doubleSided) sqft *= 2;
      totalSqFt += sqft;
      totalPerimeter += (aw * 2 + ah * 2) * qty;
    }

    // Roll material cost: applied once after summing all items
    if (material.type === 'ROLL') {
      totalMatCost = (totalLinearFeet + 2.5) * material.costLinFt;
    }
    costs.materialCost = totalMatCost;

    // Time-based calculations
    const printTimeHours = (totalSqFt / 0.83) / 60;
    let cutTimeHours = (totalPerimeter / 120) / 60;
    if (formData.complexShape) cutTimeHours *= 1.5;
    const ripTimeHours = (totalSqFt / 20.52) / 60;
    const printComputeTimeHours = (totalSqFt / 6.2) / 60;
    const totalProjectRunTimeHours = printTimeHours + cutTimeHours + ripTimeHours + printComputeTimeHours;

    const designTimeInHours = this.getTimeInHours(formData.designTime, formData.designTimeUnit);
    const laborDecalsTimeInHours = this.getTimeInHours(formData.laborDecalsTime, formData.laborDecalsTimeUnit);
    const laborFinishingTimeInHours = this.getTimeInHours(formData.laborFinishingTime, formData.laborFinishingTimeUnit);
    const laborInstallingTimeInHours = this.getTimeInHours(formData.laborInstallingTime, formData.laborInstallingTimeUnit);
    const manualOperatorTimeInHours = laborDecalsTimeInHours + laborFinishingTimeInHours + laborInstallingTimeInHours;

    costs.inkCost = totalSqFt * 0.165;
    costs.cuttingCost = cutTimeHours * 25.00;
    costs.designCost = (totalSqFt / 25) * 0.0625 * 60.00 + designTimeInHours * 60.00;
    if (formData.lamination) {
      costs.laminationCost = material.type === 'ROLL' ? totalLinearFeet * 1.02 : totalSqFt * 0.2267;
    }
    costs.equipmentCost = totalProjectRunTimeHours * 4.95;
    costs.operatorCost = (totalProjectRunTimeHours + manualOperatorTimeInHours) * 28.00;
    return costs;
  },

  getMaterialByName: function(materialName) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName('Materials');
    if (!sheet) return null;
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return null;
    const data = sheet.getRange('A2:P' + lastRow).getValues();
    for (let row of data) {
      const name = row[1];
      const primaryCategory = row[4];
      if (name && name.toString().trim() === materialName &&
          primaryCategory && primaryCategory.toString().toUpperCase().includes('PRINT')) {
        const type = row[6] ? row[6].toString().trim().toUpperCase() : 'SHEET';
        const width = parseFloat(row[7]) || 0;
        const length = parseFloat(row[8]) || 0;
        let unitCost = row[9];
        if (unitCost && typeof unitCost === 'string') {
          unitCost = parseFloat(unitCost.replace(/[^0-9.-]+/g, '')) || 0;
        } else if (typeof unitCost !== 'number') { unitCost = 0; }
        let costLinFt = 0;
        if (type === 'ROLL' && length > 0) {
          costLinFt = unitCost / length;
        }
        return { name: name.toString().trim(), type, width, height: length, costSheet: unitCost, costLinFt };
      }
    }
    return null;
  },

  getTimeInHours: function(time, unit) {
    const numTime = Number(time) || 0;
    return unit === 'Minutes' ? numTime / 60 : numTime;
  },

  // ── Fabrication Costs ─────────────────────────────────────────
  getFabricationCosts: function() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = spreadsheet.getSheetByName('FabricationLog');
    const costs = { materialTotal: 0, personnelTotal: 0, componentTotal: 0 };
    if (!logSheet) return costs;
    const values = logSheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      const formDataJson = values[i][3];
      if (!formDataJson) continue;
      try {
        const formData = JSON.parse(formDataJson);
        if (formData.materials && Array.isArray(formData.materials)) {
          formData.materials.forEach(m => { costs.materialTotal += m.total || 0; });
        }
        if (formData.personnel && Array.isArray(formData.personnel)) {
          formData.personnel.forEach(p => { costs.personnelTotal += p.total || 0; });
        }
        if (formData.components && Array.isArray(formData.components)) {
          formData.components.forEach(c => { costs.componentTotal += c.total || 0; });
        }
      } catch (error) { console.error('Error parsing FabricationLog data:', error); }
    }
    return costs;
  },

  // ── Apparel Costs (now using real pricing tiers) ──────────────
  getApparelCosts: function() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = spreadsheet.getSheetByName('ApparelLog');
    const costs = { garmentTotal: 0, totalPrintCosts: 0, screenSetupCosts: 0, additionalOptionsCosts: 0 };
    if (!logSheet) return costs;
    const values = logSheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      const formDataJson = values[i][3];
      if (!formDataJson) continue;
      try {
        const formData = JSON.parse(formDataJson);
        const calc = this.calculateApparelCosts(formData);
        costs.garmentTotal += calc.garmentTotal;
        costs.totalPrintCosts += calc.totalPrintCosts;
        costs.screenSetupCosts += calc.screenSetupCosts;
        costs.additionalOptionsCosts += calc.additionalOptionsCosts;
      } catch (error) { console.error('Error parsing ApparelLog data:', error); }
    }
    return costs;
  },

  /**
   * Calculates apparel costs using the REAL pricing tier table (mirrors ApparelIndex.html).
   */
  calculateApparelCosts: function(formData) {
    const costs = { garmentTotal: 0, totalPrintCosts: 0, screenSetupCosts: 0, additionalOptionsCosts: 0 };
    const quantity = Number(formData.quantity) || 0;
    const garmentUnitCost = Number(formData.garmentUnitCost) || 0;
    costs.garmentTotal = quantity * garmentUnitCost;

    const frontColors = Number(formData.frontColors) || 0;
    const backColors = Number(formData.backColors) || 0;
    let totalColors = frontColors + backColors;

    // Front + Back print costs using real tiers
    if (frontColors > 0) costs.totalPrintCosts += this._calculateLocationCost(frontColors, quantity, 'front');
    if (backColors > 0) costs.totalPrintCosts += this._calculateLocationCost(backColors, quantity, 'back');

    // Additional locations
    if (formData.additionalLocations && Array.isArray(formData.additionalLocations)) {
      formData.additionalLocations.forEach(loc => {
        const locColors = Number(loc.colors) || 0;
        totalColors += locColors;
        if (locColors > 0) {
          costs.totalPrintCosts += this._calculateLocationCost(locColors, quantity, loc.location || '');
        }
      });
    }

    // Screen setup: $13 per color
    costs.screenSetupCosts = totalColors * 13;

    // Additional options
    const constants = { oversize:1, colorChange:5, polyNylonSpandexMesh:0.25, metallicShimmerInk:0.50, glow:1.25, fleece:0.25, designLabor:60 };
    if (formData.additionalOptions) {
      const opts = formData.additionalOptions;
      if (opts.oversized) costs.additionalOptionsCosts += constants.oversize * quantity;
      if (opts.colorChange && opts.colorChange.enabled) {
        costs.additionalOptionsCosts += constants.colorChange * (Number(opts.colorChange.quantity) || 1);
      }
      if (opts.polyNylon) costs.additionalOptionsCosts += constants.polyNylonSpandexMesh * quantity;
      if (opts.metallicShimmer) costs.additionalOptionsCosts += constants.metallicShimmerInk * quantity;
      if (opts.glow) costs.additionalOptionsCosts += constants.glow * quantity;
      if (opts.fleece) costs.additionalOptionsCosts += constants.fleece * quantity;
      if (opts.designLabor && opts.designLabor.enabled) {
        const laborQty = Number(opts.designLabor.quantity) || 0;
        let hours = laborQty;
        if ((opts.designLabor.unit || 'hours') === 'minutes') hours = laborQty / 60;
        costs.additionalOptionsCosts += constants.designLabor * hours;
      }
    }
    return costs;
  },

  // ── BOM ────────────────────────────────────────────────────────
  createBillOfMaterials: function() {
    try {
      const ui = SpreadsheetApp.getUi();
      const TEMPLATE_ID = '1mAbdWVwLGn8v146oqmNVxxQyYO1zX3CtAUN9wXj6rWw';
      const bomData = this.collectBOMData();
      if (bomData.printing.length === 0 && bomData.fabrication.length === 0 && bomData.apparel.length === 0) {
        ui.alert('No Data Found', 'No items found in any category.', ui.ButtonSet.OK);
        return;
      }
      const projectInfo = this.getProjectInfo();
      const docUrl = this.createBOMDocument(TEMPLATE_ID, bomData, projectInfo);
      const response = ui.alert('Bill of Materials Created', 'Your BOM has been created. Open it now?', ui.ButtonSet.YES_NO);
      if (response === ui.Button.YES) {
        ui.showModelessDialog(HtmlService.createHtmlOutput(`<script>window.open("${docUrl}", "_blank"); google.script.host.close();</script>`), 'Opening Document...');
      }
    } catch (error) {
      console.error('Error creating BOM:', error);
      SpreadsheetApp.getUi().alert('Error creating BOM: ' + error.message);
    }
  },

  collectBOMData: function() {
    return {
      printing: this.collectPrintingBOM(),
      fabrication: this.collectFabricationBOM(),
      apparel: this.collectApparelBOM()
    };
  },

  collectPrintingBOM: function() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = spreadsheet.getSheetByName('PrintingLog');
    if (!logSheet) return [];
    const materialMap = new Map();
    let totalOperatorHours = 0, totalDesignHours = 0;
    const values = logSheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      const formDataJson = values[i][3];
      if (!formDataJson) continue;
      try {
        const formData = JSON.parse(formDataJson);
        const materialName = formData.materialName;
        if (materialName) {
          const materialInfo = this.calculatePrintingMaterialQuantity(formData);
          if (materialInfo && materialInfo.quantity > 0) {
            if (materialMap.has(materialName)) {
              materialMap.get(materialName).quantity += materialInfo.quantity;
            } else {
              materialMap.set(materialName, { description: materialName, quantity: materialInfo.quantity, unitOfMeasure: materialInfo.unitOfMeasure, vendor: materialInfo.vendor || '', status: '' });
            }
          }
        }
        const laborHours = this.calculatePrintingLaborHours(formData);
        totalOperatorHours += laborHours.operatorHours;
        totalDesignHours += laborHours.designHours;
      } catch (error) { console.error('Error parsing PrintingLog data:', error); }
    }
    const items = Array.from(materialMap.values());
    if (totalOperatorHours > 0) items.push({ description: 'Operator Labor', quantity: Math.round(totalOperatorHours * 100) / 100, unitOfMeasure: 'Hours', vendor: '', status: '' });
    if (totalDesignHours > 0) items.push({ description: 'Design Labor', quantity: Math.round(totalDesignHours * 100) / 100, unitOfMeasure: 'Hours', vendor: '', status: '' });
    return items;
  },

  calculatePrintingLaborHours: function(formData) {
    // Build items list: use formData.items[] if available, fall back to top-level fields
    let items = [];
    if (formData.items && Array.isArray(formData.items) && formData.items.length > 0) {
      items = formData.items.filter(it => Number(it.quantity) > 0 && Number(it.width) > 0 && Number(it.height) > 0);
    } else if (Number(formData.quantity) > 0 && Number(formData.width) > 0 && Number(formData.height) > 0) {
      items = [{ quantity: formData.quantity, width: formData.width, height: formData.height }];
    }
    if (items.length === 0) return { operatorHours: 0, designHours: 0 };

    const bleed = 0.25;
    let totalSqFt = 0, totalPerimeter = 0;

    for (const item of items) {
      const qty = Number(item.quantity), aw = Number(item.width), ah = Number(item.height);
      const artW = aw + bleed, artH = ah + bleed;
      let sqft = (artW * artH / 144) * qty;
      if (formData.doubleSided) sqft *= 2;
      totalSqFt += sqft;
      totalPerimeter += (aw * 2 + ah * 2) * qty;
    }

    const printTimeHours = (totalSqFt / 0.83) / 60;
    let cutTimeHours = (totalPerimeter / 120) / 60;
    if (formData.complexShape) cutTimeHours *= 1.5;
    const ripTimeHours = (totalSqFt / 20.52) / 60;
    const printComputeTimeHours = (totalSqFt / 6.2) / 60;
    const laborDecalsTimeInHours = this.getTimeInHours(formData.laborDecalsTime, formData.laborDecalsTimeUnit);
    const laborFinishingTimeInHours = this.getTimeInHours(formData.laborFinishingTime, formData.laborFinishingTimeUnit);
    const laborInstallingTimeInHours = this.getTimeInHours(formData.laborInstallingTime, formData.laborInstallingTimeUnit);
    const manualOperatorTimeInHours = laborDecalsTimeInHours + laborFinishingTimeInHours + laborInstallingTimeInHours;
    const totalProjectRunTimeHours = printTimeHours + cutTimeHours + ripTimeHours + printComputeTimeHours;
    const designTimeInHours = this.getTimeInHours(formData.designTime, formData.designTimeUnit);
    const baseDesignHours = (totalSqFt / 25) * 0.0625;
    return {
      operatorHours: totalProjectRunTimeHours + manualOperatorTimeInHours,
      designHours: baseDesignHours + designTimeInHours
    };
  },

  calculatePrintingMaterialQuantity: function(formData) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const matSheet = spreadsheet.getSheetByName('Materials');
    if (!matSheet) return null;
    const materialName = formData.materialName;

    // Build items list: use formData.items[] if available, fall back to top-level fields
    let items = [];
    if (formData.items && Array.isArray(formData.items) && formData.items.length > 0) {
      items = formData.items.filter(it => Number(it.quantity) > 0 && Number(it.width) > 0 && Number(it.height) > 0);
    } else if (Number(formData.quantity) > 0 && Number(formData.width) > 0 && Number(formData.height) > 0) {
      items = [{ quantity: formData.quantity, width: formData.width, height: formData.height }];
    }
    if (items.length === 0) return null;

    const bleed = 0.25, spacing = 0.25;

    // Find material in Materials sheet
    const lastRow = matSheet.getLastRow();
    const materialsData = matSheet.getRange('A2:P' + lastRow).getValues();
    let materialType = 'SHEET', materialWidth = 0, materialHeight = 0, vendor = '';
    for (let row of materialsData) {
      if (row[1] && row[1].toString().trim() === materialName) {
        materialType = row[6] ? row[6].toString().trim().toUpperCase() : 'SHEET';
        materialWidth = parseFloat(row[7]) || 0;
        materialHeight = parseFloat(row[8]) || 0;
        vendor = row[11] || '';
        break;
      }
    }

    let totalQuantity = 0;
    let unitOfMeasure = '';

    if (materialType === 'ROLL') {
      let totalLinFt = 0;
      const rw = materialWidth;
      for (const item of items) {
        const qty = Number(item.quantity), aw = Number(item.width), ah = Number(item.height);
        const artW = aw + bleed, artH = ah + bleed;
        const cP = Math.floor((rw + spacing) / (artW + spacing));
        const cL = Math.floor((rw + spacing) / (artH + spacing));
        const cols = Math.max(cP, cL, 1);
        const rowH = cL > cP ? artW : artH;
        const numRows = Math.ceil(qty / cols);
        const linIn = (numRows * rowH) + ((numRows - 1) * spacing);
        totalLinFt += linIn / 12;
      }
      totalQuantity = totalLinFt + 2.5; // Buffer applied once
      unitOfMeasure = 'Lin Feet';
    } else {
      const sw = materialWidth, sh = materialHeight;
      for (const item of items) {
        const qty = Number(item.quantity), aw = Number(item.width), ah = Number(item.height);
        const artW = aw + bleed, artH = ah + bleed;
        const fp = (artW <= sw && artH <= sh);
        const fl = (artW <= sh && artH <= sw);
        const p1 = fp ? Math.floor((sw+spacing)/(artW+spacing)) * Math.floor((sh+spacing)/(artH+spacing)) : 0;
        const p2 = fl ? Math.floor((sw+spacing)/(artH+spacing)) * Math.floor((sh+spacing)/(artW+spacing)) : 0;
        const pps = Math.max(p1, p2, 1);
        totalQuantity += Math.ceil(qty / pps);
      }
      unitOfMeasure = 'Sheets';
    }

    return { quantity: Math.ceil(totalQuantity), unitOfMeasure, vendor };
  },

  collectFabricationBOM: function() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = spreadsheet.getSheetByName('FabricationLog');
    if (!logSheet) return [];
    const materialMap = new Map(), personnelMap = new Map(), componentMap = new Map();
    const values = logSheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      const formDataJson = values[i][3];
      if (!formDataJson) continue;
      try {
        const formData = JSON.parse(formDataJson);
        if (formData.materials && Array.isArray(formData.materials)) {
          formData.materials.forEach(m => {
            const key = m.name, qty = Number(m.quantity)||0;
            if (materialMap.has(key)) { materialMap.get(key).quantity += qty; }
            else { materialMap.set(key, { description: m.name, quantity: qty, unitOfMeasure: 'Units', vendor: this.getMaterialVendor(m.name), status: '' }); }
          });
        }
        if (formData.personnel && Array.isArray(formData.personnel)) {
          formData.personnel.forEach(p => {
            const key = p.name, totalHours = (Number(p.days)||0) * (Number(p.hours)||0);
            if (personnelMap.has(key)) { personnelMap.get(key).quantity += totalHours; }
            else { personnelMap.set(key, { description: p.name, quantity: totalHours, unitOfMeasure: 'Hours', vendor: '', status: '' }); }
          });
        }
        if (formData.components && Array.isArray(formData.components)) {
          formData.components.forEach(c => {
            const key = c.description, qty = Number(c.quantity)||0;
            if (componentMap.has(key)) { componentMap.get(key).quantity += qty; }
            else { componentMap.set(key, { description: c.description, quantity: qty, unitOfMeasure: 'Units', vendor: '', status: '' }); }
          });
        }
      } catch (error) { console.error('Error parsing FabricationLog data:', error); }
    }
    return [...Array.from(materialMap.values()), ...Array.from(personnelMap.values()), ...Array.from(componentMap.values())];
  },

  getMaterialVendor: function(materialName) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const matSheet = spreadsheet.getSheetByName('Materials');
    if (!matSheet) return '';
    const data = matSheet.getRange('A2:P' + matSheet.getLastRow()).getValues();
    for (let row of data) {
      if (row[1] && row[1].toString().trim() === materialName) return row[11] || '';
    }
    return '';
  },

  collectApparelBOM: function() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = spreadsheet.getSheetByName('ApparelLog');
    if (!logSheet) return [];
    const garmentMap = new Map();
    const values = logSheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      const formDataJson = values[i][3];
      if (!formDataJson) continue;
      try {
        const formData = JSON.parse(formDataJson);
        const garmentName = formData.garment;
        const quantity = Number(formData.quantity) || 0;
        if (garmentName && quantity > 0) {
          if (garmentMap.has(garmentName)) { garmentMap.get(garmentName).quantity += quantity; }
          else { garmentMap.set(garmentName, { description: garmentName, quantity, unitOfMeasure: 'Units', vendor: '', status: '' }); }
        }
      } catch (error) { console.error('Error parsing ApparelLog data:', error); }
    }
    return Array.from(garmentMap.values());
  },

  getProjectInfo: function() {
    const ui = SpreadsheetApp.getUi();
    const projectNumber = ui.prompt('Project Number', 'Enter the project number (e.g., 19171LV):', ui.ButtonSet.OK_CANCEL);
    if (projectNumber.getSelectedButton() !== ui.Button.OK) throw new Error('BOM creation cancelled by user');
    const clientName = ui.prompt('Client Name', 'Enter the client name:', ui.ButtonSet.OK_CANCEL);
    if (clientName.getSelectedButton() !== ui.Button.OK) throw new Error('BOM creation cancelled by user');
    return {
      projectNumber: projectNumber.getResponseText(),
      clientName: clientName.getResponseText(),
      date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM dd, yyyy')
    };
  },

  createBOMDocument: function(templateId, bomData, projectInfo) {
    const templateFile = DriveApp.getFileById(templateId);
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const newFile = templateFile.makeCopy(`BOM - ${projectInfo.projectNumber} - ${timestamp}`);
    const doc = DocumentApp.openById(newFile.getId());
    const body = doc.getBody();
    body.replaceText('\\{\\{projectNumber\\}\\}', projectInfo.projectNumber);
    body.replaceText('\\{\\{clientName\\}\\}', projectInfo.clientName);
    body.replaceText('\\{\\{date\\}\\}', projectInfo.date);
    this.insertBOMSection(body, 'PRINTING', bomData.printing);
    this.insertBOMSection(body, 'FABRICATION', bomData.fabrication);
    this.insertBOMSection(body, 'APPAREL', bomData.apparel);
    doc.saveAndClose();
    return newFile.getUrl();
  },

  insertBOMSection: function(body, sectionName, items) {
    const placeholder = `{{${sectionName}_TABLE}}`;
    if (items.length === 0) { this.removeBOMSection(body, sectionName, placeholder); return; }
    const placeholderSearch = body.findText(placeholder);
    if (!placeholderSearch) return;
    const element = placeholderSearch.getElement();
    const parent = element.getParent();
    const index = body.getChildIndex(parent);
    parent.removeFromParent();
    const tableData = [['Description', 'Quantity', 'Unit of Measure', 'Vendor', 'Status']];
    items.forEach(item => { tableData.push([item.description, item.quantity.toString(), item.unitOfMeasure, item.vendor, item.status]); });
    const table = body.insertTable(index, tableData);
    this.styleBOMTable(table);
  },

  removeBOMSection: function(body, sectionName, placeholder) {
    const placeholderSearch = body.findText(placeholder);
    if (placeholderSearch) placeholderSearch.getElement().getParent().removeFromParent();
    const headingVariations = [sectionName, sectionName.charAt(0) + sectionName.slice(1).toLowerCase(), sectionName.toLowerCase()];
    for (let heading of headingVariations) {
      const headingSearch = body.findText(heading);
      if (headingSearch) {
        const parent = headingSearch.getElement().getParent();
        if (parent.asText().getText().trim() === heading) { parent.removeFromParent(); break; }
      }
    }
  },

  styleBOMTable: function(table) {
    const numRows = table.getNumRows();
    const headerRow = table.getRow(0);
    const numCols = headerRow.getNumCells();
    const DARK_GRAY_4 = '#434343', BLACK = '#000000', DARK_GRAY_2 = '#666666';
    const totalWidth = 550;
    const columnWidths = [50, 10, 15, 15, 10];
    const columnWidthsInPoints = columnWidths.map(p => (p / 100) * totalWidth);
    for (let i = 0; i < numCols && i < columnWidthsInPoints.length; i++) {
      headerRow.getCell(i).setWidth(columnWidthsInPoints[i]);
    }
    for (let i = 0; i < numCols; i++) {
      const cell = headerRow.getCell(i);
      cell.setBackgroundColor(null);
      const text = cell.getChild(0).asText();
      text.setFontFamily('Inter'); text.setFontSize(9); text.setBold(true); text.setForegroundColor(DARK_GRAY_4);
      cell.setPaddingTop(6); cell.setPaddingBottom(6); cell.setPaddingLeft(8); cell.setPaddingRight(8);
    }
    for (let i = 1; i < numRows; i++) {
      const row = table.getRow(i);
      for (let j = 0; j < numCols; j++) {
        const cell = row.getCell(j);
        const text = cell.getChild(0).asText();
        text.setFontFamily('Inter'); text.setFontSize(9); text.setBold(false); text.setForegroundColor(BLACK);
        cell.setBackgroundColor(null);
        cell.setPaddingTop(4); cell.setPaddingBottom(4); cell.setPaddingLeft(8); cell.setPaddingRight(8);
      }
    }
    table.setBorderWidth(1); table.setBorderColor(DARK_GRAY_2);
  },

  createEstimate: function() { SpreadsheetApp.getUi().alert('Create Estimate functionality coming soon!'); },
  createInvoice: function() { SpreadsheetApp.getUi().alert('Create Invoice functionality coming soon!'); }
};

// ============================================================================
// MENU CREATION
// ============================================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Production')
    .addItem('Fabrication', 'openFabricationApp')
    .addItem('Apparel', 'openApparelApp')
    .addItem('Printing', 'openPrintingApp')
    .addSeparator()
    .addItem('Edit Selected Item', 'editSelectedItem')
    .addItem('Delete Selected Block', 'deleteSelectedBlock')
    .addToUi();
  ui.createMenu('NICH Docs')
    .addItem('Create Estimate', 'createEstimate')
    .addItem('Create Invoice', 'createInvoice')
    .addItem('Create Profit & Loss', 'createProfitLoss')
    .addItem('Create Bill of Materials', 'createBillOfMaterials')
    .addToUi();
}

// ============================================================================
// GLOBAL FUNCTION WRAPPERS
// ============================================================================

function openFabricationApp() { fabricationApp.showDialog(); }
function openApparelApp() { apparelApp.showDialog(); }
function openPrintingApp() { printingApp.showDialog(); }
function getMaterials() { return fabricationApp.getMaterials(); }
function getPersonnel() { return fabricationApp.getPersonnel(); }
function getPrintingMaterials() { return printingApp.getMaterials(); }

function calculateDeliveryRoute(origin, destination) {
  try {
    const directions = Maps.newDirectionFinder()
      .setOrigin(origin).setDestination(destination)
      .setMode(Maps.DirectionFinder.Mode.DRIVING).getDirections();
    if (!directions || !directions.routes || directions.routes.length === 0) {
      return { error: 'No route found between these addresses.' };
    }
    const leg = directions.routes[0].legs[0];
    return { miles: leg.distance.value / 1609.344, minutes: leg.duration.value / 60 };
  } catch (e) {
    console.error('calculateDeliveryRoute error:', e.toString());
    return { error: e.message || 'Route calculation failed.' };
  }
}

function addFabricationToProject(fabricationData) { return fabricationApp.addToProject(fabricationData); }
function addApparelToProject(apparelData) { return apparelApp.addToProject(apparelData); }
function addPrintingToProject(printingData) { return printingApp.addToProject(printingData); }
function openFabricationAppForEdit(logId) { return fabricationApp.openForEdit(logId); }
function openApparelAppForEdit(logId) { return apparelApp.openForEdit(logId); }
function openPrintingAppForEdit(logId) { return printingApp.openForEdit(logId); }
function getLoggedFormData(logId, logSheetName) { return projectSheet.getLoggedFormData(logId, logSheetName); }

function editSelectedItem() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const activeCell = sheet.getActiveCell();
    if (activeCell.getColumn() !== 7) {
      SpreadsheetApp.getUi().alert('Edit Item', 'Please select an "Edit" cell first, then try again.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    const cellNote = activeCell.getNote();
    if (!cellNote || !cellNote.includes('LogID:')) {
      SpreadsheetApp.getUi().alert('Edit Item', 'No edit data found for this item.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    const logIdMatch = cellNote.match(/LogID:\s*([^\n\r]+)/);
    if (!logIdMatch) {
      SpreadsheetApp.getUi().alert('Edit Item', 'Could not find LogID in the selected cell.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    const logId = logIdMatch[1].trim();
    if (logId.startsWith('FAB_')) fabricationApp.openForEdit(logId);
    else if (logId.startsWith('APP_')) apparelApp.openForEdit(logId);
    else if (logId.startsWith('PRT_')) printingApp.openForEdit(logId);
    else SpreadsheetApp.getUi().alert('Edit Item', 'Unknown item type.', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    console.error('Error in editSelectedItem:', error);
    SpreadsheetApp.getUi().alert('Error', 'An error occurred: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function createEstimate() { nichDocs.createEstimate(); }
function createInvoice() { nichDocs.createInvoice(); }
function createProfitLoss() { nichDocs.createProfitLoss(); }
function createBillOfMaterials() { nichDocs.createBillOfMaterials(); }

/**
 * Deletes a block (header row + any child rows) from the spreadsheet.
 * User must select the Edit cell (Column G) of the block's header row.
 * Also removes the corresponding log entry from the log sheet.
 */
function deleteSelectedBlock() {
  try {
    const ui = SpreadsheetApp.getUi();
    const sheet = SpreadsheetApp.getActiveSheet();
    const activeCell = sheet.getActiveCell();

    // Must select Column G
    if (activeCell.getColumn() !== 7) {
      ui.alert('Delete Block', 'Please select the "Edit" cell (Column G) of the block you want to delete, then try again.', ui.ButtonSet.OK);
      return;
    }

    const cellNote = activeCell.getNote();
    if (!cellNote || !cellNote.includes('LogID:')) {
      ui.alert('Delete Block', 'No block data found for this cell. Please select a valid Edit cell.', ui.ButtonSet.OK);
      return;
    }

    const logIdMatch = cellNote.match(/LogID:\s*([^\n\r]+)/);
    if (!logIdMatch) {
      ui.alert('Delete Block', 'Could not find LogID in the selected cell.', ui.ButtonSet.OK);
      return;
    }

    const logId = logIdMatch[1].trim();
    const headerRow = activeCell.getRow();

    // Get description for confirmation
    const description = sheet.getRange(headerRow, 3).getValue() || 'this item';

    // Confirm deletion
    const response = ui.alert(
      'Delete Block',
      `Are you sure you want to delete "${description}" and all its child rows?\n\nThis cannot be undone.`,
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) return;

    // Find child rows by scanning for ParentLogID notes below the header
    const lastRow = sheet.getLastRow();
    const childRows = [];

    if (lastRow > headerRow) {
      const numRowsBelow = lastRow - headerRow;
      const notes = sheet.getRange(headerRow + 1, 7, numRowsBelow, 1).getNotes();

      for (let i = 0; i < notes.length; i++) {
        const note = notes[i][0];
        if (note && note.includes('ParentLogID: ' + logId)) {
          childRows.push(headerRow + 1 + i);
        }
      }
    }

    // Delete from bottom up: child rows first, then header
    for (let i = childRows.length - 1; i >= 0; i--) {
      sheet.deleteRow(childRows[i]);
    }
    sheet.deleteRow(headerRow);

    // Remove log entry from the appropriate log sheet
    let logSheetName = null;
    if (logId.startsWith('FAB_')) logSheetName = 'FabricationLog';
    else if (logId.startsWith('APP_')) logSheetName = 'ApparelLog';
    else if (logId.startsWith('PRT_')) logSheetName = 'PrintingLog';

    if (logSheetName) {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const logSheet = spreadsheet.getSheetByName(logSheetName);
      if (logSheet) {
        const logValues = logSheet.getDataRange().getValues();
        for (let i = logValues.length - 1; i >= 1; i--) {
          if (logValues[i][0] === logId) {
            logSheet.deleteRow(i + 1);
            break;
          }
        }
      }
    }

    ui.alert('Delete Block', `"${description}" has been deleted.`, ui.ButtonSet.OK);

  } catch (error) {
    console.error('Error in deleteSelectedBlock:', error);
    SpreadsheetApp.getUi().alert('Error', 'An error occurred while deleting: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}
