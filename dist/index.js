/* eslint-disable */
// SheetQuery library, version : https://github.com/vlucas/sheetquery/commit/9e41a1905d8d2480098b6351ff44016ac9e2cd6b

/**
 * Run new sheet query
 *
 * @param {Spreadsheet} activeSpreadsheet Specific spreadsheet to use, or will use SpreadsheetApp.getActiveSpreadsheet() if undefined\
 * @return {SheetQueryBuilder}
 */
export function sheetQuery(activeSpreadsheet = false) {
  return new SheetQueryBuilder(activeSpreadsheet);
}

/**
 * SheetQueryBuilder class - Kind of an ORM for Google Sheets
 */
class SheetQueryBuilder {
  constructor(activeSpreadsheet) {
    this.columnNames = [];
    this.headingRow = 1;
    this._sheetHeadings = [];
    this.activeSpreadsheet = activeSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  }

  select(columnNames) {
    this.columnNames = Array.isArray(columnNames) ? columnNames : [columnNames];
    return this;
  }

  /**
   * Name of spreadsheet to perform operations on
   *
   * @param {string} sheetName
   * @param {number} headingRow
   * @return {SheetQueryBuilder}
   */
  from(sheetName, headingRow = 1) {
    this.sheetName = sheetName;
    this.headingRow = headingRow;
    return this;
  }

  /**
   * Apply a filtering function on rows in a spreadsheet before performing an operation on them
   *
   * @param {Function} fn
   * @return {SheetQueryBuilder}
   */
  where(fn) {
    this.whereFn = fn;
    return this;
  }

  /**
   * Delete matched rows from spreadsheet
   *
   * @return {SheetQueryBuilder}
   */
  deleteRows() {
    const rows = this.getRows();
    let i = 0;
    rows.forEach((row) => {
      const deleteRowRange = this._sheet.getRange(row.__meta.row - i, 1, 1, row.__meta.cols);
      deleteRowRange.deleteCells(SpreadsheetApp.Dimension.ROWS);
      i += 1;
    });
    this.clearCache();
    return this;
  }

  /**
   * Update matched rows in spreadsheet with provided function
   *
   * @param {UpdateFn} updateFn
   * @return {SheetQueryBuilder}
   */
  updateRows(updateFn) {
    const rows = this.getRows();
    rows.forEach((row) => {
      const nbCol = row.__meta.cols;
      const numRow = row.__meta.row;
      const updateRowRange = this._sheet.getRange(row.__meta.row, 1, 1, row.__meta.cols);
      const richTextValues = updateRowRange.getRichTextValues();
      const updatedRow = updateFn(row);
      let arrayValues = [];
      if (updatedRow && updatedRow.__meta) {
        delete updatedRow.__meta;
        arrayValues = Object.values(updatedRow);
      } else {
        delete row.__meta;
        arrayValues = Object.values(row);
      }

      /* Update arrayValues to add formulas */
      for (let i = 1; i < nbCol + 1; i++) {
        const range = this._sheet.getRange(numRow, i, 1, 1);
        const formula = range.getFormula();

        if (formula != '') {
          arrayValues[i - 1] = formula;
        }
      }

      updateRowRange.setValues([arrayValues]);

      /* fix to preserve RichTextValues */
      for (let i = 1; i < nbCol + 1; i++) {
        const richTextValue = richTextValues[0][i - 1];
        if (richTextValue.getLinkUrl() != null) {
          const newRichTextValue = SpreadsheetApp.newRichTextValue()
            .setText(arrayValues[i - 1])
            .setLinkUrl(richTextValue.getLinkUrl())
            .build();
          const range = this._sheet.getRange(numRow, i, 1, 1);
          range.setRichTextValue(newRichTextValue);
        }
      }
    });

    this.clearCache();
    return this;
  }

  /**
   * Get Sheet object that is referenced by the current query from() method
   *
   * @return {Sheet}
   */
  getSheet() {
    if (!this._sheet) {
      this._sheet = this.activeSpreadsheet.getSheetByName(this.sheetName);
    }
    return this._sheet;
  }

  /**
   * Get values in sheet from current query + where condition
   */
  getValues() {
    if (!this._sheetValues) {
      const zh = this.headingRow - 1;
      const sheet = this.getSheet();

      if (!sheet) {
        return [];
      }

      const rowValues = [];
      const allValues = sheet.getDataRange().getValues();
      const sheetValues = allValues.slice(1 + zh);
      const numCols = sheetValues[0].length;
      const numRows = sheetValues.length;
      const headings = (this._sheetHeadings = allValues[zh] || []);
      for (let r = 0; r < numRows; r++) {
        const obj = { __meta: { row: r + (this.headingRow + 1), cols: numCols } };
        for (let c = 0; c < numCols; c++) {
          // @ts-expect-error: Headings are set already above, so possibility of an error here is nil
          obj[headings[c]] = sheetValues[r][c]; // @ts-ignore
        }
        rowValues.push(obj);
      }
      this._sheetValues = rowValues;
    }
    return this._sheetValues;
  }

  /**
   * Return matching rows from sheet query
   *
   * @return {RowObject[]}
   */
  getRows() {
    const sheetValues = this.getValues();
    return this.whereFn ? sheetValues.filter(this.whereFn) : sheetValues;
  }

  /**
   * Get array of headings in current sheet from()
   *
   * @return {string[]}
   */
  getHeadings() {
    if (!this._sheetHeadings || !this._sheetHeadings.length) {
      const zh = this.headingRow - 1;
      const sheet = this.getSheet();
      const numCols = sheet.getLastColumn();
      this._sheetHeadings = sheet.getSheetValues(1, 1, this.headingRow, numCols)[zh];
    }
    return this._sheetHeadings || [];
  }

  /**
   * Insert new rows into the spreadsheet
   * Arrays of objects like { Heading: Value }
   *
   * @param {DictObject[]} newRows - Array of row objects to insert
   * @return {SheetQueryBuilder}
   */
  insertRows(newRows) {
    const sheet = this.getSheet();
    const headings = this.getHeadings();
    newRows.forEach((row) => {
      if (!row) {
        return;
      }
      const rowValues = headings.map((heading) => {
        return (heading && row[heading]) || (heading && row[heading] === false) ? row[heading] : '';
      });
      sheet.appendRow(rowValues);
    });
    return this;
  }

  /**
   * Clear cached values, headings, and flush all operations to sheet
   *
   * @return {SheetQueryBuilder}
   */
  clearCache() {
    this._sheetValues = null;
    // this._sheetCells = null;
    this._numRows = undefined;
    this._sheetHeadings = [];
    SpreadsheetApp.flush();
    return this;
  }

  getCells() {
    const rows = this.getRows();
    const cells = [];
    rows.forEach((row) => {
      const cell = {}
      for (let i = 0; i < this._sheetHeadings.length; i++) {
        const heading = this._sheetHeadings[i];
        cell[heading] = this._sheet.getRange(row.__meta.row, i + 1, 1, 1);
      }
      cells.push(cell);
    });
    return cells;
  }

  getUrls() {
    const rows = this.getRows();
    const urls = [];
    rows.forEach((row) => {
      const url = {}
      for (let i = 0; i < this._sheetHeadings.length; i++) {
        const heading = this._sheetHeadings[i];
        const cellRichTextValue = this._sheet.getRange(row.__meta.row, i + 1, 1, 1).getRichTextValue();
        if(cellRichTextValue) url[heading] = cellRichTextValue.getLinkUrl();
      }
      urls.push(url);
    });
    return urls;
  }
}