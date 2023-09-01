"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.sheetQuery = void 0;
function getActiveSheet() {
    if (SpreadsheetApp.getActiveSpreadsheet())
        return SpreadsheetApp.getActiveSpreadsheet();
    else
        throw new Error('No active sheet for sheetQuery!');
}
/**
 * Run new sheet query
 *
 * @param {Spreadsheet} activeSpreadsheet Specific spreadsheet to use, or will use SpreadsheetApp.getActiveSpreadsheet() if undefined\
 * @return {SheetQueryBuilder}
 */
function sheetQuery(activeSpreadsheet) {
    if (activeSpreadsheet === void 0) { activeSpreadsheet = null; }
    return new SheetQueryBuilder(activeSpreadsheet);
}
exports.sheetQuery = sheetQuery;
/**
 * SheetQueryBuilder class - Kind of an ORM for Google Sheets
 */
var SheetQueryBuilder = /** @class */ (function () {
    function SheetQueryBuilder(activeSpreadsheet) {
        this.columnNames = [];
        this.headingRow = 1;
        this._sheetHeadings = [];
        this._sheetValues = [];
        this.activeSpreadsheet = activeSpreadsheet || getActiveSheet();
    }
    SheetQueryBuilder.prototype.select = function (columnNames) {
        this.columnNames = Array.isArray(columnNames) ? columnNames : [columnNames];
        return this;
    };
    /**
     * Name of spreadsheet to perform operations on
     *
     * @param {string} sheetName
     * @param {number} headingRow
     * @return {SheetQueryBuilder}
     */
    SheetQueryBuilder.prototype.from = function (sheetName, headingRow) {
        if (headingRow === void 0) { headingRow = 1; }
        this.sheetName = sheetName;
        this.headingRow = headingRow;
        return this;
    };
    /**
     * Apply a filtering function on rows in a spreadsheet before performing an operation on them
     *
     * @param {WhereFn} fn
     * @return {SheetQueryBuilder}
     */
    SheetQueryBuilder.prototype.where = function (fn) {
        this.whereFn = fn;
        return this;
    };
    /**
     * Delete matched rows from spreadsheet
     *
     * @return {SheetQueryBuilder}
     */
    SheetQueryBuilder.prototype.deleteRows = function () {
        var _this = this;
        var rows = this.getRows();
        var i = 0;
        rows.forEach(function (row) {
            if (_this._sheet && row.__meta) {
                var deleteRowRange = _this._sheet.getRange(row.__meta.row - i, 1, 1, row.__meta.cols);
                deleteRowRange.deleteCells(SpreadsheetApp.Dimension.ROWS);
                i += 1;
            }
        });
        this.clearCache();
        return this;
    };
    /**
     * Update matched rows in spreadsheet with provided function
     *
     * @param {} updateFn
     * @return {SheetQueryBuilder}
     */
    SheetQueryBuilder.prototype.updateRows = function (updateFn) {
        var _this = this;
        var rows = this.getRows();
        rows.forEach(function (row) {
            if (_this._sheet && row.__meta) {
                var nbCol = row.__meta.cols;
                var numRow = row.__meta.row;
                var updateRowRange = _this._sheet.getRange(row.__meta.row, 1, 1, row.__meta.cols);
                var richTextValues = updateRowRange.getRichTextValues();
                var updatedRow = updateFn(row);
                var arrayValues = [];
                if (updatedRow && updatedRow.__meta) {
                    delete updatedRow.__meta;
                    arrayValues = Object.values(updatedRow);
                }
                else {
                    delete row.__meta;
                    arrayValues = Object.values(row);
                }
                /* Update arrayValues to add formulas */
                for (var i = 1; i < nbCol + 1; i++) {
                    var range = _this._sheet.getRange(numRow, i, 1, 1);
                    var formula = range.getFormula();
                    if (formula != '') {
                        arrayValues[i - 1] = formula;
                    }
                }
                updateRowRange.setValues([arrayValues]);
                /* fix to preserve RichTextValues */
                for (var i = 1; i < nbCol + 1; i++) {
                    var richTextValue = richTextValues[0][i - 1];
                    if ((richTextValue === null || richTextValue === void 0 ? void 0 : richTextValue.getLinkUrl()) != null) {
                        var newRichTextValue = SpreadsheetApp.newRichTextValue()
                            .setText(arrayValues[i - 1])
                            .setLinkUrl(richTextValue.getLinkUrl())
                            .build();
                        var range = _this._sheet.getRange(numRow, i, 1, 1);
                        range.setRichTextValue(newRichTextValue);
                    }
                }
            }
        });
        this.clearCache();
        return this;
    };
    /**
     * Get Sheet object that is referenced by the current query from() method
     *
     */
    SheetQueryBuilder.prototype.getSheet = function () {
        if (!this._sheet && this.sheetName) {
            this._sheet = this.activeSpreadsheet.getSheetByName(this.sheetName);
        }
        return this._sheet;
    };
    /**
     * Get values in sheet from current query + where condition
     */
    SheetQueryBuilder.prototype.getValues = function () {
        if (!this._sheetValues) {
            var zh = this.headingRow - 1;
            var sheet = this.getSheet();
            if (!sheet) {
                return [];
            }
            var rowValues = [];
            var allValues = sheet.getDataRange().getValues();
            var sheetValues = allValues.slice(1 + zh);
            var numCols = sheetValues[0].length;
            var numRows = sheetValues.length;
            var headings = (this._sheetHeadings = allValues[zh] || []);
            for (var r = 0; r < numRows; r++) {
                var obj = { __meta: { row: r + (this.headingRow + 1), cols: numCols } };
                for (var c = 0; c < numCols; c++) {
                    // @ts-expect-error: Headings are set already above, so possibility of an error here is nil
                    obj[headings[c]] = sheetValues[r][c]; // @ts-ignore
                }
                rowValues.push(obj);
            }
            this._sheetValues = rowValues;
        }
        return this._sheetValues;
    };
    /**
     * Return matching rows from sheet query
     *
     * @return {RowObject[]}
     */
    SheetQueryBuilder.prototype.getRows = function () {
        var sheetValues = this.getValues();
        return this.whereFn ? sheetValues.filter(this.whereFn) : sheetValues;
    };
    /**
     * Get array of headings in current sheet from()
     *
     * @return {string[]}
     */
    SheetQueryBuilder.prototype.getHeadings = function () {
        if (!this._sheetHeadings || !this._sheetHeadings.length) {
            var zh = this.headingRow - 1;
            var sheet = this.getSheet();
            var numCols = sheet.getLastColumn();
            this._sheetHeadings = sheet.getSheetValues(1, 1, this.headingRow, numCols)[zh];
        }
        return this._sheetHeadings || [];
    };
    /**
     * Insert new rows into the spreadsheet
     * Arrays of objects like { Heading: Value }
     *
     * @param {RowObject[]} newRows - Array of row objects to insert
     * @return {SheetQueryBuilder}
     */
    SheetQueryBuilder.prototype.insertRows = function (newRows) {
        var sheet = this.getSheet();
        if (!sheet)
            return this;
        var headings = this.getHeadings();
        newRows.forEach(function (row) {
            if (!row) {
                return;
            }
            var rowValues = headings.map(function (heading) {
                return (heading && row[heading]) || (heading && row[heading] === false) ? row[heading] : '';
            });
            sheet === null || sheet === void 0 ? void 0 : sheet.appendRow(rowValues);
        });
        return this;
    };
    /**
     * Clear cached values, headings, and flush all operations to sheet
     *
     * @return {SheetQueryBuilder}
     */
    SheetQueryBuilder.prototype.clearCache = function () {
        this._sheetValues = [];
        this._numRows = undefined;
        this._sheetHeadings = [];
        SpreadsheetApp.flush();
        return this;
    };
    SheetQueryBuilder.prototype.getCells = function () {
        var _this = this;
        var rows = this.getRows();
        var cells = [];
        rows.forEach(function (row) {
            var cell = {};
            for (var i = 0; i < _this._sheetHeadings.length; i++) {
                var heading = _this._sheetHeadings[i];
                cell[heading] = _this._sheet.getRange(row.__meta.row, i + 1, 1, 1);
            }
            cells.push(cell);
        });
        return cells;
    };
    SheetQueryBuilder.prototype.getUrls = function () {
        var _this = this;
        var rows = this.getRows();
        var urls = [];
        rows.forEach(function (row) {
            var url = {};
            for (var i = 0; i < _this._sheetHeadings.length; i++) {
                var heading = _this._sheetHeadings[i];
                var cellRichTextValue = _this._sheet.getRange(row.__meta.row, i + 1, 1, 1).getRichTextValue();
                if (cellRichTextValue)
                    url[heading] = cellRichTextValue.getLinkUrl();
            }
            urls.push(url);
        });
        return urls;
    };
    return SheetQueryBuilder;
}());
// /* eslint-disable */
// //🚨 WARNING: Le passage en TS ne fonctionne pas. J'ai copier/coller dans dist/index.js le fichier lib_sheetQuery de gestion-formations
// /**
//  * @returns {Spreadsheet} the active spreadsheet
//  */
// function getActiveSheet() {
//     if (SpreadsheetApp.getActiveSpreadsheet()) return SpreadsheetApp.getActiveSpreadsheet();
//     else throw new Error('Error from libSheetQuery: Cannot use getActiveSpreadsheet !');
//   }
// type Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet;
// type Sheet = GoogleAppsScript.Spreadsheet.Sheet;
// type Range = GoogleAppsScript.Spreadsheet.Range;
// type Cell = { [key: string]: Range };
// type DictObject = { [key: string]: any };
// type RowValues = { __meta?: { row: number; cols: number } } & DictObject;
// type FilterFunction = (value: RowValues, index: number, array: RowValues[]) => RowValues[];
// type Url = { [key: string]: string | null };
// /**
//  * Run new sheet query
//  *
//  * @param {Spreadsheet} activeSpreadsheet Specific spreadsheet to use, or will use SpreadsheetApp.getActiveSpreadsheet() if undefined\
//  * @return {SheetQueryBuilder}
//  */
// export function sheetQuery(activeSpreadsheet = null as Spreadsheet | null) {
//   return new SheetQueryBuilder(activeSpreadsheet);
// }
// /**
//  * SheetQueryBuilder class - Kind of an ORM for Google Sheets
//  */
// class SheetQueryBuilder {
//   columnNames: string[];
//   headingRow: number;
//   _sheetHeadings: string[];
//   activeSpreadsheet: Spreadsheet;
//   sheetName: string;
//   _sheet: Sheet | null;
//   whereFn: FilterFunction;
//   _sheetValues: RowValues[];
//   constructor(activeSpreadsheet: Spreadsheet | null) {
//     this.columnNames = [];
//     this.headingRow = 1;
//     this._sheetHeadings = [];
//     this.activeSpreadsheet = activeSpreadsheet || getActiveSheet();
//     this.sheetName = '';
//     this._sheet = null;
//     this.whereFn = (value, index, array) => array;
//     this._sheetValues = [];
//   }
//   select(columnNames: string[]) {
//     this.columnNames = Array.isArray(columnNames) ? columnNames : [columnNames];
//     return this;
//   }
//   /**
//    * Name of spreadsheet to perform operations on
//    *
//    * @param {string} sheetName
//    * @param {number} headingRow
//    * @return {SheetQueryBuilder}
//    */
//   from(sheetName: string, headingRow = 1) {
//     this.sheetName = sheetName;
//     this.headingRow = headingRow;
//     return this;
//   }
//   /**
//    * Apply a filtering function on rows in a spreadsheet before performing an operation on them
//    *
//    * @param {FilterFunction} fn
//    * @return {SheetQueryBuilder}
//    */
//   where(fn: FilterFunction) {
//     this.whereFn = fn;
//     return this;
//   }
//   /**
//    * Delete matched rows from spreadsheet
//    *
//    * @return {SheetQueryBuilder}
//    */
//   deleteRows() {
//     if (!this._sheet) return this;
//     const rows = this.getRows();
//     let i = 0;
//     rows.forEach((row) => {
//       if (!row.__meta) return;
//       const deleteRowRange = this._sheet!.getRange(row.__meta.row - i, 1, 1, row.__meta.cols);
//       deleteRowRange.deleteCells(SpreadsheetApp.Dimension.ROWS);
//       i += 1;
//     });
//     this.clearCache();
//     return this;
//   }
//   /**
//    * Update matched rows in spreadsheet with provided function
//    *
//    * @param {Function} updateFn
//    * @return {SheetQueryBuilder}
//    */
//   updateRows(updateFn: Function) {
//     if (!this._sheet) return this;
//     const rows = this.getRows();
//     rows.forEach((row) => {
//       if (!row.__meta) return;
//       const nbCol = row.__meta.cols;
//       const numRow = row.__meta.row;
//       const updateRowRange = this._sheet!.getRange(row.__meta.row, 1, 1, row.__meta.cols);
//       const richTextValues = updateRowRange.getRichTextValues();
//       const updatedRow = updateFn(row);
//       let arrayValues = [];
//       if (updatedRow && updatedRow.__meta) {
//         delete updatedRow.__meta;
//         arrayValues = Object.values(updatedRow);
//       } else {
//         delete row.__meta;
//         arrayValues = Object.values(row);
//       }
//       if (!this._sheet) return this;
//       /* Update arrayValues to add formulas */
//       for (let i = 1; i < nbCol + 1; i++) {
//         const range = this._sheet.getRange(numRow, i, 1, 1);
//         const formula = range.getFormula();
//         if (formula != '') {
//           arrayValues[i - 1] = formula;
//         }
//       }
//       updateRowRange.setValues([arrayValues]);
//       /* fix to preserve RichTextValues */
//       for (let i = 1; i < nbCol + 1; i++) {
//         const richTextValue = richTextValues[0][i - 1];
//         if (richTextValue && richTextValue.getLinkUrl() != null) {
//           const newRichTextValue = SpreadsheetApp.newRichTextValue()
//             .setText(arrayValues[i - 1])
//             .setLinkUrl(richTextValue.getLinkUrl())
//             .build();
//           const range = this._sheet.getRange(numRow, i, 1, 1);
//           range.setRichTextValue(newRichTextValue);
//         }
//       }
//     });
//     this.clearCache();
//     return this;
//   }
//   /**
//    * Get Sheet object that is referenced by the current query from() method
//    *
//    * @return {Sheet}
//    */
//   getSheet() {
//     if (!this._sheet) {
//       this._sheet = this.activeSpreadsheet.getSheetByName(this.sheetName);
//     }
//     return this._sheet;
//   }
//   /**
//    * Get values in sheet from current query + where condition
//    */
//   getValues() {
//     if (!this._sheetValues) {
//       const zh = this.headingRow - 1;
//       const sheet = this.getSheet();
//       if (!sheet) {
//         return [];
//       }
//       const rowValues = [];
//       const allValues = sheet.getDataRange().getValues();
//       const sheetValues = allValues.slice(1 + zh);
//       const numCols = sheetValues[0].length;
//       const numRows = sheetValues.length;
//       const headings = (this._sheetHeadings = allValues[zh] || []);
//       for (let r = 0; r < numRows; r++) {
//         const obj = { __meta: { row: r + (this.headingRow + 1), cols: numCols } };
//         for (let c = 0; c < numCols; c++) {
//           // @ts-expect-error: Headings are set already above, so possibility of an error here is nil
//           obj[headings[c]] = sheetValues[r][c]; // @ts-ignore
//         }
//         rowValues.push(obj);
//       }
//       this._sheetValues = rowValues;
//     }
//     return this._sheetValues;
//   }
//   /**
//    * Return matching rows from sheet query
//    *
//    * @return {RowValues[]}
//    */
//   getRows(): RowValues[] {
//     const sheetValues = this.getValues();
//     return this.whereFn ? sheetValues.filter(this.whereFn) : sheetValues;
//   }
//   /**
//    * Get array of headings in current sheet from()
//    *
//    * @return {string[]}
//    */
//   getHeadings() {
//     if (!this._sheetHeadings || !this._sheetHeadings.length) {
//       const zh = this.headingRow - 1;
//       const sheet = this.getSheet();
//       if (!sheet) return [];
//       const numCols = sheet.getLastColumn();
//       this._sheetHeadings = sheet.getSheetValues(1, 1, this.headingRow, numCols)[zh];
//     }
//     return this._sheetHeadings || [];
//   }
//   /**
//    * Insert new rows into the spreadsheet
//    * Arrays of objects like { Heading: Value }
//    *
//    * @param {DictObject[]} newRows - Array of row objects to insert
//    * @return {SheetQueryBuilder}
//    */
//   insertRows(newRows: DictObject[]) {
//     const sheet = this.getSheet();
//     if (!sheet) return this;
//     const headings = this.getHeadings();
//     newRows.forEach((row) => {
//       if (!row) {
//         return;
//       }
//       const rowValues = headings.map((heading) => {
//         return (heading && row[heading]) || (heading && row[heading] === false) ? row[heading] : '';
//       });
//       sheet.appendRow(rowValues);
//     });
//     return this;
//   }
//   /**
//    * Clear cached values, headings, and flush all operations to sheet
//    *
//    * @return {SheetQueryBuilder}
//    */
//   clearCache() {
//     this._sheetValues = [];
//     this._sheetHeadings = [];
//     SpreadsheetApp.flush();
//     return this;
//   }
//   getCells() {
//     const rows = this.getRows();
//     const cells = [] as Cell[];
//     rows.forEach((row) => {
//       const cell = {} as Cell;
//       for (let i = 0; i < this._sheetHeadings.length; i++) {
//         if (!this._sheet || !row.__meta) continue;
//         const heading = this._sheetHeadings[i];
//         cell[heading] = this._sheet.getRange(row.__meta.row, i + 1, 1, 1);
//       }
//       cells.push(cell);
//     });
//     return cells;
//   }
//   getUrls() {
//     const rows = this.getRows();
//     const urls = [] as Url[];
//     rows.forEach((row) => {
//       const url = {} as Url;
//       for (let i = 0; i < this._sheetHeadings.length; i++) {
//         if (!this._sheet || !row.__meta) continue;
//         const heading = this._sheetHeadings[i];
//         const cellRichTextValue = this._sheet.getRange(row.__meta.row, i + 1, 1, 1).getRichTextValue();
//         if (cellRichTextValue) url[heading] = cellRichTextValue.getLinkUrl();
//       }
//       urls.push(url);
//     });
//     return urls;
//   }
// }
//# sourceMappingURL=index.js.map