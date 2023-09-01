/// <reference types="google-apps-script" />
declare type Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet;
declare type Sheet = GoogleAppsScript.Spreadsheet.Sheet;
declare type Range = GoogleAppsScript.Spreadsheet.Range;
declare type Cell = {
    [key: string]: Range;
};
declare type WhereFn = (row: RowObject) => boolean;
declare type UpdateFn = (row: RowObject) => RowObject;
declare type RowObject = {
    [key: string]: any;
    __meta?: {
        row: number;
        cols: number;
    };
};
/**
 * Run new sheet query
 *
 * @param {Spreadsheet} activeSpreadsheet Specific spreadsheet to use, or will use SpreadsheetApp.getActiveSpreadsheet() if undefined\
 * @return {SheetQueryBuilder}
 */
export declare function sheetQuery(activeSpreadsheet?: Spreadsheet | null): SheetQueryBuilder;
/**
 * SheetQueryBuilder class - Kind of an ORM for Google Sheets
 */
declare class SheetQueryBuilder {
    columnNames: string[];
    headingRow: number;
    _sheetHeadings: string[];
    activeSpreadsheet: Spreadsheet;
    sheetName: string | undefined;
    whereFn: WhereFn | undefined;
    _sheet: Sheet | null | undefined;
    _sheetValues: RowObject[];
    _numRows: undefined;
    constructor(activeSpreadsheet: Spreadsheet | null);
    select(columnNames: string): this;
    /**
     * Name of spreadsheet to perform operations on
     *
     * @param {string} sheetName
     * @param {number} headingRow
     * @return {SheetQueryBuilder}
     */
    from(sheetName: string, headingRow?: number): this;
    /**
     * Apply a filtering function on rows in a spreadsheet before performing an operation on them
     *
     * @param {WhereFn} fn
     * @return {SheetQueryBuilder}
     */
    where(fn: WhereFn): SheetQueryBuilder;
    /**
     * Delete matched rows from spreadsheet
     *
     * @return {SheetQueryBuilder}
     */
    deleteRows(): SheetQueryBuilder;
    /**
     * Update matched rows in spreadsheet with provided function
     *
     * @param {} updateFn
     * @return {SheetQueryBuilder}
     */
    updateRows(updateFn: UpdateFn): SheetQueryBuilder;
    /**
     * Get Sheet object that is referenced by the current query from() method
     *
     */
    getSheet(): GoogleAppsScript.Spreadsheet.Sheet | null | undefined;
    /**
     * Get values in sheet from current query + where condition
     */
    getValues(): RowObject[];
    /**
     * Return matching rows from sheet query
     *
     * @return {RowObject[]}
     */
    getRows(): RowObject[];
    /**
     * Get array of headings in current sheet from()
     *
     * @return {string[]}
     */
    getHeadings(): string[];
    /**
     * Insert new rows into the spreadsheet
     * Arrays of objects like { Heading: Value }
     *
     * @param {RowObject[]} newRows - Array of row objects to insert
     * @return {SheetQueryBuilder}
     */
    insertRows(newRows: RowObject[]): SheetQueryBuilder;
    /**
     * Clear cached values, headings, and flush all operations to sheet
     *
     * @return {SheetQueryBuilder}
     */
    clearCache(): this;
    getCells(): Cell[];
    getUrls(): {}[];
}
export {};
//# sourceMappingURL=index.d.ts.map