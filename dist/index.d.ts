/// <reference types="google-apps-script" />
declare type Spreadsheet = GoogleAppsScript.Spreadsheet.Spreadsheet;
declare type Sheet = GoogleAppsScript.Spreadsheet.Sheet;
declare type Range = GoogleAppsScript.Spreadsheet.Range;
declare type Cell = {
    [key: string]: Range;
};
declare type DictObject = {
    [key: string]: any;
};
declare type RowValues = {
    __meta?: {
        row: number;
        cols: number;
    };
} & DictObject;
declare type FilterFunction = (value: RowValues, index: number, array: RowValues[]) => RowValues[];
declare type Url = {
    [key: string]: string | null;
};
/**
 * Run new sheet query
 *
 * @param {Spreadsheet} activeSpreadsheet Specific spreadsheet to use, or will use SpreadsheetApp.getActiveSpreadsheet() if undefined\
 * @return {SheetQueryBuilder}
 */
export declare function sheetQuery(activeSpreadsheet?: GoogleAppsScript.Spreadsheet.Spreadsheet | null): SheetQueryBuilder;
/**
 * SheetQueryBuilder class - Kind of an ORM for Google Sheets
 */
declare class SheetQueryBuilder {
    columnNames: string[];
    headingRow: number;
    _sheetHeadings: string[];
    activeSpreadsheet: Spreadsheet;
    sheetName: string;
    _sheet: Sheet | null;
    whereFn: FilterFunction;
    _sheetValues: RowValues[];
    constructor(activeSpreadsheet: Spreadsheet | null);
    select(columnNames: string[]): this;
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
     * @param {FilterFunction} fn
     * @return {SheetQueryBuilder}
     */
    where(fn: FilterFunction): this;
    /**
     * Delete matched rows from spreadsheet
     *
     * @return {SheetQueryBuilder}
     */
    deleteRows(): this;
    /**
     * Update matched rows in spreadsheet with provided function
     *
     * @param {Function} updateFn
     * @return {SheetQueryBuilder}
     */
    updateRows(updateFn: Function): this;
    /**
     * Get Sheet object that is referenced by the current query from() method
     *
     * @return {Sheet}
     */
    getSheet(): GoogleAppsScript.Spreadsheet.Sheet | null;
    /**
     * Get values in sheet from current query + where condition
     */
    getValues(): RowValues[];
    /**
     * Return matching rows from sheet query
     *
     * @return {RowValues[]}
     */
    getRows(): RowValues[];
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
     * @param {DictObject[]} newRows - Array of row objects to insert
     * @return {SheetQueryBuilder}
     */
    insertRows(newRows: DictObject[]): this;
    /**
     * Clear cached values, headings, and flush all operations to sheet
     *
     * @return {SheetQueryBuilder}
     */
    clearCache(): this;
    getCells(): Cell[];
    getUrls(): Url[];
}
export {};
//# sourceMappingURL=index.d.ts.map