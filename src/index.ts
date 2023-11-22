import sqlite3 from 'sqlite3';

type Sqlite = any;

interface SqlCnx{
  db:Sqlite;
}

type SqlInfo = SqlCnx;
type SqlTable = any;
type SqlField = String;

type WhereFn = (row: SqlLine) => boolean;
type UpdateFn = (row: SqlLine) => SqlLine;

type SqlLine = {
  [key: string]: any;
  __meta?: { row: number; cols: number };
};

/* {
  id: "truc"
  driveId: "..."
  __meta:{
      row: 2,
      cols: 3
  }
} */

/**
 * Run new sheet query
 *
 * @param {SqlInfo} activeSpreadsheet Specific spreadsheet to use, or will use SpreadsheetApp.getActiveSpreadsheet() if undefined\
 * @return {SqlTableOrm}
 */
export function sheetQuery(activeSpreadsheet: SqlInfo | null = null): SqlTableOrm {
  return new SqlTableOrm(activeSpreadsheet);
}

class OrmCnx{
  private static instance: OrmCnx;
  private static db:Sqlite;

  private constructor() {
    OrmCnx.db = new sqlite3.Database("bo-v3.db"); // by process.env.DATABASE_NAME 
  }

  public static getInstance(): OrmCnx {
    if (!OrmCnx.instance) {
      OrmCnx.instance = new OrmCnx();
    }
    return OrmCnx.instance;
  }

  public static getDbCnx(): Sqlite {
    return OrmCnx.db
  }
}

/**
 * SheetQueryBuilder class - Kind of an ORM for Google Sheets
 */
class SqlTableOrm {
  columnNames: string[];
 
  allSqlFieldsStr: string[];
  dataBaseDescriptor: Sqlite | null;
  tableName: string | undefined;

  whereFn: WhereFn | undefined;
  table: SqlTable | null | undefined;
  lines: SqlLine[];
  count: undefined;

  
  constructor(sqlCnxInfo: SqlInfo | null) {
    this.columnNames = [];

    this.allSqlFieldsStr = [];
    this.lines = [];
    this.dataBaseDescriptor = OrmCnx.getInstance() 
  }

  select(columnNames: string) {
    this.columnNames = Array.isArray(columnNames) ? columnNames : [columnNames];
    return this;
  }

  /**
   * Name of spreadsheet to perform operations on
   *
   * @param {string} sheetName
   * @param {number} headingRow
   * @return {SqlTableOrm}
   */
  from(sheetName: string, headingRow: number = 1): SqlTableOrm {
    this.tableName = sheetName;
  
    return this;
  }

  /**
   * Apply a filtering function on rows in a spreadsheet before performing an operation on them
   *
   * @param {WhereFn} fn
   * @return {SqlTableOrm}
   */
  where(fn: WhereFn): SqlTableOrm {
    this.whereFn = fn;
    return this;
  }

  /**
   * Delete matched rows from spreadsheet
   *
   * @return {SqlTableOrm}
   */
  deleteRows(): SqlTableOrm {
    const rows = this.getRows();
    let i = 0;
    rows.forEach((row) => {
      if (this.table && row.__meta) {
        const deleteRowRange = this.table.getRange(row.__meta.row - i, 1, 1, row.__meta.cols);
        deleteRowRange.deleteCells(SpreadsheetApp.Dimension.ROWS);
        i += 1;
      }
    });
  
    return this;
  }

  /**
   * Update matched rows in spreadsheet with provided function
   *
   * @param {} updateFn
   * @return {SqlTableOrm}
   */
  updateRows(updateFn: UpdateFn): SqlTableOrm {
    const rows = this.getRows();
    rows.forEach((row) => {
      if (this.table && row.__meta) {

        const nbCol = row.__meta.cols;
        const numRow = row.__meta.row;
        const updateRowRange = this.table.getRange(row.__meta.row, 1, 1, row.__meta.cols);
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
          const range = this.table.getRange(numRow, i, 1, 1);
          const formula = range.getFormula();

          if (formula != '') {
            arrayValues[i - 1] = formula;
          }
        }

        updateRowRange.setValues([arrayValues]);

        /* fix to preserve RichTextValues */
        for (let i = 1; i < nbCol + 1; i++) {
          const richTextValue = richTextValues[0][i - 1];
          if (richTextValue?.getLinkUrl() != null) {
            const newRichTextValue = SpreadsheetApp.newRichTextValue()
              .setText(arrayValues[i - 1])
              .setLinkUrl(richTextValue.getLinkUrl())
              .build();
            const range = this.table.getRange(numRow, i, 1, 1);
            range.setRichTextValue(newRichTextValue);
          }
        }
      }
    });

   
    return this;
  }

  /**
   * Get Sheet object that is referenced by the current query from() method
   *
   */
  getSheet() {
    if (!this.table && this.tableName) {
      this.table = this.dataBaseDescriptor.getSheetByName(this.tableName);
    }
    return this.table;
  }

  /**
   * Get values in sheet from current query + where condition
   */
  getValues() {
    if (!this.lines) {
  
      
    }
    return this.lines;
  }

  /**
   * Return matching rows from sheet query
   *
   * @return {SqlLine[]}
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
  getHeadings(): string[] {
    if (!this.allSqlFieldsStr || !this.allSqlFieldsStr.length) {
      
    }
    return this.allSqlFieldsStr || [];
  }

  /**
   * Insert new rows into the spreadsheet
   * Arrays of objects like { Heading: Value }
   *
   * @param {SqlLine[]} newRows - Array of row objects to insert
   * @return {SqlTableOrm}
   */
  insertRows(newRows: SqlLine[]): SqlTableOrm {
   
    return this;
  }



  getCells() {
    const rows:SqlLine[] = this.getRows();
    const cells: SqlField[] = [];
    rows.forEach((row) => {
      const cell: SqlField = {}
      for (let i = 0; i < this.allSqlFieldsStr.length; i++) {
        const heading = this.allSqlFieldsStr[i];  
        cell[heading] = this.table!.getRange(row.__meta!.row, i + 1, 1, 1);
      }
      cells.push(cell);
    });
    return cells;
  }

  getUrls() {
    const rows = this.getRows();
    const urls: {}[] = [];
    rows.forEach((row) => {
      const url: { [key: string]: string } = {}
      for (let i = 0; i < this.allSqlFieldsStr.length; i++) {
        const heading = this.allSqlFieldsStr[i];
        const cellRichTextValue = this.table!.getRange(row.__meta!.row, i + 1, 1, 1).getRichTextValue();
        if (cellRichTextValue) url[heading] = cellRichTextValue.getLinkUrl()!;
      }
      urls.push(url);
    });
    return urls;
  }
}
