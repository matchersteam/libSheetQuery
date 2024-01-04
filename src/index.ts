import { Pool } from 'pg';

const pool = new Pool({
  user: '',
  host: '',
  database: '',
  password: '',
  port: 5432,
});

type WhereFn = (row: RowObject) => boolean;
type UpdateFn = (row: RowObject) => RowObject;
type RowObject = {
  [key: string]: any;
  __meta?: { row: number; cols: number };
};

/**
 * Run new sheet query
 *
 * @param {string} tableName Specific table to use
 * @return {SheetQueryBuilder}
 */
export function sheetQuery(tableName: string): SheetQueryBuilder {
  return new SheetQueryBuilder(tableName);
}

/**
 * SheetQueryBuilder class - Kind of an ORM for PostgreSQL
 */
class SheetQueryBuilder {
  columnNames: string[];
  tableName: string;
  whereFn: WhereFn | undefined;
  _tableValues: RowObject[];

  constructor(tableName: string) {
    this.columnNames = [];
    this.tableName = tableName;
    this._tableValues = [];
  }

  select(columnNames: string) {
    this.columnNames = Array.isArray(columnNames) ? columnNames : [columnNames];
    return this;
  }

  /**
   * Apply a filtering function on rows in a table before performing an operation on them
   *
   * @param {WhereFn} fn
   * @return {SheetQueryBuilder}
   */
  where(fn: WhereFn): SheetQueryBuilder {
    this.whereFn = fn;
    return this;
  }

  /**
   * Delete matched rows from table
   *
   * @return {SheetQueryBuilder}
   */
  async deleteRows(): Promise<SheetQueryBuilder> {
    const rows = await this.getRows();
    const idsToDelete = rows.filter(row => row.__meta !== undefined).map(row => row.__meta!.row);
    const deleteRowQuery = `DELETE FROM ${this.tableName} WHERE id = ANY($1::int[])`;
    await pool.query(deleteRowQuery, [idsToDelete]);
    this.clearCache();
    return this;
  }

  /**
   * Update matched rows in table with provided function
   *
   * @param {} updateFn
   * @return {SheetQueryBuilder}
   */
  async updateRows(updateFn: UpdateFn): Promise<SheetQueryBuilder> {
    const rows = await this.getRows();
    const updatedRows = rows.map(updateFn);

    const updateRowQueries = updatedRows.map((updatedRow, i) => {
      const originalRow = rows[i];
      if (originalRow.__meta) {
        const changes = Object.keys(updatedRow).reduce((acc: { [key: string]: any }, key) => {
          if (originalRow[key] !== updatedRow[key]) {
            acc[key] = updatedRow[key];
          }
          return acc;
        }, {});

        return `UPDATE ${this.tableName} SET ${Object.entries(changes).map(([col, val]) => `${col} = ${val}`).join(', ')} WHERE id = ${originalRow.__meta.row}`;
      }
      return null;
    }).filter(Boolean);

    const updateRowQuery = updateRowQueries.join('; ');
    await pool.query(updateRowQuery);

    this.clearCache();
    return this;
  }

  /**
   * Get values in table from current query + where condition
   */
  async getValues() {
    if (!this._tableValues) {
      const selectQuery = `SELECT * FROM ${this.tableName}`;
      const { rows } = await pool.query(selectQuery);
      this._tableValues = rows;
    }
    return this._tableValues;
  }

  /**
   * Return matching rows from table query
   *
   * @return {RowObject[]}
   */
  async getRows() {
    const tableValues = await this.getValues();
    return this.whereFn ? tableValues.filter(this.whereFn) : tableValues;
  }

  /**
   * Insert new rows into the table
   * Arrays of objects like { Heading: Value }
   *
   * @param {RowObject[]} newRows - Array of row objects to insert
   * @return {SheetQueryBuilder}
   */
  async insertRows(newRows: RowObject[]): Promise<SheetQueryBuilder> {
    newRows.forEach(async (row) => {
      if (!row) {
        return;
      }
      const rowValues = this.columnNames.map((heading) => {
        return (heading && row[heading]) || (heading && row[heading] === false) ? row[heading] : '';
      });
      const insertRowQuery = `INSERT INTO ${this.tableName} (${this.columnNames.join(', ')}) VALUES (${rowValues.join(', ')})`;
      await pool.query(insertRowQuery);
    });
    return this;
  }

  /**
   * Clear cached values
   *
   * @return {SheetQueryBuilder}
   */
  clearCache() {
    this._tableValues = [];
    return this;
  }
}
