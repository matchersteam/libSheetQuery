Ce projet est un fork (qui date un peu) du package [sheetQuery de Vlucas](https://github.com/vlucas/sheetquery)     

🚨 ATTENTION: Le passage en TS ne fonctionne pas. J'ai copier/coller dans dist/index.js le fichier lib_sheetQuery de gestion-formations

## Requirements

SheetQuery requires a Google Sheet with a heading row (typically the first row where the columns are named). SheetQuery
will use the heading row for all other operations, and for returning row data in key/value objects.

## Usage

SheetQuery operates on a single Sheet at a time. You can start a new query with `sheetQuery().from('SheetName')`.

### Query For Data

Data is queried based on the spreadsheet name and column headings:

```javascript
const query = sheetQuery()
  .from('Transactions')
  .where((row) => row.Category === 'Shops');

// query.getRows() => [{ Amount: 95, Category: 'Shops', Business: 'Walmart'}]
```

If your headings are on a different row than the first row, specify it as the second argument to `from`:

```javascript
const query = sheetQuery()
  .from('Transactions', 2) // For headings on row 2
  .where((row) => row.Category === 'Shops');

// query.getRows() => [{ Amount: 95, Category: 'Shops', Business: 'Walmart'}]
```

### Update Rows

Query for the rows you want to update, and then update them:

```javascript
sheetQuery()
  .from('Transactions')
  .where((row) => row.Business.toLowerCase().includes('starbucks'))
  .updateRows((row) => {
    row.Category = 'Coffee Shops';
  });
```

The `updateRows` method can either return nothing, or can return a row object with updated properties that will be saved
back to the spreadsheet row. If the updater function returns nothing/undefined, the row object that was passed in will
be used (along with any changed values that will be updated by reference).

### Delete Rows

Query for the rows you want to delete, and then delete them:

```javascript
sheetQuery()
  .from('Transactions')
  .where((row) => row.Category === 'DELETEME')
  .deleteRows();
```

Note: Be careful with this one, and always make sure to use it with a `where` filter to limit the number of rows that
will be deleted!

### Insert Rows

Rows can be inserted with SheetQuery by column heading name. No more keeping track of array index positions!

```javascript
sheetQuery()
  .from('Transactions')
  .insertRows([
    {
      Amount: -554.23,
      Name: 'BigBox, inc.',
    },
    {
      Amount: -29.74,
      Name: 'Fast-n-greasy Food Spot',
    },
  ]);
```

This can be a great way to insert rows into specific column headings without worrying about whether or not a user has
edited the spreadsheet to add their own columns, etc. that would otherwise cause inserting new rows to be painful.

SheetQuery will lookup the column headings, match them with the object keys provided, and insert a new row with an array
of values mapped to the correct index positions of the spreadsheet headings. Any heading/column values not provided will
be left blank.