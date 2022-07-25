# xl2pg
Upload an excel spreadsheet to postgres

## Usage
`python xl2pg.py spreadsheet.xlsx -d db.json --map map.json`

Arguments:
- `-m` `--map` - json file mapping postgres table fields to excel columns
- `-d` `--db` - json file specifying dbname, user and password, etc. (see https://www.postgresql.org/docs/current/libpq-connect.html#LIBPQ-PARAMKEYWORDS)

Optional arguments:
- `--clear` - clear the target table before inserting data

## Configuration

### map.json example

```json
{
  "target_schema": "public",
  "target_table": "table_name",
  "sheet": 0, // Use the first (zeroth) sheet in the excel file
  "skip_rows": 1, // Skip the headings row in the spreadsheet
  "mapping": {
    "field1": 1, // key is a field inside target_table,
    "field2": 3, // value is the corresponding column in the spreadsheet
    "field3": 5,
    ...
  }
}
```

## Dependencies
- openpyxl
- psycopg[binary]
- typing-extensions
