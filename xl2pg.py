from argparse import ArgumentParser
import json
import openpyxl
from openpyxl.worksheet._write_only import WriteOnlyWorksheet
import psycopg2
from lib import dbconfig
##

def parse_args():
    argp = ArgumentParser()
    argp.add_argument('spreadsheet', help='xlsx file to upload')
    argp.add_argument('-d', '--db', help='Database config as a json file: { "dbname": "", "user": "", "password": "" }')
    argp.add_argument('-m', '--map', required=True, help='Json mapping from excel fields to postgres table')
    argp.add_argument('--clear', action='store_true', help='perform DELETE FROM target table')
    return argp.parse_args()
args = parse_args()
##

db = dbconfig.load_or_prompt(args.db)
##

def load_map(path):
    with open(path) as file:
        return json.load(file)
map_cfg = load_map(args.map)

target_table = map_cfg['target_table']
sheet_number = map_cfg['sheet']
skip_rows = map_cfg['skip_rows']
mappings = map_cfg['mapping']
##

def generate_prepare_statement(table_name, mappings):
    field_names = []
    variables = []

    for index, field in enumerate(mappings.keys()):
        field_names.append(field)
        variables.append('$' + str(index + 1))

    return '''
PREPARE xl2pg_insert AS
INSERT INTO {table}(
    {fields}
) VALUES (
    {values}
);
    '''.format(
        table=table_name,
        fields=',\n    '.join(field_names),
        values=',\n    '.join(variables)
    )
##

prep_statement = generate_prepare_statement(target_table, mappings)
print('Generated statement:')
print(prep_statement)
print()
##

def generate_execute_statement(mappings):
    return 'EXECUTE xl2pg_insert(' + ', '.join(['%s'] * len(mappings)) + ')'
exec_statement = generate_execute_statement(mappings)
##

print('Connecting to', db.dbname)
with psycopg2.connect(dbname=db.dbname, user=db.user, password=db.password, host=db.host, port=db.port) as conn:
    with conn.cursor() as cursor:
        print('Preparing insert statement')
        cursor.execute(prep_statement)

        print('Loading spreadsheet')
        wb = openpyxl.load_workbook(args.spreadsheet)
        sheet = wb.worksheets[sheet_number]
        assert not isinstance(sheet, WriteOnlyWorksheet), 'Worksheet is write only'

        if args.clear:
            print('Clearing', target_table)
            cursor.execute(f'DELETE FROM {target_table};')


        max_row = sheet.max_row
        assert max_row, 'Cannot get max_row of spreadsheet'

        for row_index in range(skip_rows + 1, max_row + 1):
            output_row = []

            for column_index in mappings.values():
                value = sheet.cell(row_index, column_index).value
                output_row.append(value)
            print('Inserting row', row_index, 'out of', max_row, end='\r', flush=True)
            cursor.execute(exec_statement, output_row)
        print()

    conn.commit()
