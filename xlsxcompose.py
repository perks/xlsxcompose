__author__ = 'perks'

from xlutils.copy import copy
from xlrd import open_workbook
import xlsxwriter

def compose(input, output, start_row, end_row, mappings):

    START_ROW = int(start_row) + 1
    END_ROW = int(end_row) or False

    rb = open_workbook(input)
    r_sheet = rb.sheet_by_name("CLEAN")

    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet("Clients")

    columns = {}

    for col_index, mapping in enumerate(mappings):
        target_header = mapping[0]
        orig_header = mapping[1]

        for i in range(1, r_sheet.ncols):
            col_values = filter(None, r_sheet.col(i))
            col_values = map(lambda x: str(int(x.value)) if x.ctype == 2 else x.value, col_values)
            col_head = col_values[0]

            if orig_header == col_head:
                migrate_col = [target_header] + (col_values[START_ROW:END_ROW] if END_ROW else col_values[START_ROW:])
                columns.update({target_header: migrate_col})
                break

        if columns.has_key(target_header):
            for row_index, cell in enumerate(columns[target_header]):
                cell_write(worksheet, row_index, col_index, cell)
        else:
            cell_write(worksheet, 0, col_index, target_header)

    workbook.close()

def cell_write(sheet,row_index, col_index, value):
        sheet.write(row_index, col_index, value)

if __name__ == '__main__':

    import argparse

    parser = argparse.ArgumentParser(
            description='Migrate columns from one spreadsheet to columns in a new spreadsheet.'
        )
    parser.add_argument(
            '-i',
            '--input',
            help='Input .xlsx file', 
            required=True
        )
    parser.add_argument(
            '-o',
            '--output',
            help='Output .xlsx file name',
            default='xlsxcompose.xlsx'
        )
    parser.add_argument(
            '-s',
            '--start',
            help='Starting row number (Default = 0)',
            default=0
        )
    parser.add_argument(
            '-e',
            '--end',
            help='Final row number (Default = all rows)',
            default=None
        )
    parser.add_argument(
            '-m',
            '--mappings', 
            help='File with map configurations inform of TargetCol=OriginalCol',
            required=True
        )

    args = parser.parse_args()

    try:
        lines = [line.strip() for line in open(args.mappings)]
        mappings = [tuple(mapping.split("=")) for mapping in lines if mapping.split("=")[1]]
        compose(args.input, args.output, args.start, args.end, mappings)
    except Exception,e:
        print "Error parsing your column mappings"
        print e


    print "Succesfull composition:\n\t {} => {}".format(args.input, args.output)







