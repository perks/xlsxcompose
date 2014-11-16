#!/usr/bin/env python
__author__ = 'perks'

from xlutils.copy import copy
from xlrd import open_workbook
import xlsxwriter
import sys
import traceback

def compose(input, output, start_row, end_row, mappings, ss, ts):

    START_ROW = start_row 
    END_ROW = end_row

    rb = open_workbook(input)
    r_sheet = rb.sheet_by_name(ss)
    ref = [x[1] for x in mappings]
    filtered_cols = filter(lambda tup: tup[1] in ref, [(i, x) for i, x in enumerate(r_sheet.row_values(0))])


    columns = {}

    for mapping in mappings:
        target_header = mapping[0]
        orig_header = mapping[1]

        for (i, headers) in filtered_cols:
            col_values = filter(None, r_sheet.col(i))
            col_values = map(lambda x: str(int(x.value)) if x.ctype == 2 else x.value, col_values)
            col_head = col_values[0]

            if orig_header == col_head:
                columns.update({target_header: col_values})
                break


    count = 0
    while (count < 3):
        print "top loop"
        file_name = "{}_{}".format(count , output)
        workbook = xlsxwriter.Workbook(file_name)
        worksheet = workbook.add_worksheet(ts)

        for col_index, mapping in enumerate(mappings):
            target_header = mapping[0]
            orig_header = mapping[1]

            if columns.has_key(target_header):
                write_col = [target_header] + (columns[target_header][START_ROW:END_ROW] if END_ROW else columns[target_header][START_ROW:])
                for row_index, cell in enumerate(write_col):
                    cell_write(worksheet, row_index, col_index, cell)
            else:
                cell_write(worksheet, 0, col_index, target_header)

        workbook.close()
        count += 1

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
            type=int,
            default=0
        )
    parser.add_argument(
            '-e',
            '--end',
            help='Final row number (Default = all rows)',
            type=int,
            default=0
        )
    parser.add_argument(
            '-m',
            '--mappings', 
            help='File with map configurations inform of TargetCol=OriginalCol',
            required=True
        )
    parser.add_argument(
            '-ss',
            '--sourcesheet',
            help='Sheet reference in original workbook (Default=Sheet1)',
            default='Sheet1'
        )
    parser.add_argument(
            '-ts',
            '--targetsheet',
            help='Target name of sheet in workbook (Default=Sheet1)',
            default='Sheet1'
        )

    args = parser.parse_args()

    try:
        lines = [line.strip() for line in open(args.mappings)]
        mappings = [tuple(mapping.split("=")) for mapping in lines if mapping.split("=")[1]]
        compose(args.input, args.output, args.start, args.end, mappings, args.sourcesheet, args.targetsheet)
    except Exception,e:
        print traceback.format_exc()


    print "Succesfull composition:\n\t {} => {}".format(args.input, args.output)







