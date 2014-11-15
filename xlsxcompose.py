__author__ = 'perks'

from xlutils.copy import copy
from xlrd import open_workbook
import xlsxwriter

def compose(input, output, mappings):
    START_ROW = 501
    END_ROW = 1000

    rb = open_workbook(input)
    r_sheet = rb.sheet_by_name("CLEAN")

    workbook = xlsxwriter.Workbook(output)
    worksheet = workbook.add_worksheet("Clients")

    dic = {}
    print "gothere"

    for col_index, mapping in enumerate(mappings):
        target_header = mapping[0]
        orig_header = mapping[1]

        for i in range(1, r_sheet.ncols):
            col_values = filter(None, r_sheet.col(i))
            col_values = map(lambda x: str(int(x.value)) if x.ctype == 2 else x.value, col_values)
            col_head = col_values[0]

            if orig_header == col_head:
                migrate_col = [target_header] + col_values[START_ROW:END_ROW]
                dic.update({target_header: migrate_col})
                break

        if dic.has_key(target_header):
            for row_index, cell in enumerate(dic[target_header]):
                cell_write(worksheet, row_index, col_index, cell)

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
            '-m',
            '--mappings', 
            help='File with map configurations inform of TargetCol=OriginalCol',
            required=True
        )
    args = parser.parse_args()

    try:
        lines = [line.strip() for line in open(args.mappings)]
        mappings = [tuple(mapping.split("=")) for mapping in lines if mapping.split("=")[1]]
        compose(args.input, args.output, mappings)
    except Exception,e:
        print "Error parsing your column mappings"
        print e


    print "Succesfull composition:\n\t {} => {}".format(args.input, args.output)







