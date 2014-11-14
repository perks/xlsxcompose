import argparse
__author__ = 'perks'

if __name__ == '__main__':

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
            'mapping', 
            nargs='+',
            help='Comma separated column relations in the form of OriginalCol=TargetCol'
        )
    args = parser.parse_args()

    print args


