xlsxcompose.py
========

### Simple python script for rearranging and partitioning excel spreadsheets

Useful for large spreadsheets to automate the tedious copy/pasting of columns
from one sheet to another

##### For feature requests, open an issue!



---

###Documentation

####Dependencies:

* Python 2.7 >=
  * [xlrd](https://github.com/python-excel/xlrd)
  * [xlsxwriter](https://xlsxwriter.readthedocs.org/index.html)

###Usage:
```zsh
usage: xlsxcompose.py [-h] -i INPUT [-o OUTPUT] [-s START] [-e END] [-l LIMIT]
                      -m MAPPINGS [-ss SOURCESHEET] [-ts TARGETSHEET]

Migrate columns from one spreadsheet to columns in a new spreadsheet.

optional arguments:
  -h, --help            show this help message and exit
  -i INPUT, --input INPUT
                        Input .xlsx file
  -o OUTPUT, --output OUTPUT
                        Output .xlsx file name
  -s START, --start START
                        Starting row number (Default = 0)
  -e END, --end END     Final row number (Default = all rows)
  -l LIMIT, --limit LIMIT
                        Final row number to step over interval of --start to
                        --end
  -m MAPPINGS, --mappings MAPPINGS
                        File with map configurations inform of
                        TargetCol=OriginalCol
  -ss SOURCESHEET, --sourcesheet SOURCESHEET
                        Sheet reference in original workbook (Default=Sheet1)
  -ts TARGETSHEET, --targetsheet TARGETSHEET
                        Target name of sheet in workbook (Default=Sheet1)
```

## Example configuration file:

```
First Name=Guest First Name
Last Name=Guest Last Name
Status (VIP, PROMO, or 86)=NONE
Gender (MALE or FEMALE)=NONE
Phone Number=Mobile Phone
Phone Number Locale (US or INTL)=NONE
Work Phone Nuber=Home Phone
Work Phone Number Locale (US or INTL)=NONE
Email=Guest Email Address
Birthday (MM/DD/YYYY)=Birthday
Notes=Guest Notes
Private Notes=NONE
Title=NONE
Company=Guest Company
Address 1=Address 1
Address 2=Address 2
City=City
State=State
Postal Code=Zip Code
Country=Country
```

Note that if you wish to include a column that does not have a corresponding
column in the original document, then you must set it equal to `NONE` or another
field that will fail to find a match. When this happens, the target header will
be written, but the column will be empty.

Also important to know that the order of your columns is based on the order they
appear in the configuration file.

---

### Suggestions/Feedback

Always welcome! Feel free to give a shout to [@luckycevans](http://twitter.com/luckycevans)

### Copyright

The MIT License (MIT)

Copyright (c) 2014 Chris Evans

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
