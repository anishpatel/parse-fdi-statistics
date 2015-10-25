#!/usr/bin/env python3

import os
from collections import defaultdict

import xlrd
import xlwt

def isfloat(string):
  try:
    float(string)
    return True
  except ValueError:
    return False

def parse_sheet(sheet):
    region1 = sheet.cell_value(0, 0)
            
    ## Find column numbers for years ##
    r = 4  # row index
    c = 0  # column index
    year_cols = []
    years = []
    while not isfloat(sheet.cell_value(r, c)):
        c += 1
    while c < sheet.ncols and isfloat(sheet.cell_value(r, c)): 
        year_cols.append(c)
        years.append(sheet.cell_value(r, c))
        c += 1

    ## Iterate through rows ##
    r = 7
    while r < sheet.nrows:
        # Find second region name #
        c = 0
        while c < year_cols[0] and len(sheet.cell_value(r, c)) == 0:
            c += 1
        if c == year_cols[0]:
            # At empty row outside of table
            r += 6  # skip table break
            continue
        region2 = sheet.cell_value(r, c)
        assert len(region2) != 0, '{} {} {} {}'.format(sheet.name, r, c, region2)

        # Get values for this region #
        for c, year in zip(year_cols, years):
            val = sheet.cell_value(r, c)
            yield (region1, region2, year, val)

        r += 1  # go to next row

def parse_workbooks(dir_path):
    filenames = [fn for fn in sorted(os.listdir(dir_path)) if fn.endswith('.xls')]
    print('Found', len(filenames), 'workbooks')

    # e.g., data['inflows'] = [('United States', 'China', 2001, 535), ...]
    data = defaultdict(lambda: [])

    for filename in filenames:
        wb_path = os.path.join(dir_path, filename)
        with xlrd.open_workbook(wb_path) as wb:
            for sheet in wb.sheets():
                sheet_data = data[sheet.name]
                for elem in parse_sheet(sheet):
                    sheet_data.append(elem)

    return data

def write_to_workbook(data, filename):
    wb = xlwt.Workbook()

    for sheet_name, sheet_data in data.items():
        #TODO fix data cap
        sheet_data = sheet_data[:65536]
        sheet = wb.add_sheet(sheet_name)
        r = 0  # row index
        for region1, region2, year, val in sheet_data:
            sheet.write(r, 0, region1)
            sheet.write(r, 1, region2)
            sheet.write(r, 2, year)
            sheet.write(r, 3, val)
            r += 1

    wb.save(filename)


if __name__ == '__main__':

    input_dir = 'fdi-workbooks'
    output_filename = 'all_data.xls'

    print('==> Parsing data from', input_dir)
    data = parse_workbooks(input_dir)
    print('==> Parsing data done')

    print('==> Writing data to workbook', output_filename)
    write_to_workbook(data, output_filename)
    print('==> Writing data done')
