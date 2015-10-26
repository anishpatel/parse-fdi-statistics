#!/usr/bin/env python3

import os
from collections import defaultdict

def isfloat(string):
  try:
    float(string)
    return True
  except ValueError:
    return False

def parse_sheet(sheet, region1=None):
    region1 = region1 or sheet.cell_value(0, 0)
            
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
    import xlrd

    filenames = [fn for fn in sorted(os.listdir(dir_path)) if fn.endswith('.xls')]
    print('Found', len(filenames), 'workbooks')

    # e.g., data['inflows'] = [('United States', 'China', 2001, 535), ...]
    data = defaultdict(lambda: [])

    for filename in filenames:
        wb_path = os.path.join(dir_path, filename)
        with xlrd.open_workbook(wb_path) as wb:
            ## Figure out consistent name for region1 ##
            region1_candidates = set(s.cell_value(0, 0).strip() for s in wb.sheets() if s.cell_value(0, 0))
            if len(region1_candidates) == 1:
                ## Only 1 unique region name found, so use that ##
                region1 = region1_candidates.pop()
            elif len(region1_candidates) < 1:
                ## No region name found ##
                # Try to find 3-letter country code in filename
                i1 = filename.rfind('_')
                i2 = filename.rfind('.')
                if i1 >= 0 and i2 >= 0 and i1 < i2:
                    region1 = filename[i1+1:i2]
                else:
                    # Couldn't find region name, so just use filename
                    region1 = filename
            elif len(region1_candidates) > 1:
                ## Multiple different region names found ##
                # Remove any that contain commas
                region1_candidates_2 = set(r1 for r1 in region1_candidates if not ',' in r1)
                if len(region1_candidates_2) == 1:
                    # Only 1 unique region name found, so use that
                    region1 = region1_candidates_2.pop()
                elif len(region1_candidates_2) < 1:
                    # No region name without commas exists, so just pick one
                    region1 = region1_candidates.pop()
                elif len(region1_candidates_2) > 1:
                    # Multiple different region names found, so just pick one
                    region1 = region1_candidates_2.pop()
            assert region1, 'ERROR: Origin region not found for file "%s"' % filename

            for sheet in wb.sheets():
                sheet_data = data[sheet.name]
                for elem in parse_sheet(sheet, region1):
                    sheet_data.append(elem)

    return data


def write_xls(wb_data, filename):
    import xlwt
    wb = xlwt.Workbook()

    for sheet_name, rows in wb_data.items():
        #TODO fix data cap by using xlsx instead of xls
        if len(rows) > 65536:
            print('WARNING: Capping rows for sheet "%s" (xls only supports 65536 rows per sheet)' % sheet_name) 
            rows = rows[:65536]
        sheet = wb.add_sheet(sheet_name)
        for r, row in enumerate(rows):
            for c, val in enumerate(row):
                sheet.write(r, c, val)

    wb.save(filename + ('.xls' if not filename.endswith('.xls') else ''))


def write_csv(rows, filename):
    if not filename.endswith('.csv'):
        filename += '.csv'

    with open(filename, 'w') as f:
        for row in rows:
            f.write(','.join((str(v) for v in row))+'\n')


if __name__ == '__main__':

    input_dir = 'fdi-workbooks'
    output_filename_prefix = 'all_data'

    print('==> Parsing data from', input_dir)
    data = parse_workbooks(input_dir)
    print('==> Parsing data done')

    for sheet_name, rows in data.items():
        filename = '%s-%s.csv' % (output_filename_prefix, sheet_name)
        print('==> Writing data to csv', filename)
        write_csv(rows, filename)
    print('==> Writing data done')
