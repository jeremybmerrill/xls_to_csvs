#!/usr/bin/env python
# coding=utf8

import xlrd
import csv
import os, sys

def csv_from_excel(workbook_path, output_dir='.'):
  wb = xlrd.open_workbook(workbook_path)
  for sheet in wb.sheets():
    sheet_csv_path = os.path.join(output_dir, os.path.splitext(os.path.basename(workbook_path))[0] + "_" + sheet.name + ".csv")
    with open(sheet_csv_path, 'wb') as csv_file:
      csv_writer = csv.writer(csv_file, quoting=csv.QUOTE_ALL)

      for rownum in xrange(sheet.nrows):
          csv_writer.writerow(sheet.row_values(rownum))


if not sys.argv[1]:
  print "Convert an XLSX into one csv per sheet.\n\n./xlsx_to_csvs.py input_excel_path.xlsx [output_dir]"
else:
  csv_from_excel(sys.argv[1], sys.argv[2] if len(sys.argv) >= 3 else '.')
