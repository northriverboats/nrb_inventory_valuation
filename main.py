#!/usr/bin/env python

import csv
import os
from datetime import datetime
from datetime import timedelta
from excelopen import ExcelOpenDocument
import platform

"""
print("{:12.12} {:20.20} {:40.40}  {:12.12}  {:8.8} {:9.9}  {:9.9}".format(
    fields[0], fields[1], fields[2], fields[3], fields[4], fields[5],
    fields[6])
)

for row in rows:
  print("{:12.12} {:20.20} {:40.40}  {:12.4f}  {:8.8} {:9.2f} {:9.2f}".format(
         row[0], row[1], row[2], float(row[3]), row[4], float(row[5]),
         float(row[6])))
"""

inventoried = (datetime.today() - timedelta(days=15))
qtr = int(inventoried.month / 3)
ith = ['', '1st', '2nd', '3rd', '4th']
quarter = ith[qtr]
title = quarter + ' Quarter ' + str(inventoried.year)

os.getenv("LINUXXLSDIR")

if platform.system() == 'Linux':
    csv_file = os.getenv("LINUXCSVFILE")
    file_name = os.path.join(os.getenv("LINUXXLSDIR"), title)
    xlsx_file = file_name + " " + os.getenv("LINUXXLSXFILE")
else:
    csv_file = os.getenv("WINDOWSCSVFILE")
    file_name = os.path.join(os.getenv("WINDOWSXLSDIR"), title)
    xlsx_file = file_name + " " + os.getenv("WINDOWSXLSXFILE")


fields = ['Location', 'Part', 'Description', 'Qty', 'UOM', 'Cost',
          'Extended']

formats = ['General', 'General', 'General', '0.00', 'General',
           '[$$-409]#,##0.00;[RED]\-[$$-409]#,##0.00',
           '[$$-409]#,##0.00;[RED]\-[$$-409]#,##0.00']

widths = [16.25, 34.25, 80.50, 7.50, 6.50, 10, 12.75]


def read_csv_file():
    # initializing the titles and rows list
    rows = []

    with open(csv_file, 'r') as csvfile:
        # creating a csv reader object
        csvreader = csv.reader(csvfile)
        for i in range(6):
            ignore = next(csvreader)  # noqa: F841
        location = ""
        count = 0
        for row in csvreader:
            if row[0]:
                location = row[0]
                count = 0
            if row[1] == "":
                count += 1
                continue
            if count > 2:
                continue
            rows.append([location, row[1], row[2], row[10], row[12],
                         row[14].replace(",", "")[2:],
                         row[15].replace(",", "")[2:]])

    rows.sort(key=lambda l: (l[0], l[1]))
    return rows


def filterWarehouse(row):
    if row[0] not in ['Upholstery', 'Apparel']:
        return True
    else:
        return False


def filterUpholstery(row):
    if row[0] == 'Upholstery':
        return True
    else:
        return False


def filterApparel(row):
    if row[0] == 'Apparel':
        return True
    else:
        return False


def write_xlsx_file(rows):
    # create new workbook
    workbook = Workbook()
    sheet = workbook.active

    title_font = Font(name='Arial', size=10, bold=True)
    body_font = Font(name='Arial', size=10)

    for column, value in enumerate(fields, start=1):
        sheet.cell(row=1, column=column).value = value
        sheet.cell(row=1, column=column).font = title_font

    for column, width in enumerate(widths, start=65):
        sheet.column_dimensions[chr(column)].width = width

    for row, all_fields in enumerate(filter(filterWarehouse, rows), start=2):
        for column, field in enumerate(all_fields, start=1):
            if formats[column-1] == 'General':
                value = field
            else:
                value = float(field)
            cell = sheet.cell(row=row, column=column)
            cell.value = value
            cell.number_format = formats[column-1]
            cell.font = body_font
        sheet.cell(row=row, column=7).value = "=SUM(D{}*F{}".format(row, row)

    # save workbook before exiting
    print(sheet.max_row)
    workbook.save(xlsx_file)


def main():
    rows = read_csv_file()
    write_xlsx_file(rows)


if __name__ == "__main__":
        main()
