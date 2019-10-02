#!/usr/bin/env python

import csv
import os
import pprint
import fdb
from decimal import Decimal
from datetime import datetime
from datetime import timedelta
from platform import system
from excelopen import ExcelOpenDocument

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

if system() == 'Linux':
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


def read_firebird_database():
    stock = []
    con = fdb.connect(
        host=os.getenv('HOST'),
        database=os.getenv('DATABASE'),
        user=os.getenv('USER'),
        password=os.getenv("PASSWORD"),
        charset='WIN1252'
    )

    SELECT = """
    SELECT locationGroup.name AS "Group",
        COALESCE(partcost.avgcost, 0) AS averageunitcost,
        COALESCE(part.stdcost, 0) AS standardunitcost,
        locationgroup.name AS locationgroup,
        part.num AS partnumber,
        part.description AS partdescription,
        location.name AS location, asaccount.name AS inventoryaccount,
        uom.code AS uomcode, sum(tag.qty) AS qty, company.name AS company
    FROM part
        INNER JOIN partcost ON part.id = partcost.partid
        INNER JOIN tag ON part.id = tag.partid
        INNER JOIN location ON tag.locationid = location.id
        INNER JOIN locationgroup ON location.locationgroupid = locationgroup.id
        LEFT JOIN asaccount ON part.inventoryaccountid = asaccount.id
        LEFT JOIN uom ON uom.id = part.uomid
        JOIN company ON company.id = 1
    WHERE locationgroup.id IN (1)
    GROUP BY averageunitcost, standardunitcost, locationgroup, partnumber,
        partdescription, location, inventoryaccount, uomcode, company
    """

    cur = con.cursor()
    cur.execute(SELECT)

    for (group, avgcost, stdcost, locationgroup, partnum, partdescription,
         location, invaccount, uom, qty, company) in cur:
        stock.append([
            location,
            partnum,
            partdescription,
            str(Decimal(str(qty)).quantize(Decimal("1.00"))),
            uom,
            str(Decimal(str(avgcost)).quantize(Decimal("1.00"))),
        ])

    stock = sorted(stock, key=lambda k: (k[0], k[1]))
    return stock


def read_csv_file():
    # initializing the titles and rows list
    rows = []

    with open(csv_file, encoding="utf8") as csvfile:
        # creating a csv reader object
        csvreader = csv.reader(csvfile)
        for i in range(6):
            ignore = next(csvreader)  # noqa: F841
        location = ""

        count = 6
        for row in csvreader:
            count += 1
            if row[0] == "Grand Total":
                ignore = next(csvreader)  # noqa: F841
            elif row[0]:
                location = row[0]
            elif row[1]:
                rows.append([location,
                             row[1], row[2], row[10], row[12],
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
    excel = ExcelOpenDocument()
    excel.new(xlsx_file)
    title_font = excel.font(name='Arial', size=10, bold=True)
    body_font = excel.font(name='Arial', size=10)

    for column, value in enumerate(fields, start=1):
        excel.cell(row=1, column=column).value = value
        excel.cell(row=1, column=column).font = title_font

    for column, width in enumerate(widths, start=65):
        excel.set_width(chr(column), width)

    for row, all_fields in enumerate(filter(filterWarehouse, rows), start=2):
        for column, field in enumerate(all_fields, start=1):
            if formats[column-1] == 'General':
                value = field
            else:
                value = float(field.replace(",", ""))
            cell = excel.cell(row=row, column=column)
            cell.value = value
            cell.number_format = formats[column-1]
            cell.font = body_font

        excel.cell(row=row, column=7).value = "=SUM(D{}*F{})".format(row, row)
        excel.cell(row=row, column=7).font = body_font
        excel.cell(row=row, column=7).number_format = formats[6]

    row = excel.max_row() + 2
    excel.cell(row=row, column=5).value = 'Grand Total:'
    excel.cell(row=row, column=5).font = title_font
    excel.cell(row=row, column=7).value = "=SUM(G2:G{})".format(row - 2)
    excel.cell(row=row, column=7).font = title_font
    excel.cell(row=row, column=7).number_format = formats[6]

    excel.save()


def main():
    # rows = read_csv_file()
    # pp = pprint.PrettyPrinter(indent=2)
    # pp.pprint(rows[15])
    rows = read_firebird_database()
    write_xlsx_file(rows)
    # pp = pprint.PrettyPrinter(indent=2)
    # pp.pprint(rows[15])


if __name__ == "__main__":
        main()

