#!/usr/bin/python3

import xlrd
import os
import csv


def get_path():
    path = os.getcwd() + '/' + 'Co-op Dashboard.xlsx'

    return path


def open_file(path):
    book = xlrd.open_workbook(path)

    return book


def get_sheet(workbook, name):
    sheet = workbook.sheet_by_name(name)

    return sheet


def get_header(sheet):
    header = sheet.row_values(0)

    return header


def to_columns(sheet, header, ignore):
    head = ('Date', 'Category', 'Value')
    csv_write(head, 'w')
    for i in range(len(header)):
        colname = sheet.cell(0, i)
        if header[i] not in ignore:
            numrows = sheet.nrows
            for row in range(numrows):
                if row == 0:
                    pass
                else:
                    rowdata = (sheet.cell(row, 0).value,
                               colname.value,
                               sheet.cell(row, i).value)
                    print(rowdata)
                    csv_write(rowdata, 'a')


def csv_write(rowdata, wtype):
    with open('pldata.csv', wtype, newline='') as csvfile:
        writer = csv.writer(csvfile, dialect='excel')
        writer.writerow(rowdata)


name = 'PL Data'
ignore = ['DateXform', 'Date', 'Month', 'Year', 'Quarter']
path = get_path()
xlfile = open_file(path)
sheet = get_sheet(xlfile, name)
header = get_header(sheet)


def main():
    name = 'PL Data'
    ignore = ['DateXform', 'Date', 'Month', 'Year', 'Quarter']
    path = get_path()
    xlfile = open_file(path)
    sheet = get_sheet(xlfile, name)
    header = get_header(sheet)
    to_columns(sheet, header, ignore)


if __name__ == '__main__':
    main()
