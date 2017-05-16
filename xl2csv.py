#!/usr/bin/python3

import xlrd
import os
import csv


def get_path(xlsxname=None):
    '''Working directory and xlsx file name '''

    if not xlsxname:
        xlsxname = 'Co-op Dashboard.xlsx'

    path = os.getcwd() + '/' + xlsxname

    return path


def open_file(path):
    '''Create workbook object'''

    workbook = xlrd.open_workbook(path)

    return workbook


def get_sheet(workbook, name):
    '''Create sheet object by sheet name'''

    sheet = workbook.sheet_by_name(name)

    return sheet


def get_header(sheet):
    '''First row of the excel sheet becomes header in csv'''

    header = sheet.row_values(0)

    return header


def to_columns(sheet, header, ignore):
    '''Write head row, write csv in 3 columns
        date, column name, data'''

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


def csv_write(rowdata, wtype, csvname=None):
    '''With tuple or list input, write csv file row'''

    if not csvname:
        csvname = 'pldata.csv'

    try:
        with open(csvname, wtype, newline='') as csvfile:
            writer = csv.writer(csvfile, dialect='excel')
            writer.writerow(rowdata)

    except:
        print('Error writing')

    return True


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
