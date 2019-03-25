# *- coding:utf-8 -*-

import xlwt
import xlrd
from xlutils.copy import copy as xlcopy


def read_excel(file_url):
    try:
        data = xlrd.open_workbook(filename=file_url)
        return data
    except Exception as e:
        print(str(e))


def read_table(file_url):
    workbook = read_excel(file_url)
    return workbook.sheet_by_index(0)


def generate_sfa_input():

    for year in range(0, 11):
        ac_year = 2016 - year

        table = xlrd.open_workbook(filename='RFrontierInputFiles/_sfa_in' + str(ac_year) + '.xls')
        table2 = read_excel('pyDEAOutputFiles/out_dea' + str(ac_year) + '.xls').sheet_by_name('Targets')

        new_table = xlcopy(table)
        ws = new_table.get_sheet(0)

        ws.write(0, 6, 'Slack_CO2')
        ws.write(0, 7, 'Slack_WORK')
        ws.write(0, 8, 'Slack_CAPITAL')
        ws.write(0, 9, 'Slack_GDP')

        for one_row in range(1, 32):
            row = one_row
            if one_row == 26:
                continue
            if one_row > 26:
                row = one_row - 1

            row_num = 185 + (row - 1) * 6
            # Calculate CO2
            slack = table2.cell_value(row_num, 3) - table2.cell_value(row_num, 2)
            ws.write(row, 6, slack)

            # Calculate WORK
            slack = table2.cell_value(row_num + 1, 3) - table2.cell_value(row_num + 1, 2)
            ws.write(row, 7, slack)

            # Calculate CAPITAL
            slack = table2.cell_value(row_num + 2, 3) - table2.cell_value(row_num + 2, 2)
            ws.write(row, 8, slack)

            # Calculate GDP
            slack = table2.cell_value(row_num + 3, 3) - table2.cell_value(row_num + 3, 2)
            ws.write(row, 9, slack)

        new_table.save('RFrontierInputFiles/_sfa_in' + str(ac_year) + '.xls')

    return


# print(generate_sfa_input())
