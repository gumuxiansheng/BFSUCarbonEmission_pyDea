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


def sfa_result_rearrange():
    wbw = xlwt.Workbook(encoding='utf-8', style_compression=0)
    for year in range(0, 11):
        ac_year = 2016 - year
        ws = wbw.add_sheet(str(ac_year), cell_overwrite_ok=True)
        f = open('RFrontierOutputFiles/_sfa_out' + str(ac_year) + '.txt', encoding='utf-8')  # 返回一个文件对象
        line = f.readline()
        row_num = 0
        while line:
            line = f.readline()
            ws.write(row_num, 0, line)
            row_num += 1
        f.close()
    wbw.save('RFrontierOutputFiles/_sfa_out.xls')
    return


def sfa_result_xx_rearrange():
    wbw = xlwt.Workbook(encoding='utf-8', style_compression=0)
    for year in range(0, 11):
        ac_year = 2016 - year
        ws = wbw.add_sheet(str(ac_year), cell_overwrite_ok=True)
        f = open('RFrontierOutputFiles/_sfa_out_xx' + str(ac_year) + '.txt', encoding='utf-8')  # 返回一个文件对象
        line = f.readline()
        row_num = 0
        while line:
            line = f.readline()
            ws.write(row_num, 0, line)
            row_num += 1
        f.close()
    wbw.save('RFrontierOutputFiles/_sfa_out_xx.xls')
    return


def generate_3rd_dea_input():
    wb = xlrd.open_workbook(filename='RFrontierOutputFiles/sfa_results.xlsx')
    for year in range(0, 11):
        ac_year = 2016 - year
        table = wb.sheet_by_name(str(ac_year))

        wbw = xlwt.Workbook(encoding='utf-8', style_compression=0)
        ws = wbw.add_sheet('Carbon Emission', cell_overwrite_ok=True)

        ws.write(0, 1, 'CO2')
        ws.write(0, 2, 'CAPITAL')
        ws.write(0, 3, 'LABOUR')
        ws.write(0, 4, 'CO2_MAX')
        ws.write(0, 5, 'CAPITAL_MAX')
        ws.write(0, 6, 'LABOUR_MAX')

        table2 = read_excel('RFrontierInputFiles/_sfa_in' + str(ac_year) + '.xls').sheet_by_index(0)
        max_co2 = 0
        max_capital = 0
        max_labour = 0
        for row in range(1, 31):
            ws.write(row, 0, table2.cell_value(row, 0))

            # CO2
            beta_0 = table.cell_value(13, 1)
            beta_1 = table.cell_value(14, 1)
            beta_2 = table.cell_value(15, 1)
            beta_3 = table.cell_value(16, 1)
            beta_4 = table.cell_value(17, 1)
            beta_5 = table.cell_value(18, 1)

            ex_value = beta_0 + beta_1 * table2.cell_value(row, 1) + beta_2 * table2.cell_value(row, 2) \
                       + beta_3 * table2.cell_value(row, 3) + beta_4 * table2.cell_value(row, 4) \
                       + beta_5 * table2.cell_value(row, 5)
            if ex_value > max_co2:
                max_co2 = ex_value
            ws.write(row, 1, ex_value)

            # CAPITAL
            beta_0 = table.cell_value(41, 1)
            beta_1 = table.cell_value(42, 1)
            beta_2 = table.cell_value(43, 1)
            beta_3 = table.cell_value(44, 1)
            beta_4 = table.cell_value(45, 1)
            beta_5 = table.cell_value(46, 1)

            ex_value = beta_0 + beta_1 * table2.cell_value(row, 1) + beta_2 * table2.cell_value(row, 2) \
                       + beta_3 * table2.cell_value(row, 3) + beta_4 * table2.cell_value(row, 4) \
                       + beta_5 * table2.cell_value(row, 5)
            if ex_value > max_capital:
                max_capital = ex_value
            ws.write(row, 2, ex_value)

            # LABOUR
            beta_0 = table.cell_value(69, 1)
            beta_1 = table.cell_value(70, 1)
            beta_2 = table.cell_value(71, 1)
            beta_3 = table.cell_value(72, 1)
            beta_4 = table.cell_value(73, 1)
            beta_5 = table.cell_value(74, 1)

            ex_value = beta_0 + beta_1 * table2.cell_value(row, 1) + beta_2 * table2.cell_value(row, 2) \
                       + beta_3 * table2.cell_value(row, 3) + beta_4 * table2.cell_value(row, 4) \
                       + beta_5 * table2.cell_value(row, 5)
            if ex_value > max_labour:
                max_labour = ex_value
            ws.write(row, 3, ex_value)

        ws.write(1, 4, max_co2)
        ws.write(1, 5, max_capital)
        ws.write(1, 6, max_labour)

        wbw.save('RFrontierOutputFiles/_sfa_ex_' + str(ac_year) + '.xls')
    return
