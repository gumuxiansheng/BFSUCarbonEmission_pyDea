# *- coding:utf-8 -*-

import xlwt
import xlrd
from xlutils.copy import copy as xlcopy
from scipy.stats import norm
import numpy as np


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

        slack_co2_col = 10
        slack_work_col = slack_co2_col + 1
        slack_capital_col = slack_work_col + 1
        slack_gdp_col = slack_capital_col + 1

        ws.write(0, slack_co2_col, 'Slack_CO2')
        ws.write(0, slack_work_col, 'Slack_WORK')
        ws.write(0, slack_capital_col, 'Slack_CAPITAL')
        ws.write(0, slack_gdp_col, 'Slack_GDP')

        for one_row in range(1, 32):
            row = one_row
            if one_row == 26:
                continue
            if one_row > 26:
                row = one_row - 1

            row_num = 185 + (row - 1) * 6
            # Calculate CO2
            slack = table2.cell_value(row_num, 3) - table2.cell_value(row_num, 2)
            ws.write(row, slack_co2_col, slack)

            # Calculate WORK
            slack = table2.cell_value(row_num + 1, 3) - table2.cell_value(row_num + 1, 2)
            ws.write(row, slack_work_col, slack)

            # Calculate CAPITAL
            slack = table2.cell_value(row_num + 2, 3) - table2.cell_value(row_num + 2, 2)
            ws.write(row, slack_capital_col, slack)

            # Calculate GDP
            slack = table2.cell_value(row_num + 3, 3) - table2.cell_value(row_num + 3, 2)
            ws.write(row, slack_gdp_col, slack)

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
        is_entering_estimates = False
        while line:
            line = f.readline()
            if line.startswith('---'):
                is_entering_estimates = False

            if not is_entering_estimates:
                ws.write(row_num, 0, line)
            else:
                str1 = line[0:19]
                str2 = line[19:32]
                str3 = line[32:43]
                str4 = line[43:55]
                str5 = line[55:65]
                str6 = line[65:]
                ws.write(row_num, 0, str1)
                ws.write(row_num, 1, str2)
                ws.write(row_num, 2, str3)
                ws.write(row_num, 3, str4)
                ws.write(row_num, 4, str5)
                ws.write(row_num, 5, str6)

            if line.startswith('final maximum'):
                is_entering_estimates = True

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
        entering_status = 0
        while line:
            line = f.readline()
            if line.startswith('$fitted'):
                entering_status = 1
                ws.write(row_num, 0, line)
                row_num += 1
                continue
            elif line.startswith('$resid'):
                entering_status = 2
                ws.write(row_num, 0, line)
                row_num += 1
                continue
            elif line == '\n':
                entering_status = 0
                ws.write(row_num, 0, line)
                row_num += 1
                continue

            if entering_status == 0:
                ws.write(row_num, 0, line)
            elif entering_status == 1:
                str1 = line[0:6]
                str2 = line[6:]
                ws.write(row_num, 0, str1)
                ws.write(row_num, 1, str2)
            elif entering_status == 2:
                str1 = line[0:3]
                str2 = line[3:]
                ws.write(row_num, 0, str1)
                ws.write(row_num, 1, str2)
            row_num += 1
        f.close()
    wbw.save('RFrontierOutputFiles/_sfa_out_xx.xls')
    return


def generate_3rd_dea_input_cal():
    wb = xlrd.open_workbook(filename='RFrontierOutputFiles/_sfa_out_xx.xls')
    wbw = xlwt.Workbook(encoding='utf-8', style_compression=0)
    for year in range(0, 11):
        ac_year = 2016 - year
        table = wb.sheet_by_name(str(ac_year))

        ws = wbw.add_sheet(str(ac_year), cell_overwrite_ok=True)

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
            ex_value = float(table.cell_value(5 + row - 1, 1))
            if ex_value > max_co2:
                max_co2 = ex_value
            ws.write(row, 1, ex_value)

            # CAPITAL
            ex_value = float(table.cell_value(73 + row - 1, 1))
            if ex_value > max_capital:
                max_capital = ex_value
            ws.write(row, 2, ex_value)

            # LABOUR
            ex_value = float(table.cell_value(141 + row - 1, 1))
            if ex_value > max_labour:
                max_labour = ex_value
            ws.write(row, 3, ex_value)

        ws.write(1, 4, max_co2)
        ws.write(1, 5, max_capital)
        ws.write(1, 6, max_labour)

    wbw.save('RFrontierOutputFiles/_3rd_dea_input_cal.xls')  # 拟合值、残值信息
    return


def cal_vi():
    wbw = xlrd.open_workbook(filename='RFrontierOutputFiles/_sfa_out_xx.xls')  # 获取epsilon
    wbw2 = xlrd.open_workbook(filename='RFrontierOutputFiles/_sfa_out.xls')  # 获取sigma and lambda

    dst_file = xlrd.open_workbook(filename='RFrontierOutputFiles/_3rd_dea_input_cal.xls')  # 写入统一文件
    dst_file_new = xlcopy(dst_file)

    sigma_index = []
    lambda_index = []
    table_temp = wbw2.sheet_by_name(str(2016))
    for i in range(0, table_temp.nrows):
        if str(table_temp.cell_value(i, 0)).startswith('sigma '):
            sigma_index.append(i)
        elif str(table_temp.cell_value(i, 0)).startswith('lambda '):
            lambda_index.append(i)

    for year in range(0, 11):
        ac_year = 2016 - year
        table = wbw.sheet_by_name(str(ac_year))
        table2 = wbw2.sheet_by_name(str(ac_year))
        dst_ws = dst_file_new.get_sheet(str(ac_year))

        dst_ws.write(0, 7, 'CO2_Vi')
        dst_ws.write(0, 8, 'CAPITAL_Vi')
        dst_ws.write(0, 9, 'LABOUR_Vi')
        dst_ws.write(0, 10, 'CO2_Vi_MAX')
        dst_ws.write(0, 11, 'CAPITAL_Vi_MAX')
        dst_ws.write(0, 12, 'LABOUR_Vi_MAX')

        max_co2 = -200000
        max_capital = -200000
        max_labour = -200000

        print(str(ac_year))
        for row in range(1, 31):
            # CO2
            epsilon_ = float(table.cell_value(38 + row - 1, 1))
            sigma_ = float(table2.cell_value(sigma_index[0], 1))
            lambda_ = float(table2.cell_value(lambda_index[0], 1))
            norm_divide = norm.pdf(epsilon_ * lambda_ / sigma_) / (norm.cdf(epsilon_ * lambda_ / sigma_))
            if np.isnan(norm_divide):
                norm_divide = -epsilon_ * lambda_ / sigma_
            ui = lambda_ * sigma_ / (1 + lambda_ ** 2) * (norm_divide + epsilon_ * lambda_ / sigma_)
            vi = epsilon_ - ui
            if vi > max_co2:
                max_co2 = vi
            dst_ws.write(row, 7, vi)

            # CAPITAL
            epsilon_ = float(table.cell_value(106 + row - 1, 1))
            sigma_ = float(table2.cell_value(sigma_index[1], 1))
            lambda_ = float(table2.cell_value(lambda_index[1], 1))
            norm_divide = norm.pdf(epsilon_ * lambda_ / sigma_) / (norm.cdf(epsilon_ * lambda_ / sigma_))
            if np.isnan(norm_divide):
                norm_divide = -epsilon_ * lambda_ / sigma_
            ui = lambda_ * sigma_ / (1 + lambda_ ** 2) * (norm_divide + epsilon_ * lambda_ / sigma_)
            vi = epsilon_ - ui
            if vi > max_capital:
                max_capital = vi
            dst_ws.write(row, 8, vi)

            # LABOUR
            epsilon_ = float(table.cell_value(174 + row - 1, 1))
            sigma_ = float(table2.cell_value(sigma_index[2], 1))
            lambda_ = float(table2.cell_value(lambda_index[2], 1))
            norm_divide = norm.pdf(epsilon_ * lambda_ / sigma_) / (norm.cdf(epsilon_ * lambda_ / sigma_))
            if np.isnan(norm_divide):
                norm_divide = -epsilon_ * lambda_ / sigma_
            ui = lambda_ * sigma_ / (1 + lambda_ ** 2) * (norm_divide + epsilon_ * lambda_ / sigma_)
            vi = epsilon_ - ui
            if vi > max_labour:
                max_labour = vi
            dst_ws.write(row, 9, vi)

        dst_ws.write(1, 10, max_co2)
        dst_ws.write(1, 11, max_capital)
        dst_ws.write(1, 12, max_labour)

    dst_file_new.save('RFrontierOutputFiles/_3rd_dea_input_cal.xls')
    return


def generate_adjusted_dea_input():
    wb_fitted = xlrd.open_workbook(filename='RFrontierOutputFiles/_3rd_dea_input_cal.xls')  # 拟合值，残值信息
    for year in range(0, 11):
        ac_year = 2016 - year
        wb = xlrd.open_workbook(filename='pyDEAInputFiles/_dea' + str(ac_year) + '.xls')
        dst_file_new = xlcopy(wb)
        sheet = dst_file_new.get_sheet(0)
        read_sheet = wb.sheet_by_index(0)
        sheet_fitted = wb_fitted.sheet_by_name(str(ac_year))

        for row in range(1, 31):

            # CO2
            origin_value = read_sheet.cell_value(row, 3)
            fitted_surplus = sheet_fitted.cell_value(1, 4) - sheet_fitted.cell_value(row, 1)
            vi_surplus = sheet_fitted.cell_value(1, 10) - sheet_fitted.cell_value(row, 7)
            adjusted_value = origin_value + fitted_surplus + vi_surplus
            sheet.write(row, 3, adjusted_value)

            # CAPITAL
            origin_value = read_sheet.cell_value(row, 2)
            fitted_surplus = sheet_fitted.cell_value(1, 5) - sheet_fitted.cell_value(row, 2)
            vi_surplus = sheet_fitted.cell_value(1, 11) - sheet_fitted.cell_value(row, 8)
            adjusted_value = origin_value + fitted_surplus + vi_surplus
            sheet.write(row, 2, adjusted_value)

            # CAPITAL
            origin_value = read_sheet.cell_value(row, 4)
            fitted_surplus = sheet_fitted.cell_value(1, 6) - sheet_fitted.cell_value(row, 3)
            vi_surplus = sheet_fitted.cell_value(1, 12) - sheet_fitted.cell_value(row, 9)
            adjusted_value = origin_value + fitted_surplus + vi_surplus
            sheet.write(row, 4, adjusted_value)

        dst_file_new.save('pyDEAThirdStageInputFiles/_dea' + str(ac_year) + '.xls')
    return

