# *- coding:utf-8 -*-
import pyDEA.main
import xlwt
import xlrd
import os

wb = xlwt.Workbook(encoding='utf-8', style_compression=0)
ws = wb.add_sheet('Carbon Emission', cell_overwrite_ok=True)

wb_east = xlwt.Workbook(encoding='utf-8', style_compression=0)
ws_east = wb_east.add_sheet('Carbon Emission', cell_overwrite_ok=True)

wb_middle = xlwt.Workbook(encoding='utf-8', style_compression=0)
ws_middle = wb_middle.add_sheet('Carbon Emission', cell_overwrite_ok=True)

wb_west = xlwt.Workbook(encoding='utf-8', style_compression=0)
ws_west = wb_west.add_sheet('Carbon Emission', cell_overwrite_ok=True)


def read_excel(file_url):
    try:
        data = xlrd.open_workbook(filename=file_url)
        return data
    except Exception as e:
        print(str(e))


def read_table(file_url):
    workbook = read_excel(file_url)
    return workbook.sheet_by_index(0)


def split():
    for year in range(2006, 2017):
        table = read_table('pyDEAInputFiles/_dea' + str(year) + '.xls')
        row_east_index = 0
        row_middle_index = 0
        row_west_index = 0
        for row in range(0, 31):
            if row == 0:
                for col in range(1, table.ncols):
                    ws_east.write(row_east_index, col, table.cell_value(row, col))
                    ws_middle.write(row_middle_index, col, table.cell_value(row, col))
                    ws_west.write(row_west_index, col, table.cell_value(row, col))
                row_east_index += 1
                row_middle_index += 1
                row_west_index += 1
            elif row == 1 or row == 2 or row == 3 or row == 6 or row == 9 or row == 10 or row == 11 or row == 13 or row == 15 or row == 19 or row == 21:
                ws_east.write(row_east_index, 0, table.cell_value(row, 0))
                for col in range(1, table.ncols):
                    ws_east.write(row_east_index, col, table.cell_value(row, col))
                row_east_index += 1
            elif row == 4 or row == 5 or row == 7 or row == 8 or row == 12 or row == 14 or row == 16 or row == 17 or row == 18 or row == 20:
                ws_middle.write(row_middle_index, 0, table.cell_value(row, 0))
                for col in range(1, table.ncols):
                    ws_middle.write(row_middle_index, col, table.cell_value(row, col))
                row_middle_index += 1
            else:
                ws_west.write(row_west_index, 0, table.cell_value(row, 0))
                for col in range(1, table.ncols):
                    ws_west.write(row_west_index, col, table.cell_value(row, col))
                row_west_index += 1

        wb_east.save('pyDEARegionalInputFiles/East/_dea_east_' + str(year) + '.xls')
        wb_middle.save('pyDEARegionalInputFiles/Middle/_dea_middle_' + str(year) + '.xls')
        wb_west.save('pyDEARegionalInputFiles/West/_dea_west_' + str(year) + '.xls')


def calculate():
    input_files = os.listdir('pyDEAInputFiles/')
    for index in range(0, input_files.__len__()):
        f = open('pyDEAParamsFiles/' + input_files[index].replace('.xls', '') + '.txt', 'w+')
        f.write('<ABS_WEIGHT_RESTRICTIONS> {}\n\
                <DATA_FILE> {pyDEAInputFiles/' + input_files[index] + '}\n\
                <USE_SUPER_EFFICIENCY> {}\n\
                <OUTPUT_CATEGORIES> {GDP}\n\
                <NON_DISCRETIONARY_CATEGORIES> {}\n\
                <OUTPUT_FILE> {pyDEAOutputFiles/out' + input_files[index] + '}\n\
                <CATEGORICAL_CATEGORY> {}\n\
                <VIRTUAL_WEIGHT_RESTRICTIONS> {}\n\
                <PRICE_RATIO_RESTRICTIONS> {}\n\
                <WEAKLY_DISPOSAL_CATEGORIES> {}\n\
                <MULTIPLIER_MODEL_TOLERANCE> {0}\n\
                <MAXIMIZE_SLACKS> {}\n\
                <INPUT_CATEGORIES> {CAPITAL;CO2;WORK}\n\
                <PEEL_THE_ONION> {}\n\
                <RETURN_TO_SCALE> {both}\n\
                <DEA_FORM> {env}\n\
                <ORIENTATION> {input}')
        f.close()
        pyDEA.main.main('pyDEAParamsFiles/' + input_files[index].replace('.xls', '') + '.txt', output_format='xlsx')

    return


def calculate_regional():
    for orient in range(0, 3):
        orient_str = 'East'
        if orient == 1:
            orient_str = 'Middle'
        elif orient == 2:
            orient_str = 'West'

        input_files = os.listdir('pyDEARegionalInputFiles/' + orient_str + '/')
        for index in range(0, input_files.__len__()):
            f = open('pyDEARegionalParamsFiles/' + input_files[index].replace('.xls', '') + '.txt', 'w+')
            f.write('<ABS_WEIGHT_RESTRICTIONS> {}\n\
                    <DATA_FILE> {pyDEARegionalInputFiles/' + orient_str + '/' + input_files[index] + '}\n\
                    <USE_SUPER_EFFICIENCY> {}\n\
                    <OUTPUT_CATEGORIES> {GDP}\n\
                    <NON_DISCRETIONARY_CATEGORIES> {}\n\
                    <OUTPUT_FILE> {pyDEARegionalOutputFiles/' + orient_str + '/out' + input_files[index] + '}\n\
                    <CATEGORICAL_CATEGORY> {}\n\
                    <VIRTUAL_WEIGHT_RESTRICTIONS> {}\n\
                    <PRICE_RATIO_RESTRICTIONS> {}\n\
                    <WEAKLY_DISPOSAL_CATEGORIES> {}\n\
                    <MULTIPLIER_MODEL_TOLERANCE> {0}\n\
                    <MAXIMIZE_SLACKS> {}\n\
                    <INPUT_CATEGORIES> {CAPITAL;CO2;WORK}\n\
                    <PEEL_THE_ONION> {}\n\
                    <RETURN_TO_SCALE> {both}\n\
                    <DEA_FORM> {env}\n\
                    <ORIENTATION> {input}')
            f.close()
            pyDEA.main.main('pyDEARegionalParamsFiles/' + input_files[index].replace('.xls', '') + '.txt',
                            output_format='xlsx')

    return


def arrange_regional():
    for orient in range(0, 3):
        orient_str = 'East'
        if orient == 1:
            orient_str = 'Middle'
        elif orient == 2:
            orient_str = 'West'
        tables = []
        for year in range(2006, 2017):
            index = year - 2006
            tables.append(read_table('pyDEARegionalOutputFiles/' + orient_str + '/out_dea_' + orient_str.lower() + '_' + str(year) + '.xls'))

            for row in range(0, tables[index].nrows):
                if index == 0:
                    ws.write(row, 0, tables[index].cell_value(row, 0))

                if row == 0:
                    ws.write(row, index + 1, str(year) + '年')
                else:
                    ws.write(row, index + 1, tables[index].cell_value(row, 1))

        wb.save('各省各年份碳排放效率对比表_' + orient_str + '.xls')
    return


def arrange():
    tables = []
    for year in range(2006, 2017):
        index = year - 2006
        tables.append(read_table('pyDEAOutputFiles/out_dea' + str(year) + '.xls'))
        for row in range(0, 64):
            if index == 0:
                ws.write(row, 0, tables[index].cell_value(row, 0))

            if row == 0:
                ws.write(row, index + 1, str(year) + '年')
            else:
                ws.write(row, index + 1, tables[index].cell_value(row, 1))

    wb.save('各省各年份碳排放效率对比表.xlsx')
    return
