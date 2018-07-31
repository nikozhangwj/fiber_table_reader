# encoding: utf-8
# /usr/bin/python3
# author : niko.zhang

import xlrd
import csv
import time


class TablePraser(object):

    def table_reader(self, xls_path, keyword='房号'):
        workbook = xlrd.open_workbook(xls_path)
        sheet_names = workbook.sheet_names()
        data = {'room' : []}
        for name in sheet_names:
            sheet = workbook.sheet_by_name(name)
            print('正在读取表{0}'.format(name))
            data['room'].append(name)
            for i in range(0, sheet.nrows):
                if keyword in sheet.row_values(i):
                    for k in sheet.row_values(i):
                        if k == keyword or k == '':
                            continue
                        else:
                            data['room'].append(k)
            print('表{0}读取完成\n'.format(name))
            #time.sleep(0.5)
        return data

    def odb_reader(self, xls_path, keyword = 'ODB序号'):
        workbook = xlrd.open_workbook(xls_path)
        sheet_names = workbook.sheet_names()
        data = {'locate': [] , 'fiber_type' : [] , 'odb_num' : [] , 'building_floor' : []}
        for name in sheet_names:
            sheet = workbook.sheet_by_name(name)
            for i in range(0, sheet.nrows):
                if keyword == sheet.row_values(i)[0]:
                    data['locate'].append(sheet.row_values(i+1)[1].split(' ')[0])
                    data['fiber_type'].append(sheet.row_values(i+1)[1].split(' ')[5])
                    data['odb_num'].append(sheet.row_values(i+1)[1].split(' ')[6])
                    data['building_floor'].append(sheet.row_values(i+1)[1].split(' ')[-1])
        return data

    def table_writer(self, _csv_file, data):
        with open(_csv_file,'a',newline='') as csvfile:
            skywriter = csv.writer(csvfile,dialect='excel')
            for key,values in data.items():
                skywriter.writerow(['\n'+key])
                for value in values:
                    skywriter.writerow([value])

        return True


if __name__ == '__main__':
    praser = TablePraser()
    input_file_path = input('Please in put xls path:')
    _out_put = input('Please setup output path:')
    _select_input = input('which function you want to use readtable/readodbinfo:')
    if _select_input == 'readtable':
        data = praser.table_reader(input_file_path)
        praser.table_writer(_out_put,data)
    elif _select_input == 'readodbinfo':
        data = praser.odb_reader(input_file_path)
        praser.table_writer(_out_put,data)
    else:
        print('errors')
