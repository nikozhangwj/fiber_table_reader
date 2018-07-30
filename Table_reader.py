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
        data = []
        for name in sheet_names:
            sheet = workbook.sheet_by_name(name)
            print('正在读取表{0}'.format(name))
            data.append(name)
            for i in range(0, sheet.nrows):
                if keyword in sheet.row_values(i):
                    for k in sheet.row_values(i):
                        if k == keyword or k == '':
                            continue
                        else:
                            data.append(k)
            print('表{0}读取完成\n'.format(name))
            #time.sleep(0.5)
        return data


    def table_writer(self, _csv_file, data):
        with open(_csv_file,'w',newline='') as csvfile:
            skywriter = csv.writer(csvfile,dialect='excel')
            for room in data:
                skywriter.writerow([room])
        return True


if __name__ == '__main__':
    try:
        praser = TablePraser()
        input_file_path = input('请输入解析XLS文件路径:')
        output_file_path = input('请输入输出文件路径:')
        data = praser.table_reader(input_file_path)
        praser.table_writer(output_file_path, data)
    except BaseException as e:
        print(e,'异常情况')
