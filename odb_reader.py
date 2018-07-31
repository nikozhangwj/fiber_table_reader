# encoding: utf-8
# /usr/bin/python3
# author : niko.zhang

import xlrd
import csv



workbook = xlrd.open_workbook('E:\巴伦文档\光纤检测\金利通金融中心\金利通光纤到户工程端子表.xls')
sheet_names = workbook.sheet_names()
keyword = 'ODB序号'
for name in sheet_names:
    sheet = workbook.sheet_by_name(name)
    for i in range(0, sheet.nrows):
        if keyword == sheet.row_values(i)[0]:
            #print(sheet.row_values(i+1))
            #print(sheet.row_values(i+1)[0])
            locate = sheet.row_values(i+1)[1].split(' ')[0]
            fiber_type = sheet.row_values(i+1)[1].split(' ')[5]
            odb_num = sheet.row_values(i+1)[1].split(' ')[6]
            building_floor = sheet.row_values(i+1)[1].split(' ')[-1]
            with open('E:\巴伦文档\光纤检测\金利通金融中心\金利通光纤到户工程info.csv','a',newline='') as csvfile:
                skywriter = csv.writer(csvfile,dialect='excel')
                skywriter.writerow([locate, fiber_type, odb_num, building_floor])
