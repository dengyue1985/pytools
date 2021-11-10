# -*- coding:utf-8 -*-
'''
    CSV格式文件转换成excel文件
    支持单个文件转换和批量转换
    可以指定csv文件分隔符，默认分隔符为tab
'''
import pandas as pd
import argparse
import os

def parse_args():
    """解析命令行传入参数"""
    usage = """
Usage Example: 
#转个单个csv文件
python3 csv2xlsx.py --csv="/tmp/test.csv" --xlsx="/tmp/test.xlsx" --sep=','

#批量转换csv文件
python3 csv2xlsx.py --csv="/tmp/*.csv" --sep=','

Description:
    Convert file from csv to xlsx
    """    
    # 创建解析对象并传入描述
    parser = argparse.ArgumentParser(description=usage,formatter_class=argparse.RawTextHelpFormatter)
    #参数 csv
    parser.add_argument('--csv',dest='csvfile',required=True,action='store',metavar='CsvFile',help='The csv file')
    #参数 xlsx
    parser.add_argument('--xlsx',dest='xlsxfile',required=False,action='store',metavar='XlsxFile',help='The xlsx file. Default: same as csv file name')
    #参数 sep #csv文件分隔符
    parser.add_argument('--sep',dest='sep',required=False,action='store',metavar='Separator',help='Field separator in CSV file. Default:tab')
    args = parser.parse_args()
    return args

def csv_to_xlsx(csvfile,xlsxfile,csvsep):
    csv = pd.read_csv(csvfile,encoding='gbk',sep=csvsep,engine='python',dtype='str')
    csv.to_excel(xlsxfile,index=False)

def read_path(csvpath):
    filelist = os.listdir(csvpath)
    csvlist = [x for x in filelist if x[-4:]=='.csv']
    return csvlist
def main():
    args = parse_args()
    xlsxfile = args.xlsxfile
    csvfile = args.csvfile
    sep = args.sep
    if csvfile[-5:] == '*.csv':
        csvpath=csvfile[:-5]
        csvlist = read_path(csvpath)
        for i in csvlist:
            csvfile = csvpath + i
            xlsxfile=csvfile.replace('csv','xlsx')
            csv_to_xlsx(csvfile,xlsxfile,sep)
    else:
        if not xlsxfile:
            xlsxfile=csvfile.replace('csv','xlsx')
        csv_to_xlsx(csvfile,xlsxfile,sep)
if __name__ == '__main__':
    main()