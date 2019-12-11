import requests
import os
import json
import socket
import getpass
import xlrd
from urllib import request


def filter_excel(workbook, column_name=0, by_name='Sheet2'):
    table = workbook.sheet_by_name('Sheet2')  # 获得表格
    total_rows = table.nrows  # 拿到总共行数
    # total_cols = table.ncols  # 拿到总共列数
    # 某一行数据 ['姓名', '用户名', '联系方式', '密码']
    # columns = table.row_values(column_name)
    # excel_list = []
    # excel_list = []
    d = PatchDownload()
    Compute_name = getpass.getuser()
    dir_path = r'C:\Users\%s\Desktop\Pdfs' % Compute_name
    isExists = os.path.exists(dir_path)
    if not isExists:
        os.makedirs(dir_path)
    for i in range(0, total_rows):
        d.downReport(table.row_values(i)[1], table.row_values(i)[
                     0], table.row_values(i)[3], dir_path)
        # print(table.row_values(i)[0])


def read_file(file_url):
    try:
        data = xlrd.open_workbook(file_url)
        filter_excel(data)
    except Exception as e:
        print(str(e))




class PatchDownload():
    def __init__(self):
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36',
            'Referer': 'http://172.20.0.201:8989/'}
        self.main_url = 'http://172.20.0.201:8989/'
        self.session = requests.Session()
        self.session.headers = self.headers
        # self.filePath = 'D:\\1.xls'

    # def getPersList(self, orderCode, hosCode):
    #     url = self.main_url + 'ma_publiccenter/weChat/order/orderDetail?orderCode=%s' % orderCode + \
    #         '&&hosCode=%s' % hosCode
    #     r = requests.get(url)
    #     data = json.loads(r.text)
    #     # print(data)
    #     if data['resCode'] != '00':
    #         print(data['msg'])
    #         return []
    #     else:
    #         print(data['msg'])
    #         groupExtList = data['order']['planExt']['groupExtList']
    #         newPersList = []
    #         for groupItem in groupExtList:
    #             if groupItem['persList'] != None:
    #                 for person in groupItem['persList']:
    #                     if int(person["status"]) >= 1090 and int(person["status"]) <= 1120:
    #                         if person["cardNo"] not in self.list:
    #                             newPersList.append(person)
    #         print("未下载的报告：--------------")
    #         print(len(newPersList))
    #         return newPersList

    def downReport(self, name, examinationNo, cardNo, dir_path):
        url = (self.main_url + 'ma_publiccenter/report/getReportFile/{}/{}').format(
            examinationNo[0:7], examinationNo)
        path = dir_path + os.sep + examinationNo + '_' + name + '.pdf'  # 文件路径
        print('正在下载 %s的报告...' % name)
        request.urlretrieve(url, path)  # 下载文件
        print('%s的报告已完成下载' % name, '√√√√√√')

    # def downReports(self):
    #     Compute_name = getpass.getuser()
    #     # newPersList = self.getPersList(orderCode, hosCode)
    #     dir_path = r'C:\Users\%s\Desktop\Pdfs' % Compute_name
    #     isExists = os.path.exists(dir_path)
    #     if not isExists:
    #         os.makedirs(dir_path)
    #     for person in newPersList:
    #         self.downReport(
    #             person['name'], person['examinationNo'], person['cardNo'], dir_path)

    #     print('---------------体检报告下载完毕----------------')
    #     print('----------------------------------------------')
    #     print('----------------------------------------------')


read_file('D:\\1.xls')

