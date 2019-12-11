import requests
import os
import json
import socket
import getpass
from urllib import request

class PatchDownload():
    def __init__(self):
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36',
            'Referer': 'http://172.20.0.201:8989/'}
        self.main_url = 'http://172.20.0.201:8989/'
        self.session = requests.Session()
        self.session.headers = self.headers
        filePath = 'F:\\桐庐报告\\2019桐君街道企退体检-11\\'
        self.list = []
        for file in os.listdir(filePath):
            self.list.append(file[-22:-4])

    def getPersList(self, orderCode, hosCode):
        url = self.main_url + 'ma_publiccenter/weChat/order/orderDetail?orderCode=%s'% orderCode + '&&hosCode=%s'% hosCode
        r = requests.get(url)
        data = json.loads(r.text)
        # print(data)
        if data['resCode'] != '00':
            print(data['msg'])
            return []
        else:
            print(data['msg'])
            groupExtList = data['order']['planExt']['groupExtList']
            newPersList = []
            for groupItem in groupExtList:
                if groupItem['persList'] != None:
                    for person in groupItem['persList']:
                        if int(person["status"]) >= 1090 and int(person["status"]) <= 1120:
                            if person["cardNo"] not in self.list:
                                newPersList.append(person)
            print("未下载的报告：--------------")   
            print(len(newPersList))
            return newPersList

    def downReport(self, name, examinationNo, cardNo, dir_path):
        url = (self.main_url + 'ma_publiccenter/report/getReportFile/{}/{}').format(
            examinationNo[0:7], examinationNo)
        path = dir_path + os.sep + name + '_' + cardNo + '.pdf'  # 文件路径
        print('正在下载 %s的报告...' % name)
        request.urlretrieve(url, path)  # 下载文件
        print('%s的报告已完成下载' % name, '√√√√√√')

    def downReports(self):
        Compute_name = getpass.getuser()
        newPersList = self.getPersList(orderCode, hosCode)
        dir_path = r'C:\Users\%s\Desktop\Pdfs' % Compute_name
        isExists = os.path.exists(dir_path)
        if not isExists:
            os.makedirs(dir_path)
        for person in newPersList:
            self.downReport(
                person['name'], person['examinationNo'], person['cardNo'], dir_path)
        print('---------------体检报告下载完毕----------------')
        print('----------------------------------------------')
        print('----------------------------------------------')

d = PatchDownload()
orderCode = input("请输入订单号:")
hosCode = input("请输入分院编码:")
# startNo = int(input("请输入第几份开始下载:"))
# count = int(input("请输入份数:"))
d.downReports()
