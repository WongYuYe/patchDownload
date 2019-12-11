import requests, os, json, socket, getpass
from urllib import request

class PatchDownload():
    def __init__(self):
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36',
            'Referer': 'http://172.20.0.201:8989/'}
        self.main_url='http://172.20.0.201:8989/'
        self.session = requests.Session()
        self.session.headers=self.headers
        self.hosCodes = [ '苍南-0001001', '美生-0001002', '韩诺-0001003', '桐庐-0001004', '西湖-0001005', '武义-0001006' ]
        print('--------------   hosCode列表如下   --------------')
        for hosCodeItem in self.hosCodes:
            print('--------------   ' + hosCodeItem + '   --------------')

    def getPersList(self, orderCode, hosCode, startNo, count):
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
                            newPersList.append(person)
            print('此订单下共有%d份报告'% len(newPersList))      
            end = startNo + count
            print('开始下载第%d份报告'% startNo,'到第%d份报告'% (end - 1))          
            modifyNewPersList = newPersList[startNo: end]
            # for i in range(len(newPersList)):
            #     if (startNo >= i and len(modifyNewPersList) <= count):
            #         modifyNewPersList.append(newPersList[i])
            # print(len(modifyNewPersList))        
            return modifyNewPersList

    def downReport(self, name, examinationNo, cardNo, dir_path):
        url = (self.main_url + 'ma_publiccenter/report/getReportFile/{}/{}').format(examinationNo[0:7], examinationNo)
        path = dir_path + os.sep + name + '_' + cardNo + '.pdf'  # 文件路径
        print('正在下载 %s的报告...'% name)
        request.urlretrieve(url, path)  # 下载文件
        print('%s的报告已完成下载'% name, '√√√√√√')

    def downReports(self, orderCode, hosCode, startNo, count):
        Compute_name = getpass.getuser()
        newPersList = self.getPersList(orderCode, hosCode, startNo, count)
        dir_path = r'C:\Users\%s\Desktop\Pdfs'% Compute_name
        isExists = os.path.exists(dir_path)
        if not isExists:
            os.makedirs(dir_path)
        for person in newPersList:
            self.downReport(person['name'], person['examinationNo'], person['cardNo'], dir_path)
        print('---------------体检报告下载完毕----------------')    
        print('----------------------------------------------')
        print('----------------------------------------------')

        orderCode = input("请输入订单号:")
        hosCode = input("请输入分院编码:")
        startNo = int(input("请输入第几份开始下载:"))
        count = int(input("请输入份数:"))
        d.downReports(orderCode, hosCode, startNo, count)
        # input()


d = PatchDownload()
orderCode = input("请输入订单号:")
hosCode = input("请输入分院编码:")
startNo = int(input("请输入第几份开始下载:"))
count = int(input("请输入份数:"))
 
d.downReports(orderCode, hosCode, startNo, count)