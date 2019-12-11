# 增加分组
import requests, os, sys, json, socket, getpass
import threading
import time
from urllib import request
# import tkinter as tk
from tkinter import ttk
from tkinter import *
import tkinter.messagebox
root= Tk()
root.title('批量下载体检报告')
root.geometry('640x640') # 这里的乘号不是 * ，而是小写英文字母 x

class MyThread(threading.Thread):
    def __init__(self, func, *args):
        super().__init__()
        
        self.func = func
        self.args = args
        
        self.setDaemon(True)
        self.start()    # 在这里开始
        
    def run(self):
        self.func(*self.args)
# class MyThread(threading.Thread):
#     def __init__(self, func, *args):
#         super().__init__()
#         # self.__flag = threading.Event()     # 用于暂停线程的标识
#         # self.__flag.set()       # 设置为True
#         # self.__running = threading.Event()      # 用于停止线程的标识
#         # self.__running.set()      # 将running设置为True

#         self.func = func
#         self.args = args
        
#         self.setDaemon(True)
#         self.start()

#     def run(self):
#         while self.__running.isSet():
#             self.__flag.wait()      # 为True时立即返回, 为False时阻塞直到内部的标识位为True后返回
#             print(time.time())
#             time.sleep(1)

#     def pause(self):
#         self.__flag.clear()     # 设置为False, 让线程阻塞

#     def resume(self):
#         self.__flag.set()    # 设置为True, 让线程停止阻塞

#     def stop(self):
#         self.__flag.set()       # 将线程从暂停状态恢复, 如何已经暂停的话
#         self.__running.clear()        # 设置为False        
#         txt.insert(0.0, '\n-----------------批量下载停止！！！-----------------')

class PatchDownload():
    def __init__(self):
        threading.Thread.__init__(self)
        self.headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36',
            'Referer': 'http://172.20.0.201:8989/'}
        self.main_url='http://172.20.0.201:8989/'
        self.session = requests.Session()
        self.session.headers=self.headers
        self.hosCodes = [ '苍南-0001001', '美生-0001002', '韩诺-0001003', '桐庐-0001004', '西湖-0001005', '武义-0001006' ]
        # print('--------------   hosCode列表如下   --------------')
        # for hosCodeItem in self.hosCodes:
        #     print('--------------   ' + hosCodeItem + '   --------------')

    def getPersList(self, orderCode, planGroupName, hosCode):
        url = self.main_url + 'ma_publiccenter/weChat/order/orderDetail?orderCode=%s'% orderCode + '&&hosCode=%s'% hosCode
        r = requests.get(url)
        data = json.loads(r.text)
        if data['resCode'] != '00':
            self.dialog(data['msg'])
            # return None
        else:
            # self.dialog(data['msg'])
            groupExtList = data['order']['planExt']['groupExtList']
            newPersList = []
            for groupItem in groupExtList:
                if planGroupName:
                    planGroupNameList = planGroupName.split('/')
                    if groupItem['persList'] != None:
                        if groupItem['planGroupName'] in planGroupNameList:
                            for person in groupItem['persList']:
                                if int(person["status"]) >= 1090 and int(person["status"]) <= 1120:
                                    newPersList.append(person)
                else:
                    if groupItem['persList'] != None:
                        for person in groupItem['persList']:
                            if int(person["status"]) >= 1090 and int(person["status"]) <= 1120:
                                newPersList.append(person)        
            # print('此订单下共有%d份报告'% len(newPersList))
            txt.insert(0.0, '\n共有%d份已出报告'% len(newPersList))    
            return newPersList

    def downReport(self, name, examinationNo, cardNo, dir_path):
        url = (self.main_url + 'ma_publiccenter/report/getReportFile/{}/{}').format(examinationNo[0:7], examinationNo)
        path = dir_path + os.sep + name + '_' + cardNo + '.pdf'  # 文件路径
        txt.insert(0.0, '\n正在下载 %s的报告...'% name)
        request.urlretrieve(url, path, self.callback)  # 下载文件
    def callback(self, a, b, c):
        '''回调函数
        @a: 已经下载的数据块
        @b: 数据块的大小
        @c: 远程文件的大小
        '''
        per = 100.0 * a * b / c 
        if per > 100: 
            per = 100
            txt.insert(0.0, '\n下载完成√√√√√√')
        # print('%.2f%%' % per)
        # time.sleep(5)

    def downReports(self, orderCode, planGroupName, hosCode):
        Compute_name = getpass.getuser()
        newPersList = self.getPersList(orderCode, planGroupName, hosCode)
        dir_path = r'C:\Users\%s\Desktop\Pdfs'% Compute_name
        isExists = os.path.exists(dir_path)
        if not isExists:
            os.makedirs(dir_path)
        if newPersList:
            if len(newPersList) > 0:
                for person in newPersList:
                    self.downReport(person['name'], person['examinationNo'], person['cardNo'], dir_path)
                # print('---------------体检报告下载完毕----------------')    
                # print('----------------------------------------------')
                # print('----------------------------------------------')
                txt.insert(0.0, '\n---------------体检报告下载完毕----------------')
            else:
                txt.insert(0.0, '\n---------------无已出体检报告----------------')
                # print('--------------------无已出体检报告--------------------')
           
        
    def dialog(self, msg):
        tkinter.messagebox.askokcancel('error', msg)
        # if answer:
        #     lb.config(text='已确认')
        # else:
        #     lb.config(text='已取消')

d = PatchDownload()

lb1 = Label(root, text='订单号')
lb1.place(relx=0.1, rely=0.1, relwidth=0.3, relheight=0.1)
inp1 = Entry(root)
inp1.place(relx=0.1, rely=0.2, relwidth=0.3, relheight=0.1)
lb2 = Label(root, text='分组名称，多个以/隔开（可不填）')
lb2.place(relx=0.6, rely=0.1, relwidth=0.3, relheight=0.1)
inp2 = Entry(root)
inp2.place(relx=0.6, rely=0.2, relwidth=0.3, relheight=0.1)
txt = Text(root)
txt.place(rely=0.6, relheight=0.4)

var = StringVar()
values = [
    {'name': '苍南', 'hosCode': '0001001'},
    {'name': '美生', 'hosCode': '0001002'},
    {'name': '韩诺', 'hosCode': '0001003'},
    {'name': '桐庐', 'hosCode': '0001004'},
    {'name': '西湖', 'hosCode': '0001005'},
]
# hosCode = ''
def selected(event):
    # hosCode = values[comb.current()]['hosCode']
    # print(values[comb.current()]['hosCode'])
    txt.insert(0.0, '\n切换到%s体检中心，'%values[comb.current()]['name'] + '体检中心编码%s'%values[comb.current()]['hosCode'])

hos = [ '苍南', '美生', '韩诺', '桐庐', '西湖' ]
comb = ttk.Combobox(root,textvariable=var,values=hos)
comb.place(relx=0.1,rely=0.4,relwidth=0.3, relheight=0.1)
comb.current(0)
comb.bind('<<ComboboxSelected>>',selected)
lb2 = Label(root, text='体检中心')
lb2.place(relx=0.1, rely=0.3, relwidth=0.3, relheight=0.1)

# def stop():
#     try:
#         sys.exit(0)
#     except:
#         txt.insert(0.0, '\n-----------------批量下载停止！！！-----------------')
#     finally:
#         print('cleanup')
#         # txt.insert(0.0, '\n-----------------批量下载停止！！！-----------------')


btn1 = Button(root, text='批量下载', command=lambda: MyThread(d.downReports, inp1.get(), inp2.get(), values[comb.current()]['hosCode']))
btn1.place(relx=0.6, rely=0.4, relwidth=0.14, relheight=0.1)
# btn2 = Button(root, text='停止', command=lambda: MyThread().stop())
# btn2.place(relx=0.76, rely=0.4, relwidth=0.14, relheight=0.1)

root.mainloop()
