__author__ = 'huayang'


import tkinter.messagebox
from tkinter import *
#from tkinter.ttk import *
from tkinter.ttk import Combobox
from tkinter.ttk import Treeview
import datetime
import threading
import pickle
import configparser
import time
#from time import time, localtime, strftime
import os 
import pywinauto
import tushare as ts

import urllib




from winguiauto import *
from win32gui import *

def _normaliseText(controlText):
    '''Remove '&' characters, and lower case.
    Useful for matching control text.'''
    return controlText.lower().replace('&', '')

g_msgindex = 0





WorkingDir = os.getcwd()
#WorkingDir = "/"
now = datetime.datetime.now()   
otherStyleTime = now.strftime("%Y_%m_%d %H_%M_%S")

cfg_file = WorkingDir + "\cfg.ini"
#print(cfg_file)
outputfile =  WorkingDir + "\log\log_" +  otherStyleTime + ".txt"
windatafile =  WorkingDir + "\log\windos_"  + otherStyleTime + ".txt"


if not os.path.exists(WorkingDir + "\log"):
    print ("建立log目录")
    
    os.mkdir(WorkingDir + "\log")
    #x = input('')
    #sys.exit()



def WriteLogFIle(word ="[LOG]"):
    #global outputfile
    f=open(outputfile,'w')
    f.write(str(word)+'\n')
    f.close()

WriteLogFIle()

def IniRead(config_file_path, field, key):      
    cf = configparser.ConfigParser()    
    try:    
        cf.read(config_file_path)   
        result = cf.get(field, key)   
    except Exception as e:    
    #except  e:     
        print(e)
        WriteLogFIle(e)
        #sys.exit(1)     
    return result   
  
def IniWrite(config_file_path, field, key, value):      
    
    global g_msgindex
    g_msgindex = g_msgindex+1

    cf = configparser.ConfigParser()    
    try:    
        cf.read(config_file_path)   
        cf.set(field, key, value)   
        cf.write(open(config_file_path,'w'))    
    except Exception as e:           
        
        print(e)
        WriteLogFIle(e)
        #sys.exit(1)    
    return True     

broker = 'stock'
broker_name = '券商'


send_error = False

is_start = False
              
is_monitor = True
set_stock_info = []
order_msg = []
actual_stock_info = []


data_file = ""
main_title = ""
main_class = ""

#num_rows = int(IniRead(cfg_file, "common", "stock_numbers"))  


try:
   num_rows = int(IniRead(cfg_file, "common", "stock_numbers"))  
except Exception as e:
   print("配置项读取异常。请核对配置文件： cfg.ini") 
   x=input('')
   sys.exit()


is_ordered = [1] * num_rows  # 1：未下单  0：已下单下单sure
tip_msg = ['']  * num_rows
order_result = [''] * num_rows
order_retry = [0] * num_rows


pre_order_result = [''] * num_rows
hwnds_top_level = []
balance_index = 0   #index in cfg file of balance
account_index = 0   #index in cfg file of account
g_hwnds_position = 0 #hwnd of position
account_name =""
b_init_ready = False
positionList_ev = []

#********************************************

#*****************************************************
#check3.select()
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email import encoders

#from email.header import Header

import smtplib

send_mail = IniRead(cfg_file, "common", "send_mail")  
if send_mail  == '1':

    from_addr_cfg = IniRead(cfg_file, "common", "from_addr")  

    password_cfg = IniRead(cfg_file, "common", "password")  

    to_addr1 = IniRead(cfg_file, "common", "to_addr")  
    to_addr_cfg = eval(to_addr1)
    print ("to_addr_cfg")
    print (to_addr_cfg)
    smtp_server_cfg = IniRead(cfg_file, "common", "smtp_server") 


from_addr = '13688122988@139.com'
#from_addr = '13816994970@139.com'

password = 'Xiaolil1'

to_addr2 = '1742157288@qq.com'

#to_addr = ['13688122988@139.com',to_addr2]
to_addr = ['13688122988@139.com','a@abc.com']

#to_addr = ['13816994970@139.com',to_addr2]
 
smtp_server = 'smtp.139.com'

weichat_url = 'http://sc.ftqq.com/SCU7959T154ea232ece87eb21f6652bc534649d95903d8351de58.send'
#http://sc.ftqq.com/SCU7959T154ea232ece87eb21f6652bc534649d95903d8351de58.send

   

def post(url, data=None, headers=None, timeout=2, decode='utf-8'):
    #rt = HttpReturn()
    if headers is None:
        headers = {}
    post_data = urllib.parse.urlencode(data).encode(decode)
    try:
        req = urllib.request.Request(url, post_data, headers)
        hr = urllib.request.urlopen(req, timeout=timeout)
        #rt = HttpReturn()
        #rt.obj = rt
        text = hr.read().decode('utf-8')
        
        status = hr.status
    finally:
        return text ,status                
        #return text

def sendWechat(text,desp):

       
        params = {'text':text, 'desp':desp}
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/31.0.1650.57 Safari/537.36'}

        text,status = post(weichat_url,params,headers)
        
        return text,status


def mailmsg(header,content,attfile = ""):
    

    #send wechat
    text,status = sendWechat(header,content)
    '''

    if attfile == "":
        msg = MIMEText(content, 'plain', 'utf-8')
    else:
   
  
        msg = MIMEMultipart()
        content1 = MIMEText(content, 'plain', 'utf-8')
        msg.attach(content1)
  
        #attfile = 'C:\\Python35-32\\work\\pyautotrade_tdx-master\\cfg.ini'
        basename = os.path.basename(attfile)
        fp = open(attfile, 'rb')
        att = MIMEText(fp.read(), 'base64', 'utf-8')
        att["Content-Type"] = 'application/octet-stream'
        att.add_header('Content-Disposition', 'attachment',filename=('gbk', '', basename))
        encoders.encode_base64(att)
        msg.attach(att)  
    '''
    '''
    msg = MIMEText(content, 'plain', 'utf-8')
    #msg['From'] = broker+'@139.com'
    msg['From'] = broker+account_name +'@'+broker+'.com'
    msg['To'] = ','.join(to_addr)
    #msg['To'] = to_addr
    msg['Subject'] = header

    #send wechat
    text,status = sendWechat(header,content)
    
    
    try:


        server = smtplib.SMTP(smtp_server,25)
        server.set_debuglevel(1)
        server.login(from_addr, password)
    #server.sendmail(from_addr, [to_addr], msg.as_string())
        server.sendmail(from_addr, to_addr, msg.as_string())
    #print('发送成功')
        server.close()
    except Exception as e:
        Log(str(e))
    '''

    if send_mail  == '1':
        msg1 = MIMEText(content, 'plain', 'utf-8')
        #msg1['From'] = broker+account_name+from_addr_cfg
        
        msg1['From'] = broker+account_name +'@'+broker+'.com'
        msg1['To'] = ','.join(to_addr_cfg)
        

        msg1['Subject'] = header




        try:


            server = smtplib.SMTP(smtp_server_cfg,25)
            #server.set_debuglevel(1)
            server.login(from_addr_cfg, password_cfg)
    
            server.sendmail(from_addr_cfg, to_addr_cfg, msg1.as_string())

            server.close()
        except Exception as e:
            Log(str(e))
            server.close()
    

def LogData(data):
    f=open(windatafile,'a')
    f.write(data) 
    f.write('\n')    
    f.close

def Log(message = "I am here!",pre = "DBG"):
   #def _DebugBuffer(message = "I am here!",pre = "DBG"):
   #time = @HOUR & ":" & @MIN & ":" & @SEC & ":" & @MSEC & " "
    #global g_debug_to_buffer
    now = datetime.datetime.now()

    otherStyleTime = now.strftime("%Y-%m-%d %H.%M.%S_")
   
    IniWrite(outputfile, "LOG", "NO" + str(g_msgindex) + " " + otherStyleTime +  " " + pre , message)
    
    print (otherStyleTime + "  " + pre)
    print  (message)


def getConfigData():
    '''
    读取配置文件参数
    :return:双向委托界面下，控件的数量
    '''
    brokers = IniRead(cfg_file, "common", "brokers")  
    brokers_lsit = brokers.split(",")
    #print(brokers_lsit)

    numChildWindows = IniRead(cfg_file, "common", "numChildWindows")  

    
    numChildWindows_list = numChildWindows.split(",")
    #print(numChildWindows_list)
    #numAccountChildWindows = int(IniRead(cfg_file, broker, "numAccountChildWindows"))  


    #accountIndex = IniRead(cfg_file, "common", 'accountIndex')
    balanceIndex = IniRead(cfg_file, "common", 'balanceIndex')
    balanceIndex_list = balanceIndex.split(",")

    positionIndex =  IniRead(cfg_file, "common", 'positionIndex')
    
    positionIndex_list = positionIndex.split(",")

    positionList =  IniRead(cfg_file, "common", 'positionList')
    
    positionList_list = positionList.split(";")
    #print(positionIndex_list)


    return positionList_list,brokers_lsit,positionIndex_list,balanceIndex_list, numChildWindows_list,
        



def getConfigDataoold(broker):
    '''
    读取配置文件参数
    :return:双向委托界面下，控件的数量
    '''
    
    numChildWindows = int(IniRead(cfg_file, broker, "numChildWindows"))  
    #numAccountChildWindows = int(IniRead(cfg_file, broker, "numAccountChildWindows"))  
 
    data_file = IniRead(cfg_file, broker, 'data_file')
 
    main_title = IniRead(cfg_file, broker, 'main_title')
 
    main_class = IniRead(cfg_file, broker, 'main_class')

    accountIndex = IniRead(cfg_file, broker, 'accountIndex')
    balanceIndex = IniRead(cfg_file, broker, 'balanceIndex')


    return accountIndex,balanceIndex, \
        numChildWindows, \
        main_title,main_class,data_file

# def pickHwndOfControls(top_hwnd, num_child_windows):
#     cleaned_hwnd_controls = []
#     hwnd_controls = findSpecifiedWindows(top_hwnd, num_child_windows)
#     for Hwnd, text_name, class_name in hwnd_controls:
#         if class_name in ('Button', 'Edit'):
#             cleaned_hwnd_controls.append((Hwnd, text_name, class_name))
#     return cleaned_hwnd_controls


def getRunningMoney(sub_hwnds):
    '''
    :param sub_hwnds: 双向委托操作界面下的控件句柄列表
    :return:可用资金
    '''
    return getWindowText(sub_hwnds[12][1])


def pre_buy( code, stop_price, quantity):

    #need_money = stop_price * quantity
    need_money = float(stop_price) * float(quantity)

    msg = "ok"
    if float(getBanlence(1)) < need_money:
        msg = "资金不足"
        #print(msg)
    
    return msg

def buy(sub_hwnds, code, stop_price, quantity):
    '''
    买函数，自动填写3个Edit控件，及点击买入按钮。
    :param sub_hwnds: 双向委托操作界面下的控件句柄列表
    :param code: 股票代码，字符串
    :param stop_price: 涨停市价， 字符串
    :param quantity: 买入股票 数量，字符串
    :return:
   

    '''


    hwnd = win32gui.GetParent(sub_hwnds[27][0])
        #print(win32gui.SetForegroundWindow(hwnd))
        #print(win32gui.BringWindowToTop(hwnd))
        #print("EnableWindow(1381152,True)") 
    win32gui.EnableWindow(hwnd,True)
        
        #print('win32gui.ShowWindow(1381152,"SW_RESTORE")')
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    time.sleep(0.3)

    selectComboboxItem(sub_hwnds[19][0],0)
    time.sleep(0.3)
    setEditText(sub_hwnds[0][0], code)
    #print(sub_hwnds[1][0])
    setEditText(sub_hwnds[1][0], stop_price)
    setEditText(sub_hwnds[3][0], quantity)
  
    time.sleep(0.3)
    click(sub_hwnds[5][0])

    
    time.sleep(0.3)


def pre_sell( code, stop_price, quantity):
    
    ownshare = 0
    newposition = getPosition()

    for eachshare in  newposition :
        if eachshare[0] == code:
           share =eachshare[3]
           
           #ownshare = int(share[0:share.find('.')])
           ownshare = float(share)
    msg = "ok" 
    if float(quantity) > ownshare:
        msg = "股票不足"
    
        #print(msg)
    
    return msg

def sell(sub_hwnds, code, stop_price, quantity):
    '''
    买函数，自动填写3个Edit控件，及点击卖出按钮。
    :param sub_hwnds: 双向委托操作界面下的控件句柄列表
    :param code: 股票代码，字符串
    :param stop_price: 涨停市价， 字符串
    :param quantity: 卖出股票 数量，字符串
    :return:
    '''
    hwnd = win32gui.GetParent(sub_hwnds[27][0])
        #print(win32gui.SetForegroundWindow(hwnd))
        #print(win32gui.BringWindowToTop(hwnd))
        #print("EnableWindow(1381152,True)") 
    win32gui.EnableWindow(hwnd,True)
        
        #print('win32gui.ShowWindow(1381152,"SW_RESTORE")')
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    time.sleep(0.3)

    selectComboboxItem(sub_hwnds[45][0],0)
    time.sleep(0.3)
    setEditText(sub_hwnds[24][0], code)
    setEditText(sub_hwnds[25][0], stop_price)
    setEditText(sub_hwnds[27][0], quantity)
    time.sleep(0.3)

    click(sub_hwnds[29][0])
    time.sleep(0.3)
def pre_order(code, stop_prices, quantity, direction):
    '''
    模拟买卖函数
    :param top_hwnd: 顶层窗口句柄
    :param sub_hwnds: 双向委托操作界面下的控件句柄列表
    :param code: 股票代码， 字符串
    :param stop_prices: 涨跌停价格，字符串
    :param quantity: 买卖数量，字符串
    :param direction: 交易方向，字符串
    :return:
    '''


    if direction == 'B':
        
        
        msg = pre_buy(code, stop_prices[0], quantity)
        #if not msg == "ok":
          # print(msg + "PreBuy :"+" code "+code+" | price "+stop_prices[0]+" | quant "+quantity)

    if direction == 'S':
        
        
        msg = pre_sell(code, stop_prices[1], quantity)
        #if not msg == "ok":
        #   print(msg + "PreSell:"+" code "+code+" | price "+stop_prices[1]+" | quant "+quantity)
    return msg


def order(top_hwnd, sub_hwnds, code, stop_prices, quantity, direction):
    '''
    买卖函数
    :param top_hwnd: 顶层窗口句柄
    :param sub_hwnds: 双向委托操作界面下的控件句柄列表
    :param code: 股票代码， 字符串
    :param stop_prices: 涨跌停价格，字符串
    :param quantity: 买卖数量，字符串
    :param direction: 交易方向，字符串
    :return:
    '''
    global send_error

    try:
     
     if direction == 'B':
        
        Log("Buy :"+" code "+code+" | price "+stop_prices[0]+" | quant "+quantity)
        buy(sub_hwnds, code, stop_prices[0], quantity)
     if direction == 'S':
        
        Log("Sell:"+" code "+code+" | price "+stop_prices[1]+" | quant "+quantity)
        sell(sub_hwnds, code, stop_prices[1], quantity)
     time.sleep(0.3)
        
    
    except Exception as e:
            Log(str(e))      
           
            if not send_error:
                if direction == 'B':
        
                    msg ="Buy :"+" code "+code+" | price "+stop_prices[0]+" | quant "+quantity     
                if direction == 'S':
        
                    msg = "Sell:"+" code "+code+" | price "+stop_prices[1]+" | quant "+quantity
        
                mailmsg(broker+account_name+"下单失败，请重启程序！",str(e)+"\n"+msg)
                print("send error")
                send_error = True
              
            return "0"


def searchSpectialWindow(hwnd,wantedText=None, wantedClass=None):


        results = []

        lev1_win =searchLowerLevelWindows(hwnd)
        
               
        i = 0
        for (lev1_hwd,lev1_title,lev1_class) in lev1_win:
            
            LogData(str(i)+"  ********************************* "+ str(lev1_hwd)+" "+lev1_title+" "+lev1_class)
                        
            if wantedText and \
                    not _normaliseText(wantedText) in _normaliseText(lev1_title):
                continue
            if wantedClass and \
                    not lev1_class == wantedClass:
                continue
            #results.append(lev1_hwd)
            
            lev2_win = searchLowerLevelWindows(lev1_hwd)
            jj = 0        
            for (lev2_hwd,lev2_title,lev2_class) in lev2_win:
               #if i == 0 and jj >320:
                if log:
                 #Log(str(i)+"  *** "+str(jj)+ "  @@@@@@@@@@@@@@@@@ "+str(lev2_hwd)+" "+lev2_title+" "+lev2_class)
                 LogData(str(i)+"  *** "+str(jj)+ "  @@@@@@@@@@@@@@@@@ "+str(lev2_hwd)+" "+lev2_title+" "+lev2_class)
               
                if wantedText and \
                    not _normaliseText(wantedText) in _normaliseText(lev2_title):
                    continue
                if wantedClass and \
                    not lev2_class == wantedClass:
                    continue

                results.append(lev2_hwd)                        
                lev3_win = searchLowerLevelWindows(lev2_hwd)
                kk = 0        
                for (lev3_hwd,lev3_title,lev3_class) in lev3_win:
                 # if i == 0 and jj ==320:
                    if log:
                        LogData(str(i)+"  *** "+str(jj)+ "  @@@ "+str(kk)+ "  &&& "+str(lev3_hwd)+" "+lev3_title+" "+lev3_class)
                    if wantedText and \
                        not _normaliseText(wantedText) in _normaliseText(lev2_title):
                        continue
                    if wantedClass and \
                        not lev2_class == wantedClass:
                        continue

                    results.append(lev3_hwd) 
                    return results                       
                               
                    kk = kk+1      
                results.append(lev2_hwd) 
                jj = jj+1
            results.append(lev1_hwd) 
            i = i+1      
        return results

def searchWindows(hwnd,log = False):

        resultsLev1 = []
        resultsLev2 = []
        resultsLev3 = []


        lev1_win =searchLowerLevelWindows(hwnd)
        
               
        i = 0
        #resultsLev1.append(0)  
        resultsLev1 = [] 
        
        for (lev1_hwd,lev1_title,lev1_class) in lev1_win:
            
            LogData(str(i)+"  ********************************* "+ str(lev1_hwd)+" "+lev1_title+" "+lev1_class)
                        
            #results.append(lev1_hwd)
            
            lev2_win = searchLowerLevelWindows(lev1_hwd)
            jj = 0 
            resultsLev2 = []
            #resultsLev2.append(0)         
            for (lev2_hwd,lev2_title,lev2_class) in lev2_win:
               #if i == 0 and jj >320:
               if log:
                 #Log(str(i)+"  *** "+str(jj)+ "  @@@@@@@@@@@@@@@@@ "+str(lev2_hwd)+" "+lev2_title+" "+lev2_class)
                 LogData(str(i)+"  *** "+str(jj)+ "  @@@@@@@@@@@@@@@@@ "+str(lev2_hwd)+" "+lev2_title+" "+lev2_class)
               
               #resultsLev2.append(lev1_hwd)
               lev3_win = searchLowerLevelWindows(lev2_hwd)
               kk = 0        
               resultsLev3 = []
               #resultsLev3.append(0)  
               for (lev3_hwd,lev3_title,lev3_class) in lev3_win:
                 # if i == 0 and jj ==320:
                  if log:
                    LogData(str(i)+"  *** "+str(jj)+ "  @@@ "+str(kk)+ "  &&& "+str(lev3_hwd)+" "+lev3_title+" "+lev3_class)
                  #resultsLev3.append(lev3_hwd)  
                  resultsLev3.append([lev3_hwd,lev3_title,lev3_class])  
                  kk = kk+1      
               resultsLev2.append(resultsLev3)
               
               jj = jj+1
            resultsLev1.append(resultsLev2)     
            
            i = i+1      
        return resultsLev1


def findSpectificHwnd(wantedClass = "ComboBox"):

       
 
        for i in range(len(hwnds_top_level)):
        #for i in range(56):

            if len(hwnds_top_level[i]) == 0:
             #   print("continue")
                continue
            temp1 =  hwnds_top_level[i][0]
            #print("temp1")
            
            if len(hwnds_top_level[i][0]) == 0:
                #print("len(hwnds_top_level[i][0]) == 0:")
                #print("continue")
                continue
            for j in range(len(temp1)):
            #for j in range(50,56):

                if len(temp1)<3:
                    continue
                temp2 = temp1[j]
                print(temp2)   
                #if temp2[2] == 'ComboBox':
                if temp2[2] == wantedClass:
                    
                    print(i,j)
                    return i,j
def findSpectificHwndCTN(wantedClass = "MHPToolBar",wantedText = None):

        #ToolBar MHPToolBar
        result = [] 
        for i in range(len(hwnds_top_level)):
        #for i in range(56):

            if len(hwnds_top_level[i]) == 0:
             #   print("continue")
                continue
            temp1 =  hwnds_top_level[i][0]
            #print("temp1")
            
            if len(temp1) == 0:
                #print("len(hwnds_top_level[i][0]) == 0:")
                #print("continue")
                continue
            for j in range(len(temp1)):

                temp2 = temp1[j]
                #print(temp2)   
                
                if temp2[2] == wantedClass:

                    if wantedText :
                        if  _normaliseText(wantedText) in _normaliseText(temp2[1]):
                            #print(i,j)
                            result.append([i,j])
                            #return i,j    
                    else:


                        #print(i,j)
                        result.append([i,j])
                        #return i,j    
        return result


def getAccount():
        
        
        print (account_index)
        hwnd_account_info = eval("hwnds_top_level"+account_index+"[0]")
        print(hwnd_account_info)
        print(getComboboxItems(hwnd_account_info)[0])
        #result = []
        text = getComboboxItems(hwnd_account_info)[0]
        text1 = text.split(" ")

        return text1


def getBanlence(index=1):


    global send_error
    result = []

    try:  
        hwnd_banlance_info = eval("hwnds_top_level"+balance_index +"[0]")
        text = getWindowText(hwnd_banlance_info)
        #send_error = False

           
        
        #hwnd_banlance_info = eval("hwnds_top_level"+balance_index +"[0]")
        #print(hwnd_banlance_info)
        
        
        text1 = text.split("  ")

        for text2 in text1:
            text3 = text2.split(":")
            result.append(text3)
        
       

        if index == -1:
                return result
        return result[index][1]
        
    
    except Exception as e:
            Log(str(e))      
           
            if not send_error:
              mailmsg(broker+account_name+"余额获取失败。请重启程序！",str(e))
              send_error = True
              
            return "0"


def getPositionStr():


        new_position = getPosition()

        num_stocks = len(new_position)
        stocks_position = ""
        for i in range(num_stocks):
          #  print("i is %d " % i)
            stock_position1 = []
            for j in range(4):
         #       print(j)
                stock_position1.append(str(new_position[i][j]))

                #print (readListViewItems(g_hwnds_position,j)[i].decode('gbk'))
            stocks_position= stocks_position+str(stock_position1)+"\n\n"
               #stocks_position.append(str(stock_position1))
        #print(stocks_position)
        return stocks_position



def getPosition():

    global send_error
    try:  
        
        hwnd = win32gui.GetParent(g_hwnds_position)
        #hwnd = g_hwnds_position
        
        #print(win32gui.BringWindowToTop(hwnd))
        #print("EnableWindow(1381152,True)") 
        win32gui.EnableWindow(hwnd,True)
        
        #print('win32gui.ShowWindow(1381152,"SW_RESTORE")')
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        #win32gui.SetForegroundWindow(hwnd)

        time.sleep(0.3)





        num_stocks = len(readListViewItems(g_hwnds_position,0))
        
        #print(num_stocks)
        
        

        stocks_position = []
        #print(positionList_ev)
        #for i in positionList_ev:
         #   print(i)
        for i in range(num_stocks):
          #  print("i is %d " % i)
            stock_position = []
            #for j in range(10):
            for j in positionList_ev:

                #print(j)
                stock_position.append(readListViewItems(g_hwnds_position,j)[i].decode('gbk'))

                #print (readListViewItems(g_hwnds_position,j)[i].decode('gbk'))
            stocks_position.append(stock_position) 
        #print(stocks_position)
        return stocks_position
    except Exception as e:
        Log(str(e))      
           
        if not send_error:
            mailmsg(broker+account_name+"仓位获取失败。请重启程序！",str(e))
            send_error = True
              

        return [["000000","null","0","0"]]

def searcHwndOfClient(top_hwnd,numChildWindows,position_index_offset,balance_index_offset):
    
        global hwnds_top_level,account_index,balance_index,g_hwnds_position
        
        sub_hwnds = findSpecifiedWindows(top_hwnd, numChildWindows)  
        
        if sub_hwnds == 0:
           tkinter.messagebox.showerror('错误', '请先点击“对买对卖”，然后重启程序!!!  ')
           sys.exit()
        
        #print("sub_hwnds")
        #print(sub_hwnds[0:9])
        #print("searching windows...")

        ok = False
        while not ok:
            try:  

                hwnds_top_level = searchWindows(top_hwnd,log = False) 
                #print("here")

                try:   

                    if broker == 'guangfa': 
                        #x,y = findSpectificHwndCT("MHPToolBar","ToolBar")

                        x,y = findSpectificHwndCTN("MHPToolBar","ToolBar")[0]
                        account_index = "["+ str(x) +"][0][+" + str(y+1) +"]"
                        

                    else:
                        #x,y = findSpectificHwndCT("ComboBox")
                        x,y = findSpectificHwndCTN("ComboBox")[0]            
                        account_index = "["+ str(x) +"][0][+" + str(y) +"]"


                 
                    account = getAccount()
                    #print (account)
                    
                    if broker == 'haitong' or broker == 'shenwan': 


                         account_name = account[1]+account[3]
                    elif broker == 'moni':      

                         account_name = account[0]+account[1]        


                    elif broker == 'guangfa':      

                         account_name = account[0]
                    else:     
                         account_name = account[4]+account[3]        
                    #print("account_name")
                    print(account_name)

                except Exception as e:
                    print (str(e))
                    print ("\n请确认 " +  broker_name +" 客户端登录成功，再重启程序。按任意键程序将退出。")
                    input (":")
                    sys.exit()
            

                x,y = findSpectificHwndCTN("Static","余额")[balance_index_offset]
                balance_index = "["+ str(x) +"][0][+" + str(y) +"]"     

                #print("Find hwnd of position...")

                g_hwnds_position = findControls(top_hwnd,"List1",'SysListView32')          
                print(g_hwnds_position[0:9])

                g_hwnds_position = findControls(top_hwnd,"List1",'SysListView32')[position_index_offset]
            

                #print("getBanlence(1)")
                print(getBanlence(1))

                #print("getPosition()")
                print(getPosition())
                #测试交易对买对卖
                #setEditText(sub_hwnds[27][0], "0")
                
                ok = True

            except Exception as e:
                print (str(e))
                print ("\n请在 " +  broker_name +" 客户端中，先点击“查询--->资金股份”，再点击“对卖对卖”，然后按任意键继续...")
                input (":")
                

        return account_name,sub_hwnds
        



def tradingInit():
    '''

    :return:顶层窗口句柄，双向委托操作界面下的控件句柄列表
    '''
    global data_file,main_title,main_class,\
        hwnds_top_level,\
        balance_index,account_index,\
        g_hwnds_position,\
        account_name,\
        b_init_ready,\
        broker ,broker_name,\
        positionList_ev

    
    

    positionList_list,brokers_lsit,positionIndex_list,balanceIndex_list, numChildWindows_list = getConfigData()

    print("支持的客户端列表：\n")
    for eachbroker in brokers_lsit:
       print(eachbroker)
    #top_hwnd,found_title = findTopStockWindows(broker_name)
    #top_hwnd = findSpecifiedTopWindow(wantedClass=main_class)
    #main_title = "证券"
    top_hwnd_all,found_title_all_initial = findTopStockWindows(brokers_lsit)
    #print (found_title_all_initial)
    #print (top_hwnd_all)




    found_title_all_initial_short = [ " " for i in range(len(found_title_all_initial))]

    for i in range(len(found_title_all_initial)):
        tempstr = found_title_all_initial[i]
        found_title_all_initial_short[i] = tempstr[0:tempstr.find("V")+5]

    #print (found_title_all_initial_short)

    found_title_all = list(set(found_title_all_initial_short))
    if len(found_title_all)==0:
        print("\n未发现通达信客户端！按任意键程序将退出。")
        x = input("")
        sys.exit()

    choose = '0'
    top_hwnd = top_hwnd_all[0]
    found_title =  found_title_all[0]

    #找到同一个券商多个客户端对应的窗口句柄
    hwndForTopWindowAll = []
    for i in range(len(found_title_all)):
        hwndForTopWindowOne = []
        tempstr = found_title_all[i]
        for j in range(len(found_title_all_initial)):
           tempstr1 = found_title_all_initial[j] 
           if tempstr in tempstr1 :
               hwndForTopWindowOne.append(top_hwnd_all[j])
        hwndForTopWindowAll.append(hwndForTopWindowOne)
    print(found_title_all)
    print(hwndForTopWindowAll)

    
    if found_title_all == 0:
        tkinter.messagebox.showerror('错误', '请先打开交易软件:  ' + main_title +' , 然后重启程序!!!  ')
        sys.exit()
        
    if len(found_title_all) > 0:
       choose = "-1"
       while not choose.isdigit() or \
             not ( choose.isdigit() and int(choose) in range(len(found_title_all))):
           print('\n\n请选择对应的券商客户端。（输入对应序号）\n如果需要的客户端不在列表中，请启动对应客户端，并登录:\n') 
           for i in range(len(found_title_all)):
               print('%d : %s:\n' % (i,str(found_title_all[i])))


           choose = input(':')

       print(choose)     
       #top_hwnd = top_hwnd_all[int(choose)]
       top_hwnd_one = hwndForTopWindowAll[int(choose)]
       print(top_hwnd_one)
       found_title =  found_title_all[int(choose)]
       print(found_title)
       if  '模拟交易' in found_title:
            broker = 'moni'
            broker_name = '模拟交易'
       elif '平安证券'  in found_title:
            broker = 'pingan'
            broker_name = '平安证券' 
       elif '海通证券'  in found_title:
            broker = 'haitong'
            broker_name = '海通证券' 
       elif '广发证券'  in found_title:
            broker = 'guangfa'
            broker_name = '广发证券' 
       elif '华泰证券'  in found_title:
            broker = 'huatai'
            broker_name = '华泰证券' 
       elif '国金太阳'  in found_title:
            broker = 'guojin'
            broker_name = '国金证券' 
       elif '申万宏源'  in found_title:
            broker = 'shenwan'
            broker_name = '申万宏源'  
       elif '通达信网上交易'  in found_title:
            broker = 'yinhe'
            broker_name = '银河证券'              
       else:
            broker = 'unknow'
            broker_name = '未知证券公司' 
    
    print(broker)  
    print(broker_name)  

    print(numChildWindows_list) 

    for i in range(len(brokers_lsit)):
#        print(found_title)
#        print(brokers_lsit[i])

        if found_title == brokers_lsit[i]:

            numChildWindows = int(numChildWindows_list[i])
            position_index_offset = int(positionIndex_list[i])
            balance_index_offset = int(balanceIndex_list[i])
            positionList_ev = eval(positionList_list[i])
            main_title = found_title
    print(main_title) 
    print(numChildWindows)



    #if top_hwnd == 0:
    if False:
        tkinter.messagebox.showerror('错误', '请先打开交易软件:  ' + main_title +' , 然后重启程序!!!  ')
        sys.exit()
        return top_hwnd, []
    else:
        #
        print("Closing pop windows ...")
        closePopupWindows(top_hwnd)  

        
        '''
        app = pywinauto.application.Application()
        #w_handle = pywinauto.findwindows.find_windows(title="国金太阳至强版V7.22 - [组合图-创业板]", class_name='TdxW_MainFrame_Class')[0]
        
        app.connect(handle=top_hwnd)
        client = app.window(handle=top_hwnd)
        print("client")
        print(client)
        ctrl = client['TreeView']  #SysTreeView32
        print("ctrl = client['TreeView']")
        print(ctrl)         
        treeviewitem = ctrl.GetItem(['查询', '资金股份']).ClickInput() 
        time.sleep(1)
        ctrl.GetItem(['对买对卖']).ClickInput() 
        '''
        all_accounnt = []
        subhwnd_list = []
        if len(top_hwnd_one) > 0:



            for i in range(len(top_hwnd_one)):
                account,subhwnd = searcHwndOfClient(top_hwnd_one[i],numChildWindows,position_index_offset,balance_index_offset)
                all_accounnt.append(account)
                subhwnd_list.append(subhwnd) 
            print (all_accounnt)

            choose = "-1"
            while not choose.isdigit() or \
                not ( choose.isdigit() and int(choose) in range(len(top_hwnd_one))):

                print('\n\n请选择对应的账户：\n ') 
                for i in range(len(top_hwnd_one)):
                    print('%d : %s:' % (i,all_accounnt[i]))   
                choose = input(':')
        else:
            choose = "0"
        
        print(int(choose))
        
        top_hwnd = top_hwnd_one[int(choose)]
        
        print(top_hwnd)
        
       
        #sub_hwnds = subhwnd_list[int(choose)]
        
        #account_name = all_accounnt[int(choose)]
        

        account_name,sub_hwnds =  searcHwndOfClient(top_hwnd,numChildWindows,position_index_offset,balance_index_offset)
        #print("sub_hwnds")
        #print(sub_hwnds)

        data_file =  account_name+".dat"




       
        b_init_ready = True
        
    return top_hwnd, sub_hwnds


def pickCodeFromItems(items_info):
    '''
    提取股票代码
    :param items_info: UI下各项输入信息
        stock_codes.append(item[0])
    :return:股票代码列表
    '''
    stock_codes = []
    for item in items_info:
        stock_codes.append(item[0])
    return stock_codes

def getsettings():

    seetings =""  
    for row in range(len(set_stock_info)):
                  
        if set_stock_info[row][6] \
                        and set_stock_info[row][1] and set_stock_info[row][2] > 0 \
                        and set_stock_info[row][3] and set_stock_info[row][4] :
            seetings = seetings + str(set_stock_info[row][0]) +"  "  \
                                + str(set_stock_info[row][1]) +"  "  \
                                + str(set_stock_info[row][2]) +"  "  \
                                + str(set_stock_info[row][3]) +"  "  \
                                + str(set_stock_info[row][4]) +"  "  \
                                + str(set_stock_info[row][5])+"\n\n" 
    return seetings
                    
                    



def getStockData(items_info):
    '''
    获取股票实时数据
    :param items_info:UI下各项输入信息
    :return:股票实时数据
    '''
    code_name_price = []
    stock_codes = pickCodeFromItems(items_info)
    try:
        df = ts.get_realtime_quotes(stock_codes)
        df_len = len(df)
        for stock_code in stock_codes:
            is_found = False
            for i in range(df_len):
                actual_code = df['code'][i]
                if stock_code == actual_code:
                    actual_name = df['name'][i]
                    pre_close = float(df['pre_close'][i])
                    if 'ST' in actual_name:
                        highest = str(round(pre_close * 1.04999, 2))
                        lowest = str(round(pre_close * 0.95001, 2))
                        code_name_price.append((actual_code, actual_name, df['price'][i], (highest, lowest)))
                    else:
                        highest = str(round(pre_close * 1.09999, 2))
                        lowest = str(round(pre_close * 0.90001, 2))
                        code_name_price.append((actual_code, actual_name, df['price'][i], (highest, lowest)))
                    is_found = True
                    break
            if is_found is False:
                code_name_price.append(('', '', '', ('', '')))
    except:
        code_name_price = [('', '', '', ('', ''))] * num_rows  # 网络不行，返回空
    return code_name_price


def SellNull(sub_hwnds):
    #Log("Sell null to keep active ...")
    global send_error
    try:
        



        setEditText(sub_hwnds[27][0], "0")
        time.sleep(0.3)
        click(sub_hwnds[29][0])
        time.sleep(0.3)

    except Exception as e:
        Log(str(e))      
           
        if not send_error:
              mailmsg(broker+account_name+"心跳失败。请重启程序！",str(e))
              send_error = True
     


top_hwnd, sub_hwnds = tradingInit() 

def monitor():
    '''
    实时监控函数
    :return:
    '''
    is_send = False
    is_send_afternoon = False
    global actual_stock_info, order_msg, is_ordered, set_stock_info,pre_order_result, \
           top_hwnd, sub_hwnds,\
           tip_msg
    count = 0
    #top_hwnd, sub_hwnds = tradingInit()

    old_banlance = getBanlence(1)
    old_position = getPositionStr()
    old_settings = getsettings()
    print (old_settings)




    #    tkinter.messagebox.showerror('错误', "07:10:00")

    #while True:‘’


    #	testsell(sub_hwnds)
    # 如果top_hwnd为零，直接终止循环
    while is_monitor and top_hwnd:
        if count % 60 == 0:

            SellNull(sub_hwnds)
            msg = closePopupWindows(top_hwnd)  
                #global g_debug_to_buffer
            now = datetime.datetime.now()
            otherStyleTime = now.strftime("%Y-%m-%d %H:%M:%S")        
            print (otherStyleTime)
    
            time.sleep(1)
        time.sleep(3)

        new_banlance = getBanlence(1)
        new_position = getPositionStr()
        new_settings = getsettings()
        #print (new_settings)
        #print (old_settings)
        #print (old_settings == new_settings)        

        if not is_send :
            if datetime.datetime.now().time() > \
               datetime.datetime.strptime('09:00:00', '%H:%M:%S').time() \
               and datetime.datetime.now().time() < \
               datetime.datetime.strptime('9:30:00', '%H:%M:%S').time() :
                
                print("Setings in the morning...")
                mailmsg(broker+account_name+" 仓位和设置" ,\
                    "余额：" +str (new_banlance) +"\n\n" +\
                    "持仓：" +str (new_position)  +"\n"+ \
                    "设置：" +str (new_settings)  +"\n\n")
                is_send = True

        if not is_send_afternoon :
            if datetime.datetime.now().time() > \
               datetime.datetime.strptime('12:30:00', '%H:%M:%S').time() \
               and datetime.datetime.now().time() < \
               datetime.datetime.strptime('13:00:00', '%H:%M:%S').time() :
                
                print("Setings in the afternoon...")




                mailmsg(broker+account_name+" 仓位和设置" ,\
                    "余额：" +str (new_banlance) +"\n\n" +\
                    "持仓：" +str (new_position)  +"\n"+ \
                    "设置：" +str (new_settings)  +"\n\n")


                is_send_afternoon = True      

        #if not new_banlance == old_banlance or not new_position == old_position:
        if not new_banlance == old_banlance \
           or not new_settings == old_settings:
           #or not new_position == old_position \
           

            if datetime.datetime.now().time() > \
               datetime.datetime.strptime('09:00:00', '%H:%M:%S').time() \
               and datetime.datetime.now().time() < \
               datetime.datetime.strptime('15:30:00', '%H:%M:%S').time() :
                



                mailmsg(broker+account_name+" 提醒" ,\
                 "旧余额：" + str  (old_banlance) +"\n\n"+\
                 "新余额：" +str (new_banlance) +"\n\n" +\
                 "旧持仓：" + str  (old_position) +"\n"+ \
                 "新持仓：" +str (new_position)  +"\n"+ \
                 "旧设置：" +str (old_settings)  +"\n\n"+ \
                 "新设置：" +str (new_settings)  +"\n\n")

                 
            old_banlance = new_banlance
            old_position = new_position
            old_settings = new_settings 



        count += 1
        if is_start:
            actual_stock_info = getStockData(set_stock_info)

            
            for row, (actual_code, actual_name, actual_price, stop_prices) in enumerate(actual_stock_info):

                if is_start and actual_code \
                        and set_stock_info[row][1] and set_stock_info[row][2] > 0 \
                        and set_stock_info[row][3] and set_stock_info[row][4] :
                        
                 
                        result = pre_order(actual_code, stop_prices,
                              set_stock_info[row][4], set_stock_info[row][3])
            
                        #print(order_msg)
                        pre_order_result[row] = result
                        #print(pre_order_result[row])
 

            # print('actual_stock_info', actual_stock_info)
            for row, (actual_code, actual_name, actual_price, stop_prices) in enumerate(actual_stock_info):
                if is_start and actual_code and is_ordered[row] == 1 \
                        and set_stock_info[row][6] \
                        and set_stock_info[row][1] and set_stock_info[row][2] > 0 \
                        and set_stock_info[row][3] and set_stock_info[row][4] \
                        and datetime.datetime.now().time() > set_stock_info[row][5]:
                    if is_start and set_stock_info[row][1] == '>' and float(actual_price) > set_stock_info[row][2] \
                        and pre_order_result[row] == "ok":
                        dt = datetime.datetime.now()
                        order(top_hwnd, sub_hwnds, actual_code, stop_prices,
                              set_stock_info[row][4], set_stock_info[row][3])
                        
                        print("order >  ******************")
                        msg,msg1 = closePopupWindows(top_hwnd)  
                        Log(msg1)
                        Log(msg)

                        
                        length = msg.find(",")+msg.find("，")+1
                        tip_msg[row] = msg
                        #result = msg[0:length]
                        if "合同号" in msg: 
                            result = "下单成功"
                        elif "原因" in msg: 
                            result = "下单失败"
                        else:
                            result = "结果未知"

                        
                        
                        
                        if   order_retry[row]  < 3:
                            print(order_retry[row])
                            if set_stock_info[row][3] == "B":
                                mailmsg(set_stock_info[row][3] +"|"+ actual_code +" P|"+stop_prices[0]+" M|"+set_stock_info[row][4],msg+'\n' + broker+account_name)
                            else:    
                                mailmsg(set_stock_info[row][3] +"|"+ actual_code +" P|"+stop_prices[1]+" M|"+set_stock_info[row][4],msg+'\n' + broker+account_name)


                        order_msg.append(
                            (dt.strftime('%x'), dt.strftime('%X'), actual_code,
                             actual_name, set_stock_info[row][3],
                             actual_price, set_stock_info[row][4], '已下单',result))
                        #print(order_msg)
                        order_result[row] = result

                        is_ordered[row] = 0
                        
                        if msg == "null":
                             Log("order failed!")
                             order_retry[row] =   order_retry[row] + 1
                             print(order_retry[row])

                             is_ordered[row] = 1
                        else:
                             order_retry[row] = 0

                    if is_start and set_stock_info[row][1] == '<' and float(actual_price) < set_stock_info[row][2] \
                        and pre_order_result[row] == "ok":
                        dt = datetime.datetime.now()
                        order(top_hwnd, sub_hwnds, actual_code, stop_prices,
                              set_stock_info[row][4], set_stock_info[row][3])


                        print("order <  ******************")
                        msg,msg1 = closePopupWindows(top_hwnd)  
                        Log(msg1)
                        Log(msg)

                        
                        #length = msg.find(",")+msg.find("，")+1
                        #result = msg[0:length]
                        
                        length = msg.find(",")+msg.find("，")+1
                        tip_msg[row] = msg
                        #result = msg[0:length]
                        if "合同号" in msg: 
                            result = "下单成功"
                        elif "原因" in msg: 
                            result = "下单失败"
                        else:
                            result = "结果未知"


                        if   order_retry[row]  < 3:
                            print(order_retry[row])
                            if set_stock_info[row][3] == "B":
                                mailmsg(set_stock_info[row][3] +"|"+ actual_code +"|"+stop_prices[0]+"|"+set_stock_info[row][4],msg+'\n' + broker+account_name)
                            else:    
                                mailmsg(set_stock_info[row][3] +"|"+ actual_code +"|"+stop_prices[1]+"|"+set_stock_info[row][4],msg+'\n' + broker+account_name)


                        order_msg.append(
                            (dt.strftime('%x'), dt.strftime('%X'), actual_code,
                             actual_name, set_stock_info[row][3],
                             actual_price, set_stock_info[row][4], '已下单',result))
                        #print(order_msg)
                        order_result[row] = result

                        is_ordered[row] = 0
                        
                        if msg == "null":
                             Log("order failed!")
                             order_retry[row] =   order_retry[row] + 1
                             print(order_retry[row])

                             is_ordered[row] = 1
                        else:
                             order_retry[row] = 0


class ToolTip( Toplevel ):
    """
    Provides a ToolTip widget for Tkinter.
    To apply a ToolTip to any Tkinter widget, simply pass the widget to the
    ToolTip constructor
    """ 
    def __init__( self, wdgt,index, msg=None, msgFunc=None, delay=1, follow=True ):
        """
        Initialize the ToolTip
        
        Arguments:
          wdgt: The widget this ToolTip is assigned to
          msg:  A static string message assigned to the ToolTip
          msgFunc: A function that retrieves a string to use as the ToolTip text
          delay:   The delay in seconds before the ToolTip appears(may be float)
          follow:  If True, the ToolTip follows motion, otherwise hides
        """
        self.msg = msg

        self.wdgt = wdgt
        self.parent = self.wdgt.master                                          # The parent of the ToolTip is the parent of the ToolTips widget
        Toplevel.__init__( self, self.parent, bg='black', padx=1, pady=1 )      # Initalise the Toplevel
        self.withdraw()                                                         # Hide initially
        self.overrideredirect( True )                                           # The ToolTip Toplevel should have no frame or title bar
        
        self.msgVar = StringVar()                                               # The msgVar will contain the text displayed by the ToolTip        
        if msg == None:                                                         
            self.msgVar.set( 'No message provided' )
        else:
            self.msgVar.set( msg )
        self.msgFunc = msgFunc
        self.delay = delay
        self.follow = follow
        self.visible = 0
        self.lastMotion = 0
        Message( self, textvariable=self.msgVar, bg='#FFFFDD',
                 aspect=1000 ).grid()                                           # The test of the ToolTip is displayed in a Message widget
        self.wdgt.bind( '<Enter>', self.spawn, '+' )                            # Add bindings to the widget.  This will NOT override bindings that the widget already has
        self.wdgt.bind( '<Leave>', self.hide, '+' )
        self.wdgt.bind( '<Motion>', self.move, '+' )
        
    def spawn( self, event=None ):
        """
        Spawn the ToolTip.  This simply makes the ToolTip eligible for display.
        Usually this is caused by entering the widget
        
        Arguments:
          event: The event that called this funciton
        """
        self.visible = 1
        self.after( int( self.delay * 1000 ), self.show )                       # The after function takes a time argument in miliseconds
        
    def show( self ):
        """
        Displays the ToolTip if the time delay has been long enough
        """
        if self.visible == 1 and time.time() - self.lastMotion > self.delay:
            self.visible = 2
        if self.visible == 2:
            self.deiconify()
            
    def move( self, event ):
        """
        Processes motion within the widget.
        
        Arguments:
          event: The event that called this function
        """
        self.lastMotion = time.time()
        if self.follow == False:                                                # If the follow flag is not set, motion within the widget will make the ToolTip dissapear
            self.withdraw()
            self.visible = 1
        self.geometry( '+%i+%i' % ( event.x_root+10, event.y_root+10 ) )        # Offset the ToolTip 10x10 pixes southwest of the pointer
        try:
            
            self.msgVar.set( self.msgFunc() )                                   # Try to call the message function.  Will not change the message if the message function is None or the message function fails
        except:
            self.msgVar.set( self.msg )                                  
            #print ("error:msgFunc()")
            pass
        self.after( int( self.delay * 1000 ), self.show )
            
    def hide( self, event=None ):
        """
        Hides the ToolTip.  Usually this is caused by leaving the widget
        
        Arguments:
          event: The event that called this function
        """
        self.visible = 0
        self.withdraw()


class StockGui:
    def __init__(self):


        self.window = Tk()

        self.window.title(broker_name +  "  账号： "+account_name)
        self.window.resizable(0, 0)

        #enable_states = [tk.IntVar()] * num_rows

        frame1 = Frame(self.window)
        frame1.pack(padx=10, pady=10)

        Label(frame1, text="代码", width=6, justify=LEFT).grid(
            row=1, column=1, padx=5, pady=5)
        Label(frame1, text="股票名称", width=8, justify=LEFT).grid(
            row=1, column=2, padx=5, pady=5)
        Label(frame1, text="当前价", width=6, justify=LEFT).grid(
            row=1, column=3, padx=5, pady=5)
        Label(frame1, text="关系", width=4, justify=LEFT).grid(
            row=1, column=4, padx=5, pady=5)
        Label(frame1, text="目标价", width=8, justify=LEFT).grid(
            row=1, column=5, padx=5, pady=5)
        Label(frame1, text="方向", width=4, justify=LEFT).grid(
            row=1, column=6, padx=5, pady=5)
        Label(frame1, text="数量", width=8, justify=LEFT).grid(
            row=1, column=7, padx=5, pady=5)
        Label(frame1, text="时间可选", width=8, justify=LEFT).grid(
            row=1, column=8, padx=5, pady=5)
        Label(frame1, text="状态", width=6, justify=LEFT).grid(
            row=1, column=9, padx=5, pady=5)
        Label(frame1, text="结果", width=10, justify=LEFT).grid(
            row=1, column=10, padx=5, pady=5)

        Label(frame1, text="开关", width=4, justify=LEFT).grid(
            row=1, column=11, padx=5, pady=5)


        '''

        Label(frame1, text="代码", width=6, justify=CENTER).grid(
            row=1, column=1, padx=5, pady=5)
        Label(frame1, text="股票名称", width=8, justify=CENTER).grid(
            row=1, column=2, padx=5, pady=5)
        Label(frame1, text="当前价", width=6, justify=CENTER).grid(
            row=1, column=3, padx=5, pady=5)
        Label(frame1, text="关系", width=4, justify=CENTER).grid(
            row=1, column=4, padx=5, pady=5)
        Label(frame1, text="目标价", width=8, justify=CENTER).grid(
            row=1, column=5, padx=5, pady=5)
        Label(frame1, text="方向", width=4, justify=CENTER).grid(
            row=1, column=6, padx=5, pady=5)
        Label(frame1, text="数量", width=8, justify=CENTER).grid(
            row=1, column=7, padx=5, pady=5)
        Label(frame1, text="时间可选", width=8, justify=CENTER).grid(
            row=1, column=8, padx=5, pady=5)
        Label(frame1, text="状态", width=6, justify=CENTER).grid(
            row=1, column=9, padx=5, pady=5)
        Label(frame1, text="结果", width=10, justify=CENTER).grid(
            row=1, column=10, padx=5, pady=5)

        Label(frame1, text="开关", width=4, justify=CENTER).grid(
            row=1, column=11, padx=5, pady=5)
        '''
        self.rows =  num_rows
        self.cols = 11

        self.variable = []
        self.tip_msg = []
        '''
        for row in range(self.rows):
            self.variable.append([])
            for col in range(self.cols-1):
                temp = StringVar()
                self.variable[row].append(temp)
            temp = IntVar()
            self.variable[row].append(temp)
                for row in range(self.rows):
            Entry(frame1, textvariable=self.variable[row][0],
                  width=6).grid(row=row + 2, column=1, padx=5, pady=5)
            Entry(frame1, textvariable=self.variable[row][1], state=DISABLED,
                  width=8).grid(row=row + 2, column=2, padx=5, pady=5)
            Entry(frame1, textvariable=self.variable[row][2], state=DISABLED,
                  width=6).grid(row=row + 2, column=3, padx=5, pady=5)
            Combobox(frame1, values=('<', '>'), textvariable=self.variable[row][3],
                     width=2).grid(row=row + 2, column=4, padx=5, pady=5)
            Spinbox(frame1, from_=0, to=1000, textvariable=self.variable[row][4],
                    increment=0.01, width=6).grid(row=row + 2, column=5, padx=5, pady=5)
            Combobox(frame1, values=('B', 'S'), textvariable=self.variable[row][5],
                     width=2).grid(row=row + 2, column=6, padx=5, pady=5)
            Spinbox(frame1, from_=0, to=100000, textvariable=self.variable[row][6],
                    increment=100, width=6).grid(row=row + 2, column=7, padx=5, pady=5)
            Entry(frame1, textvariable=self.variable[row][7],
                  width=8).grid(row=row + 2, column=8, padx=5, pady=5)
            Entry(frame1, textvariable=self.variable[row][8], state=DISABLED,
                  width=6).grid(row=row + 2, column=9, padx=5, pady=5)
            Entry(frame1, textvariable=self.variable[row][9], state=DISABLED,
                  width=10).grid(row=row + 2, column=10, padx=5, pady=5)
            #Checkbutton(frame1, textvariable=str(self.variable[row][10]), variable=self.variable[row],
            Checkbutton(frame1, textvariable=str(self.variable[row][10]), variable=self.variable[row][10],
                  width=2).grid(row=row + 2, column=11, padx=5, pady=5)    
        '''
        self.ctl = [[0 for i in range(self.cols)] for i in range(self.rows)]
        
        for row in range(self.rows):
            self.variable.append([])
            
            temp = StringVar()
            self.tip_msg.append(temp)

            for col in range(self.cols-1):
                temp = StringVar()
                self.variable[row].append(temp)
            temp = IntVar()
            self.variable[row].append(temp)
        
        result_ctl_list=[]
        self.tip_ctl_list=[]
        
        follow = True
        #msgFunc = self.show_tip
        msgFunc = None
        delay = 0
        self.tip_msg = [" "] * self.rows

        for row in range(self.rows):
            self.ctl[row][0] =  Entry(frame1, textvariable=self.variable[row][0],
                  width=6)
            self.ctl[row][0].grid(row=row + 2, column=1, padx=5, pady=5)
            self.ctl[row][1] =  Entry(frame1, textvariable=self.variable[row][1], state=DISABLED,
                  width=8)
            self.ctl[row][1].grid(row=row + 2, column=2, padx=5, pady=5)
            self.ctl[row][2] = Entry(frame1, textvariable=self.variable[row][2], state=DISABLED,
                  width=6)
            self.ctl[row][2].grid(row=row + 2, column=3, padx=5, pady=5)
            self.ctl[row][3] = Combobox(frame1, values=('<', '>'), textvariable=self.variable[row][3],
                     width=2)
            self.ctl[row][3].grid(row=row + 2, column=4, padx=5, pady=5)
            self.ctl[row][4] = Spinbox(frame1, from_=0, to=1000, textvariable=self.variable[row][4],
                    increment=0.01, width=6)
            self.ctl[row][4].grid(row=row + 2, column=5, padx=5, pady=5)
            self.ctl[row][5] = Combobox(frame1, values=('B', 'S'), textvariable=self.variable[row][5],
                     width=2)
            self.ctl[row][5].grid(row=row + 2, column=6, padx=5, pady=5)
            self.ctl[row][6] = Spinbox(frame1, from_=0, to=100000, textvariable=self.variable[row][6],
                    increment=100, width=6)
            self.ctl[row][6].grid(row=row + 2, column=7, padx=5, pady=5)
            self.ctl[row][7] = Entry(frame1, textvariable=self.variable[row][7],
                  width=8)
            self.ctl[row][7].grid(row=row + 2, column=8, padx=5, pady=5)
            self.ctl[row][8] = Entry(frame1, textvariable=self.variable[row][8], state=DISABLED,
                  width=6)
            self.ctl[row][8].grid(row=row + 2, column=9, padx=5, pady=5)
            

            #self.ctl[row][9] = Entry(frame1, textvariable=self.variable[row][9], state=DISABLED,
             #     width=10,bg = self.ctlclr[row][9].get())

            self.ctl[row][9] = Entry(frame1, textvariable=self.variable[row][9], state=DISABLED,
                 width=10)
            result_ctl_list.append(self.ctl[row][9])
            tipctl = ToolTip( result_ctl_list[-1], index = row,msg=self.tip_msg[row], msgFunc=msgFunc, follow=follow, delay=delay)
            self.tip_ctl_list.append(tipctl)
            #ToolTip( result_ctl_list[-1], msg=self.tip_msg[row], msgFunc=msgFunc, follow=follow, delay=delay)
            
            #self.ctl[row][9] = Label(frame1, bg = "red", width=8, justify=LEFT)
            #self.mytest = Label(frame1, bg = "red", width=8, justify=LEFT)
           # print(self.ctl[row][9])

            self.ctl[row][9].grid(row=row + 2, column=10, padx=5, pady=5)

            #Checkbutton(frame1, textvariable=str(self.variable[row][10]), variable=self.variable[row],
            self.ctl[row][10] = Checkbutton(frame1, textvariable=str(self.variable[row][10]), variable=self.variable[row][10],
                  width=2)
            self.ctl[row][10].grid(row=row + 2, column=11, padx=5, pady=5)    
        
            #print(self.ctl[row])  

        

        frame3 = Frame(self.window)
        frame3.pack(padx=10, pady=10)
        self.start_bt = Button(frame3, text="开始", command=self.start)
        self.start_bt.pack(side=LEFT)
        self.set_bt = Button(frame3, text='重置买卖', command=self.setFlags)
        self.set_bt.pack(side=LEFT)
        Button(frame3, text="历史记录", command=self.displayHisRecords).pack(side=LEFT)
        #Button(frame3, text='保存', command=self.save).pack(side=LEFT)
        Button(frame3, text='持仓查询', command=self.position).pack(side=LEFT)
        Button(frame3, text='资金查询', command=self.cash).pack(side=LEFT)
        #Button(frame3, text='保存', command=self.save).pack(side=LEFT)
        self.save_bt = Button(frame3, text='保存', command=self.save)
        self.save_bt.pack(side=LEFT)
        self.load_bt = Button(frame3, text='载入', command=self.load)
        self.load_bt.pack(side=LEFT)
        self.load()
        self.window.protocol(name="WM_DELETE_WINDOW", func=self.close)
        self.window.after(100, self.updateControls)
        self.window.mainloop()

    def displayHisRecords(self):
        '''
        显示历史信息
        :return:
        '''
        global order_msg
        tp = Toplevel()
        tp.title('历史记录')
        tp.resizable(0, 1)
        scrollbar = Scrollbar(tp)
        scrollbar.pack(side=RIGHT, fill=Y)
        col_name = ['日期', '时间', '证券代码', '证券名称', '方向', '价格', '数量', '备注','结果']
        tree = Treeview(
            tp, show='headings', columns=col_name, height=30, yscrollcommand=scrollbar.set)
        tree.pack(expand=1, fill=Y)
        scrollbar.config(command=tree.yview)
        for name in col_name:
            tree.heading(name, text=name)
            tree.column(name, width=70, anchor=CENTER)

        for msg in order_msg:
            tree.insert('', 0, values=msg)



    def cash(self):
        '''
        持仓查询
        :return:
        '''
        #showinfo(title='',message='') 
        msg = getBanlence(-1)
        output = ""
        for item in msg :
            output = output + str(item) + "\n"

        print(msg)
        tkinter.messagebox.showinfo('资金情况',output)


    def position(self):
        '''
        持仓查询
        :return:
        '''
        tp1 = Toplevel()
        tp1.title('持仓记录')
        tp1.resizable(0, 1)
        scrollbar = Scrollbar(tp1)
        scrollbar.pack(side=RIGHT, fill=Y)
        col_name = ['证券代码', '证券名称', '股份余额', '可用股份', '成本价', '当前价','最新市值', '浮动盈亏','盈亏比例']
        tree = Treeview(
            tp1, show='headings', columns=col_name, height=10, yscrollcommand=scrollbar.set)
        tree.pack(expand=1, fill=Y)
        scrollbar.config(command=tree.yview)
        for name in col_name:
            tree.heading(name, text=name)
            tree.column(name, width=70, anchor=CENTER)
        newposition = getPosition()
        for msg in newposition:
            tree.insert('', 0, values=msg)

    def save(self):
        '''
        保存设置
        :return:
        '''
        global set_stock_info, order_msg, actual_stock_info
        set_stock_info = self.getItems()
        #with open('stockInfo.dat', 'wb') as fp:
        try:
            f=open(data_file,'rb')
        except Exception as e:
            Log(str(e)) 
            f=open(data_file,'w')
            f.close()


        with open(data_file, 'wb') as fp:
            pickle.dump(set_stock_info, fp)
            # pickle.dump(actual_stock_info, fp)
            pickle.dump(order_msg, fp)
        self.save_bt['state'] = DISABLED
        self.save_bt.configure(bg="white")

    def load(self):
        '''
        载入设置
        :return:
        '''
        global set_stock_info, order_msg, actual_stock_info
        
        try:
            f=open(data_file,'rb')
        except Exception as e:
            Log(str(e)) 
            f=open(data_file,'w')
            f.close()
            self.save()


        with open(data_file, 'rb') as fp:
            set_stock_info = pickle.load(fp)
            # actual_stock_info = pickle.load(fp)
            order_msg = pickle.load(fp)
            
        for row in range(self.rows):
            for col in range(self.cols):
                if col == 0:
                    self.variable[row][col].set(set_stock_info[row][0])
                elif col == 3:
                    self.variable[row][col].set(set_stock_info[row][1])
                elif col == 4:
                    self.variable[row][col].set(set_stock_info[row][2])
                elif col == 5:
                    self.variable[row][col].set(set_stock_info[row][3])
                elif col == 6:
                    self.variable[row][col].set(set_stock_info[row][4])
                elif col == 7:
                    temp = set_stock_info[row][5].strftime('%X')
                    if temp == '00:00:00':
                        self.variable[row][col].set('')
                    else:
                        self.variable[row][col].set(temp)
                elif col == 10:
                    self.variable[row][col].set(set_stock_info[row][6])
        self.save_bt['state'] = DISABLED
        self.save_bt.configure(bg="white")


    def setFlags(self):
        '''
        重置买卖标志
        :return:
        '''
        global is_start, is_ordered
        if is_start is False:
            is_ordered = [1] * num_rows

    def updateControls(self):
        '''
        实时股票名称、价格、状态信息
        :return:
        '''
        global set_stock_info, actual_stock_info, is_start,is_ordered

        set_stock_info_old = set_stock_info
        set_stock_info = self.getItems()
        if not set_stock_info == set_stock_info_old:
            rows =len(set_stock_info)
            try:
                for i in range(rows):
                    if not set_stock_info[i] == set_stock_info_old[i]:
                       #print(i)
                       is_ordered[i] = 1
            except Exception as e:    
        
                print(e)

            self.save_bt['state'] = NORMAL
            self.save_bt.configure(bg="yellow")


        if is_start:
           #print('actual_stock_info', actual_stock_info)
            for row, (actual_code, actual_name, actual_price, _) in enumerate(actual_stock_info):
                self.variable[row][1].set(actual_name)
                self.variable[row][2].set(str(actual_price))
                #print actual_code
                if actual_code \
                        and set_stock_info[row][1] and set_stock_info[row][2] > 0 \
                        and set_stock_info[row][3] and set_stock_info[row][4] :


                #if actual_code:
                    set_stock_info[row][6]=self.variable[row][10].get()  
                    if set_stock_info[row][6]: #opened
                        
                        if is_ordered[row] == 1:
                            self.variable[row][8].set('监控中')
                            self.ctl[row][8].configure(bg="green")
                            self.variable[row][9].set(pre_order_result[row])

                            self.tip_ctl_list[row].msg = pre_order_result[row]

                            if  pre_order_result[row] == "ok":
                                self.ctl[row][9].configure(bg="green")
                            else:
                                self.ctl[row][9].configure(bg="red")
                            #self.mytest.configure(bg="red")
                        elif is_ordered[row] == 0:
                            self.variable[row][8].set('已下单')
                            self.ctl[row][8].configure(bg="yellow")
                            
                            self.tip_ctl_list[row].msg = tip_msg[row]
                            
                            self.variable[row][9].set(order_result[row])
                            if order_result[row] == "下单成功":
                                self.ctl[row][9].configure(bg="green")
                            elif order_result[row] == "下单失败":
                                self.ctl[row][9].configure(bg="red")
                            else:
                                self.ctl[row][9].configure(bg="yellow")
                            

                    else:
                      self.variable[row][8].set('已关闭')
                      self.ctl[row][8].configure(bg="yellow")
                      self.variable[row][9].set(pre_order_result[row])
                      if  pre_order_result[row] == "ok":
                                self.ctl[row][9].configure(bg="green")
                      else:
                                self.ctl[row][9].configure(bg="red")
                      is_ordered[row] = 1    #to be reppened                

                      self.tip_ctl_list[row].msg = pre_order_result[row]
                      
                            


                else:
                    self.variable[row][8].set('')
                    self.variable[row][9].set('')
                    self.ctl[row][8].configure(bg="white")
                    self.ctl[row][9].configure(bg="white")
                    self.tip_ctl_list[row].msg = "信息不完整"



        self.window.after(1000, self.updateControls)

    def start(self):
        '''
        启动停止
        :return:
        '''
        global is_start,set_stock_info
        if is_start is False:
            is_start = True
        else:
            is_start = False

        if is_start:
            set_stock_info = self.getItems()
            # print(set_stock_info)
            self.start_bt['text'] = '停止'
          #  self.ctl[0][9].configure(bg="green")
           # self.ctl[1][9].configure(bg="blue")
         #   self.ctl[1][9].configure(bg="red")

            self.set_bt['state'] = DISABLED
            self.load_bt['state'] = DISABLED

            for row in range(self.rows):
                for col in range(self.cols):
                    self.ctl[row][8]['state'] = NORMAL
                    self.ctl[row][9]['state'] = NORMAL
                    if not col == 8 and not col == 9:
                        self.ctl[row][col]['state'] = DISABLED

        else:
            self.start_bt['text'] = '开始'
            self.set_bt['state'] = NORMAL
            self.load_bt['state'] = NORMAL
            for row in range(self.rows):
                for col in range(self.cols):
                    if not col == 1 and not col == 2:
                        self.ctl[row][col]['state'] = NORMAL


    def close(self):
        '''
        关闭程序时，停止monitor线程
        :return:
        '''
        global is_monitor
        is_monitor = False
        self.window.quit()

    def getItems(self):
        '''
        获取UI上用户输入的各项数据，
        '''
        #global set_stock_info
        set_stock_info = []

        # 获取买卖价格数量输入项等
        for row in range(self.rows):
            set_stock_info.append([])
            for col in range(self.cols-1):
                temp = self.variable[row][col].get().strip()
                if col == 0:
                    if len(temp) == 6 and temp.isdigit():  # 判断股票代码是否为6位数
                        set_stock_info[row].append(temp)
                    else:
                        set_stock_info[row].append('')
                elif col == 3:
                    if temp in ('>', '<'):
                        set_stock_info[row].append(temp)
                    else:
                        set_stock_info[row].append('')
                elif col == 4:
                    try:
                        price = float(temp)
                        if price > 0:
                            set_stock_info[row].append(price)  # 把价格转为数字
                        else:
                            set_stock_info[row].append(0)
                    except ValueError:
                        set_stock_info[row].append(0)
                elif col == 5:
                    if temp in ('B', 'S'):
                        set_stock_info[row].append(temp)
                    else:
                        set_stock_info[row].append('')
                elif col == 6:
                    if temp.isdigit() and int(temp) >= 100:
                        set_stock_info[row].append(str(int(temp) // 100 * 100))
                    else:
                        set_stock_info[row].append('')
                elif col == 7:
                    try:
                        set_stock_info[row].append(datetime.datetime.strptime(temp, '%H:%M:%S').time())
                    except ValueError:
                        set_stock_info[row].append(datetime.datetime.strptime('09:31:00', '%H:%M:%S').time())
            
                    #temp = self.variable[row][self.cols-1].get()
            temp = self.variable[row][self.cols-1].get()
            #print(temp)
            set_stock_info[row].append(temp)  
        return set_stock_info              
            


if __name__ == '__main__':
    
    t1 = threading.Thread(target=StockGui)    
    t2 = threading.Thread(target=monitor)
    t2.start()
    while not b_init_ready:
       print(".")
       time.sleep(2)
             
    t1.start()
    t1.join()
    t2.join()


    
    
    

