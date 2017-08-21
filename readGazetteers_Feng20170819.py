
# coding: utf-8

# In[ ]:



# In[7]:


#coding=utf-8
from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

import argparse

#_______以上为读取google sheets所用函数库———

import os
import csv
import requests
import codecs
import time
import random
import sys
from datetime import datetime
from collections import deque


from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.header import Header
import smtplib

Cookie = '*******'
url = ["http://server.wenzibase.com.ezp-prod1.hul.harvard.edu/downWords.action?code=8030&libID=","&buyId=***&bookID=","&pageID="]
database_dict = {"古籍库" : "***", "基本古籍库" : "***", "方志库": "***"}

#以下sendEmail函数使用了Jason Hu的Python SMTP 发送带附件电子邮件代码
def sendEmail(authInfo, fromAdd, toAdd, subject, plainText, att_file):

        strFrom = fromAdd
        strTo = '; '.join(toAdd)

        server = authInfo.get('server')
        smtpPort = 587
        sslPort = 465
        user = authInfo.get('user')
        passwd = authInfo.get('password')

        if not (server and user and passwd) :
                print ('incomplete login info, exit now')
                return

        # 设定root信息
        msgRoot = MIMEMultipart('related')
        msgRoot['Subject'] = subject
        msgRoot['From'] = '%s<%s>' % (Header('爱如生检索记录前后三页全文', 'utf-8'), strFrom)
        msgRoot['To'] = strTo

        # 邮件正文内容
        msgText = MIMEText(plainText, 'plain', 'utf-8')
        msgRoot.attach(msgText)
        
        msgAlternative = MIMEMultipart('alternative')
        msgRoot.attach(msgAlternative)
        
        
        # mail_msg = """
        #         <p>Python 邮件发送测试...</p>
        #         <p><a href="http://www.runoob.com">菜鸟教程链接</a></p>
        #         <p>图片演示：</p>
        #         <p><img src="cid:pic_attach"></p>
        # """
        # msgAlternative.attach(MIMEText(mail_msg, 'html', 'utf-8'))

        #设定内置图片信息
        fp = open(att_file, 'rb')
        msgImage = MIMEText(fp.read())
        msgImage["Content-Type"] = 'application/octet-stream'
        #filename可自定义，供邮件中显示
        print(att_file)
        msgImage["Content-Disposition"] = 'attachment; filename="{}"'.format(att_file.encode('utf-8'))
        fp.close()
        msgImage.add_header('Content-ID', '<pic_attach>')
        msgAlternative.attach(msgImage)

        try:
                #发送邮件
                #smtp = smtplib.SMTP()
                #smtp.connect(server, smtpPort)
                #ssl加密方式，通信过程加密，邮件数据安全
                smtp = smtplib.SMTP_SSL()
                smtp.connect(server, sslPort)

                #设定调试级别，依情况而定
                # smtp.set_debuglevel(1)
                smtp.login(user, passwd)
                smtp.sendmail(strFrom, toAdd, msgRoot.as_string())
                smtp.quit()
                print ("邮件发送成功!")
        except Exception, e:
                print ("失败：" + str(e))

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Sheets API Python Quickstart'



def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'sheets.googleapis.com-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def google_sheets(pointer):
    """Shows basic usage of the Sheets API.

    Creates a Sheets API service object and prints the names and majors of
    students in a sample spreadsheet:
    https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit
    """
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

    spreadsheetId = '*******'
    
    rangeName = 'Raw!A'+str(pointer)+':J100000'
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    else:
        return(values)
       #print('Name, Major:')
          #for row in values:
            # Print columns A and E, which correspond to indices 0 and 4.
            #print('%s, %s' % (row[0], row[1]))


    
#Cookie = '_ga=GA1.2.115038926.1457945531; hulaccess=1.3|76.119.237.142|20170813213841EDT|pin|91303817|harvard|FAS-100688|OFFI|1206|hul-prod; user_OpenURL="http://sfx.hul.harvard.edu:80/sfx_local/"; ezproxyezpprod1=5ZzTxwV4XHUUWYm; JSESSIONID=DB55511E24F621C901B42FA07821A7D2'
#url = ["http://server.wenzibase.com.ezp-prod1.hul.harvard.edu/downWords.action?code=8030&libID=223&buyId=184&bookID=","&pageID="]    
#抓取文本函数
def read(url,header):
    global err_count
    global err_read_count
    err = '读取数据错误'
    #type(err)
    r = requests.get(url,headers=header)
    str=r.text[0:6]
    #print(len(r.text))
    #print(r.text[0:7])
    #利用递归再次抓取同一记录，处理“读取数据错误”问题
    if r.text[0:6].encode('utf-8') == err:
        #print('y')
        err_read_count = err_read_count + 1
        if err_read_count > 10:
            print('Can not read')
            return('Can not read')
        return read(url,header)
    #当遇到错误时，利用递归重复读取一定次数，如仍不能解决则退出
    if 'harvard' in r.text or 'Apache'in r.text or r.status_code <> requests.codes.ok:
        err_count = err_count + 1
        if err_count == 10:
            print('网络出现问题的时间：')
            print(unicode(datetime.now()))
            sys.exit()
        return read(fullurl, header)
    return(r.text)    

def list2csv(list,file):
    #global count
    with open(file,'ab') as f:
        for i in range(0,len(list)):
            #print(list[i])
            w=csv.writer(f)
            w.writerow([list[i][0].encode("utf-8"),list[i][1].encode("utf-8"),list[i][2].encode("utf-8"),list[i][3].encode("utf-8"),list[i][4].encode("utf-8"),list[i][5].encode("utf-8"),list[i][6].encode("utf-8"),list[i][7].encode("utf-8"),list[i][8].encode("utf-8"),list[i][9].encode("utf-8")])

def get_last_row(csv_filename):
    with open(csv_filename, 'r') as f:
        try:
            lastrow = deque(csv.reader(f), 1)[0]
        except IndexError:  # empty file
            lastrow = None
        return lastrow
#发送邮件至Harvard用户邮箱    
def mail_to_harvard_user(toAdd, att_file):
    authInfo = {}
    authInfo['server'] = '******'
    authInfo['user'] = '******'
    authInfo['password'] = '******'
    fromAdd = '******'
    #toAdd = [row[2]]
    subject = 'Email Subject'
    plainText = 'Email Content'
    #att_file = row[0]+' in '+row[1]+'.csv'
    print(att_file)
    sendEmail(authInfo, fromAdd, toAdd, subject, plainText, att_file)


#------------主程序----------------
#读取上一次成功下载的书号与页码
#with open('log.txt','rb') as csvlog:
#    log = csv.reader(csvlog)
#    for line in log:
#        row_last_time = line
#        print('row last time processed is:')
#       print(row_last_time)
#按照下载列表逐行处理
list_request = []
contents_list = []
list_request_count = 0
#google sheets表指针
#pointer = 0
#【读取google sheets表指针，从下一条记录开始读取】
#with open('working_log.txt', 'rb') as csvfile:
    #spamreader = csv.reader(csvfile)
    
#lines = spamreader.readlines()
lastline = get_last_row('working_log.txt')
pointer = lastline[1]
print(pointer)
begin_record_pointer = int(pointer) + 1
list_request = google_sheets(begin_record_pointer)
#print(list_request)
#google sheets表指针
if list_request == None:
    print('没有新的请求可以处理')
    sys.exit()
for row in list_request:
    pointer = int(pointer) + 1
    list_request_count =list_request_count + 1
    #如果是读取数据第一行
    if row == list_request[0]:
        email_addr = row[1].encode('utf-8')
        print(email_addr)
        database_ID = database_dict[row[2].encode('utf-8')]
        print(database_ID)
        timestampe_last = row[0]
    #如果一个用户提交数据已结束
    elif row[0] <> timestampe_last:
        list2csv(contents_list, email_addr+timestampe_last+'.csv')
        #email()
        mail_to_harvard_user(email_addr, email_addr+timestampe_last+'.csv')
        #【记录已完成行数】
        date = unicode(datetime.now())
        with open('working_log.txt','ab') as log:
            w = csv.writer(log)
            w.writerow([date,pointer-1])                                
        email_addr = row[1]
        database_ID = database_dict[row[2].encode('utf-8')]
        timestampe_last = row[0]
        #count = 0
        contents_list = []
    #如果仍然是同一用户的数据
    else:
        #print('row will be deal with is:')
        #print(row)
        #【如果与上一本书相同，且page_num_begin<page_num_end_last则将后者作为起始页码
        page_num_begin  = int(row[2]) - 3
        page_num_end = int(row[2]) + 3
        contents = ''
        for page in range(page_num_begin, page_num_end+1):
            err_count = 0
            err_read_count =0
            if len(row[1]) < 5:
                page_ID = row[1].zfill(5)
            fullurl = url[0]+database_ID+url[1]+page_ID.decode('utf-8')+url[2]+str(page)
            print(fullurl)
            header = {'Accept':'*/*','Accept-Language':'zh-CN','Referer':'http://server.wenzibase.com.ezp-prod1.hul.harvard.edu/userReadAction.action?prId=77&page=fz12','x-flash-version':'25,0,0,171','Content-Type':'application/x-amf','Accept-Encoding':'gzip, deflate','User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/51.0.2704.79 Safari/537.36 Edge/14.14393','Cookie':Cookie}
            #headers = {'Accept':'*/*','Accept-Language':'zh-CN','Referer':'http://server.wenzibase.com.ezp-prod1.hul.harvard.edu/book/fz.swf','x-flash-version':'25,0,0,171','Content-Type':'application/x-amf','Accept-Encoding':'gzip, deflate','User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/51.0.2704.79 Safari/537.36 Edge/14.14393','Cookie':'hul_aeon=HtxEVg/6n5indqVB71vVF3DB1VSwqPE=; PDS_HANDLE=96201721849156754967906852281455; _ga=GA1.2.115038926.1457945531; _gid=GA1.2.1577009755.1497801061; PRIMO_RT=; hulaccess=1.3|76.119.237.142|20170618214024EDT|pin|91303817|harvard|FAS-100688|OFFI|1060|hul-prod; user_OpenURL="http://sfx.hul.harvard.edu:80/sfx_local/"; ezproxyezpprod1=o5Wlkd5XyDRGw5y; JSESSIONID=560C3C5A3E7B641B99B0EDA0B21A3D20'}
            #r = requests.get(fullurl,headers=header) 
            r = read(fullurl, header)
            #print(r)
            contents = contents + r + '【'.decode('utf-8')+str(page)+'】'.decode('utf-8')
        #print(contents)
        contents_list.append([row[1],row[2],row[3],row[4],row[5],row[6],row[7],row[8],row[9],contents])
        #print(contents_list)
        #row.append(contents)
        #count = count+1
        #print(list_request[1][10].encode('utf-8'))
        if list_request_count == len(list_request):
            list2csv(contents_list, email_addr+timestampe_last+'.csv')
            mail_to_harvard_user(email_addr, email_addr+timestampe_last+'.csv')
            #【记录已完成行数】
            date = unicode(datetime.now())
            with open('working_log.txt','ab') as log:
                w = csv.writer(log)
                w.writerow([date,pointer])   
            contents_list = []
                
            
            
    


