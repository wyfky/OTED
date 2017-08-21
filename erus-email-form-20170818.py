
# coding: utf-8

# In[ ]:


from __future__ import print_function
import httplib2
import os

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

import argparse

#Using url, unique id, amf and xml library
import urllib2
import uuid
import pyamf
from pyamf import remoting
from pyamf.flex import messaging
import xml.etree.cElementTree as ET
import json
import csv
import time
import re
from datetime import datetime
from collections import deque
import sys
    
#以下sendEmail函数使用了Jason Hu的Python SMTP 发送带附件电子邮件代码

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.header import Header
import smtplib


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
        msgRoot['From'] = '%s<%s>' % (Header('爱如生数据库检索结果', 'utf-8'), strFrom)
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

def google_sheet(pointer):
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

    spreadsheetId = '******'
    rangeName = 'Form responses 1!B'+str(pointer)+':D100000'
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    else:
        return(values)
    '''    
        print('Name, Major:')
        for row in values:
            # Print columns A and E, which correspond to indices 0 and 4.
            print('%s, %s' % (row[0], row[1]))
    '''

    
#Function of tansfer list to csv
def list2csv(list,file):
    with open(file,'ab') as f:
        for i in range(0,len(list)):
            w=csv.writer(f)
            #For dealing with Chinese char in python2, encode every list element to utf-8, and then write to a line of csv 
            w.writerow([list[i][0].encode("utf-8"),list[i][1].encode("utf-8"),list[i][2].encode("utf-8"),list[i][3].encode("utf-8"),list[i][4].encode("utf-8"),list[i][5].encode("utf-8"),list[i][6].encode("utf-8"),list[i][7].encode("utf-8"),list[i][8].encode("utf-8")])
    return
# Function of communcation with erus database using the framework of pyamf           
def commun(keyword, BeginNo, database_ID): 
    # 构建flex.messaging.messages.RemotingMessage信息
    msg= messaging.RemotingMessage(messageId=str(uuid.uuid1()).upper(),clometOd=None,operation='sedCmd',destination='callBackMethod',clientId=str(uuid.uuid1()).upper(),timeTolive=0,timestamp=0)
    msg.body = ['<Code>7300</Code><TAG>0</TAG><LibID>'+database_ID+'</LibID><KeyWord>'+keyword+'</KeyWord><FTag></FTag><Class></Class><Name></Name><Author></Author><Years></Years><SimpTrad>TRUE</SimpTrad><AlloNorm>TRUE</AlloNorm><PageRecNum>24</PageRecNum><RecBeginNo>'+BeginNo+'</RecBeginNo>', 1]
    msg.headers['DSEndpoint']='my-amf'
    msg.headers['DSId']=str(uuid.uuid1()).upper()

    # 按AMF协议编码
    req = remoting.Request('null', body=(msg,))
    env = remoting.Envelope(amfVersion=pyamf.AMF3)
    env.bodies = [('/1',req)]
    data = bytes(remoting.encode(env).read())

    # 提交请求
    url = 'http://server.wenzibase.com.ezp-prod1.hul.harvard.edu/messagebroker/amf'
    headers = {'Accept':'*/*','Accept-Language':'zh-CN','Referer':'http://server.wenzibase.com.ezp-prod1.hul.harvard.edu/book/fz.swf','x-flash-version':'25,0,0,171','Content-Type':'application/x-amf','Accept-Encoding':'gzip, deflate','User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/51.0.2704.79 Safari/537.36 Edge/14.14393','Cookie':'*******'}
    req = urllib2.Request(url,data,headers)

    # 解析返回数据
    opener = urllib2.build_opener()

    # 解码AMF协议返回的数据
    resp = remoting.decode(opener.open(req).read())
    resp_body = resp.bodies[0][1].body.body
    # 如果爱如生数据库出错，则利用递归重新读取该页数据
    if resp_body==None: 
        print("y")
        return commun(keyword, BeginNo, database_ID)
    
    #type(resp_body)
    #print(resp_body)
    resp_body_string = resp_body.encode("utf-8")
    #print(resp_body_string)
    return resp_body_string
#获取日志文件最后一行内容
def get_last_row(csv_filename):
    with open(csv_filename, 'r') as f:
        try:
            lastrow = deque(csv.reader(f), 1)[0]
        except IndexError:  # empty file
            lastrow = None
        return lastrow
#写日志文件
def write_log(csv_log, pointer):
    date = unicode(datetime.now())
    with open(csv_log,'ab') as log:
        w = csv.writer(log)
        w.writerow([date,pointer])   
#发送邮件至Harvard用户邮箱
def mail_to_harvard_user(toAdd, att_file):
    authInfo = {}
    authInfo['server'] = '******'
    authInfo['user'] = '******'
    authInfo['password'] = '*******'
    fromAdd = '******'
    #toAdd = [row[2]]
    subject = 'Email Subject'
    plainText = 'Email Content'
    #att_file = row[0]+' in '+row[1]+'.csv'
    print(att_file)
    sendEmail(authInfo, fromAdd, toAdd, subject, plainText, att_file)
 


#以下为主程序
def main(word, database_name):
    keyword = word
    database_dict = {"古籍库" : "***", "方志库": "***"}
    database_ID = database_dict[database_name.encode('utf-8')]
    #获取爱如生数据库返回由amf转化而来的字符串
    str_amf=''
    #读取第一页，从中得到记录总数
    str_amf=commun(keyword,"1",database_ID)
    root=ET.fromstring(str_amf)
    #读取返回记录总数TotalNum
    TotalNum = root[5].text
    print(TotalNum)
    #定义存放记录的list
    record_list=[]
    for num in range(0,(int(TotalNum)/24)+1):
        #if num<>286:
            # i为开始记录号
            i = num*24+1
            print(num)
            str_amf=commun(keyword,str(i),database_ID)
            #过滤XML中的非法字符
            str_amf_filter=re.sub("[\x00-\x08\x0b-\x0c\x0e-\x1f]+","",str_amf)
            root=ET.fromstring(str_amf_filter)
            #读取Item内容并存入record_list列表
            for elem in root.iterfind('List/Item'):
                ID = elem.findtext('ID')
                PageID = elem.findtext('PageID')
                VolNo = elem.findtext('VolNo')
                Name = elem.findtext('Name')
                Author = elem.findtext('Author')
                Years = elem.findtext('Years')
                VolNum = elem.findtext('VolNum')
                TextVerInfo = elem.findtext('TextVerInfo')
                Exam = elem.findtext('Exam')
                record={}
                record = [ID, PageID, VolNo, Name, Author, Years, VolNum, TextVerInfo, Exam]
                #record = [ID, PageID, Name, Author, Years, VolNum, Exam]
                record_list.append(record)
            time.sleep(0.5)
        #print(record_list)
    list2csv(record_list,word+' in '+database_name+'.csv')
    record_list=[]
if __name__ == '__main__' :
    keyword_list = []
    lastline = get_last_row('working_log_search.txt')
    pointer = lastline[1]
    print(pointer)
    begin_record_pointer = int(pointer) + 1
    keyword_list = google_sheet(begin_record_pointer)
    if keyword_list == None:
        print('没有新的请求可以处理')
        sys.exit()
    for row in keyword_list:
        #判断是否已经有相应的搜索结果缓存
        attfile = row[0]+' in '+row[1]+'.csv'
        if os.path.isfile(attfile):
            mail_to_harvard_user(row[2], attfile)
            pointer = int(pointer) + 1
            write_log('working_log_search.txt',pointer)
            continue
        main(row[0],row[1])
        mail_to_harvard_user(row[2], attfile)
        pointer = int(pointer) + 1
        write_log('working_log_search.txt',pointer)
        

            
        



# In[ ]:



