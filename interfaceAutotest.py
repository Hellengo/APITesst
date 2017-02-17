#!/usr/bin/env python
# -*- coding:utf-8 -*-

#Todo：接口自动化测试

import json
import random
import time
import re
import logging
import os
import sys
import requests
reload(sys)
sys.setdefaultencoding('utf-8')
import xlrd
import smtplib
from email.mime.text import MIMEText

log_file = os.path.join(os.getcwd(), './log/liveappapi.log')
log_format = '[%(asctime)s] [%(levelname)s] %(message)s'
logging.basicConfig(format=log_format, filename=log_file, filemode='w', level=logging.DEBUG)
console = logging.StreamHandler()
console.setLevel(logging.DEBUG)
formatter = logging.Formatter(log_format)
console.setFormatter(formatter)
logging.getLogger('').addHandler(console)


#获取并执行测试用例
def runTest(testCaseFile):
    testCaseFile = './TestCase/TestCase.xlsx'
    testCaseFile = os.path.join(os.getcwd(), testCaseFile)
    if not os.path.exists(testCaseFile):
        logging.error('测试用例文件不存在！！！')
        sys.exit()
    testCase = xlrd.open_workbook(testCaseFile)
    table = testCase.sheet_by_index(0)
    errorCase = []
    correlationDict = {}
    correlationDict['${session}'] = None
    for i in range(1, table.nrows):
        correlationDict['${randomEmail}'] = ''.join(random.sample('abcdefghijklmnopqrstuvwxyz', 6)) + '@automation.test'
        correlationDict['${randomTel}'] = '186' + str(random.randint(10000000, 99999999))
        correlationDict['${timestamp}'] = int(time.time())
        if table.cell(i, 10).value.replace('\n', '').replace('\r', '') != 'Yes':
            continue
        num = str(int(table.cell(i, 0).value)).replace('\n', '').replace('\r', '')
        api_purpose = table.cell(i, 1).value.replace('\n', '').replace('\r', '')
        api_host = table.cell(i, 2).value.replace('\n', '').replace('\r', '')
        request_url = table.cell(i, 3).value.replace('\n', '').replace('\r', '')
        request_method = table.cell(i, 4).value.replace('\n', '').replace('\r', '')
        request_data_type = table.cell(i, 5).value.replace('\n', '').replace('\r', '')
        request_data = table.cell(i, 6).value.replace('\n', '').replace('\r', '')
        encryption = table.cell(i, 7).value.replace('\n', '').replace('\r', '')
        check_point = table.cell(i, 8).value
        correlation = table.cell(i, 9).value.replace('\n', '').replace('\r', '').split(';')
        for key in correlationDict:
            if request_url.find(key) > 0:
                request_url = request_url.replace(key, str(correlationDict[key]))
        if request_data_type == 'Form':
            dataFile = request_data
            if os.path.exists(dataFile):
                fopen = open(dataFile, encoding='utf-8')
                request_data = fopen.readline()
                print request_data
                fopen.close()
            for keyword in correlationDict:
                if request_data.find(keyword) > 0:
                    request_data = request_data.replace(keyword, str(correlationDict[keyword]))
                continue
        elif request_data_type == 'Data':
            dataFile = request_data
            if os.path.exists(dataFile):
                fopen = open(dataFile, encoding='utf-8')
                request_data = fopen.readline()
                fopen.close()
            for keyword in correlationDict:
                if request_data.find(keyword) > 0:
                    request_data = request_data.replace(keyword, str(correlationDict[keyword]))
            request_data = request_data.encode('utf-8')

        status, resp = interfaceTest(num, api_purpose, api_host, request_url, request_data, check_point, request_method, request_data_type, correlationDict['${session}'])
        if status != 200:
            errorCase.append((num + ' ' + api_purpose, str(status), request_url, resp))
            continue
        for j in range(len(correlation)):
            param = correlation[j].split('=')
            if len(param) == 2:
                if param[1] == '' or not re.search(r'^\[', param[1]) or not re.search(r'\]$', param[1]):
                    logging.error(num + ' ' + api_purpose + ' 关联参数设置有误，请检查[Correlation]字段参数格式是否正确！！！')
                    continue
                value = resp
                for key in param[1][1:-1].split(']['):
                    try:
                        temp = value[int(key)]
                    except:
                        try:
                            temp = value[key]
                        except:
                            break
                    value = temp
                correlationDict[param[0]] = value
                #print value
    return errorCase


# 接口测试
def interfaceTest(num, api_purpose, api_host, request_url, request_data, check_point, request_method, request_data_type, session):
    headers = {'Content-Type': 'application/x-www-form-urlencoded',
               'X-Requested-With': 'XMLHttpRequest',
               'Connection': 'keep-alive',
               #'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/43.0.2357.134 Safari/537.36'}
               'User-Agent': 'iOS/10.000000;iPhone 5c(A1507)',
                'Cookie': ''}

    if session is not None:
        headers['Cookie'] = 'sessionid=' + session

    if request_method == 'POST':
        r = requests.post(api_host+request_url, data=request_data, headers=headers, verify = './ssl/＊.crt')##可设置成false
    elif request_method == 'GET':
        r = requests.get(api_host+request_url, verify = './ssl/＊.crt')
    else:
        logging.error(num + ' ' + api_purpose + ' HTTPS请求方法错误，请确认[Request Method]字段是否正确！！！')
        return 400, request_method

    status = r.status_code
    resp = r.text

    if status == 200:
        resp = resp.decode('unicode_escape')
        if re.search(check_point, str(resp)):
            logging.info(num + ' ' + api_purpose + ' 成功, ' + str(status) + ', ' + str(resp))
            try:
                respdata = json.loads(resp.replace('\n', ''))
            except Exception as e:
                print(e)
                respdata = {'error': str(e)}
            return status, respdata
        else:
            logging.error(num + ' ' + api_purpose + ' 失败！！！, [ ' + str(status) + ' ], ' + str(resp))
            return 2001, resp
    else:
        logging.error(num + ' ' + api_purpose + ' 失败！！！, [ ' + str(status) + ' ], ' + str(resp))
        return status, resp.decode('raw_unicode_escape')
        
# 发送通知邮件
def sendMail(text):
    sender = ''
    receiver = []
    mailToCc = []
    subject = '[AutomantionTest]接口自动化测试报告通知'
    smtpserver = ''
    username = ''
    password = ''

    msg = MIMEText(text, 'html', 'utf-8')
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = ';'.join(receiver)
    msg['Cc'] = ';'.join(mailToCc)
    smtp = smtplib.SMTP()
    smtp.connect(smtpserver, 587)
    smtp.starttls()
    smtp.login(username, password)
    smtp.sendmail(sender, receiver + mailToCc, msg.as_string())
    smtp.quit()
    
# 用云片发送短信
class SmsLogics(object):
    APP_KEY = '＊＊＊'
    SENDURL = '＊＊＊'

    @classmethod
    def send_sms_to_user(cls, mobilephone, msg):
        msg = u'线上接口监控：' + msg
        params = {'apikey': cls.APP_KEY,
                  'mobile': mobilephone,
                  'text': msg}
        res = requests.post(cls.SENDURL, data=params)
        t = res.json()
        if t.get('code', 500):
            emsg = t.get('msg') or t.get('detail', 'error')
            e = Exception(emsg)
            return emsg or 'error'
        return None

# 生成报告
def main():
    errorTest = runTest('./TestCase/TestCase.xlsx')
    if len(errorTest) > 0:
        html = '<html><body><h3 style="text-align:center;">接口自动化扫描,共有 ' + str(len(errorTest)) + ' 个异常接口 </h3>' + '</p><table><tr><th style="width:100px;">接口</th><th style="width:50px;">状态</th><th style="width:100px;">接口地址</th><th>接口返回值</th></tr><hr/>'
        for test in errorTest:
            html = html + '<tr><td>' + test[0] + '</td><td>' + test[1] + '</td><td>' + test[2] + '</td><td>' + test[3]  + '</td></tr>'
        html = html + '</table></body></html>'
        f = open("Report/Report.html", 'w')
        f.truncate()
        f.write(html)
        f.close()
        sendMail(html)
    elif len(errorTest) > 3:
        SmsLogics.send_sms_to_user('', u'线上＊＊接口挂了')

if __name__ == '__main__':
    main()
