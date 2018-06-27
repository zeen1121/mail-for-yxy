# 最终版本 注意发送及收件地址

import pymysql #Python3的mysql模块，Python2 是mysqldb
import os
import datetime #定时发送，以及日期
import shutil #文件操作
import smtplib #邮件模块
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
import time
from email.mime.application import MIMEApplication
##import xlwt #excel写入
import string,re


os.system('title 批量发送邮件脚本')
#第三方SMTP服务
##mail_host = "smtp.exmail.qq.com"           # 设置服务器
##mail_to  = ['569051642@qq.com']     #接收邮件列表,是list,不是字符串
##mail_host = "smtp.qq.com"           # 设置服务器




def GetAddr(AddrString):#输入收件人列表
    mail_to_list = re.findall('<.*?>',AddrString)#匹配邮箱地址列表
    mail_to=[]
    for i in mail_to_list:
        mail_to.append(i.replace('<','').replace('>',''))  #去除<>
    return mail_to#返回邮箱地址列表
        



def SendMail(to_addr,cc_addr,subject,content,attfile):
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = '杨晓泳 <11030@ake.com.cn>'
    
    to_list = to_addr.split(',')  #收件人列表
    cc_list = cc_addr.split(',')  #抄送人列表

    msg['To'] = to_addr
    msg['cc'] = cc_addr
    

    mail_to = GetAddr(to_addr+';'+cc_addr)
    
##    mail_to = ['569051642@qq.com']  #测试用
    
    
    msg.attach(MIMEText(content,'plain','gb2312'))
    #正文

    part = MIMEText(open('附件/'+attfile, 'rb').read(), 'base64', 'gb2312')
    part["Content-Type"] = 'application/octet-stream'
    part.add_header('Content-Disposition', 'attachment', filename=('gbk','',attfile))
    msg.attach(part)
    #附件

    try:
        with open('lib/sender.txt') as ss:
            for line in ss:
                mail_user,mail_pwd = line.strip().split('\t')
                break
##        mail_host = "smtp.qq.com"                    # 测试用
        mail_host = "smtp.exmail.qq.com"           # 实际版
        smtpObj = smtplib.SMTP_SSL(mail_host, 465)
        smtpObj.login(mail_user, mail_pwd)
        smtpObj.sendmail(mail_user,mail_to, msg.as_string())
        smtpObj.quit()
        today = datetime.date.today().strftime("%Y-%m-%d")
        nowtime = time.strftime("%Y-%m-%d %H:%M:%S")
        log = open('发送结果/发送报告 %s.txt'%today,'a')
        print('\n%s\t发送成功\t%s\n'%(attfile,nowtime))
        log.write('%s\t发送成功\t%s\n'%(attfile,nowtime))
        log.close()
    except Exception as err:
        nowtime = time.strftime("%Y-%m-%d %H:%M:%S")
        today = datetime.date.today().strftime("%Y-%m-%d")
        log = open('发送结果/发送报告 %s.txt'%today,'a')
        log.write('%s\t发送失败\t%s\n'%(attfile,nowtime))
        log.close()
        err_log = open('错误日志.txt','a')
        print('\n%s\t发送失败\t%s\t%s\n'%(attfile,err,nowtime))
        err_log.write('%s\t发送失败\t%s\t%s\n'%(attfile,err,nowtime))
        err_log.close()



        #发送


def ShowMail(to_addr,cc_addr,subject,content,attfile):
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = '杨晓泳 <11030@ake.com.cn>'
    
    to_list = to_addr.split(',')  #收件人列表
    cc_list = cc_addr.split(',')  #抄送人列表

    msg['To'] = to_addr
    msg['cc'] = cc_addr
    

    
    #附件
    print(attfile)
    print('From:'+msg['From'])
    print('To:'+msg['To'])
    print('Cc:'+msg['cc'])
    print('Subject:'+msg['Subject']+'\n')
    print(content)
    print('\n')
    
    


if __name__ == "__main__":
    show_file = open('发送目录.txt')
    for line in show_file:
        if '@' not in line and '-' not in line:continue
        FileName,GroupName,ToAddr,CcAddr,Subject,Title,Content,Month,Dist = line.split('\t')
        content2 = Title+'\n    '+Content.replace('MM',Month.strip()).replace('DD',Dist.strip())
        ShowMail(ToAddr,CcAddr,Subject,content2,FileName)
    show_file.close()
    
    check = input('是否发送（y/n）?')
    if check == 'y' or check == 'Y':
        send_file = open('发送目录.txt')
        for line in send_file:
            try:
                if '@' not in line and '-' not in line:continue
                FileName,GroupName,ToAddr,CcAddr,Subject,Title,Content,Month,Dist = line.split('\t')
                content2 = Title+'\n    '+Content.replace('MM',Month.strip()).replace('DD',Dist.strip())
                SendMail(ToAddr,CcAddr,Subject,content2,FileName)
            except Exception as err2:
                nowtime = time.strftime("%Y-%m-%d %H:%M:%S")
                err_log = open('错误日志.txt','a')
                err_log.write('%s\t发送失败\t%s\t%s\n'%(FileName,err2,nowtime))
                err_log.close()
                continue
        send_file.close()
    

        input('\n\n发送完成，确定键退出')
    
                
        
        

    
