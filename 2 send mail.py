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
import string

#第三方SMTP服务
##mail_host = "smtp.exmail.qq.com"           # 设置服务器

mail_host = "smtp.qq.com"           # 设置服务器
mail_user = "569**@qq.com"        # 用户名
mail_pwd  = "*****"




def SendMail(to_addr,subject,content,attfile):
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = mail_user
    mail_to = to_addr.split(',')
    msg['To'] = ','.join(mail_to)
    msg.attach(MIMEText(content,'plain','gb2312'))
    #正文

    part = MIMEText(open('附件/'+attfile, 'rb').read(), 'base64', 'gb2312')
    part["Content-Type"] = 'application/octet-stream'
    part.add_header('Content-Disposition', 'attachment', filename=('gbk','',attfile))
    msg.attach(part)
    #附件

    try:
        smtpObj = smtplib.SMTP_SSL(mail_host, 465)
        smtpObj.login(mail_user, mail_pwd)
        smtpObj.sendmail(mail_user,mail_to, msg.as_string())
        smtpObj.quit()
        today = datetime.date.today().strftime("%Y-%m-%d")
        nowtime = time.strftime("%Y-%m-%d %H:%M:%S")
        log = open('发送结果/发送报告 %s.txt'%today,'a')
        log.write('%s\t发送成功\t%s\n'%(attfile,nowtime))
        log.close()
    except Exception as err:
        nowtime = time.strftime("%Y-%m-%d %H:%M:%S")
        today = datetime.date.today().strftime("%Y-%m-%d")
        log = open('发送结果/发送报告 %s.txt'%today,'a')
        log.write('%s\t发送失败\t%s\n'%(attfile,nowtime))
        log.close()
        err_log = open('错误日志.txt','a')
        err_log.write('%s\t发送失败\t%s\t%s\n'%(attfile,err,nowtime))
        err_log.close()

        #发送
    
    
    



if __name__ == "__main__":
    send_file = open('发送目录.txt')
    for line in send_file:
        try:
            if '文件' in line:continue
            FileName,GroupName,RecName,ToAddr,Subject,Title,Content,Month,Dist = line.split('\t')
            content2 = Title+'\n    '+Content.replace('MM',Month.strip()).replace('DD',Dist.strip())
            SendMail(ToAddr,Subject,content2,FileName)
        except Exception as err2:
            nowtime = time.strftime("%Y-%m-%d %H:%M:%S")
            err_log = open('错误日志.txt','a')
            err_log.write('%s\t发送失败\t%s\t%s\n'%(FileName,err2,nowtime))
            err_log.close()
            continue
    send_file.close()
    
                
        
        

    
