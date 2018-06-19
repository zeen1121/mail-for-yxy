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
##mail_user = "yangxy@ake.com.cn"        # 用户名
##mail_pwd  = "DZ8RhnZLmSef3bRf"      # 口令,QQ邮箱是输入授权码，在qq邮箱设置 里用验证过的手机发送短信获得，不含空格
##mail_to  = ['569051642@qq.com']     #接收邮件列表,是list,不是字符串
mail_host = "smtp.qq.com"           # 设置服务器
mail_user = "569051642@qq.com"        # 用户名
mail_pwd  = "xsqhrqoogxnmbcgi"




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
    
    
    
##attfile='附件/安锋元宝原始数据.xlsx'
##basename= os.path.basename(attfile) 
##print(basename)
###邮件内容
##msg = MIMEMultipart()      # 邮件正文
##msg['Subject'] = "测试邮件"     # 邮件标题
##msg['From'] = mail_user        # 发件人
##msg['To'] = ','.join(mail_to)         # 收件人，必须是一个字符串
##msg.attach(MIMEText('张经理你好！\n    这是3月华南大区费用预算执行情况明细表，请接收查阅，如有疑问，请及时联系。谢谢！', 'plain', 'gb2312'))
##
##
##part =  MIMEText(open(attfile, 'rb').read(), 'base64', 'gb2312')
##part["Content-Type"] = 'application/octet-stream' 
##part.add_header('Content-Disposition', 'attachment', filename=('gbk','',attfile.split('/')[1]))
##
##msg.attach(part)


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
    
                
        
        

    
