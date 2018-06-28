# 正式版本 注意发送及收件地址



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
import string,re

os.system('title 批量发送邮件脚本')
#第三方SMTP服务

def GetAddr(AddrString):#输入收件人列表
    mail_to_list = re.findall('<.*?>',AddrString)#匹配邮箱地址列表
    mail_to=[]
    for i in mail_to_list:
        mail_to.append(i.replace('<','').replace('>',''))  #去除<>
    return mail_to#返回邮箱地址列表

def SendMail(to_addr,cc_addr,subject,title,content,attfile):
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg['From'] = '杨晓泳 <11030@ake.com.cn>'
    
    msg['To'] = to_addr
    if '@' in cc_addr:
        msg['cc'] = cc_addr  
        mail_to = GetAddr(to_addr+';'+cc_addr)
    else:
        mail_to = GetAddr(to_addr)

    content_html = '''<div style="position:relative;"><div>&nbsp;</div>
<div>%s</div>
<div>&nbsp; &nbsp; &nbsp; %s</div><div><div style="FONT-FAMILY: Arial Narrow; COLOR: #909090; FONT-SIZE: 12px"><br><br><br>------------------</div>
<div style="FONT-FAMILY: Verdana; COLOR: #000; FONT-SIZE: 14px">
<div>
<div style="LINE-HEIGHT: 35px; MARGIN: 20px 0px 0px; WIDTH: 305px; HEIGHT: 35px" class="logo"><img src="https://exmail.qq.com/cgi-bin/viewfile?type=logo&amp;domain=ake.com.cn"></div>
<div style="MARGIN: 10px 0px 0px" class="c_detail">
<h4 style="LINE-HEIGHT: 28px; MARGIN: 0px; ZOOM: 1; FONT-SIZE: 14px; FONT-WEIGHT: bold" class="name">杨晓泳</h4>
<p style="LINE-HEIGHT: 22px; MARGIN: 0px; COLOR: #a0a0a0" class="position">税务会计</p>
<p style="LINE-HEIGHT: 22px; MARGIN: 0px; COLOR: #a0a0a0" class="department">广东艾科技术股份有限公司/财务部</p>
<p style="LINE-HEIGHT: 22px; MARGIN: 0px; COLOR: #a0a0a0" class="phone"></p>
<p style="LINE-HEIGHT: 22px; MARGIN: 0px; COLOR: #a0a0a0" class="addr"><span onmouseover="QMReadMail.showLocationTip(this)" class="readmail_locationTip" onmouseout="QMReadMail.hideLocationTip(this)" over="0" style="z-index: auto;">广东省佛山市南海区桂城深海路17号A区三号楼三楼</span></p></div></div></div></div>
<div>&nbsp;</div>
<div><tincludetail></tincludetail></div></div>'''%(title,content)
    
    msg.attach(MIMEText(content_html,_subtype='html', _charset='gb2312'))
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
    
##    to_list = to_addr.split(',')  #收件人列表
##    cc_list = cc_addr.split(',')  #抄送人列表

    msg['To'] = to_addr
    if '@' in cc_addr:
        msg['cc'] = cc_addr
    
    print(attfile)
    print('From:'+msg['From'])
    print('To:'+msg['To'])
    if '@' in cc_addr:
        print('Cc:'+msg['cc'])
    print('Subject:'+msg['Subject']+'\n')
    print(content)
    print('\n------------------------------------------------------------------------\n')

    
if __name__ == "__main__":
    show_file = open('发送目录.txt')
    for line in show_file:
        if '@' not in line:continue
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
                content2 =Content.replace('MM',Month.strip()).replace('DD',Dist.strip())
                SendMail(ToAddr,CcAddr,Subject,Title,content2,FileName)
            except Exception as err2:
                nowtime = time.strftime("%Y-%m-%d %H:%M:%S")
                err_log = open('错误日志.txt','a')
                err_log.write('%s\t发送失败\t%s\t%s\n'%(FileName,err2,nowtime))
                err_log.close()
                continue
        send_file.close()
    

        input('\n\n发送完成，确定键退出')
    
                
        
        

    
