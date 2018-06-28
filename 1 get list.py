import os
import os.path
import time,datetime

def month_minus(n):##输入差值输出n个月前的yyyy-mm格式日期字符串
    d = datetime.datetime.now()
    if d.month - n > 0:
        results = datetime.datetime(d.year, d.month-n, 1, 0, 0, 0).strftime('%m')
    else:
        results = datetime.datetime(d.year-1, d.month-n+12, 1, 0, 0, 0).strftime('%m')
    return results





#-------读取配置文件------------

file_dic={}
with open('lib/config.txt') as config:
    for i in config:
        if '@' not in i:continue
        l = i.strip().split('\t')
        key_word = l[0]
        if key_word in file_dic:continue
        file_dic[key_word]={}
        file_dic[key_word]['to_list'] = l[1]
        file_dic[key_word]['cc_list'] = l[2]
        file_dic[key_word]['title'] = l[3]
        file_dic[key_word]['content'] = l[4]
        file_dic[key_word]['dist'] = l[5]

#-------写待发送文件------------
now = datetime.datetime.now().strftime('%Y-%m-%d %H_%M_%S')
w = open('待发送 %s.txt'%now,'a')
rootdir = "./附件"
for parent,dirnames,filenames in os.walk(rootdir):
    for filename in filenames:
        groupName = '-'
        to_list = '-'
        cc_list = '-'
        groupName = '-'
        subject = '-'
        title = '-'
        content = '-'
        month = '-'
        dist = '-'
        for key_word in file_dic:
            if key_word in filename:
                groupName = file_dic[key_word]['dist']
                to_list = file_dic[key_word]['to_list']
                cc_list = file_dic[key_word]['cc_list']
                subject = filename.split('.')[0]
                title = file_dic[key_word]['title']
                content = file_dic[key_word]['content']
                month = month_minus(1)
                dist = file_dic[key_word]['dist']
        w.write(('%s\t'*8+'%s\n')%(filename,groupName,to_list,cc_list,subject,title,content,month,dist))
            
w.close()            
            
        



