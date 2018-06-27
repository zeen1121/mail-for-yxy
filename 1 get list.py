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


rootdir = "./附件"


group_dic={}
group_file = open('lib/group.txt')
for line in group_file:
    if '@' not in line:continue
    gname,to_list,cc_list = line.strip().split('\t')
    group_dic[gname]=[to_list,cc_list]

##print(group_dic)
now = datetime.datetime.now().strftime('%Y-%m-%d %H_%M_%S')
w = open('待发送 %s.txt'%now,'a')

for parent,dirnames,filenames in os.walk(rootdir):
    for filename in filenames:
        if 'BEMS华南大区预算执行情况明细表' in filename:
            groupName = 'BEMS华南大区'
            to_list = group_dic[groupName][0]
            cc_list = group_dic[groupName][1]
            subject = filename.split('.')[0]
            title = '张经理你好！'
            content = '这是MM月DD费用预算执行情况明细表，请接收查阅，如有疑问，请及时联系。谢谢！'
            month = month_minus(1)
            dist = '华南大区'
        else:
            to_list = '-'
            cc_list = '-'
            groupName = '-'
            subject = '-'
            title = '-'
            content = '-'
            month = '-'
            dist = '-'
        w.write(('%s\t'*8+'%s\n')%(filename,groupName,to_list,cc_list,subject,title,content,month,dist))
            
w.close()            
         
        
