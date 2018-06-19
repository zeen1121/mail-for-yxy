import os
import os.path
rootdir = "./附件"

file_object = open('train_list.txt','w')

def GetGroup():
    

for parent,dirnames,filenames in os.walk(rootdir):
    for filename in filenames:
        print(filename)
        
        file_object.write(filename+ '\n')
file_object.close()
