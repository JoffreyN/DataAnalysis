import os,xlwt,sys,time
path=os.getcwd()+'\\'
filename=input('输入文件名(脚本所在文件夹)：')

'''
v0.9
初步编写完成
'''
def jindutiao(i,count,word='处理数据'):
    m='-'*int((i+1)/(count)*50)
    sys.stdout.write('正在{}：'.format(word)+'|'+m+'>'+' '*(50-int((i+1)/(count)*50))+'|'+str(int((i+1)/(count)*100))+"%\r")
    sys.stdout.flush()
    #time.sleep(0.01)

def getallLines(path):
    print('\n正在读取文件...')
    with open(path,'r',encoding='utf-8',errors='ignore') as f:
        allLines=f.readlines()
    print('\n')
    return allLines

def getallList(allLines):
    strLines=''.join(allLines)
    allList=strLines.split('='*54)
    count=len(allList)
    #print(count)
    for i in range(count):
        if len(allList[i])==36:
            allList[i]=''
        jindutiao(i,count,'生成列表')
    print('\n')
    print('正在准备进行下一步处理...\n')
    while '' in allList:
        allList.remove('')
    while '\n\n\n\n' in allList:
        allList.remove('\n\n\n\n')
    while '\n\n\n' in allList:
        allList.remove('\n\n\n')
    #print(allList)
    return allList

def onestrTodic(onestr):
    dic={}
    deln=onestr.strip('\n')
    onelist=deln.split('\n')
    onelist[0]=onelist[0].replace(' ',':',1)
    while '' in onelist:
        onelist.remove('')
    if not onelist[-1].startswith('Accept:'):
        if 'Accept: */*' in onelist:
            onelist[onelist.index('Accept: */*')+1]='\n'.join(onelist[onelist.index('Accept: */*')+1:])
            del onelist[onelist.index('Accept: */*')+2:]
            onelist[-1]='PostData:'+onelist[-1]
    for i in onelist:
        l=i.split(':',1)
        if len(l)==1:
            l.append(',')      
        dic[l[0]]=l[1]
    return dic

def getallDic(allList):
    allDic=[]
    for i in allList:
        allDic.append(onestrTodic(i))
        jindutiao(allList.index(i),len(allList),'生成字典')
    #print(allDic)
    return allDic

def saveExcel(allDic):
    #line=len(allDic)
    row=['GET','POST','PostData','Host','Accept', 'Accept-Charset', 'Accept-Encoding',
         'Accept-Language', 'Accept-Ranges', 'Authorization', 'Cache-Control',
         'Connection', 'Cookie', 'Content-Length', 'Content-Type', 'Date',
         'Expect', 'From', 'If-Match', 'If-Modified-Since', 'If-None-Match',
         'If-Range', 'If-Unmodified-Since', 'Max-Forwards', 'Pragma',
         'Proxy-Authorization', 'Range', 'Referer', 'TE', 'Upgrade',
         'User-Agent', 'Via', 'Warning','Acunetix-Aspect','Acunetix-Aspect-Password',
         'Acunetix-Aspect-Queries','otherKeys']
    excel=xlwt.Workbook()
    sheet=excel.add_sheet('sheet',cell_overwrite_ok=True)
    z=0
    for i in range(len(allDic)):        
        for j in range(len(row)):
            z+=1
            if i==0:
                sheet.write(0,j,row[j])
                if row[j] in allDic[i]:
                    sheet.write(i+1,j,allDic[i][row[j]].strip())
                else:
                    sheet.write(i+1,j,'')
            elif row[j] in allDic[i]:
                sheet.write(i+1,j,allDic[i][row[j]].strip())
                #print(allDic[i][row[j]])
            else:
                sheet.write(i+1,j,'')
            jindutiao(z,len(allDic)*len(row),'处理数据并写入文件')
        
        for k in list(allDic[i].keys()):
            if k not in row:
                sheet.write(i+1,row.index(row[-1]),k+':\t')
                #print(k+':'+allDic[i][k].strip())
                sheet.write(i+1,row.index(row[-1])+1,allDic[i][k].strip())
            else:
                continue
    print('\n')
    excel.save(path+filename.split('.')[0]+'.xls')
    print('处理完成！文件保存至{}'.format(path+filename.split('.')[0]+'.xls'))

def main():       
    filepath=path+filename
    allLines=getallLines(filepath)
    allList=getallList(allLines)
    allDic=getallDic(allList)
    saveExcel(allDic)
    #print(data)
if __name__=='__main__':
    main()
