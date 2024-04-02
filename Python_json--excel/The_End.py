import json
import tablib
import time

# json.text文件的格式： [{"a":1},{"a":2},{"a":3},{"a":4},{"a":5}]

# 获取ｊｓｏｎ数据
with open('sum.json', 'r') as f:
    rows = json.load(f)

# 将json中的key作为header, 也可以自定义header（列名）
header=tuple([ i for i in rows[0].keys()])

data = []
# 循环里面的字典，将value作为数据写入进去
for row in rows:
    body = []
    for v in row.values():
        body.append(v)
    data.append(tuple(body))

data = tablib.Dataset(*data,headers=header)

open('haha1.xls', 'wb').write(data.xls)



import pandas as pd

df = pd.read_excel(r'haha1.xls')#默认读取工作簿中第一个工作表，默认第一行为表头
data=df.values#获取整个工作表数据
with open('end.json','w')as f:
    f.write('[')


    for i in range(len(data)-1):
        for j in range(len(data[i][0])):
            if data[i][0][j]!= '[' and data[i][0][j]!= ']':
                f.write((data[i][0][j]))
        f.write(',')
        f.write('\n')
    for i in range(len(data[(len(data))-1][0])):
        if data[(len(data))-1][0][i] != '[' and data[(len(data))-1][0][(len(data))] != ']':
            f.write((data[(len(data))-1][0][i]))






list1=[]
a = json.load(open("end.json", encoding='utf8'))
for i in range(0,len(a)-1,2):
    x1=a[i]['name']
    y1=a[i]['value']
    x2=a[i+1]['name']
    y2=a[i+1]['value']
    dic={}
    dic[x1]=y1
    dic[x2]=y2
    list1.append(dic)
with open('end.json','w')as f:
    for i in range(len(str(list1))):
        if str(list1)[i]=="'":
            f.write('"')
            continue
        f.write(str(list1)[i])






# 获取ｊｓｏｎ数据
with open('end.json', 'r') as f:
    rows = json.load(f)

# 将json中的key作为header, 也可以自定义header（列名）
header=tuple([ i for i in rows[0].keys()])

data = []
# 循环里面的字典，将value作为数据写入进去
for row in rows:
    body = []
    for v in row.values():
        body.append(v)
    data.append(tuple(body))

data = tablib.Dataset(*data,headers=header)
t=time.localtime()
year=t.tm_year
mon=t.tm_mon
day=t.tm_mday
hour=t.tm_hour
min=t.tm_min
sec=t.tm_sec
open(f'{year};{mon};{day};{hour};{min};{sec}.xls', 'wb').write(data.xls)