
#!/usr/bin/env python
# coding: utf-8



import requests
import time
from datetime import datetime,timezone
import pandas as pd
import hashlib



from datetime import datetime
import tzlocal  # $ pip install tzlocal

unix_timestamp = time.time()   #加密时间
local_timezone = tzlocal.get_localzone() # get pytz timezone
local_time = datetime.fromtimestamp(unix_timestamp, local_timezone)
current_time = local_time.strftime("%a %b %d %H:%M:%S %Z %Y")




#MD5加密
a = {'appId':'tutu'}
b = { 'appSecretKey': 'XXXXXXXXX'} 
c = {'timestamp':current_time }  
d = {'param':'{}'}

a1 = '='.join(x + '=' + y for x, y in a.items())
b1 = '='.join(x + '=' + y for x, y in b.items())
c1 = '='.join(x + '=' + y for x, y in c.items())
d1 = '='.join(x + '=' + y for x, y in d.items())
password = a1+ '&'+ b1 + '&' + d1 + '&' + c1+'&'
md5 = hashlib.md5(password.encode('utf-8')).hexdigest()


#链接接口
data = { 
      "param": "{}",
    "appId": "tutu",
    "timestamp": str((round(unix_timestamp*1000))),
    "sign": md5
}
bdtodaydata = requests.post('XXXXXXXXXXXX',json = data)
bdtoday = bdtodaydata.json()



#数据处理

today = pd.DataFrame(bdtoday['data'])
business = today.loc[today['displayName'].isin(['王为红','张国才','姜涛07','万鹏01'])]
business =  business.sort_values(by=['primaryTotal','weeklyTotal','monthlyTotal'], ascending=[False,False,False])
monthlytotal = sum(business['monthlyTotal'])
weeklyTotal  = sum(business['weeklyTotal'])



employees = pd.read_excel('C:/Users/liumingxin01/Desktop/上传临时表/20-花名册上传.xlsx')
bd = employees.loc[employees.status=='在职',['name','position','region','area']]
bd = bd.loc[bd.area!='MVP项目',['name','position','region','area']]
bd = bd.loc[bd.name!='白煜10',['name','position','region','area']]


# 个人排名表


#关联架构
bd_today = pd.merge(bd,today,how='left',left_on=['name'],right_on=['displayName'])
bd_today = bd_today.fillna(0) 
#添加总计
bd_today['todaytotal']= bd_today.primaryTotal+bd_today.middleTotal+bd_today.highTotal
#不同架构分区排序
bd_today['rank'] = bd_today.groupby('position')['todaytotal'].rank(ascending = False)
bd_today['ranktotal']=bd_today['todaytotal'].rank(ascending=False, method='first') 
#自定义排序规则
bd_today['position'] = pd.Categorical(bd_today['position'], ["BD", "区域经理", "大区经理"])
bd_today['region'] = pd.Categorical(bd_today['region'], ["第一大区", "第二大区", "第三大区","第四大区","初中大区","高中大区"])
bd_today =  bd_today.sort_values(by=['position','rank','monthlyTotal','region'], ascending=[True,True,False,True])
#保留整数
bd_today['ranktotal'] = bd_today['ranktotal'].astype('int64')
bd_today['todaytotal'] = bd_today['todaytotal'].astype('int64')
#取前三名
rank_first=bd_today[bd_today.ranktotal==1][['name','todaytotal']]
rank_second = bd_today[bd_today.ranktotal==2][['name','todaytotal']]
rank_third = bd_today[bd_today.ranktotal==3][['name','todaytotal']]
#保留排名
rank_today = bd_today.sort_values(by=['ranktotal'], ascending=[True])
rank_today = rank_today.dropna( how='all', subset=['position'])


# 区域&大区表

#区域表
area_today = pd.DataFrame(bd_today.groupby('area').sum())
area_today = area_today.sort_values(by=['todaytotal','monthlyTotal'], ascending=[False,False])
area_today['rank_area'] = area_today['todaytotal'].rank(ascending = False)
area_today = area_today.reset_index(level=0)
#区域取前三名
area_first=area_today[area_today.rank_area==1][['area','todaytotal']]
area_second = area_today[area_today.rank_area==2][['area','todaytotal']]
area_third = area_today[area_today.rank_area==3][['area','todaytotal']]

#大区表
region_today = pd.DataFrame(bd_today.groupby('region').sum())
region_today = region_today.sort_values(by=['todaytotal','monthlyTotal'], ascending=[False,False])
region_today['rank_region'] = region_today['todaytotal'].rank(ascending = False)
region_today = region_today.reset_index(level=0)
rename = ['BD姓名', '架构', '大区', '区域', 'displayName','本月招生','本周招生','今日小学','今日初中','今日高中','今日总计','排名','总排名']
bd_today.columns = rename
bd_today['排名'] = bd_today['排名'].astype('int64')
bd_today['本周招生'] = bd_today['本周招生'].astype('int64')
bd_today['今日小学'] = bd_today['今日小学'].astype('int64')
bd_today['今日初中'] = bd_today['今日初中'].astype('int64')
bd_today['今日高中'] = bd_today['今日高中'].astype('int64')
bd_today['今日总计'] = bd_today['今日总计'].astype('int64')
bd_today['今日总计'] = bd_today['今日总计'].astype('int64')
bd_today = bd_today.dropna( how='all', subset=['架构'])
bd_today_excel = bd_today.loc[:,['排名','架构','BD姓名','大区', '区域','本月招生','本周招生','今日小学','今日初中','今日高中','今日总计']]
bd_today = bd_today.loc[:,['排名','架构','BD姓名','大区', '区域','本周招生','今日小学','今日初中','今日高中','今日总计']]
bd_today_excel_fouth = bd_today_excel.loc[bd_today_excel.大区=='第四大区']


bd_today_excel.to_excel('个人招生明细.xlsx',index = False)
bd_today_excel_fouth.to_excel('第四大区招生明细.xlsx',index=False)



#可视化排名
import pandas as pd
import numpy as np
import dataframe_image as dfi
import plotly_express as px
import plotly.graph_objects as go  # 方法1：go.Table
from plotly.colors import n_colors
import matplotlib
import matplotlib.pyplot as plt
from kaleido.scopes.plotly import PlotlyScope
import plotly.graph_objects as go
import os
from PIL import Image





map_color = {"BD":"red", "区域经理":"green", "大区经理":"#0071C7"}
bd_today["color"] = bd_today["架构"].map(map_color)
cols_to_show = ['排名','架构','BD姓名','大区', '区域','本周招生','今日小学','今日初中','今日高中','今日总计']
text_color = []
n = len(bd_today)
for col in cols_to_show:
    if col!='架构':
        text_color.append(["black"] * n)
    else:
        text_color.append(bd_today["color"].to_list())
        
fill_color = []
n = len(bd_today)
for col in cols_to_show:
    if col!='今日总计':
        fill_color.append(['#EFEFEF']*n)
    else:
        fill_color.append(['#F6E9A4']*n)
        
        
        
        

CELL_HEIGHT = 30
layout = go.Layout( autosize=True, 
                   margin={'l': 0, 'r': 0, 't': 0, 'b': 0},
                  height=CELL_HEIGHT * (len(bd_today.排名) + 1),
                  title='<b>Bold</b> <i>animals</i>')

fig = go.Figure(layout=layout,
                data=[go.Table(
    columnorder = [1,2,3,4,5,6,7,8,9,10], 
    columnwidth = [300,500,500,500,500,500,500,500,500,500],
    header=dict(values=['排名','架构','BD姓名','大区', '区域','本周招生','今日小学','今日初中','今日高中','今日总计'],   # 表头：字典形式
                line_color="#FCFFFF",  # 表头线条颜色
                fill_color="#FB0200",  # 表头填充色
                font=dict(color='white', size=22, family = 'SimHei'),
                height = 35,
                align= 'center' # 文本显示位置 'left', 'center', 'right'
               ), 
        cells=dict(values=[bd_today.排名,
                           bd_today.架构,
                          bd_today.BD姓名,
                          bd_today.大区,
                          bd_today.区域,
                          bd_today.本周招生,
                          bd_today.今日小学,
                          bd_today.今日初中,
                          bd_today.今日高中,
                          bd_today.今日总计],# 第二列元素

               line_color="#FCFFFF",  # 单元格线条颜色
               fill_color=fill_color,  # 单元格填充色
            align='center',
            font=dict(color=text_color, size=19, family = 'SimHei'),
                   height = CELL_HEIGHT
               # 文本显示位置
              ))])
fig.write_image("bd_today.jpg",scale = 2,width=1000)


# 区域&大区可视化



rename1 = [ '区域/大区','本月招生', '本周招生', '今日小学', '今日初中', '今日高中','今日总计','rank','ranktotal','排名']
area_today.columns = rename1
area_today = area_today.loc[:,['排名','区域/大区','本月招生','本周招生','今日小学','今日初中','今日高中','今日总计']]


#区域保留整数
area_today['排名'] = area_today['排名'].astype('int64')
area_today['本月招生'] = area_today['本月招生'].astype('int64')
area_today['本周招生'] = area_today['本周招生'].astype('int64')
area_today['今日小学'] = area_today['今日小学'].astype('int64')
area_today['今日初中'] = area_today['今日初中'].astype('int64')
area_today['今日高中'] = area_today['今日高中'].astype('int64')
area_today['今日总计'] = area_today['今日总计'].astype('int64')
area_today['架构'] = '区域'
col = area_today.pop("架构")
area_today.insert(0, col.name, col)
col = area_today.pop("排名")
area_today.insert(0, col.name, col)



rename2 = [ '区域/大区','本月招生', '本周招生', '今日小学', '今日初中', '今日高中','今日总计','rank','ranktotal','排名']
region_today.columns = rename2
region_today = region_today.loc[:,['排名','区域/大区','本月招生','本周招生','今日小学','今日初中','今日高中','今日总计']]




#大区保留整数
region_today['排名'] = region_today['排名'].astype('int64')
region_today['本月招生'] = region_today['本月招生'].astype('int64')
region_today['本周招生'] = region_today['本周招生'].astype('int64')
region_today['今日小学'] = region_today['今日小学'].astype('int64')
region_today['今日初中'] = region_today['今日初中'].astype('int64')
region_today['今日高中'] = region_today['今日高中'].astype('int64')
region_today['今日总计'] = region_today['今日总计'].astype('int64')
region_today['架构'] = '大区'
col = region_today.pop("架构")
region_today.insert(0, col.name, col)
col = region_today.pop("排名")
region_today.insert(0, col.name, col)



#整合列表
total_today = pd.concat([area_today, region_today], ignore_index=True, sort=False)


map_color1 = {"区域":"red","大区":"#0071C7"}
total_today["color"] = total_today["架构"].map(map_color1)
cols_to_show = ['排名','架构','区域/大区','本月招生','本周招生','今日小学','今日初中','今日高中','今日总计']


#循环填色
text_color1 = []
n = len(total_today)
for col in cols_to_show:
    if col!='架构':
        text_color1.append(["black"] * n)
    else:
        text_color1.append(total_today["color"].to_list())
        
fill_color1 = []
n = len(total_today)
for col in cols_to_show:
    if col!='今日总计':
        fill_color1.append(['#EFEFEF']*n)
    else:
        fill_color1.append(['#F6E9A4']*n)
        
        
        
        

CELL_HEIGHT = 30
layout = go.Layout( autosize=True, 
                   margin={'l': 0, 'r': 0, 't': 0, 'b': 0},
                  height=CELL_HEIGHT * (len(total_today.排名) + 1),
                  title='<b>Bold</b> <i>animals</i>')

fig = go.Figure(layout=layout,
                data=[go.Table(
    columnorder = [1,2,3,4,5,6,7,8,9,10,11], 
    columnwidth = [300,500,500,500,500,500,500,500,500,500,500],
    header=dict(values=['排名','架构','区域/大区','本月招生','本周招生','今日小学','今日初中','今日高中','今日总计'],   # 表头：字典形式
                line_color="#FCFFFF",  # 表头线条颜色
                fill_color="#FB0200",  # 表头填充色
                font=dict(color='white', size=22, family = 'SimHei'),
                height = 35,
                align= 'center' # 文本显示位置 'left', 'center', 'right'
               ), 
        cells=dict(values=[total_today.排名,
                           total_today.架构,
                          total_today['区域/大区'],
                          total_today.本月招生,
                          total_today.本周招生,
                          total_today.今日小学,
                          total_today.今日初中,
                          total_today.今日高中,
                          total_today.今日总计],# 第二列元素

               line_color= "#FCFFFF",  # 单元格线条颜色
               fill_color= fill_color1,  # 单元格填充色
            align='center',
            font=dict(color=text_color1, size=19, family = 'SimHei'),
                   height = CELL_HEIGHT
               # 文本显示位置
              ))])
fig.write_image("total_today.jpg",scale = 2,width=1000)



from requests_toolbelt.multipart.encoder import MultipartEncoder
#图片上传密匙
unix_timestamp = time.time()
a = {'signType':'md5'}
b = { "timestamp": str(int(unix_timestamp))} 
c = {'bizCode':"yuanzishuju" }  
secret = 'XXXXXXXX'
a1 = '='.join(x + '=' + y for x, y in a.items())
b1 = '='.join(x + '=' + y for x, y in b.items())
c1 = '='.join(x + '=' + y for x, y in c.items())
password = c1+ '&'+ a1 + '&' + b1 +secret
p_md5 = hashlib.md5(password.encode('utf-8')).hexdigest()


#图片上传
url = 'XXXXXXXXXXX/upload'
headers = {
    "bizCode":"XXXXX",
    "signType":"md5",
    "timestamp":str(int(unix_timestamp)),
    "sign":p_md5
}
payload={'type': 'image'}
files=[
  ('file',('bd_today.jpg',open('bd_today.jpg','rb'),'image/jpge'))
]
pic = requests.post(url,files = files, headers = headers,data=payload)
picsend = pic.json()
print(pic.text)


# In[131]:


from requests_toolbelt.multipart.encoder import MultipartEncoder
#图片上传密匙

unix_timestamp = time.time()
a = {'signType':'md5'}
b = { "timestamp": str((round(unix_timestamp*1000)))} 
c = {'bizCode':"XXXXXXXX" }  
secret = 'XXXXXXXXXXXXX'
a1 = '='.join(x + '=' + y for x, y in a.items())
b1 = '='.join(x + '=' + y for x, y in b.items())
c1 = '='.join(x + '=' + y for x, y in c.items())
password = c1+ '&'+ a1 + '&' + b1 +secret
md5_r = hashlib.md5(password.encode('utf-8')).hexdigest()
mediacode = picsend['data']['mediaCode']



#机器人接口
url = 'https://internal-ei.baijia.com/ei-serve-management-logic/internal/teamRobot/sign/sendMessage?key=XXXXXXXXXXXXXXX'

headers = {
    "bizCode":"XXXXXXX",
    "signType":"md5",
   "timestamp":str((round(unix_timestamp*1000))),
    "sign":md5_r
}

params = { 
      "toDomains": ["@all"],
    "messageType": 1,
        "image":{
        "mediaCode":mediacode
    }  
}
send = requests.post(url,json = params, headers = headers)


# In[132]:



#标语发送

mark = { 
      "toDomains": ["@all"],
    "messageType": 24,
        "markdown":{
        "body": ('\n'.join([
    "__<font color='black'>今日截止目前0元招生排名:</font>__",
            "__<font color='red'>第一名:</font> "+ rank_today.iloc[0,0]+" "+str(rank_today.iloc[0,10])+"__", 
            "__<font color='red'>第二名:</font> "+ rank_today.iloc[1,0]+" "+str(rank_today.iloc[1,10])+"__",
            "__<font color='red'>第三名:</font> "+ rank_today.iloc[2,0]+" "+str(rank_today.iloc[2,10])+"__",
            "__<font color='red'>第四名:</font> "+ rank_today.iloc[3,0]+" "+str(rank_today.iloc[3,10])+"__",
            "__<font color='red'>第五名:</font> "+ rank_today.iloc[4,0]+" "+str(rank_today.iloc[4,10])+"__",
            "__<font color='red'>第六名:</font> "+ rank_today.iloc[5,0]+" "+str(rank_today.iloc[5,10])+"__",
            "__<font color='red'>第七名:</font> "+ rank_today.iloc[6,0]+" "+str(rank_today.iloc[6,10])+"__", 
            "__<font color='red'>第八名:</font> "+ rank_today.iloc[7,0]+" "+str(rank_today.iloc[7,10])+"__",
            "__<font color='red'>第九名:</font> "+ rank_today.iloc[8,0]+" "+str(rank_today.iloc[8,10])+"__",
            "__<font color='red'>第十名:</font> "+ rank_today.iloc[9,0]+" "+str(rank_today.iloc[9,10])+"__"]))
            
            
            
                
    }  
}
send = requests.post(url,json = mark, headers = headers)


# 区域招生发送

# In[133]:


from requests_toolbelt.multipart.encoder import MultipartEncoder
#图片上传密匙
unix_timestamp = time.time()
a = {'signType':'md5'}
b = { "timestamp": str(int(unix_timestamp))} 
c = {'bizCode':"XXXXXXXXXXXXXXXX" }  
secret = 'XXXXXXXXXXXXXXXXXX'
a1 = '='.join(x + '=' + y for x, y in a.items())
b1 = '='.join(x + '=' + y for x, y in b.items())
c1 = '='.join(x + '=' + y for x, y in c.items())
password = c1+ '&'+ a1 + '&' + b1 +secret
p_md5 = hashlib.md5(password.encode('utf-8')).hexdigest()


#图片上传
url = 'XXXXXXXXXXXXXXXXXX/upload'
headers = {
    "bizCode":"XXXXXXXXXXX",
    "signType":"md5",
    "timestamp":str(int(unix_timestamp)),
    "sign":p_md5
}
payload={'type': 'image'}
files=[
  ('file',('total_today.jpg',open('total_today.jpg','rb'),'image/jpge'))
]
pic = requests.post(url,files = files, headers = headers,data=payload)
picsend = pic.json()
print(pic.text)


# In[134]:


from requests_toolbelt.multipart.encoder import MultipartEncoder
#图片上传密匙

unix_timestamp = time.time()
a = {'signType':'md5'}
b = { "timestamp": str((round(unix_timestamp*1000)))} 
c = {'bizCode':"XXXXXXXXXXXXXX" }  
secret = 'XXXXXXXXXXXXXXXXXX'
a1 = '='.join(x + '=' + y for x, y in a.items())
b1 = '='.join(x + '=' + y for x, y in b.items())
c1 = '='.join(x + '=' + y for x, y in c.items())
password = c1+ '&'+ a1 + '&' + b1 +secret
md5_r = hashlib.md5(password.encode('utf-8')).hexdigest()
mediacode = picsend['data']['mediaCode']



#机器人接口
url = 'https://internal-ei.baijia.com/ei-serve-management-logic/internal/teamRobot/sign/sendMessage?key=XXXXXXXXXXXXXXXXXX'

headers = {
    "bizCode":"XXXXXXXXXXXXXXX",
    "signType":"md5",
    "timestamp":str((round(unix_timestamp*1000))),
    "sign":md5_r
}

params = { 
      "toDomains": ["@all"],
    "messageType": 1,
        "image":{
        "mediaCode":mediacode
    }  
}
send = requests.post(url,json = params, headers = headers)


# In[135]:



#标语发送

mark = { 
      "toDomains": ["@all"],
    "messageType": 24,
        "markdown":{
        "body": ('\n'.join([
    "__<font color='black'>今日截止目前0元区域/大区排名:</font>__",
            "__<font color='red'>第一名:</font> "+ area_first.iloc[0,0]+" "+str(area_first.iloc[0,1])+"__", 
            "__<font color='red'>第二名:</font> "+ area_second.iloc[0,0]+" "+str(area_second.iloc[0,1])+"__",
            "__<font color='red'>第三名:</font> "+ area_third.iloc[0,0]+" "+str(area_third.iloc[0,1])+"__"]))
            
            
            
                
    }  
}
send = requests.post(url,json = mark, headers = headers)


# In[5]:


from requests_toolbelt.multipart.encoder import MultipartEncoder
unix_timestamp = time.time()
a = {'signType':'md5'}
b = { "timestamp": str((round(unix_timestamp*1000)))} 
c = {'bizCode':"XXXXXXXXXXXXXXXXXXXXXX" }  
secret = 'XXXXXXXXXXXXXXXXXXXXXXX'
a1 = '='.join(x + '=' + y for x, y in a.items())
b1 = '='.join(x + '=' + y for x, y in b.items())
c1 = '='.join(x + '=' + y for x, y in c.items())
password = c1+ '&'+ a1 + '&' + b1 +secret
md5_r = hashlib.md5(password.encode('utf-8')).hexdigest()
url = 'https://internal-ei.baijia.com/ei-serve-management-logic/internal/teamRobot/sign/sendMessage?key=XXXXXXXXXXXXXXXXX'
headers = {
    "bizCode":"XXXXXXXXXXXXXXXXXXXXXX",
    "signType":"md5",
   "timestamp":str((round(unix_timestamp*1000))),
    "sign":md5_r
}
mark = { 
      "toDomains": ["@all"],
    "messageType": 24,
        "markdown":{
        "body": ('\n'.join([
    "__<font color='black'>商务渠道招生情况:</font>__" ,
            "__<font color='red'>本月总计:  "+ str(monthlytotal) +"</font>__", 
            "__<font color='red'>本周总计:  "+ str(weeklyTotal) +"</font>__",
            str(business.iloc[0,0])+" 今日小学领课:  "+"__<font color='red'>"+str(business.iloc[0,3])+"</font>__",
            str(business.iloc[1,0])+" 今日小学领课:  "+"__<font color='red'>"+str(business.iloc[1,3])+"</font>__",
            str(business.iloc[2,0])+" 今日小学领课:  "+"__<font color='red'>"+str(business.iloc[2,3])+"</font>__",
            str(business.iloc[3,0])+" 今日小学领课:  "+"__<font color='red'>"+str(business.iloc[3,3])+"</font>__",
        ]
        ))
    }  
}
send = requests.post(url,json = mark, headers = headers)


# In[51]:


from requests_toolbelt.multipart.encoder import MultipartEncoder
#图片上传密匙
unix_timestamp = time.time()
a = {'signType':'md5'}
b = { "timestamp": str(int(unix_timestamp))} 
c = {'bizCode':"XXXXXXXXXXXXXXXXXXXX" }  
secret = 'XXXXXXXXXXXXXXXXXXXXXXXX'
a1 = '='.join(x + '=' + y for x, y in a.items())
b1 = '='.join(x + '=' + y for x, y in b.items())
c1 = '='.join(x + '=' + y for x, y in c.items())
password = c1+ '&'+ a1 + '&' + b1 +secret
p_md5 = hashlib.md5(password.encode('utf-8')).hexdigest()


#图片上传
url = 'XXXXXXXXXXXXXXXXXXX/upload'
headers = {
    "bizCode":"yuanzishuju",
    "signType":"md5",
    "timestamp":str(int(unix_timestamp)),
    "sign":p_md5
}
payload={'type': 'file'}
files=[
  ('file',('个人招生明细.xlsx',open('个人招生明细.xlsx','rb'),'xlsx'))
]
pic = requests.post(url,files = files, headers = headers,data=payload)
picsend = pic.json()
print(pic.text)







# In[52]:


from requests_toolbelt.multipart.encoder import MultipartEncoder

unix_timestamp = time.time()
a = {'signType':'md5'}
b = { "timestamp": str((round(unix_timestamp*1000)))} 
c = {'bizCode':"XXXXXXXXXXXXXXXXXXXXXX" }  
secret = 'XXXXXXXXXXXXXXXXX'
a1 = '='.join(x + '=' + y for x, y in a.items())
b1 = '='.join(x + '=' + y for x, y in b.items())
c1 = '='.join(x + '=' + y for x, y in c.items())
password = c1+ '&'+ a1 + '&' + b1 +secret
md5_r = hashlib.md5(password.encode('utf-8')).hexdigest()
mediacode = picsend['data']['mediaCode']


# In[53]:





#机器人接口
url = 'https://internal-ei.baijia.com/ei-serve-management-logic/internal/teamRobot/sign/sendMessage?key=XXXXXXXXXXXXXXXXXX'

headers = {
    "bizCode":"XXXXXXXXXXXXXXX",
    "signType":"md5",
   "timestamp":str((round(unix_timestamp*1000))),
    "sign":md5_r
}

params = { 
      "toDomains": ["@all"],
    "messageType": 6,
        "file":{
        "mediaCode":mediacode
    }  
}
send = requests.post(url,json = params, headers = headers)


# In[ ]:


#图片上传
url = 'XXXXXXXXXXXXXXXXXXXX/upload'
headers = {
    "bizCode":"XXXXXXXXXXXXXXXXX",
    "signType":"md5",
    "timestamp":str(int(unix_timestamp)),
    "sign":p_md5
}
payload={'type': 'file'}
files=[
  ('file',('第四大区招生明细.xlsx',open('第四大区招生明细.xlsx','rb'),'xlsx'))
]
pic = requests.post(url,files = files, headers = headers,data=payload)
picsend = pic.json()
print(pic.text)







# In[52]:


from requests_toolbelt.multipart.encoder import MultipartEncoder

unix_timestamp = time.time()
a = {'signType':'md5'}
b = { "timestamp": str((round(unix_timestamp*1000)))} 
c = {'bizCode':"XXXXXXXXXXXXXXXXXXXXX" }  
secret = 'XXXXXXXXXXXXXXXXXXXXXX'
a1 = '='.join(x + '=' + y for x, y in a.items())
b1 = '='.join(x + '=' + y for x, y in b.items())
c1 = '='.join(x + '=' + y for x, y in c.items())
password = c1+ '&'+ a1 + '&' + b1 +secret
md5_r = hashlib.md5(password.encode('utf-8')).hexdigest()
mediacode = picsend['data']['mediaCode']


# In[53]:





#机器人接口
url = 'https://internal-ei.baijia.com/ei-serve-management-logic/internal/teamRobot/sign/sendMessage?key=XXXXXXXXXXXXXXXXXX'

headers = {
    "bizCode":"XXXXXXXXXXXXXXXXXXXXXXX",
    "signType":"md5",
   "timestamp":str((round(unix_timestamp*1000))),
    "sign":md5_r
}

params = { 
      "toDomains": ["@all"],
    "messageType": 6,
        "file":{
        "mediaCode":mediacode
    }  
}
send = requests.post(url,json = params, headers = headers)

