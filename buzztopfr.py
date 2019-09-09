# -*- coding: utf-8 -*-
"""
Created on Wed Aug 28 15:26:14 2019

@author: yzhao
"""

import requests
import json
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header


'''Using Different API in order to get Buzz information from different Country. 1 is always represent top buzzing,2 is represent 24h top volume'''

buzz_eu=[]
buzz_as=[]
top_eu=[]
top_as=[]

'------------------------------EUROPE-------------------------------------------'
'---------France------------'
url_fr1='https://api.dev.tradingcentral.com/articlecounts/v3/trends?entityGroups=FR&limit=5&lang=en&fieldsfilter=**&token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJjbGllbnRJZCI6InJjLTUiLCJyY1VzZXJJZCI6IjE0NzM0MzY5MCIsImRldiI6dHJ1ZSwiaW50ZXJuYWwiOnRydWUsImlhdCI6MTU0ODcxMDQyMX0.YEiyxu68zmlmBjCdXQen-KHTJhWoRqJBVLWr8WXzTPU'
r_fr1=requests.get(url_fr1)
text_fr1=r_fr1.text
buzz_eu.append(text_fr1)

url_fr2='https://api.dev.tradingcentral.com/articlecounts/v3/volume?entitygroups=FR&rollup=1day&limit=5&lang=en&fieldsfilter=**&token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJjbGllbnRJZCI6InJjLTUiLCJyY1VzZXJJZCI6IjE0NzM0MzY5MCIsImRldiI6dHJ1ZSwiaW50ZXJuYWwiOnRydWUsImlhdCI6MTU0ODcxMDQyMX0.YEiyxu68zmlmBjCdXQen-KHTJhWoRqJBVLWr8WXzTPU'
r2_fr2=requests.get(url_fr2)
text_fr2=r2_fr2.text
top_eu.append(text_fr2)

'----------Germany----------'
url_de1='https://api.dev.tradingcentral.com/articlecounts/v3/trends?entityGroups=DE&limit=5&lang=en&token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJjbGllbnRJZCI6InJjLTUiLCJyY1VzZXJJZCI6IjE0NzM0MzY5MCIsImRldiI6dHJ1ZSwiaW50ZXJuYWwiOnRydWUsImlhdCI6MTU0ODcxMDQyMX0.YEiyxu68zmlmBjCdXQen-KHTJhWoRqJBVLWr8WXzTPU'
r_de1=requests.get(url_de1)
text_de1=r_de1.text
buzz_eu.append(text_de1)

url_de2='https://api.dev.tradingcentral.com/articlecounts/v3/volume?entitygroups=DE&rollup=1day&limit=5&lang=en&token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJjbGllbnRJZCI6InJjLTUiLCJyY1VzZXJJZCI6IjE0NzM0MzY5MCIsImRldiI6dHJ1ZSwiaW50ZXJuYWwiOnRydWUsImlhdCI6MTU0ODcxMDQyMX0.YEiyxu68zmlmBjCdXQen-KHTJhWoRqJBVLWr8WXzTPU'
r_de2=requests.get(url_de2)
text_de2=r_de2.text
top_eu.append(text_de2)


'-----------UK--------------'
url_uk1='https://api.dev.tradingcentral.com/articlecounts/v3/trends?entityGroups=UK&limit=5&lang=en&token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJjbGllbnRJZCI6InJjLTUiLCJyY1VzZXJJZCI6IjE0NzM0MzY5MCIsImRldiI6dHJ1ZSwiaW50ZXJuYWwiOnRydWUsImlhdCI6MTU0ODcxMDQyMX0.YEiyxu68zmlmBjCdXQen-KHTJhWoRqJBVLWr8WXzTPU'
r_uk1=requests.get(url_uk1)
text_uk1=r_uk1.text
buzz_eu.append(text_uk1)

url_uk2='https://api.dev.tradingcentral.com/articlecounts/v3/volume?entitygroups=UK&rollup=1day&limit=5&lang=en&token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJjbGllbnRJZCI6InJjLTUiLCJyY1VzZXJJZCI6IjE0NzM0MzY5MCIsImRldiI6dHJ1ZSwiaW50ZXJuYWwiOnRydWUsImlhdCI6MTU0ODcxMDQyMX0.YEiyxu68zmlmBjCdXQen-KHTJhWoRqJBVLWr8WXzTPU'
r_uk2=requests.get(url_uk2)
text_uk2=r_uk2.text
top_eu.append(text_uk2)


'---------Italy-------------'
url_it1='https://api.dev.tradingcentral.com/articlecounts/v3/trends?entityGroups=IT&limit=5&lang=en&token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJjbGllbnRJZCI6InJjLTUiLCJyY1VzZXJJZCI6IjE0NzM0MzY5MCIsImRldiI6dHJ1ZSwiaW50ZXJuYWwiOnRydWUsImlhdCI6MTU0ODcxMDQyMX0.YEiyxu68zmlmBjCdXQen-KHTJhWoRqJBVLWr8WXzTPU'
r_it1=requests.get(url_it1)
text_it1=r_it1.text
buzz_eu.append(text_it1)

url_it2='https://api.dev.tradingcentral.com/articlecounts/v3/volume?entitygroups=IT&rollup=1day&limit=5&lang=en&token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJjbGllbnRJZCI6InJjLTUiLCJyY1VzZXJJZCI6IjE0NzM0MzY5MCIsImRldiI6dHJ1ZSwiaW50ZXJuYWwiOnRydWUsImlhdCI6MTU0ODcxMDQyMX0.YEiyxu68zmlmBjCdXQen-KHTJhWoRqJBVLWr8WXzTPU'
r_it2=requests.get(url_it2)
text_it2=r_it2.text
top_eu.append(text_it2)


'--------Switzerland--------'
url_sw1='https://api.dev.tradingcentral.com/articlecounts/v3/trends?entityGroups=CH&limit=5&lang=en&token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJjbGllbnRJZCI6InJjLTUiLCJyY1VzZXJJZCI6IjE0NzM0MzY5MCIsImRldiI6dHJ1ZSwiaW50ZXJuYWwiOnRydWUsImlhdCI6MTU0ODcxMDQyMX0.YEiyxu68zmlmBjCdXQen-KHTJhWoRqJBVLWr8WXzTPU'
r_sw1=requests.get(url_sw1)
text_sw1=r_sw1.text
buzz_eu.append(text_sw1)

url_sw2='https://api.dev.tradingcentral.com/articlecounts/v3/volume?entitygroups=CH&rollup=1day&limit=5&lang=en&token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJjbGllbnRJZCI6InJjLTUiLCJyY1VzZXJJZCI6IjE0NzM0MzY5MCIsImRldiI6dHJ1ZSwiaW50ZXJuYWwiOnRydWUsImlhdCI6MTU0ODcxMDQyMX0.YEiyxu68zmlmBjCdXQen-KHTJhWoRqJBVLWr8WXzTPU'
r_sw2=requests.get(url_sw2)
text_sw2=r_sw2.text
top_eu.append(text_sw2)


'-----------Spain-----------'
url_es1='https://api.dev.tradingcentral.com/articlecounts/v3/trends?entityGroups=ES&limit=5&lang=en&token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJjbGllbnRJZCI6InJjLTUiLCJyY1VzZXJJZCI6IjE0NzM0MzY5MCIsImRldiI6dHJ1ZSwiaW50ZXJuYWwiOnRydWUsImlhdCI6MTU0ODcxMDQyMX0.YEiyxu68zmlmBjCdXQen-KHTJhWoRqJBVLWr8WXzTPU'
r_es1=requests.get(url_es1)
text_es1=r_es1.text
buzz_eu.append(text_es1)

url_es2='https://api.dev.tradingcentral.com/articlecounts/v3/volume?entitygroups=ES&rollup=1day&limit=5&lang=en&token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJjbGllbnRJZCI6InJjLTUiLCJyY1VzZXJJZCI6IjE0NzM0MzY5MCIsImRldiI6dHJ1ZSwiaW50ZXJuYWwiOnRydWUsImlhdCI6MTU0ODcxMDQyMX0.YEiyxu68zmlmBjCdXQen-KHTJhWoRqJBVLWr8WXzTPU'
r_es2=requests.get(url_es2)
text_es2=r_es2.text
top_eu.append(text_es2)


'----------------------------------APAC--------------------------------------'
'------------China----------'
url_cn1='https://api.dev.tradingcentral.com/articlecounts/v3/trends?entityGroups=CN&limit=5&lang=en&token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJjbGllbnRJZCI6InJjLTUiLCJyY1VzZXJJZCI6IjE0NzM0MzY5MCIsImRldiI6dHJ1ZSwiaW50ZXJuYWwiOnRydWUsImlhdCI6MTU0ODcxMDQyMX0.YEiyxu68zmlmBjCdXQen-KHTJhWoRqJBVLWr8WXzTPU'
r_cn1=requests.get(url_cn1)
text_cn1=r_cn1.text
buzz_as.append(text_cn1)

url_cn2='https://api.dev.tradingcentral.com/articlecounts/v3/volume?entitygroups=CN&rollup=1day&limit=5&lang=en&token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJjbGllbnRJZCI6InJjLTUiLCJyY1VzZXJJZCI6IjE0NzM0MzY5MCIsImRldiI6dHJ1ZSwiaW50ZXJuYWwiOnRydWUsImlhdCI6MTU0ODcxMDQyMX0.YEiyxu68zmlmBjCdXQen-KHTJhWoRqJBVLWr8WXzTPU'
r_cn2=requests.get(url_cn2)
text_cn2=r_cn2.text
top_as.append(text_cn2)


'--------Hong Kong----------'
url_hk1='https://api.dev.tradingcentral.com/articlecounts/v3/trends?entityGroups=HK&limit=5&lang=en&token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJjbGllbnRJZCI6InJjLTUiLCJyY1VzZXJJZCI6IjE0NzM0MzY5MCIsImRldiI6dHJ1ZSwiaW50ZXJuYWwiOnRydWUsImlhdCI6MTU0ODcxMDQyMX0.YEiyxu68zmlmBjCdXQen-KHTJhWoRqJBVLWr8WXzTPU'
r_hk1=requests.get(url_hk1)
text_hk1=r_hk1.text
buzz_as.append(text_hk1)

url_hk2='https://api.dev.tradingcentral.com/articlecounts/v3/volume?entitygroups=HK&rollup=1day&limit=5&lang=en&token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJjbGllbnRJZCI6InJjLTUiLCJyY1VzZXJJZCI6IjE0NzM0MzY5MCIsImRldiI6dHJ1ZSwiaW50ZXJuYWwiOnRydWUsImlhdCI6MTU0ODcxMDQyMX0.YEiyxu68zmlmBjCdXQen-KHTJhWoRqJBVLWr8WXzTPU'
r_hk2=requests.get(url_hk2)
text_hk2=r_hk2.text
top_as.append(text_hk2)


'-----------Japan-----------'
url_jp1='https://api.dev.tradingcentral.com/articlecounts/v3/trends?entityGroups=JP&limit=5&lang=en&token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJjbGllbnRJZCI6InJjLTUiLCJyY1VzZXJJZCI6IjE0NzM0MzY5MCIsImRldiI6dHJ1ZSwiaW50ZXJuYWwiOnRydWUsImlhdCI6MTU0ODcxMDQyMX0.YEiyxu68zmlmBjCdXQen-KHTJhWoRqJBVLWr8WXzTPU'
r_jp1=requests.get(url_jp1)
text_jp1=r_jp1.text
buzz_as.append(text_jp1)

url_jp2='https://api.dev.tradingcentral.com/articlecounts/v3/volume?entitygroups=JP&rollup=1day&limit=5&lang=en&token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJjbGllbnRJZCI6InJjLTUiLCJyY1VzZXJJZCI6IjE0NzM0MzY5MCIsImRldiI6dHJ1ZSwiaW50ZXJuYWwiOnRydWUsImlhdCI6MTU0ODcxMDQyMX0.YEiyxu68zmlmBjCdXQen-KHTJhWoRqJBVLWr8WXzTPU'
r_jp2=requests.get(url_jp2)
text_jp2=r_jp2.text
top_as.append(text_jp2)


'-----------USA-------------'
url_us1='https://api.dev.tradingcentral.com/articlecounts/v3/trends?entityGroups=US&limit=5&lang=en&token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJjbGllbnRJZCI6InJjLTUiLCJyY1VzZXJJZCI6IjE0NzM0MzY5MCIsImRldiI6dHJ1ZSwiaW50ZXJuYWwiOnRydWUsImlhdCI6MTU0ODcxMDQyMX0.YEiyxu68zmlmBjCdXQen-KHTJhWoRqJBVLWr8WXzTPU'
r_us1=requests.get(url_us1)
text_us1=r_us1.text
buzz_as.append(text_us1)

url_us2='https://api.dev.tradingcentral.com/articlecounts/v3/volume?entitygroups=US&rollup=1day&limit=5&lang=en&token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJjbGllbnRJZCI6InJjLTUiLCJyY1VzZXJJZCI6IjE0NzM0MzY5MCIsImRldiI6dHJ1ZSwiaW50ZXJuYWwiOnRydWUsImlhdCI6MTU0ODcxMDQyMX0.YEiyxu68zmlmBjCdXQen-KHTJhWoRqJBVLWr8WXzTPU'
r_us2=requests.get(url_us2)
text_us2=r_us2.text
top_as.append(text_us2)


'--------Singapore----------'
url_sg1='https://api.dev.tradingcentral.com/articlecounts/v3/trends?entityGroups=SG&limit=5&lang=en&token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJjbGllbnRJZCI6InJjLTUiLCJyY1VzZXJJZCI6IjE0NzM0MzY5MCIsImRldiI6dHJ1ZSwiaW50ZXJuYWwiOnRydWUsImlhdCI6MTU0ODcxMDQyMX0.YEiyxu68zmlmBjCdXQen-KHTJhWoRqJBVLWr8WXzTPU'
r_sg1=requests.get(url_sg1)
text_sg1=r_sg1.text
buzz_as.append(text_sg1)

url_sg2='https://api.dev.tradingcentral.com/articlecounts/v3/volume?entitygroups=SG&rollup=1day&limit=5&lang=en&token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJjbGllbnRJZCI6InJjLTUiLCJyY1VzZXJJZCI6IjE0NzM0MzY5MCIsImRldiI6dHJ1ZSwiaW50ZXJuYWwiOnRydWUsImlhdCI6MTU0ODcxMDQyMX0.YEiyxu68zmlmBjCdXQen-KHTJhWoRqJBVLWr8WXzTPU'
r_sg2=requests.get(url_sg2)
text_sg2=r_sg2.text
top_as.append(text_sg2)

'--------Australia----------'
url_au1='https://api.dev.tradingcentral.com/articlecounts/v3/trends?entityGroups=AU&limit=5&lang=en&token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJjbGllbnRJZCI6InJjLTUiLCJyY1VzZXJJZCI6IjE0NzM0MzY5MCIsImRldiI6dHJ1ZSwiaW50ZXJuYWwiOnRydWUsImlhdCI6MTU0ODcxMDQyMX0.YEiyxu68zmlmBjCdXQen-KHTJhWoRqJBVLWr8WXzTPU'
r_au1=requests.get(url_au1)
text_au1=r_au1.text
buzz_as.append(text_au1)

url_au2='https://api.dev.tradingcentral.com/articlecounts/v3/volume?entitygroups=AU&rollup=1day&limit=5&lang=en&token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJjbGllbnRJZCI6InJjLTUiLCJyY1VzZXJJZCI6IjE0NzM0MzY5MCIsImRldiI6dHJ1ZSwiaW50ZXJuYWwiOnRydWUsImlhdCI6MTU0ODcxMDQyMX0.YEiyxu68zmlmBjCdXQen-KHTJhWoRqJBVLWr8WXzTPU'
r_au2=requests.get(url_au2)
text_au2=r_au2.text
top_as.append(text_au2)


EUROPE=['France','Germany','UK','Italy','Switzerland','Spain']
ASPA=['China','Hong Kong','Japan','USA','Singapore','Australia']



'-------------Generate Excel File-----------------------'
'Buzzing Europe'
finallist=[]
for i in range(0,len(buzz_eu)):
    element=buzz_eu[i]
    diclist=json.loads(element)['articleCounts']
    dicdic={}
    country=EUROPE[i]
    for item in diclist:
        dicdic['Country']=country
        dicdic['Name']=item['entity']['name']
        dicdic['NewsNumber']=item['news']
        dicdic['BlogsNumber']=item['blogs']
        finallist.append(dicdic)
        dicdic={}

df=pd.DataFrame(finallist)
df=df[['Country','Name','NewsNumber','BlogsNumber']]
df.to_excel(r'/home/ec2-user/yuxuan/BuzzDailyReport/Report/EuropeBuzzing.xlsx')
    
'Top Europe'
finallist=[]
for i in range(0,len(top_eu)):
    element=top_eu[i]
    diclist=json.loads(element)['articleCounts']
    dicdic={}
    country=EUROPE[i]
    for item in diclist:
        dicdic['Country']=country
        dicdic['Name']=item['entity']['name']
        dicdic['NewsNumber']=item['news']
        dicdic['BlogsNumber']=item['blogs']
        finallist.append(dicdic)
        dicdic={}

df=pd.DataFrame(finallist)
df=df[['Country','Name','NewsNumber','BlogsNumber']]
df.to_excel(r'/home/ec2-user/yuxuan/BuzzDailyReport/Report/EuropeTop.xlsx')

'Buzzing ASPA'
finallist=[]
for i in range(0,len(buzz_as)):
    element=buzz_as[i]
    diclist=json.loads(element)['articleCounts']
    dicdic={}
    country=ASPA[i]
    for item in diclist:
        dicdic['Country']=country
        dicdic['Name']=item['entity']['name']
        dicdic['NewsNumber']=item['news']
        dicdic['BlogsNumber']=item['blogs']
        finallist.append(dicdic)
        dicdic={}

df=pd.DataFrame(finallist)
df=df[['Country','Name','NewsNumber','BlogsNumber']]
df.to_excel(r'/home/ec2-user/yuxuan/BuzzDailyReport/Report/ASPABuzzing.xlsx')
    
'Top ASPA'
finallist=[]
for i in range(0,len(top_as)):
    element=top_as[i]
    diclist=json.loads(element)['articleCounts']
    dicdic={}
    country=ASPA[i]
    for item in diclist:
        dicdic['Country']=country
        dicdic['Name']=item['entity']['name']
        dicdic['NewsNumber']=item['news']
        dicdic['BlogsNumber']=item['blogs']
        finallist.append(dicdic)
        dicdic={}

df=pd.DataFrame(finallist)
df=df[['Country','Name','NewsNumber','BlogsNumber']]
df.to_excel(r'/home/ec2-user/yuxuan/BuzzDailyReport/Report/ASPATop.xlsx')




'-------------------------send mail------------------------------------------------------------------' 
sender = 'zhaoyuxuanfr@gmail.com'
receivers = ['yuxuan.zhao@tradingcentral.com']  # 接收邮件，可设置为你的QQ邮箱或者其他邮箱
 
#创建一个带附件的实例
message = MIMEMultipart()
message['From'] = Header("TcMarketBuzz", 'utf-8')
message['To'] =  Header("Analysts", 'utf-8')
subject = 'BuzzDailyReport'
message['Subject'] = Header(subject, 'utf-8')
 
#邮件正文内容
message.attach(MIMEText('Hello Everyone\nBuzz Daily Report,this mail is automatically sent.....', 'plain', 'utf-8'))
 
# 构造附件1，传送当前目录下的 test.txt 文件
att1 = MIMEText(open(r'/home/ec2-user/yuxuan/BuzzDailyReport/Report/ASPABuzzing.xlsx', 'rb').read(), 'base64', 'utf-8')
att1["Content-Type"] = 'application/octet-stream'
# 这里的filename可以任意写，写什么名字，邮件中显示什么名字
att1["Content-Disposition"] = 'attachment; filename="ASPABuzzing.xlsx"'
message.attach(att1)
 
# 构造附件2，传送当前目录下的 runoob.txt 文件
att2 = MIMEText(open(r'/home/ec2-user/yuxuan/BuzzDailyReport/Report/ASPATop.xlsx', 'rb').read(), 'base64', 'utf-8')
att2["Content-Type"] = 'application/octet-stream'
att2["Content-Disposition"] = 'attachment; filename="ASPATop.xlsx"'
message.attach(att2)

# 构造附件1，传送当前目录下的 test.txt 文件
att3 = MIMEText(open(r'/home/ec2-user/yuxuan/BuzzDailyReport/Report/EuropeBuzzing.xlsx', 'rb').read(), 'base64', 'utf-8')
att3["Content-Type"] = 'application/octet-stream'
# 这里的filename可以任意写，写什么名字，邮件中显示什么名字
att3["Content-Disposition"] = 'attachment; filename="EuropeBuzzing.xlsx"'
message.attach(att3)
 
# 构造附件2，传送当前目录下的 runoob.txt 文件
att4 = MIMEText(open(r'/home/ec2-user/yuxuan/BuzzDailyReport/Report/EuropeTop.xlsx', 'rb').read(), 'base64', 'utf-8')
att4["Content-Type"] = 'application/octet-stream'
att4["Content-Disposition"] = 'attachment; filename="EuropeTop.xlsx"'
message.attach(att4)


mail_host='smtp.gmail.com'
mail_port=587
mail_user='zhaoyuxuanfr1@gmail.com'
mail_pass='z8757256'

flag_mail_ssl=0
#try:
if flag_mail_ssl:
    server=smtplib.SMTP_SSL(mail_host, mail_port)  # 发件人邮箱中的SMTP服务器
else:
    server=smtplib.SMTP(mail_host, mail_port)#
server.ehlo()
server.starttls()
server.login(mail_user,mail_pass)  
#    smtpObj = smtplib.SMTP('localhost')
server.sendmail(sender, receivers, message.as_string())
print("sent succuss")
server.quit()
#except smtplib.SMTPException:
#    print('ERROR!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!')