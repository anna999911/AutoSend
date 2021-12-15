#!/usr/bin/env python
# coding: utf-8

# In[ ]:


# python
# coding: utf-8
# author: Jason kuo 2021/4/26

# In[2]:

# 所需套件
#     excel2img(excel截圖) docx(製作word檔) docx2pdf(word檔轉pdf) smtplib email(自動寄信) pandas(解析excel)

import excel2img
from docx import Document
from docx.shared import Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pandas as pd
from docx2pdf import convert
import time
import sys
import os
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# log紀錄時間
path_c = 'complete.txt'
path_e = 'error.txt'

file_name= input("請輸入檔案名稱:") 

f = open(path_c, 'a')
f.write(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())+'\n')
f.close()
f = open(path_e, 'a')
f.write(time.strftime('%Y-%m-%d %H:%M:%S', time.localtime())+'\n')
f.close()

xls = pd.ExcelFile(file_name)
# name_list為excel工作表裡有的名字
name_list=xls.sheet_names
date= input("請輸入執行日:(請以以下格式:2021/**/01)")
#first_pic=input("請輸入第一張截圖網址:")
#second_pic=input("請輸入第二張截圖網址:")
#third_pic=input("請輸入第三張截圖網址:")
#fourth_pic=input("請輸入第四張截圖網址:")
#fifth_pic=input("請輸入第五張截圖網址:")

first_pic="https://i.imgur.com/3qd8Az9.jpg"
second_pic="https://i.imgur.com/YgCxfQO.jpg"
third_pic="https://i.imgur.com/28C06eU.jpg"
fourth_pic="https://i.imgur.com/yFANPpN.jpg"

# 計算正在轉第幾個檔案
count=0

print("程式開始執行")

# 製作對帳單(先截圖 再放入word檔)
for k in name_list:
    try:
        # 截圖excel中的明細 (範圍A1:L23)
        excel2img.export_img(file_name, k+".bmp", "", k+"!A1:L23")
        df = pd.read_excel(file_name, sheet_name=k) 
        # 信件檔案
        filename = "sample_letter.docx"
        doc = Document(filename)
        doc.add_picture(k+".bmp", width=Cm(19.5))   
        time.sleep(1)

        doc.save(str(date[5:7])+'_pic_'+k+'.docx')
        print(str(count+1)+"/"+str(len(name_list)))
        count+=1
        print(str(date[5:7])+'_pic_'+k+'.docx  截圖OK')
        f = open(path_c, 'a',encoding='utf-8')
        f.write(k+'.docx  截圖OK\n')
        f.close()
        time.sleep(1)
        
    except:
        print('!!!!!!!!!!!!!!!!'+k+'picture   有問題!!!!!!!!!!!!!!!!!')
        f = open(path_e, 'a',encoding='utf-8')
        f.write(k+'照片截圖   有問題\n')
        f.close()
        s=sys.exc_info()
        print ("Error '%s' happened on line %d" % (s[1],s[2].tb_lineno))
        print(str(count+1)+"/"+str(len(name_list)))
        count+=1
        continue
print("============截圖結束 開始轉成pdf===============")        
count=0   
# # 將製作好的word檔 轉換為pdf檔
for w in name_list:
    try:
        convert(str(date[5:7])+'_pic_'+w+'.docx','數寶數位資產對帳單_'+w+'.pdf')
        print(w+'.pdf  OK')
        f = open(path_c, 'a',encoding='utf-8')
        f.write(w+'pdf OK\n')
        f.close()
        print(str(count+1)+"/"+str(len(name_list)))
        count+=1
        os.remove(str(date[5:7])+'_pic_'+w+'.docx')
        os.remove(w+'.bmp')
        time.sleep(3)
    except:
        print('!!!!!!!!!!!!!!!!'+w+'pdf   有問題!!!!!!!!!!!!!!!!!')
        f = open(path_e, 'a',encoding="utf-8")
        f.write(w+'轉換pdf   有問題\n')
        f.close()
        print(str(count+1)+"/"+str(len(name_list)))
        count+=1
        time.sleep(5)
print("=============轉為pdf結束 開始寄信==============")      
count=0

# 讀取excel資料 改信件的數字 並寄出信件
for z in name_list:
    count_whichmonth=0
    try:
        print(z)
        df = pd.read_excel(file_name, sheet_name=z)
        name="顧客"
        name1=df.iat[3,2]#name
        if len(name1)>1:
            name=name1[-2:]

        for i in range(13):
            count_whichmonth+=1
            if str(df.iat[9+i,9])=="nan":
                totalnumber=round(df.iat[9+i-1,9],2)   
                break
                #累積報酬

        for i in range(13):
            if str(df.iat[9+i,10])=="nan":
                totalrate_p=round(df.iat[9+i-1,10],4)*100
                break
                #累積報酬率

        for i in range(13):
            if str(df.iat[9+i,11])=="nan":
                if str(df.iat[9+i+1,11])=="nan":
                    year_rate=round(df.iat[9+i-1,11],4)*100
                    break
                #累積年報酬率

        dictionary = {"888name888":name,"888date888":date, "888all_earn888": name,
                        "888all_earn888":str(format(totalnumber, '0,.2f')),
                        "888all_earn_percent888":str(format(float('%.2f'%totalrate_p), '0,.2f')),
                        "888yearly_earn888":str(format(float('%.2f'%year_rate), '0,.2f')),
                        "888firstpic888":first_pic,
                        "888secondpic888":second_pic,
                        "888thirdpic888":third_pic,
                        "888fourthpic888":fourth_pic
                        }               
        print(name)

        COMMASPACE = ', '
        sender = 'j965553@gmail.com'
        gmail_password = 'Khbwweiobkasoxsw'
    #     recipients = ['ryanbo1982@gmail.com','ryan.hsu@shubaoex.com']
        recipients = ['chenann0420@gmail.com']
        #recipients = ['shubaoex2018@gmail.com']
        # recipients = ['j965553j@gmail.com']



        # 建立郵件主題
        outer = MIMEMultipart()
        subject='【數寶數位資產管理_比特幣定投服務】每月對帳單_'+name1+'_第'+str(count_whichmonth-1)+'期'
        subject_re=subject.replace("12","十二").replace("11","十一").replace("10","十").replace("9","九").replace("8","八").replace("7","七").replace("6","六").replace("5","五").replace("4","四").replace("3","三").replace("2","二").replace("1","一")
        subject_re=subject_re+'_累積報酬率為'+str(format(float('%.2f'%totalrate_p), '0,.2f'))+'%'
        outer['Subject'] =subject_re
        print(outer['Subject'])
        outer['To'] = COMMASPACE.join(recipients)
        outer['From'] = sender
        outer.preamble = 'You will not see this in a MIME-aware mail reader.\n'

        # 檔案位置 在windows底下記得要加上r 如下 要完整的路徑
        attachments = [r'C:/Users/數寶數位資產管理/Documents/'+'數寶數位資產對帳單_'+z+'.pdf','數寶數位資產管理_產品單張DM.pdf','比特幣每月定期定額投資 2017年 _ 2021年以來每年的定投報酬率.pdf']
        # HTML檔案位置
        path = 'C:/Users/數寶數位資產管理/desktop/202110/letter_html_BTC.html'
        htmlfile = open(path, 'r', encoding='utf-8')
        htmlhandle = htmlfile.read()
        index_html=htmlhandle
        #替代html內文字
        for word, initial in dictionary.items():
            index_html = index_html.replace(word, initial)

        # 處理我們的文字 MIMEtext
        mine_html = MIMEText(_text=index_html, _subtype="html", _charset="UTF8")
        outer.attach(mine_html)
        for file in attachments:
            try:
                with open(file, 'rb') as fp:
                    print ('讀到檔案了')
                    msg = MIMEBase('application', "octet-stream")
                    msg.set_payload(fp.read())
                    encoders.encode_base64(msg)
                    msg.add_header('Content-Disposition', 'attachment', filename=os.path.basename(file))
                    outer.attach(msg)
            except:
                print("無法打開你的附加檔案. Error: ", sys.exc_info()[0])
                raise

            composed = outer.as_string()

        # 寄送EMAIL
        try:
            with smtplib.SMTP('smtp.gmail.com', 587) as s:
                s.ehlo()
                s.starttls()
                s.ehlo()
                s.login(sender, gmail_password)
                s.sendmail(sender, recipients, composed)
                s.close()
                print("Email 寄出了歐!")
                f = open(path_c, 'a',encoding="utf-8")
                f.write(z+'Email 寄出\n')
                f.close()
                print(str(count+1)+"/"+str(len(name_list)))
                count+=1
                time.sleep(1)
        except:
            print("無法寄出Email. Error: ", sys.exc_info()[0])
            time.sleep(3)
            raise
    except:
        print("!!!!!!!!!"+z+"email有問題 !!!!!!!!!!!!!!!!!!")
        f = open(path_e, 'a',encoding="utf-8")
        f.write(z+'寄出email   有問題\n')
        f.close()
        print("無法寄出Email. Error: ", sys.exc_info()[0]) 
print("==============全部皆已完成=======================")      

