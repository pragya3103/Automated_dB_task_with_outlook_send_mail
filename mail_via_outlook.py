# -*- coding: utf-8 -*-
"""
Created on Mon Mar 22 13:36:52 2021

@author: pragyaja

"""



import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

import ibm_db_dbi
import pandas as pd
import time
from datetime import datetime, timedelta


timestr = time.strftime("%d_%m_%Y")
day = time.strftime("%A")

#date_to_check = ""

if day == 'Thursday':
    date_to_check = datetime.today() - timedelta(days=3)

if day == "Monday":
    date_to_check = datetime.today() - timedelta(days=4)


#date_to_check = "2021-03-15"
date_to_check = date_to_check.strftime("%Y-%m-%d 00:00:00.0")

#date = time.strftime("%Y-%m-%d 00:00:00.0")

connection = ibm_db_dbi.connect()
sql = "sql query"

df = pd.read_sql(sql, connection)
df.to_excel('path'+timestr+'.xlsx')

filepath = 'path'+timestr+'.xlsx'

fromaddr = ""
toaddr = "" 

msg = MIMEMultipart()

cc = "x,y"
rcpt = cc.split(",")  + [toaddr]

msg['From'] = fromaddr
msg['To'] = toaddr
msg['Cc'] = cc
msg['Subject'] = "filename_"+timestr


body = ''' BODY '''



msg.attach(MIMEText(body, 'plain'))

filename = "SSR176.xlsx"
attachment = open(filepath, "rb")

part = MIMEBase('application', 'octet-stream')
part.set_payload((attachment).read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

msg.attach(part)

server = smtplib.SMTP('smtp.outlook.com', 587)
server.starttls()
server.login("user id", "pass")
text = msg.as_string()
server.sendmail(fromaddr, rcpt, text)
server.quit()
