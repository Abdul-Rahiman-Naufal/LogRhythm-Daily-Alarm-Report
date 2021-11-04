# Developed by abdulrahimannaufal@gmail.com

import pyodbc 
import smtplib
import subprocess
import re
import sys
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import xml.etree.ElementTree as ET
from pathlib import Path

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


conn = pyodbc.connect('Driver={SQL Server};'
                      'Server=ServerName;'
                      'Database=LogRhythm_Alarms;'
					  'UID=username;'
					  'PWD=password;'
                      'Trusted_Connection=yes;')
					  
cursor = conn.cursor()


query="""
select Name, AutoClosed,New,OpenAlarm,Reported,Resolved,FalsePositive,Monitor
from
(
select AlarmName.Name,count(AlarmName.Name) as AlarmCount,
(Convert(nvarchar, Case When AlarmStatus = 4 Then 'AutoClosed' When AlarmStatus = 8 Then 'Reported' When AlarmStatus = 6 Then 'Resolved' When AlarmStatus = 5 Then 'FalsePositive' When AlarmStatus = 0 Then 'New' When AlarmStatus = 1 Then 'OpenAlarm' when AlarmStatus = 9 Then 'Monitor' End)) as AlarmStatus
from [LogRhythm_Alarms].[dbo].[Alarm] as Alarm join LogRhythmEMDB.dbo.AlarmRule as AlarmName on Alarm.AlarmRuleID=AlarmName.AlarmRuleID
where 
Alarm.AlarmDate BETWEEN getdate()-DATEADD(hour,37,0) AND  getdate()-DATEADD(hour,13,0)
group by AlarmName.Name,AlarmStatus

) d
pivot
(
  max(AlarmCount)
  for AlarmStatus in (AutoClosed,New,Reported,Resolved,FalsePositive,OpenAlarm,Monitor)
) piv
"""
cursor.execute(query)

table_row_autoclosed=""
table_row_stoppage=""
table_row_na_closed=""

colAutoClose=0
colNew=0
colOpen=0
colReported=0
colResolved=0
colFalse=0
colMonitor=0

TotalAlarms=0
grandTotal=0

for row in cursor:
  if(str(row[1])=="None"):
    colAutoClose=0
  else:
    colAutoClose=int(row[1])

  if(str(row[2])=="None"):
    colNew=0
  else:
    colNew=int(row[2])

  if(str(row[3])=="None"):
    colOpen=0
  else:
    colOpen=int(row[3])

  if(str(row[4])=="None"):
    colReported=0
  else:
    colReported=int(row[4])

  if(str(row[5])=="None"):
    colResolved=0
  else:
    colResolved=int(row[5])

  if(str(row[6])=="None"):
    colFalse=0
  else:
    colFalse=int(row[6])

  if(str(row[7])=="None"):
    colMonitor=0
  else:
    colMonitor=int(row[7])

  TotalAlarms=colAutoClose+colNew+colOpen+colReported+colResolved+colFalse+colMonitor
  
  if ("log stoppage" in str(row[0]).lower() or "logrhythm" in str(row[0]).lower()):
    table_row_stoppage=table_row_stoppage+"<tr><td> "+row[0]+"</td><td> "+str(TotalAlarms)+"</td></tr>"
  elif(colAutoClose==0):
    grandTotal=grandTotal+TotalAlarms
    table_row_na_closed=table_row_na_closed+"<tr><td> "+str(row[0])+"</td><td> "+str(TotalAlarms)+"</td><td>"+str(colNew)+" </td> <td>"+str(colOpen)+" </td><td> "+str(colReported)+"</td> <td>"+str(colResolved)+" </td> <td>"+str(colFalse)+" </td> <td>"+str(colMonitor)+" </td></tr>"      
  else:
    table_row_autoclosed=table_row_autoclosed+"<tr><td> "+row[0]+"</td><td> "+str(TotalAlarms)+"</td></tr>"	
    
cursor.close()
conn.close()



sender_address = "sender@email.com"
receiver_address= "Receipients@email.om"


receiver_address=list(receiver_address.split(",")) 

html = """\

<html>

  <head>
<style>
table, th, td {
  border: 2px solid black;
  border-collapse: collapse;
}
</style>
  </head>

  <body>

    <p>Dear Team,
    <br>
      <br>
     
     Below rules were triggered for previous day.
      
    </p>
	
  <p><b>Non auto closed alarms:"""+str(grandTotal)+"""</b></p>
	<table>
      <tr><th> &nbsp; Rule Name &nbsp;</th> <th>&nbsp; Total &nbsp;</th> <th>&nbsp; New &nbsp;</th> <th>&nbsp; Open &nbsp;</th> <th>&nbsp; Reported &nbsp;</th> <th>&nbsp; Resolved &nbsp;</th> <th>&nbsp; False Positive &nbsp;</th> <th>&nbsp; Monitor &nbsp;</th></tr>

	"""+table_row_na_closed+"""
	</table>

</br>

<p><b>Auto Closed alarms:</b></p>
	<table>
      <tr><th> &nbsp; Rule Name &nbsp;</th> <th>&nbsp; Count &nbsp;</th></tr>

	"""+table_row_autoclosed+"""
	</table>

</br>


<p><b>Stoppage and LogRhythm alarms:</b></p>
<table>
      <tr><th> &nbsp; Rule Name &nbsp;</th> <th>&nbsp; Count &nbsp;</th></tr>

	"""+table_row_stoppage+"""
	</table>

  </body>

</html>

"""

mail_content = html

message = MIMEMultipart()
message['From'] = sender_address
message['To'] = ", ".join(receiver_address)
message['Subject'] = "Daily Alarm Report"
message.attach(MIMEText(mail_content, 'html'))

session = smtplib.SMTP("SMTPServerName")
text = message.as_string()
session.sendmail(sender_address, receiver_address, text)
session.quit()
print('Mail Sent')
