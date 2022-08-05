from pywinauto import application
import time
import os

os.system('taskkill /IM ncStarter* /F /T')
os.system('taskkill /IM CpStart* /F /T') 
os.system('taskkill /IM DibServer* /F /T') 
os.system('wmic process where "name like \'%ncStarter%\'" call terminate') 
os.system('wmic process where "name like \'%CpStart%\'" call terminate') 
os.system('wmic process where "name like \'%DibServer%\'" call terminate')

time.sleep(5)        

app = application.Application()
app.start('C:\CREON\STARTER\\coStarter.exe /prj:cp /id:tjgus03 /pwd:ksh0924! /pwdcert:seohyunkim0924!! /autostart')
time.sleep(60)
