@echo off
Rem batch file to initiate sanity test
Rem UiRobot.exe location
cd %userprofile%\AppData\Local\UiPath\app-20.4.0-beta0472
start /min UiRobot.exe execute --file "C:\Users\ANIRUDHGIREY\Documents\GitHub\pcmt_uipath\project.json"
