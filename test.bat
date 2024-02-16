C:
cd C:\TestManager
powershell set-executionpolicy Bypass -scope CurrentUser
powershell set-executionpolicy Bypass -scope process
powershell set-executionpolicy RemoteSigned -scope process
reg add HKLM\Software\Microsoft\Windows\CurrentVersion\Run /v "TestManager" /t REG_SZ /d "C:\TestManager\test.bat" /f
reg delete "HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Run" /v "TestManager" /f
reg import "C:\TestManager\logon.reg"
powershell start-process "C:\TestManager\TestManager.exe" -verb RunAs
exit