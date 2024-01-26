C:
cd C:\TestManager
powershell set-executionpolicy Bypass -scope CurrentUser
powershell set-executionpolicy Bypass -scope process
powershell set-executionpolicy RemoteSigned -scope process
reg add HKCU\Software\Microsoft\Windows\CurrentVersion\Run /v "TestManager" /t REG_SZ /d "C:\TestManager\test.bat" /f
powershell start-process "C:\TestManager\TestManager.exe" -verb RunAs
exit