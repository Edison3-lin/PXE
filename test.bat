@REM powershell start-process c:\Users\edison\Downloads\NewRepo-master\TestManager\TestManager\bin\Debug\TestManager.exe -verb RunAs
powershell set-executionpolicy Bypass -scope CurrentUser
powershell set-executionpolicy Bypass -scope process
@REM powershell start-process c:\TestManager\TestManager.exe -WindowStyle Hidden -verb RunAs
reg add HKCU\Software\Microsoft\Windows\CurrentVersion\Run /v "TestManager.exe" /t REG_SZ /d "C:\TestManager\TestManager.exe" /f
powershell start-process .\TestManager.exe -verb RunAs
exit