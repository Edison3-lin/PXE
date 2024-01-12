C:
cd C:\TestManager
powershell set-executionpolicy Bypass -scope CurrentUser
powershell set-executionpolicy Bypass -scope process
powershell set-executionpolicy RemoteSigned -scope process
reg add HKCU\Software\Microsoft\Windows\CurrentVersion\Run /v "TestManager" /t REG_SZ /d "C:\TestManager\test.bat" /f
powershell start-process "C:\TestManager\TM1007.exe" -verb RunAs

@REM powershell set-executionpolicy Bypass -scope CurrentUser
@REM powershell set-executionpolicy Bypass -scope process
@REM powershell Start-Process -FilePath "C:\TestManager\Service_Start.bat" -verb RunAs
exit