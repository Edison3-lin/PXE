powershell set-executionpolicy Bypass -scope CurrentUser
powershell set-executionpolicy Bypass -scope process
powershell start-process "C:\TestManager\TM1004.exe" -verb RunAs
@REM powershell Start-Process -FilePath "C:\TestManager\TM1004.exe" -ArgumentList "MyTestAll2.dll" -Verb RunAs
exit