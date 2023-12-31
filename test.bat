powershell set-executionpolicy Bypass -scope CurrentUser
powershell set-executionpolicy Bypass -scope process
@REM powershell start-process "C:\TestManager\TM1006.exe" -verb RunAs
powershell Start-Process -FilePath "C:\TestManager\TM1006.exe" -ArgumentList "Template.dll 150" -Verb RunAs
exit