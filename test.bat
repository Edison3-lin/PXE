powershell set-executionpolicy Bypass -scope CurrentUser
powershell set-executionpolicy Bypass -scope process
powershell start-process "C:\TestManager\TM1003.exe" -verb RunAs
exit