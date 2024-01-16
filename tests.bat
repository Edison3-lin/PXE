powershell set-executionpolicy Bypass -scope CurrentUser
powershell set-executionpolicy Bypass -scope process
powershell start-process "C:\TestManager\TM1007b1.exe" -verb RunAs
exit