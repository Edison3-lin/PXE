powershell set-executionpolicy Bypass -scope CurrentUser
powershell set-executionpolicy Bypass -scope process
powershell start-process "C:\TestManager\TM1002.exe" -verb RunAs
exit