powershell set-executionpolicy Bypass -scope CurrentUser
powershell set-executionpolicy Bypass -scope process
powershell start-process "C:\TestManager\TM1006.exe" -verb RunAs
exit