powershell set-executionpolicy Bypass -scope CurrentUser
powershell set-executionpolicy Bypass -scope process
powershell stop-process -Name TM1003 -Force
exit