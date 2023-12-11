powershell set-executionpolicy Bypass -scope CurrentUser
powershell set-executionpolicy Bypass -scope process
powershell Start-Process -FilePath "C:\TestManager\Service_in.bat" -verb RunAs
exit