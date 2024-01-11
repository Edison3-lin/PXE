powershell set-executionpolicy Bypass -scope CurrentUser
powershell set-executionpolicy Bypass -scope process
powershell set-executionpolicy RemoteSigned -scope process
powershell Start-Process -FilePath "C:\TestManager\Service_Start.bat" -verb RunAs
exit