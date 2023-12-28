powershell set-executionpolicy Bypass -scope CurrentUser
powershell set-executionpolicy Bypass -scope process
powershell start-process "C:\TestManager\TM1005.exe" -verb RunAs
@REM powershell Start-Process -FilePath "C:\TestManager\TM1005.exe" -ArgumentList "image_installation_driver_default.dll" -Verb RunAs
exit