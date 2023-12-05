@REM 請使用Administrator環境執行
 .\InstallUtil.exe .\TMservice.exe
 
@REM 卸載 service 再安裝修改過的 service
sc stop TMservice
.\InstallUtil.exe /u .\TMservice.exe
.\InstallUtil.exe .\TMservice.exe