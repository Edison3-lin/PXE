@REM 請使用Administrator環境執行
c:\TestManager\TestManager\TMservice\bin\Debug\InstallUtil.exe c:\TestManager\TestManager\TMservice\bin\Debug\TMservice.exe
 
@REM 卸載 service 再安裝修改過的 service
sc stop TMservice
c:\TestManager\TestManager\TMservice\bin\Debug\InstallUtil.exe /u c:\TestManager\TestManager\TMservice\bin\Debug\TMservice.exe