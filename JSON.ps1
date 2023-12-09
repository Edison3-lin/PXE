    # Server.json
    $configPath = ".\Server.json"
    $config = Get-Content -Raw -Path $configPath | ConvertFrom-Json
    $ftpServer  = $config.ftpServer
    $username   = $config.username
    $password   = $config.password
    $Database   = $config.Database
    $DBserver   = $config.DBserver 
    $DBuserName = $config.DBuserName
    $DBpassword = $config.DBpassword
    
    # TR_Result.json
    $TRPath = ".\TR_Result.json"
    $TRconfig = Get-Content -Raw -Path $TRPath | ConvertFrom-Json
    $TCM_ID     = $TRconfig.TCM_ID 
    $TR_ID      = $TRconfig.TR_ID 
    $TestResult = $TRconfig.$TestResult
    $TestStatus = $TRconfig.$TestStatus
    $Text_Log_File_Path = $TRconfig.Text_Log_File_Path
    $Test_TimeOut       = $TRconfig.Test_TimeOut
