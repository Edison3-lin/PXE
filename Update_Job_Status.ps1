################## Build Log function ##########################
function process_log($log)
{
   $timestamp = Get-Date -Format "[yyyy-MM-dd HH:mm:ss] "
   $timestamp+$log | Add-Content $logfile
}
function result_log($log)
{
    Add-Content -Path $outputfile -Value (Get-Date -Format "[yyyy-MM-dd HH:mm:ss] ")
    $timestamp+$log | Add-Content $outputfile
}

$file = Get-Item $PSCommandPath
$Directory = Split-Path -Path $PSCommandPath -Parent
$baseName = $file.BaseName
$logfile = $Directory+'\'+$baseName+"_process.log"
$tempfile = $Directory+'\temp.log'
$outputfile = $Directory+'\'+$baseName+'_result.log'

Import-Module SQLPS -DisableNameChecking

#Config
$configPath = ".\Server.json"
$config = Get-Content -Raw -Path $configPath | ConvertFrom-Json
$Database   = $config.Database
$DBserver   = $config.DBserver 
$DBuserName   = $config.DBuserName
$DBpassword   = $config.DBpassword

# TR_Result.json
$TRPath = ".\TR_Result.json"
$TRconfig = Get-Content -Raw -Path $TRPath | ConvertFrom-Json
$TCM_ID = $TRconfig.TCM_ID 
$TR_ID = $TRconfig.TR_ID 
$TestResult = $TRconfig.TestResult 
$TestStatus = $TRconfig.TestStatus 
# $Text_Log_File_Path = $TRconfig.Text_Log_File_Path

# System.Data.SqlClient
# Add-Type -Path "c:\TestManager\ItemDownload\System.Data.dll"

$SqlConn = New-Object System.Data.SqlClient.SqlConnection
$SqlConn.ConnectionString = "Data Source=$DBserver;Initial Catalog=$Database;user id=$DBuserName;pwd=$DBpassword"
try {
    $SqlConn.open()
}
catch {
    process_log "!!!<Exception>: $($_.Exception.Message)"
    return "Unconnected_"
}

Write-Host $TestResult $TestStatus $TR_ID
# return

$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.connection = $SqlConn
# $UUID = Get-WmiObject Win32_ComputerSystemProduct | Select-Object -ExpandProperty UUID 


# Read SQL data
if($TestStatus -ne "PXE BOOT")
{
    $sqlCmd.CommandText = 
    "
        update Test_Result 
        set    TR_Test_Result   = '$TestResult' , 
               TR_Excute_Status = '$TestStatus' ,
               TR_TestEndTime   = SYSDATETIME()
        where  TR_ID = '$TR_ID'
    "
    $NULL = $SqlCmd.executenonquery()

    $TRconfig.TCM_ID = ""
    $TRconfig.TR_ID = ""
    $TRconfig.TestResult = "Pass"
    $TRconfig.TestStatus = "DONE"
    $updatedJson = $TRconfig | ConvertTo-Json -Depth 10
    $updatedJson | Set-Content -Path $TRPath

    # $sqlCmd.CommandText = 
    # "
    #     update Test_Control_Main 
    #     set    TCM_Status = 'Done',  
    #            TCM_FinishDate = SYSDATETIME() 
    #     where TCM_ID = '$TR_ID'
    # "
    # $NULL = $SqlCmd.executenonquery()
}

#Close DB
$SqlConn.close()
return 0