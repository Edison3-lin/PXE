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

function CreateDir($directoryName)
{
    $ftpPath = $ftpServer + $directoryName
    $ftpRequest = [System.Net.FtpWebRequest]::Create($ftpPath)
    $ftpRequest.Credentials = New-Object System.Net.NetworkCredential($username, $password)
    $ftpRequest.Method = [System.Net.WebRequestMethods+Ftp]::MakeDirectory
    # process_log $ftpPath
    try {
        $ftpResponse = $ftpRequest.GetResponse()
    } catch {
        Write-Host "ftpResponse error!!"
    } finally {
        if ($ftpResponse -ne $null) {
            $ftpResponse.Close()
        }
    }
}

$file = Get-Item $PSCommandPath
$Directory = Split-Path -Path $PSCommandPath -Parent
$Directory += '\MyLog'
$baseName = $file.BaseName
$logfile = $Directory+'\'+$baseName+"_process.log"
$tempfile = $Directory+'\temp.log'
$outputfile = $Directory+'\'+$baseName+'_result.log'

Import-Module SQLPS -DisableNameChecking

#Config
$configPath = ".\Server.json"
$config = Get-Content -Raw -Path $configPath | ConvertFrom-Json

$ftpServer = $config.ftpServer
$username = $config.username
$password = $config.password

$Database   = $config.Database
$DBserver   = $config.DBserver 
$DBuserName   = $config.DBuserName
$DBpassword   = $config.DBpassword

$TRPath = ".\TR_Result.json"
$TRconfig = Get-Content -Raw -Path $TRPath | ConvertFrom-Json


$SqlConn = New-Object System.Data.SqlClient.SqlConnection
$SqlConn.ConnectionString = "Data Source=$DBserver;Initial Catalog=$Database;user id=$DBuserName;pwd=$DBpassword"
try {
    $SqlConn.open()
}
catch {
    process_log "Waiting 5 sec for DB connected !!!"
    return "Unconnected_"
}

$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.connection = $SqlConn

$UUID = Get-WmiObject Win32_ComputerSystemProduct | Select-Object -ExpandProperty UUID 
$TCM_ID = $NULL
$TR_ID = $NULL
$TA_Execute_Path = $NULL

# Read SQL data
$sqlCmd.CommandText = "
select 
    A.TCM_ID 'Control_ID'
    ,A.TCM_Status 'Control_Test_Status'
    ,A.TCM_Name 'Control_Test_Task_Name'
    ,B.TR_ID 'Test_Result_ID'
    ,B.TR_Test_Result 'Test_Result'
    ,UPPER(B.TR_Excute_Status) 'Test_Test_Status'
    ,B.TR_TestStartTime 'Test_Start_Time'
    ,C.TMD_Name 'Test_Name'
    ,C.TMD_Desc 'Test_Description'
    ,C.TMD_TimeOut_Second 'Test_TimeOut'
    ,D.TA_Desc 'Automation_Tool_Desc'
    ,D.TA_Name 'Automation_Tool_Name'
    ,D.TA_Execute_Path 'Execute_Path'
from Test_Control_Main A
    ,Test_Result B
    ,Test_Measurements_Def C
    ,Test_Automation D
    , Test_Phase_Def E
where 
A.DP_UUID = '$UUID' and
(UPPER(A.TCM_Status) != 'DROP' and UPPER(A.TCM_Status) != 'DONE' and UPPER(A.TCM_Status) != 'OK' 
and  UPPER(A.TCM_Status) != 'ERROR' or A.TCM_Status is null)
            and B.TCM_ID = A.TCM_ID
and A.TCM_FinishDate is null
and (UPPER(B.TR_Excute_Status) != 'DONE' or B.TR_Excute_Status is null)
    and B.TMD_ID = C.TMD_ID
    and C.TA_ID = D.TA_ID
    and C.TPD_ID = E.TPD_ID
order by A.TCM_ID
    ,E.TPD_Priority
    ,C.TMD_Priority
"
$adapter = New-Object System.Data.SqlClient.SqlDataAdapter $sqlCmd
$dataset = New-Object System.Data.DataSet
$NULL = $adapter.Fill($dataSet)
# if($dataSet.Tables[0].Rows.Count -eq 0)
# {
#     $TRconfig.TCM_ID = ""
#     $TRconfig.TR_ID = ""
#     $updatedJson = $TRconfig | ConvertTo-Json -Depth 10
#     $updatedJson | Set-Content -Path $TRPath
#     $SqlConn.close()
#     return $NULL
# }
for ($i=0; $i -lt $dataSet.Tables[0].Rows.Count; $i++)
{
    $Test_Result = $dataSet.Tables[0].Rows[$i][4]
    if( $Test_Result -ne "Done" )
    {
        $TCM_ID = ($dataSet.Tables[0].Rows[$i][0])
        $TR_ID = ($dataSet.Tables[0].Rows[$i][3])
        $TA_Execute_Path = $dataSet.Tables[0].Rows[$i][12]
        $TRconfig.TCM_ID = $TCM_ID
        $TRconfig.TR_ID = $TR_ID
        $updatedJson = $TRconfig | ConvertTo-Json -Depth 10
        $updatedJson | Set-Content -Path $TRPath
        break
    }    
}
$SqlConn.close()
if($dataSet.Tables[0].Rows.Count -ne 0)
{
    $directoryName = "/Test_Log/$UUID"
    CreateDir($directoryName)
    $directoryName = "/Test_Log/$UUID/$TCM_ID"
    CreateDir($directoryName)
    $directoryName = "/Test_Log/$UUID/$TCM_ID/$TR_ID"
    CreateDir($directoryName)
    return $TA_Execute_Path
}    

return $NULL
