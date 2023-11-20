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
$baseName = $file.BaseName
$logfile = $Directory+'\'+$baseName+"_process.log"
$tempfile = $Directory+'\temp.log'
$outputfile = $Directory+'\'+$baseName+'_result.log'

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

$SqlConn = New-Object System.Data.SqlClient.SqlConnection
$SqlConn.ConnectionString = "Data Source=$DBserver;Initial Catalog=$Database;user id=$DBuserName;pwd=$DBpassword"
try {
    $SqlConn.open()
}
catch {
    process_log "!!!<Exception>: $($_.Exception.Message)"
    return "Unconnected_"
}

$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.connection = $SqlConn

$UUID = Get-WmiObject Win32_ComputerSystemProduct | Select-Object -ExpandProperty UUID 
$TR_IDs = $NULL
$programs = $NULL

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

for ($i=0; $i -lt $dataSet.Tables[0].Rows.Count; $i++)
{
    $Test_Result = $dataSet.Tables[0].Rows[$i][4]
    if( $Test_Result -ne "Done" )
    {
        $programs += $dataSet.Tables[0].Rows[$i][12]+','
        $TR_IDs += ($dataSet.Tables[0].Rows[$i][3]).ToString()+','
        $programs += ($dataSet.Tables[0].Rows[$i][3]).ToString()+','
        process_log $programs $TR_IDs
    }    
}

$SqlConn.close()

if( ($programs -eq $NULL) -or ($programs -eq "") )
{
    return $NULL
}
else
{
    $UUID = Get-WmiObject Win32_ComputerSystemProduct | Select-Object -ExpandProperty UUID 

    $directoryName = "/Test_Log/$UUID"
    CreateDir($directoryName)

    process_log $TR_IDs

    $temp = $TR_IDs -replace "[ ]", ""
    $TR_ID = $temp.split(',')
    $filtered = $TR_ID | Where-Object { $_ -ne "" }
    foreach ($subDir in $filtered) {
        $directoryName = "/Test_Log/$UUID/$subDir"
        process_log $directoryName
        CreateDir($directoryName)
    }
}

$temp = $programs -replace "[ ]", ""
$Program = $temp.split(',')
$filtered = $Program | Where-Object { $_ -ne "" }
return $filtered

