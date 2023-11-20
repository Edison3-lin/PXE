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


#Build connect object
$SqlConn = New-Object System.Data.SqlClient.SqlConnection
 
#Connect MSSQL
$SqlConn.ConnectionString = "Data Source=$DBserver;Initial Catalog=$Database;user id=$DBuserName;pwd=$DBpassword"
 
#Open DB
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
$programs = $NULL

$sqlCmd.CommandText = "
select 
  TCM.TCM_ID
  ,TCM.TCM_Status
  ,TR.TR_ID
  ,TAC.TAC_ID
  ,TAC.TAC_Table_Index
 , TAC.TAC_Table_Name
 ,TAC.TAC_Table_Coulmn
  from (
	select TOP (1) 
		B.*
	from DUT_Profile A, Test_Control_Main B
	where A.DP_UUID = '$UUID'
		and A.DP_UUID = B.DP_UUID
		and B.TCM_Name = 'Image Flash'
		and B.TCM_Status is null
	order by B.TCM_CreateDate desc
	) TCM 
	, Test_Result TR 
	, Test_Automation_Config TAC
 where TCM.TCM_ID = TR.TCM_ID
       and TR.TAC_ID = TAC.TAC_ID 
	   and TAC.TAC_Table_Index is not null
"

$adapter = New-Object System.Data.SqlClient.SqlDataAdapter $sqlCmd
$dataset = New-Object System.Data.DataSet
$NULL = $adapter.Fill($dataSet)

for ($i=0; $i -lt $dataSet.Tables[0].Rows.Count; $i++)
{
    # $TCM_Status = $dataSet.Tables[0].Rows[$i][1]
    $programs = @("common_bios_pxeboot_default.dll", "0")
}

#Close Database
$SqlConn.close()

return $programs
