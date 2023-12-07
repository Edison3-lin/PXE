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
$Directory += '\MyLog'
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

$TRPath = ".\TR_Result.json"
$TRconfig = Get-Content -Raw -Path $TRPath | ConvertFrom-Json

#Build connect object
$SqlConn = New-Object System.Data.SqlClient.SqlConnection
 
#Connect MSSQL
$SqlConn.ConnectionString = "Data Source=$DBserver;Initial Catalog=$Database;user id=$DBuserName;pwd=$DBpassword"
 
#Open DB
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
  ,TR.TR_Excute_Status
  from (
	select TOP (1) 
		B.*
	from DUT_Profile A, Test_Control_Main B
	where A.DP_UUID = '$UUID'
		and A.DP_UUID = B.DP_UUID
		and B.TCM_Name = 'Image Flash'
		and (B.TCM_Status is null or B.TCM_Status = 'Running')
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
    $TCM_ID = ($dataSet.Tables[0].Rows[$i][0])
    $TCM_Status = ($dataSet.Tables[0].Rows[$i][1])
    $TR_ID = ($dataSet.Tables[0].Rows[$i][2])
    $TR_Excute_Status = ($dataSet.Tables[0].Rows[$i][7])
    $TRconfig.TCM_ID = $TCM_ID
    $TRconfig.TR_ID = $TR_ID
    $updatedJson = $TRconfig | ConvertTo-Json -Depth 10
    $updatedJson | Set-Content -Path $TRPath

    # $TCM_Status = $dataSet.Tables[0].Rows[$i][1]
    $programs = "common_bios_pxeboot_default.dll"

    # $TR_Excute_Status = $NULL
    # 第一次PXE boot (NULL)
    if ( '' -eq $TCM_Status )
    {
        $TRconfig.TestStatus = "New"
        $updatedJson = $TRconfig | ConvertTo-Json -Depth 10
        $updatedJson | Set-Content -Path $TRPath

        $sqlCmd.CommandText = 
        "update Test_Result 
        set    TR_Excute_Status = 'Running'
        where  TR_ID = '$TR_ID'"
        $NULL = $SqlCmd.executenonquery()

        $sqlCmd.CommandText = 
        "update Test_Control_Main 
        set    TCM_Status = 'Running'
        where  TCM_ID = '$TCM_ID'"
        $NULL = $SqlCmd.executenonquery()
    }
    # 非第一次跑PXE (Running)
    elseif ( $TR_Excute_Status -eq "Running" ) 
    {
        $sqlCmd.CommandText = 
        "update Test_Control_Main 
        set    TCM_Status = 'DONE'
        where  TCM_ID = '$TCM_ID'"
        $NULL = $SqlCmd.executenonquery()
        # image patch 有寫"Done"到TR_Resulte.json
        if( $TRconfig.TestStatus -eq "DONE" )
        {
            $sqlCmd.CommandText = 
            "update Test_Result 
            set    TR_Excute_Status = 'DONE',
                   TR_Test_Result = 'Pass'
            where  TR_ID = '$TR_ID'"
            $NULL = $SqlCmd.executenonquery()
            $TRconfig.TestResult = "Pass"
        }
        else 
        {
            # image 沒有完成
            $sqlCmd.CommandText = 
            "update Test_Result 
            set    TR_Excute_Status = 'DONE',
                   TR_Test_Result = 'Fail'
            where  TR_ID = '$TR_ID'"
            $NULL = $SqlCmd.executenonquery()
            $TRconfig.TestResult = "Fail"
        }  
        $TRconfig.TestStatus = "DONE"
        $updatedJson = $TRconfig | ConvertTo-Json -Depth 10
        $updatedJson | Set-Content -Path $TRPath
        $programs = $NULL
    }
}

#Close Database
$SqlConn.close()

return $programs
