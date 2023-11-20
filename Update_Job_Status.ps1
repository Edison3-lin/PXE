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

#Config
$configPath = ".\Server.json"
$config = Get-Content -Raw -Path $configPath | ConvertFrom-Json
$Database   = $config.Database
$DBserver   = $config.DBserver 
$DBuserName   = $config.DBuserName
$DBpassword   = $config.DBpassword

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

$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.connection = $SqlConn
$UUID = Get-WmiObject Win32_ComputerSystemProduct | Select-Object -ExpandProperty UUID 

# Read SQL data
if($args[0] -eq "common_bios_pxeboot_default.dll")
{

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

        $TCM_ID = $dataSet.Tables[0].Rows[$i][0]
        $TCM_Status = $dataSet.Tables[0].Rows[$i][1]
        # process_log $TCM_Status $TCM_ID
        $sqlCmd.CommandText = "
        update Test_Control_Main set TCM_Status = 'Done' where TCM_ID = $TCM_ID"
        $NULL = $SqlCmd.executenonquery()
    }
}
else
{
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
        $TR_ID = $dataSet.Tables[0].Rows[$i][3]
        if($TR_ID.ToString() -eq $args[0])
        {
            $TCM_Status = $dataSet.Tables[0].Rows[$i][1]
            $sqlCmd.CommandText = "
            update Test_Result set TR_Test_Result = '$($args[1])' , 
                                TR_Excute_Status = 'Done',
                                TR_TestEndTime = SYSDATETIME()
            where TR_ID = '$TR_ID'"
            $NULL = $SqlCmd.executenonquery()

            if( $TCM_Status -ne "Done")
            {
                $sqlCmd.CommandText = "update Test_Control_Main set TCM_Status = 'Done',  TCM_FinishDate = SYSDATETIME() where TCM_ID = '$TR_ID'"
                $NULL = $SqlCmd.executenonquery()
            }
        }    
    }
}

#Close DB
$SqlConn.close()
return 0