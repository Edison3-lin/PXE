#配置資訊
$Database   = 'SIT_TEST'
$Server     = '"192.168.10.1"'
$UserName   = 'Captain001'
$Password   = 'Captaintest2023@SIT'


#建立連線對像
$SqlConn = New-Object System.Data.SqlClient.SqlConnection
 
#使用賬號連線MSSQL
$SqlConn.ConnectionString = "Data Source=$Server;Initial Catalog=$Database;user id=$UserName;pwd=$Password"
 
#打開數據庫連線
try {
    echo "Connecting to Database..."
    $SqlConn.open()
}
catch {
    echo "Cannot connect to Database"
    return "Unconnected"
}

$SqlCmd = New-Object System.Data.SqlClient.SqlCommand
$SqlCmd.connection = $SqlConn

$UUID = Get-WmiObject Win32_ComputerSystemProduct | Select-Object -ExpandProperty UUID 

echo "Query for Bios config..."

# Read SQL data
$sqlCmd.CommandText = "select * from BIOS_Update A where BU_ID = (
    select TAC.TAC_Table_Index
    FROM (
        select TOP (1) 
        B.*
        from DUT_Profile A, Test_Control_Main B
        where A.DP_UUID = '$UUID'
        and A.DP_UUID = B.DP_UUID
        and B.TCM_Name = 'BIOS Update'
        and B.TCM_Status = 'Running'
        order by B.TCM_ID desc
        )
    TCM, Test_Result TR, Test_Automation_Config TAC
    where TCM.TCM_ID = TR.TCM_ID
    and TR.TAC_ID = TAC.TAC_ID
    and TAC.TAC_Table_Index is not null
)
"

$adapter = New-Object System.Data.SqlClient.SqlDataAdapter $sqlCmd
$dataset = New-Object System.Data.DataSet
$NULL = $adapter.Fill($dataSet)
if($dataset.Tables[0].Rows.Count -eq 0)
{
    #print out a warning message of no data found
    echo "!!!<Warning>: No data found in database"
    return $false
}

$NULL = $adapter.Fill($dataSet)
$Bios = $dataset.Tables[0].Rows[0].BU_BIOS_Ver_To
$SqlConn.close()
$Target_Version = Split-Path $Bios -Leaf

$Current_Version = Get-WmiObject -Class Win32_BIOS | Select-Object -ExpandProperty SMBIOSBIOSVersion
$Current_Version = $Current_Version -replace 'V', '' -replace '\.'

if ($Target_Version -eq $Current_Version)
{
	echo "Bios update succeeded!"
	$true
} else {
	echo "Bios update failed!"
	$false
}

