. .\FTP.ps1
. .\LOG.ps1
. .\JSON.ps1


    ### Create log file ###
    $file = Get-Item $PSCommandPath
    $Directory = Split-Path -Path $PSCommandPath -Parent
    $Directory += '\MyLog'
    $baseName = $file.BaseName
    $logfile = $Directory+'\'+$baseName+"_process.log"
    $outputfile = $Directory+'\'+$baseName+'_result.log'

    #Build connect object
    $SqlConn = New-Object System.Data.SqlClient.SqlConnection
    $SqlConn.ConnectionString = "Data Source=$DBserver;Initial Catalog=$Database;user id=$DBuserName;pwd=$DBpassword"
 
    # Try to open the connection, wait up to 30 seconds
    $timeout = 30
    $timer = [System.Diagnostics.Stopwatch]::StartNew()

    while ($SqlConn.State -ne 'Open' -and $timer.Elapsed.TotalSeconds -lt $timeout) {
        try {
            # Open connection
            $SqlConn.Open()
            Start-Sleep -Seconds 1
        } catch {
            # If the connection fails to open, catch the exception and continue waiting.
            # process_log "Error opening connection: $_"
        }
    }

    $timer.Stop()

    # Check connection status
    if ($SqlConn.State -eq 'Open') {
        process_log "Connection opened successfully!"
    } else {
        process_log "Connection failed to open within the specified timeout."
        return "Unconnected_"
    }

    # #Open DB
    # try {
    #     $SqlConn.open()
    # }
    # catch {
    #     process_log "Waiting 5 sec for DB connected !!!"
    #     return "Unconnected_"
    # }

    # process_log $SqlConn.State

    $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
    $SqlCmd.connection = $SqlConn

    $UUID = Get-WmiObject Win32_ComputerSystemProduct | Select-Object -ExpandProperty UUID 
    $programs = $NULL

    $sqlCmd.CommandText = "
        select 
            TCM.TCM_ID,
            TCM.TCM_Status,
            TR.TR_ID,
            TAC.TAC_ID,
            TAC.TAC_Table_Index,
            TAC.TAC_Table_Name,
            TAC.TAC_Table_Coulmn,
            TR.TR_Excute_Status
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
        # First time PXE boot (NULL)
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
        # Not the first time PXE (Running)
        elseif ( $TR_Excute_Status -eq "Running" ) 
        {
            $sqlCmd.CommandText = 
            "update Test_Control_Main 
            set    TCM_Status = 'DONE'
            where  TCM_ID = '$TCM_ID'"
            $NULL = $SqlCmd.executenonquery()
            # image patch write "Done" to TR_Resulte.json
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
                # 'Image Flash' not finish
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
