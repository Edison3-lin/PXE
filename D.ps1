. .\FTP.ps1
. .\LOG.ps1
. .\JSON.ps1
. .\DATABASE.ps1

    $UUID = Get-WmiObject Win32_ComputerSystemProduct | Select-Object -ExpandProperty UUID 

    $MySqlCmd1 = "
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

    $MySqlCmd2 = 
    "
        update Test_Result 
        set    TR_Test_Result   = '$TestResult',
               TR_Excute_Status = '$TestStatus' ,
               TR_TestEndTime   = SYSDATETIME()
        where  TR_ID = 3
    "
    
    $dataSet = DATABASE "update" $MySqlCmd2
    
    # # if($dataSet -eq "Unconnected_")
    # if($dataSet -eq "xxxxxx")
    # {
        Write-Host $TestResult
    #     return 11
    # }
    # for ($i=0; $i -le 12; $i++)
    # {
    #     Write-Host ($dataSet.Tables[0].Rows[0][$i])
    # # $dataSet.Tables[0].Rows.Count
    # }

    return 0
