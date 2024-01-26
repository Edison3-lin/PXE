. .\FunAll.ps1

    ### Create log file ###
    $file = Get-Item $PSCommandPath
    $Directory = Split-Path -Path $PSCommandPath -Parent
    $Directory += '\MyLog'
    $baseName = $file.BaseName
    $logfile = $Directory+'\'+$baseName+"_process.log"

    $UUID = Get-WmiObject Win32_ComputerSystemProduct | Select-Object -ExpandProperty UUID 
    $ExecuteDll = $NULL

    $MySqlCmd1 = "
            select 
                TCM.TCM_ID,
                TCM.TCM_Status,
                TR.TR_ID,
                TAC.TAC_ID,
                TAC.TAC_Table_Index,
                TAC.TAC_Table_Name,
                TAC.TAC_Table_Coulmn,
                TR.TR_Excute_Status,
                TCM.TCM_Name
            from (
                select TOP (1) 
                    B.*
                from DUT_Profile A, Test_Control_Main B
                where A.DP_UUID = '$UUID'
                    and A.DP_UUID = B.DP_UUID
                    and (B.TCM_Name = 'Image Flash' or B.TCM_Name = 'BIOS Update')
                    and (B.TCM_Status = 'Running')
                order by B.TCM_CreateDate desc
                ) TCM 
                , Test_Result TR 
                , Test_Automation_Config TAC
            where TCM.TCM_ID = TR.TCM_ID
                and TR.TAC_ID = TAC.TAC_ID 
                and TAC.TAC_Table_Index is not null
        "
    $dataSet = DATABASE "read" $MySqlCmd1
        
    if ( $dataSet.Tables[0].Rows.Count -eq 0 )
    {
        $MySqlCmd2 = "
            select 
                TCM.TCM_ID,
                TCM.TCM_Status,
                TR.TR_ID,
                TAC.TAC_ID,
                TAC.TAC_Table_Index,
                TAC.TAC_Table_Name,
                TAC.TAC_Table_Coulmn,
                TR.TR_Excute_Status,
                TCM.TCM_Name
            from (
                select TOP (1) 
                    B.*
                from DUT_Profile A, Test_Control_Main B
                where A.DP_UUID = '$UUID'
                    and A.DP_UUID = B.DP_UUID
                    and (B.TCM_Name = 'Image Flash' or B.TCM_Name = 'BIOS Update')
                    and (B.TCM_Status is null)
                order by B.TCM_CreateDate desc
                ) TCM 
                , Test_Result TR 
                , Test_Automation_Config TAC
            where TCM.TCM_ID = TR.TCM_ID
                and TR.TAC_ID = TAC.TAC_ID 
                and TAC.TAC_Table_Index is not null
        "
        $dataSet = DATABASE "read" $MySqlCmd2
    }

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

        if(($dataSet.Tables[0].Rows[$i][8]) -eq "BIOS Update") {
            $ExecuteDll = "common_bios_pxeboot_default.dll"
        }
        else {
            $ExecuteDll = "common_image_pxeboot_default.dll"
        }
        process_log "Got DB job: $ExecuteDll"

        # First time PXE boot (NULL)
        if ( '' -eq $TCM_Status )
        {
            $TRconfig.TestStatus = "New"
            $TRconfig.TestResult = ""
            $updatedJson = $TRconfig | ConvertTo-Json -Depth 10
            $updatedJson | Set-Content -Path $TRPath

            $MySqlCmd = "
                    update Test_Result 
                    set    TR_Excute_Status = 'Running',
                           TR_TestStartTime   = SYSDATETIME()
                    where  TR_ID = '$TR_ID'

                    update Test_Control_Main 
                    set    TCM_Status = 'Running',
                           TCM_CreateDate = SYSDATETIME()
                    where  TCM_ID = '$TCM_ID'
                "
            process_log "TCM_Status/TR_Excute_Status: Running"
            DATABASE "update" $MySqlCmd    
        }
        # Not the first time PXE (Running)
# (EdisonLin-20240110-1)         elseif ( $TR_Excute_Status -eq "Running" ) 
        elseif ( ($TRconfig.TestStatus).ToUpper() -eq "DONE" ) 
        {
            $directoryPath = "c:\\TestManager\\ItemDownload"
            $items = Get-ChildItem -Path $directoryPath
            # check common_bios_pxeboot_default.dll exist?
            if ($items.Length -eq 0) 
            {
                if(($dataSet.Tables[0].Rows[$i][8]) -eq "BIOS Update") {
                    $ExecuteDll = "common_bios_pxeboot_default.dll"
                }
                else {
                    $ExecuteDll = "common_image_pxeboot_default.dll"
                }
            } 

            $MySqlCmd = "
                update Test_Control_Main 
                set    TCM_Status = 'DONE',
                       TCM_FinishDate = SYSDATETIME() 
                where  TCM_ID = '$TCM_ID'

                update Test_Result 
                set    TR_Excute_Status = '$($TRconfig.TestStatus)',
                       TR_Test_Result = '$($TRconfig.TestResult)',
                       TR_TestEndTime   = SYSDATETIME()
                    where  TR_ID = '$TR_ID'
            "
            process_log "job finish, update all status by TR.json"
            DATABASE "update" $MySqlCmd    
            $ExecuteDll = $NULL

        }
    }

return $ExecuteDll
