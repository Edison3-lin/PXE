. .\FunAll.ps1

    ### Create log file ###
    $file = Get-Item $PSCommandPath
    $Directory = Split-Path -Path $PSCommandPath -Parent
    $Directory += '\MyLog'
    $baseName = $file.BaseName
    $logfile = $Directory+'\'+$baseName+"_process.log"
    $outputfile = $Directory+'\'+$baseName+'_result.log'

    # Import-Module SQLPS -DisableNameChecking

    $UUID = Get-WmiObject Win32_ComputerSystemProduct | Select-Object -ExpandProperty UUID 
    $TCM_ID = $NULL
    $TR_ID = $NULL
    $TA_Execute_Path = $NULL

    $MySqlCmd = "
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
            and UPPER(A.TCM_Status) != 'ERROR' or A.TCM_Status is null)
            and B.TCM_ID = A.TCM_ID
            and A.TCM_FinishDate is null
            and (UPPER(B.TR_Excute_Status) != 'DONE' or B.TR_Excute_Status is null)
            and (UPPER(B.TR_Excute_Status) != 'DROP' or B.TR_Excute_Status is null)
            and (UPPER(B.TR_Excute_Status) != 'ABORT' or B.TR_Excute_Status is null)
            and B.TMD_ID = C.TMD_ID
            and C.TA_ID = D.TA_ID
            and C.TPD_ID = E.TPD_ID
            order by A.TCM_ID
                ,E.TPD_Priority
                ,C.TMD_Priority
    "

    $dataSet = DATABASE "read" $MySqlCmd

    # Write-Host $dataSet.Tables[0].Rows.Count
    $TCM_list = New-Object System.Collections.Generic.List[Object]
    for ($i=0; $i -lt $dataSet.Tables[0].Rows.Count; $i++) 
    {
        $TCM_list.Add($dataSet.Tables[0].Rows[$i][0])
    }    

    if($TCM_list -eq '')
    {
        process_log "Got TCM_list: $TCM_list"
    }

    if($dataSet -ne "Unconnected_") 
    {
       for ($i=0; $i -lt $dataSet.Tables[0].Rows.Count; $i++) 
       {
#//(EdisonLin-20240111-1)            $Test_Result = $dataSet.Tables[0].Rows[$i][4]
           $Test_Status = $dataSet.Tables[0].Rows[$i][5]
           if( '' -eq $Test_Status ) 
           {
               $TRconfig.TestStatus = "New"
               $TRconfig.TestResult = ""
               $TCM_ID = ($dataSet.Tables[0].Rows[$i][0])

               # All TCM_ID done? 
               $TCM_list.Remove($TCM_ID)

               if( $TCM_list.Count -eq 0 ){
                   $TRconfig.TCM_Done = $true
               }
               elseif ($TCM_ID -notin $TCM_list) {
                   $TRconfig.TCM_Done = $true
               }
               else {
                   $TRconfig.TCM_Done = $false
               }

               $TR_ID = ($dataSet.Tables[0].Rows[$i][3])
               $TA_Execute_Path = $dataSet.Tables[0].Rows[$i][12]
               $TRconfig.TCM_ID = $TCM_ID
               $TRconfig.TR_ID = $TR_ID
               $TRconfig.Test_TimeOut = $dataSet.Tables[0].Rows[$i][9]
               $updatedJson = $TRconfig | ConvertTo-Json -Depth 10
               $updatedJson | Set-Content -Path $TRPath
               process_log "Update TR.json for new job"
               break
           }    
       }
    }
    else {
        process_log "Unconnected_"
        return "Unconnected_"
    }

    # If get any job, Create LOG directory on FTP
    if($dataSet.Tables[0].Rows.Count -ne 0)
    {
        process_log "Create FTP log directory: /Test_Log/$UUID, /Test_Log/$UUID/$TCM_ID, /Test_Log/$UUID/$TCM_ID/$TR_ID"
        $directoryName = "/Test_Log/$UUID"
        CreateDir($directoryName)
        $directoryName = "/Test_Log/$UUID/$TCM_ID"
        CreateDir($directoryName)
        $directoryName = "/Test_Log/$UUID/$TCM_ID/$TR_ID"
        CreateDir($directoryName)
        $directoryName = "/Test_Log/$UUID/$TCM_ID/$TR_ID/ScreenShot"
        CreateDir($directoryName)
        $directoryName = "/Test_Log/$UUID/$TCM_ID/$TR_ID/Test_Log"
        CreateDir($directoryName)
        process_log $TA_Execute_Path
        return $TA_Execute_Path
    }    

return $NULL
