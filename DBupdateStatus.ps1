. .\FunAll.ps1

    $file = Get-Item $PSCommandPath
    $Directory = Split-Path -Path $PSCommandPath -Parent
    $Directory += '\MyLog'
    $baseName = $file.BaseName
    $logfile = $Directory+'\'+$baseName+"_process.log"

    # Check TR_Result.json data
    if( $NULL -eq $TRconfig.TCM_ID ) 
    {
        process_log "Not found 'TCM_ID' in TR.json"
        return $NULL
    }
    if( $NULL -eq $TRconfig.TR_ID ) 
    {
        process_log "Not found 'TR_ID' in TR.json"
        return $NULL
    }
    if( $NULL -eq $TRconfig.TestStatus ) 
    {
        process_log "Not found 'TestStatus' in TR.json"
        return $NULL
    }
    if( $NULL -eq $TRconfig.TestResult ) 
    {
        process_log "Not found 'TestResult' in TR.json"
        return $NULL
    }
    if( $NULL -eq $TRconfig.TCM_Done ) 
    {
        process_log "Not found 'TCM_Done' in TR.json"
        return $NULL
    }

    if ($null -eq $args[0])
    {
        # job finished
        if ($TRconfig.TestStatus -eq "Running")
        {
            if ( $($TRconfig.TestResult).ToUpper() -notin @("PASS","FAIL","ERROR","ABORT","WAIT","CONTINUE") )
            {
                $TRconfig.TestResult = "Error"
                $TestResult = "Error"
            }    
            $TRconfig.TestStatus = "Done"
            $updatedJson = $TRconfig | ConvertTo-Json -Depth 10
            $updatedJson | Set-Content -Path $TRPath
            $MySqlCmd = "
                    update  Test_Result 
                    set     TR_Test_Result   = '$($TRconfig.TestResult)',
                            TR_Excute_Status = 'Done',
                            TR_TestEndTime   = SYSDATETIME()
                    where   TR_ID = '$($TRconfig.TR_ID)'
            "
            process_log "DLL didn't update TR.json"
            process_log "TR_Excute_Status: Done ,TR_Excute_Status: Error"                    
            DATABASE "update" $MySqlCmd    

        }
        else 
        {
            $MySqlCmd = "
                    update  Test_Result 
                    set     TR_Test_Result   = '$($TRconfig.TestResult)',
                            TR_Excute_Status = '$($TRconfig.TestStatus)' ,
                            TR_TestEndTime   = SYSDATETIME()
                    where   TR_ID = '$($TRconfig.TR_ID)'
                    "
            process_log "TR_Test_Result/TR_Excute_Status by TR.json"                    
            DATABASE "update" $MySqlCmd    
        }
        # job finished

        $MySqlCmd = "
            update  Test_Control_Main 
            set     TCM_Status     = 'DONE',  
                    TCM_FinishDate = SYSDATETIME() 
            where   TCM_ID = '$($TRconfig.TCM_ID)'
            "
        process_log "TCM_Status: Done after Running"                    
        DATABASE "update" $MySqlCmd    
    }
    else
    {            
        # job didn't finish
        process_log "TestManager update TestStatus : $($args[0])"
        $TRconfig.TestStatus = $args[0]
        $updatedJson = $TRconfig | ConvertTo-Json -Depth 10
        $updatedJson | Set-Content -Path $TRPath

        $MySqlCmd = "
                update  Test_Result 
                set     TR_Excute_Status = '$($args[0])' ,
                        TR_TestStartTime   = SYSDATETIME()
                where   TR_ID = '$($TRconfig.TR_ID)'
                "
        process_log "TR_Excute_Status: Running"                    
        DATABASE "update" $MySqlCmd    
    }


    if( $TRconfig.TCM_Done ) {
        $MySqlCmd = "
                    update  Test_Control_Main 
                    set     TCM_Status     = 'DONE',  
                            TCM_FinishDate = SYSDATETIME() 
                    where   TCM_ID = '$($TRconfig.TCM_ID)'
                    "
        process_log "TCM_Status: TCM all Done"                    
        DATABASE "update" $MySqlCmd    
    }

return 0