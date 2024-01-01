. .\FunAll.ps1

    $file = Get-Item $PSCommandPath
    $Directory = Split-Path -Path $PSCommandPath -Parent
    $Directory += '\MyLog'
    $baseName = $file.BaseName
    $logfile = $Directory+'\'+$baseName+"_process.log"
    $tempfile = $Directory+'\temp.log'
    $outputfile = $Directory+'\'+$baseName+'_result.log'

    if ($null -eq $args[0])
    {
        if ($TestStatus -eq "Running")
        {
            process_log "Test Item didn't update TR_Test_Result,TR_Excute_Status to Json !!!!!"
            if ( $TestResult -notin @("Pass","Fail","Error","Abort","Wait","Continue") )
            {
                $TRconfig.TestResult = "Error"
                $TestResult = $TRconfig.TestResult
            }    
            $TRconfig.TestStatus = "Done"
            $updatedJson = $TRconfig | ConvertTo-Json -Depth 10
            $updatedJson | Set-Content -Path $TRPath
            $MySqlCmd = "
                    update  Test_Result 
                    set     TR_Test_Result   = '$TestResult',
                            TR_Excute_Status = 'Done',
                            TR_TestEndTime   = SYSDATETIME()
                    where   TR_ID = '$TR_ID'
            "
        }
        else 
        {
            process_log "Update TR_Test_Result,TR_Excute_Status by Json: $TestResult $TestStatus"
            $MySqlCmd = "
                    update  Test_Result 
                    set     TR_Test_Result   = '$TestResult',
                            TR_Excute_Status = '$TestStatus' ,
                            TR_TestEndTime   = SYSDATETIME()
                    where   TR_ID = '$TR_ID'
                    "
                #     update  Test_Control_Main 
                #     set     TCM_Status     = 'DONE',  
                #             TCM_FinishDate = SYSDATETIME() 
                #     where   TCM_ID = '$TR_ID'
        }
    }
    else
    {            
        process_log "Update TR_Excute_Status by Args[0]: $($args[0])"
        $TRconfig.TestStatus = $args[0]
        $updatedJson = $TRconfig | ConvertTo-Json -Depth 10
        $updatedJson | Set-Content -Path $TRPath

        $MySqlCmd = "
                update  Test_Result 
                set     TR_Excute_Status = '$($args[0])' ,
                        TR_TestStartTime   = SYSDATETIME()
                where   TR_ID = '$TR_ID'
                "
    }

    DATABASE "update" $MySqlCmd    

    if( $TRconfig.TCM_Done ) {
        $MySqlCmd = "
                    update  Test_Control_Main 
                    set     TCM_Status     = 'DONE',  
                            TCM_FinishDate = SYSDATETIME() 
                    where   TCM_ID = '$TCM_ID'
                    "
        DATABASE "update" $MySqlCmd    
    }

return 0