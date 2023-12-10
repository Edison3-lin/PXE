. .\FunAll.ps1

    $file = Get-Item $PSCommandPath
    $Directory = Split-Path -Path $PSCommandPath -Parent
    $baseName = $file.BaseName
    $logfile = $Directory+'\'+$baseName+"_process.log"
    $tempfile = $Directory+'\temp.log'
    $outputfile = $Directory+'\'+$baseName+'_result.log'

    $MySqlCmd = "
            update  Test_Result 
            set     TR_Test_Result   = '$TestResult',
                    TR_Excute_Status = '$TestStatus' ,
                    TR_TestEndTime   = SYSDATETIME()
            where   TR_ID = '$TR_ID'

            update  Test_Control_Main 
            set     TCM_Status     = 'DONE',  
                    TCM_FinishDate = SYSDATETIME() 
            where   TCM_ID = '$TR_ID'
        "

    DATABASE "update" $MySqlCmd    

return 0