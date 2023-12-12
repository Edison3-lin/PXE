. .\FunAll.ps1

    $file = Get-Item $PSCommandPath
    $Directory = Split-Path -Path $PSCommandPath -Parent
    $Directory += '\MyLog'
    $baseName = $file.BaseName
    $logfile = $Directory+'\'+$baseName+"_process.log"
    $tempfile = $Directory+'\temp.log'
    $outputfile = $Directory+'\'+$baseName+'_result.log'

    # local and FTP directory
    $UUID = Get-WmiObject Win32_ComputerSystemProduct | Select-Object -ExpandProperty UUID 
    $remoteFilePath = "/Test_Log/$UUID/$TCM_ID/$TR_ID/"

    foreach ($log in $Text_Log_File_Path) {
        $f = Get-Item $log
        $file = $f.Name
        $destinationFilePath = $remoteFilePath+$file
        try {
            process_log "$sourceFilePath to $ftpServer$destinationFilePath"
            FTP "$ftpServer$destinationFilePath" up "$log"
        }
        catch {
            process_log "ERROR!!! <$log> upload failed !!!"
            return $false
        }
        process_log "<$log> upload to $ftpServer$destinationFilePath"
    }

    process_log  "======Upload finished======"

return $true
