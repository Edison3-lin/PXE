. .\FTP.ps1
. .\LOG.ps1
. .\JSON.ps1

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

    # Create a WebClient object and set credentials
    $webClient = New-Object System.Net.WebClient
    $webClient.Credentials = New-Object System.Net.NetworkCredential($username, $password)

    foreach ($log in $Text_Log_File_Path) {
        $f = Get-Item $log
        $file = $f.Name
        $sourceFilePath = $log
        $destinationFilePath = $remoteFilePath+$file
        try {
            process_log "$sourceFilePath to $ftpServer$destinationFilePath"
            $webClient.UploadFile("$ftpServer$destinationFilePath", $sourceFilePath)
        }
        catch {
            process_log "!!!<$f>: $($_.Exception.Message)"
        }
        process_log "<$f> upload to $localFilePath"
    }

    process_log  "======Upload finished======"
    # Release WebClient
    $webClient.Dispose()

return 0
