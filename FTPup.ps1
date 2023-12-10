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

    ### Upload files ###
    $files = Get-ChildItem -Path $args[1]

    # Distinguish between files and directories
    foreach ($file in $files) {
        if ($file.PSIsContainer) {
            process_log "Directory: $($file.FullName)"
        } else {
            process_log "$($args[1])\$file -> $ftpServer/$($args[0])/$file"
            FTP "$ftpServer/$($args[0])/$file" up "$($args[1])\$file"
        }
    }
