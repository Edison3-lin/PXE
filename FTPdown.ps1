. .\FunAll.ps1

    ### Create log file ###
    $file = Get-Item $PSCommandPath
    $Directory = Split-Path -Path $PSCommandPath -Parent
    $Directory += '\MyLog'
    $baseName = $file.BaseName
    $logfile = $Directory+'\'+$baseName+"_process.log"
    $outputfile = $Directory+'\'+$baseName+'_result.log'

    # if (Test-Path -Path $args[1] -PathType Container) {
    #     process_log "ERROR!!! <$($args[1])> not a directory !!!"
    #     return 0
    # }    

    ### Download files ###
    # $detailsList = ftp "$ftpServer/Test_Item/" list 
    $detailsList = FTP "$ftpServer$($args[0])" list 

    # Distinguish between files and directories
    foreach ($details in $detailsList) {
        # Parse the details and get the properties of each item
        $splitDetails = $details -split "\s+"
        $permissions = $splitDetails[0]
        $name = $splitDetails[-1]

        # Files or Directories?
        if ($permissions -like "d*") {
            process_log "Directory: $name"
        } else {
            try {
                process_log "$ftpServer$($args[0])$name -> $($args[1])$name"
                FTP "$ftpServer$($args[0])$name" down "$($args[1])$name"
            }
            catch {
                process_log "ERROR!!! <$name> download failed !!!"
            }

        }
    }


