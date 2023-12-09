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

    ### Download files ###
    # $detailsList = ftp "$ftpServer/Test_Item/" list 
    $detailsList = ftp "$ftpServer/$($args[0])" list 

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
            process_log "$ftpServer/$($args[0])/$name -> $($args[1])\$name"
            ftp "$ftpServer/$($args[0])/$name" down "$($args[1])\$name"
        }
    }


