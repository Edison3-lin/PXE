. .\FunAll.ps1

    ### Create log file ###
    $file = Get-Item $PSCommandPath
    $Directory = Split-Path -Path $PSCommandPath -Parent
    $Directory += '\MyLog'
    $baseName = $file.BaseName
    $logfile = $Directory+'\'+$baseName+"_process.log"
    $outputfile = $Directory+'\'+$baseName+'_result.log'

 
    $remoteFile = "MyTestAll.dll"    
    # TestManager pass $remoteFile
    $attDir = $remoteFile.split('.')[0]
    $ftpPath = "/Test_Item/$attDir/"
    $localPath = "c:\\TestManager\\ItemDownload\\"
    if (-not (Test-Path -Path $localPath -PathType Container)) {
        New-Item -Path $localPath -ItemType Directory
    }

   # Download MD5 file
    try {
        FTP "$ftpServer/Test_Item/$attDir/Items.md5" down "$($localPath)Items.md5"
    }
    catch {
        process_log "Can't found out MD5 file"
        return $false
    }

    ### Download files ###
    $detailsList = FTP "$ftpServer/Test_Item/$attDir/" list 

    # Distinguish between files and directories
    foreach ($details in $detailsList) {
        # Parse the details and get the properties of each item
        $splitDetails = $details -split "\s+"
        $permissions = $splitDetails[0]
        $name = $splitDetails[-1]
        if( $name -eq "Items.md5" ) {
            continue
        }
        # Files or Directories?
        if ($permissions -like "d*") {
            process_log "Directory: $name"
        } else {
            try {
                process_log "$ftpServer/Test_Item/$attDir/$name -> $localPath$name"
                do {
                    FTP "$ftpServer/Test_Item/$attDir/$name" down "$localPath$name"
                    $f = Get-Item "$localPath$name"
                    $DownOK = CheckMD5 $f ".\\ItemDownload\\Items.md5"
                } while ( -not $DownOK )
            }
            catch {
                process_log "ERROR!!! <$name> download failed !!!"
            }

        }
    }

process_log  "======Download finished======"
return $true
