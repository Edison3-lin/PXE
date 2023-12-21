. .\FunAll.ps1

    ### Create log file ###
    $file = Get-Item $PSCommandPath
    $Directory = Split-Path -Path $PSCommandPath -Parent
    $Directory += '\MyLog'
    $baseName = $file.BaseName
    $logfile = $Directory+'\'+$baseName+"_process.log"
    $outputfile = $Directory+'\'+$baseName+'_result.log'

    # TestManager pass $remoteFile
    process_log "remoteFile: $remoteFile"
    $remoteFile = "common_image_pxeboot_default.dll"
    $attDir = $remoteFile.split('.')[0]
    process_log "attDir: $attDir"
    $ftpPath = "/Test_Item/$attDir/"
    $localPath = "c:\\TestManager\\ItemDownload\\"
    if (-not (Test-Path -Path $localPath -PathType Container)) {
        New-Item -Path $localPath -ItemType Directory
    }

   # Download MD5 file
    try {
        process_log "$ftpServer/Test_Item/$attDir/Items.md5 -> $($localPath)Items.md5"
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
        if( $name -eq "MD5.ps1" ) {
            continue
        }
        # Files or Directories?
        if ($permissions -like "d*") {
            process_log "Directory: $name"
        } else {
            try {
                process_log "$ftpServer/Test_Item/$attDir/$name -> $localPath$name"
                for( $i = 0; $i -lt 5; $i++) {
                    FTP "$ftpServer/Test_Item/$attDir/$name" down "$localPath$name"
                    $f = Get-Item "$localPath$name"
                    $DownOK = CheckMD5 $f ".\\ItemDownload\\Items.md5"
                    if($DownOK) {
                        process_log "  <$name> MD5 OK!"
                        break
                    }
                }
                if(-not $DownOK) {                
                    process_log "!!!!!! <$name> MD5 ERROR !!!!!!"
                    return $false
                }    
            }
            catch {
                process_log "ERROR!!! <$name> download failed !!!"
            }

        }
    }

process_log  "======MD5 veriry OK! Download finished======"
return $true
