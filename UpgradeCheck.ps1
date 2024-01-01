. .\FunAll.ps1
function Get_Version() {
        $files = Get-ChildItem -Path ".\" -File
        foreach ($file in $files) {
            if($file.Name -like "TM????.exe")
            {
                $f = $file.Name.Substring(2, 4) 
                # process_log "TestManager ver. $f"
                return $f
            }
        }
        return $null    
    }
function TM_Version($version) {
        $ftpDirectory = "/Captain_Tool/TestManager/"
        $detailsList = FTP "$ftpServer$ftpDirectory" list 

        foreach ($details in $detailsList) {
            $splitDetails = $details -split "\s+"
            $permissions = $splitDetails[0]
            $name = $splitDetails[-1]

            # Files or Directories?
            if ($permissions -like "d*") {
                if([int]$name -gt [int]$version )
                {
                    process_log "FTP=$([int]$name) ,  Now=$([int]$version)"
                    return $true
                }
            }    
        }
        return $false
    }

    ### Create log file ###
    $file = Get-Item $PSCommandPath
    $Directory = Split-Path -Path $PSCommandPath -Parent
    $baseName = $file.BaseName
    $logfile = $Directory+'\'+$baseName+"_process.log"

    if ( Test-Path -Path "c:\\TestManager\\UT.ps1" -PathType Leaf ) {
        Remove-Item "c:\\TestManager\\UT.ps1"
        Remove-Item "c:\\TestManager\\UT_process.log"
    }

    $version = Get_Version

    try {
        $CheckUpdate =  TM_Version($version)
    }
    catch {
        Write-Host "Directory not exist?"
        return $false
    }

return $CheckUpdate
