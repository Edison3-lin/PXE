function MakeMD5() {
        # Get all files 
        $files = Get-ChildItem -Path ".\" -File
        $MD5 = ""
        $MD5file = ".\\Items.md5"
        if (Test-Path -Path $MD5file -PathType Leaf) {
            Remove-Item $MD5file
        }
        foreach ($file in $files) {
            if($file.Name -eq "Items.md5") {
                continue
            }
            if($file.Name -eq "MD5.ps1") {
                continue
            }
            $MD5 = Get-FileHash -Path $file.FullName -Algorithm MD5 | Select-Object -ExpandProperty Hash
            $MD5 | Add-Content $MD5file
        }
    }

function CheckMD5($MD5file) {
        $files = Get-ChildItem -Path ".\" -File
        foreach ($file in $files) {
            if($file.Name -eq "Items.md5") {
                continue
            }
            if($file.Name -eq "1.ps1") {
                continue
            }
            $MD5s = Get-Content -Path $MD5file
            $MD5 = Get-FileHash -Path $file.FullName -Algorithm MD5 | Select-Object -ExpandProperty Hash
            if($MD5 -notin $MD5s) {
                Write-Host $file.Name
                return $false
            }
        }
        return $true
    }    

    MakeMD5
    # $a = CheckMD5(".\Items.md5")
    # Write-Host "$($a)Items.md5"
