Function ftp($ftpurl,$do,$filename,$DownPath) { 
    # ftp 伺服器位址，使用者名，密碼，操作（上傳up/下載down/清單list），檔案名，下載路徑
    # 示例：ftp ftp://10.10.98.91/ up C:\Windows\setupact.txt
        if ($do -eq "up")
        {
            $fileinf=New-Object System.Io.FileInfo("$filename")
            $upFTP = [system.net.ftpwebrequest] [system.net.webrequest]::create("$ftpurl"+$fileinf.name)
            $upFTP.Credentials = New-Object System.Net.NetworkCredential("$username","$password")
            $upFTP.Method=[system.net.WebRequestMethods+ftp]::UploadFile
            $upFTP.KeepAlive=$false
            $sourceStream = New-Object System.Io.StreamReader($fileInf.fullname)
            $fileContents = [System.Text.Encoding]::UTF8.GetBytes($sourceStream.ReadToEnd())
            $sourceStream.Close();
            $upFTP.ContentLength = $fileContents.Length;
            $requestStream = $upFTP.GetRequestStream();
            $requestStream.Write($fileContents, 0, $fileContents.Length);
            $requestStream.Close();
            $response =$upFTP.GetResponse();
            $response.StatusDescription
            $response.Close();
        }
        if ($do -eq "down")
        {
            $downFTP = [system.net.ftpwebrequest] [system.net.webrequest]::create("$ftpurl"+"$filename")
            $downFTP.Credentials = New-Object System.Net.NetworkCredential("$username","$password")
            $response = $downFTP.getresponse()
            $stream=$response.getresponsestream()
            $buffer = new-object System.Byte[] 2048
            $outputStream=New-Object System.Io.FileStream("$DownPath","Create")
            while(($readCount = $stream.Read($buffer, 0, 1024)) -gt 0){
                $outputStream.Write($buffer, 0, $readCount)
            }
            $outputStream.Close()
            $stream.Close()
            $response.Close() 
            if(Test-Path  $DownPath){echo "DownLoad successful"}
        }
        if ($do -eq "list")
        {
            $listFTP = [system.net.ftpwebrequest] [system.net.webrequest]::create("$ftpurl")
            $listFTP.Credentials = New-Object System.Net.NetworkCredential("$username","$password")
            # $listFTP.Method=[system.net.WebRequestMethods+ftp]::listdirectorydetails
            $listFTP.Method=[system.net.WebRequestMethods+ftp]::listdirectory
            $response = $listFTP.getresponse()
            $stream = New-Object System.Io.StreamReader($response.getresponsestream(),[System.Text.Encoding]::UTF8)
            $files = @()
            while(-not $stream.EndOfStream){
                $files += $stream.ReadLine()
            }
            $stream.Close()
            $response.Close()     
            return $files
        }
    }

    $configPath = ".\Server.json"
    $config = Get-Content -Raw -Path $configPath | ConvertFrom-Json
    $ftpServer    = $config.ftpServer
    $username  = $config.username
    $password  = $config.password

    # ftp $ftpServer down "/Test_Item/TestCase.exe" "C:\TestManager\TestCase.exe" 
    $f = ftp "$ftpServer/Test_Item/" list 

    Write-Host $f



