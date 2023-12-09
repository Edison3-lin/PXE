Function ftp($ftpurl,$do,$filename) { 
    # ftp 伺服器位址，使用者名，密碼，操作（上傳up/下載down/清單list），檔案名，下載路徑
    # 示例：ftp ftp://10.10.98.91/ up C:\Windows\setupact.txt
        if ($do -eq "up")
        {
            $ftpRequest = [System.Net.WebRequest]::Create("$ftpurl")
            $ftpRequest.Method = [System.Net.WebRequestMethods+Ftp]::UploadFile
            $ftpRequest.Credentials = New-Object System.Net.NetworkCredential($username, $password)
            $fileBytes = [System.IO.File]::ReadAllBytes($filename)
            $requestStream = $ftpRequest.GetRequestStream()
            $requestStream.Write($fileBytes, 0, $fileBytes.Length)
            $requestStream.Close()
            $response = $ftpRequest.GetResponse()
            $response.Close()
        }    
        if ($do -eq "down")
        {
            $downFTP = [system.net.ftpwebrequest] [system.net.webrequest]::create($ftpurl)
            $downFTP.Credentials = New-Object System.Net.NetworkCredential("$username","$password")
            $response = $downFTP.getresponse()
            $stream=$response.getresponsestream()
            $buffer = new-object System.Byte[] 2048
            $outputStream=New-Object System.Io.FileStream("$filename","Create")
            while(($readCount = $stream.Read($buffer, 0, 1024)) -gt 0){
                $outputStream.Write($buffer, 0, $readCount)
            }
            $outputStream.Close()
            $stream.Close()
            $response.Close() 
            # if(Test-Path  $filename){echo "DownLoad successful"}
        }
        if ($do -eq "list")
        {
            $listFTP = [system.net.ftpwebrequest] [system.net.webrequest]::create("$ftpurl")
            $listFTP.Credentials = New-Object System.Net.NetworkCredential("$username","$password")
            $listFTP.Method=[system.net.WebRequestMethods+ftp]::listdirectorydetails
            # $listFTP.Method=[system.net.WebRequestMethods+ftp]::listdirectory
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

function CreateDir($directoryName) {
        $ftpPath = $ftpServer + $directoryName
        $ftpRequest = [System.Net.FtpWebRequest]::Create($ftpPath)
        $ftpRequest.Credentials = New-Object System.Net.NetworkCredential($username, $password)
        $ftpRequest.Method = [System.Net.WebRequestMethods+Ftp]::MakeDirectory
        try {
            $ftpResponse = $ftpRequest.GetResponse()
        } catch {
            Write-Host "ftpResponse error!!"
        } finally {
            if ($ftpResponse -ne $null) {
                $ftpResponse.Close()
            }
        }
    }
