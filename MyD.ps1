Function ftp($ftpurl,$username,$password,$do,$filename,$DownPatch) { 
    # ftp 服务器地址，用户名，密码，操作（上传up/下载down/列表list），文件名，下载路径
    # 示例：ftp ftp://10.10.98.91/ ftpuser shenzhen up C:\Windows\setupact.txt
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
            $outputStream=New-Object System.Io.FileStream("$DownPatch","Create")
            while(($readCount = $stream.Read($buffer, 0, 1024)) -gt 0){
                $outputStream.Write($buffer, 0, $readCount)
            }
            $outputStream.Close()
            $stream.Close()
            $response.Close() 
            if(Test-Path  $DownPatch){echo "DownLoad successful"}
        }
        if ($do -eq "list")
        {
            $listFTP = [system.net.ftpwebrequest] [system.net.webrequest]::create("$ftpurl")
            $listFTP.Credentials = New-Object System.Net.NetworkCredential("$username","$password")
            $listFTP.Method=[system.net.WebRequestMethods+ftp]::listdirectorydetails
            $response = $listFTP.getresponse()
            $stream = New-Object System.Io.StreamReader($response.getresponsestream(),[System.Text.Encoding]::UTF8)
            while(-not $stream.EndOfStream){
                $stream.ReadLine()
            }
            $stream.Close()
            $response.Close()     
        }
    }
    ftp "ftp://127.0.0.1:21/" "sit001" "sit1234" down "Test_Item/TestCase.exe" "C:\TestManager\TestCase.exe" 