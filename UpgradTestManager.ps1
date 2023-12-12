function FTP_Download_TM() {
        $ftpDirectory = "/TestManager/"
        $commonFilePath = ".\"

        # # 删除c:\TestManager目錄下所有file
        # $files = Get-ChildItem -Path $commonFilePath -File
        # foreach ($file in $files) {
        #     if($file.Name -ne "UpdateT.ps1")
        #     {
        #         Remove-Item $file.FullName -Force
        #     }    
        # }

        # Get all files 
        $files = Get-ChildItem -Path ".\" -File
        foreach ($file in $files) {
            if($file.Name -like "TM1*") {
                Remove-Item $file.FullName -Force
            }
        }

        $ftpRequest = [System.Net.FtpWebRequest]::Create("$ftpServer$ftpDirectory")
        $ftpRequest.Credentials = New-Object System.Net.NetworkCredential($username, $password)
        $ftpRequest.Method = [System.Net.WebRequestMethods+Ftp]::ListDirectory

        # Get FTP directory list
        $ftpResponse = $ftpRequest.GetResponse()
        $ftpStream = $ftpResponse.GetResponseStream()
        $ftpReader = New-Object System.IO.StreamReader($ftpStream)
        $directoryListing = $ftpReader.ReadToEnd()
        $dir = $directoryListing -split "`r`n"

        Start-Process -FilePath "C:\TestManager\Service_Stop.bat" -NoNewWindow -Wait

        $ftpDirectory = "/TestManager/$($dir[-2])/"
        $ftpRequest = [System.Net.FtpWebRequest]::Create("$ftpServer$ftpDirectory")
        $ftpRequest.Credentials = New-Object System.Net.NetworkCredential($username, $password)
        $ftpRequest.Method = [System.Net.WebRequestMethods+Ftp]::ListDirectory
        $ftpResponse = $ftpRequest.GetResponse()
        $ftpStream = $ftpResponse.GetResponseStream()
        $ftpReader = New-Object System.IO.StreamReader($ftpStream)

        while (-not $ftpReader.EndOfStream) {
            $fileName = $ftpReader.ReadLine()
            try {
                # Download the file
                Write-Host  "Download.... $fileName"
                $webClient.DownloadFile("$ftpServer$ftpDirectory$fileName", "$commonFilePath$fileName")
            }
            catch {
                Write-Host "!!!<$fileName>: $($_.Exception.Message)"
            }
        }    
        $ftpReader.Close()
        $ftpStream.Close()
        $ftpResponse.Close()

        Start-Process -FilePath "C:\TestManager\Service_Start.bat" -NoNewWindow -Wait
        return
    }

    $configPath = ".\Server.json"
    $config = Get-Content -Raw -Path $configPath | ConvertFrom-Json
    $ftpServer = $config.ftpServer
    $username = $config.username
    $password = $config.password

    # Create a WebClient object and set credentials
    $webClient = New-Object System.Net.WebClient
    $webClient.Credentials = New-Object System.Net.NetworkCredential($username, $password)

    try {
        FTP_Download_TM
    }
    catch {
        Write-Host "Directory not exist?"
    }

    Write-Host  "@@@ Update finished @@@"
    # Release WebClient
    $webClient.Dispose()
return