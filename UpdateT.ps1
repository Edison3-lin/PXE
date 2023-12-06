function Down_Common()
{
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

    # 获取目录下的所有文件
    $files = Get-ChildItem -Path ".\" -File
    foreach ($file in $files) {
        if($file.Name -like "TM1*")
        {
            # Write-Host $file.FullName
            Remove-Item $file.FullName -Force
        }
    }

    $ftpRequest = [System.Net.FtpWebRequest]::Create("$ftpServer$ftpDirectory")
    $ftpRequest.Credentials = New-Object System.Net.NetworkCredential($username, $password)
    $ftpRequest.Method = [System.Net.WebRequestMethods+Ftp]::ListDirectory

    # 获取FTP服务器上的目录列表
    $ftpResponse = $ftpRequest.GetResponse()
    $ftpStream = $ftpResponse.GetResponseStream()
    $ftpReader = New-Object System.IO.StreamReader($ftpStream)
    $directoryListing = $ftpReader.ReadToEnd()
    $dir = $directoryListing -split "`r`n"
    # if( [int]$dir[-2] -gt [int]$version)
    # {

#(EdisonLin-20231206-)>>
Start-Process -FilePath "C:\TestManager\Service_out.bat" -NoNewWindow -Wait
# Start-Sleep -Seconds 2
#(EdisonLin-20231206-)<<


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
        # }
        # Start-Process -FilePath ".\Test.bat" -Wait
#(EdisonLin-20231206-)>>
# Start-Sleep -Seconds 5
Start-Process -FilePath "C:\TestManager\Service_in.bat" -NoNewWindow -Wait
#(EdisonLin-20231206-)<<

    }    
    $ftpReader.Close()
    $ftpStream.Close()
    $ftpResponse.Close()
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
    Down_Common
}
catch {
    Write-Host "Directory not exist?"
}

Write-Host  "@@@ Update finished @@@"
# Release WebClient
$webClient.Dispose()
