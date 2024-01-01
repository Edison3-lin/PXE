Function process_log($log) {
    $timestamp = Get-Date -Format "[yyyy-MM-dd HH:mm:ss] "
    $timestamp+$log | Add-Content $logfile
}

Function FTP_Download_TM() {
        $ftpDirectory = "/Captain_Tool/TestManager/"
        $commonFilePath = ".\"

        # delete c:\TestManager all files
        # $files = Get-ChildItem -Path $commonFilePath -File
        # foreach ($file in $files) {
        #     if ($file.Name -ne "UT.ps1")
        #     {
        #         if ($file.Name -notlike "TMservice*") 
        #         { 
        #             if ($file.Name -notlike "Service*") 
        #             {
        #                 if ($file.Name -notlike "InstallUtil*") 
        #                 {
        #                     # Remove-Item $file.FullName -Force
        #                 }
        #             }    
        #         }    
        #     }    
        # }

        # Get all files 
        $files = Get-ChildItem -Path ".\" -File
        foreach ($file in $files) {
            if($file.Name -like "TM1*") {
                process_log "delete.. $($file.FullName)" 
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
 process_log "--- Service_Stop ---" 
        Start-Process -FilePath "C:\TestManager\Service_Stop.bat" -NoNewWindow -Wait
        $ftpDirectory = "/Captain_Tool/TestManager/$($dir[-2])/"
 process_log "/Captain_Tool/TestManager/$($dir[-2])/"
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
                process_log "Download.. $ftpServer$ftpDirectory$fileName"
                $webClient.DownloadFile("$ftpServer$ftpDirectory$fileName", "$commonFilePath$fileName")
            }
            catch {
                process_log "!!!<$fileName>: $($_.Exception.Message)"
            }
        }    
        $ftpReader.Close()
        $ftpStream.Close()
        $ftpResponse.Close()
        
        process_log "=== Service_Start ==="
        Start-Process -FilePath "C:\TestManager\Service_Start.bat" -NoNewWindow -Wait
        return
    }

    $configPath = ".\Server.json"
    $config = Get-Content -Raw -Path $configPath | ConvertFrom-Json
    $ftpServer = $config.ftpServer
    $username = $config.username
    $password = $config.password

    ### Create log file ###
    $file = Get-Item $PSCommandPath
    $Directory = Split-Path -Path $PSCommandPath -Parent
    $baseName = $file.BaseName
    $logfile = $Directory+'\'+$baseName+"_process.log"

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