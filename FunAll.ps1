# .\FTPdown.ps1 "/Test_Item/" "c:\\TestManager\\ItemDownload\\"
# .\FTPup.ps1 "/Test_Item/" "c:\\TestManager\\ItemDownload\\"
# $dataSet = DATABASE "read" $MySqlCmd1
# DATABASE "update" $MySqlCmd2

Function DATABASE($do, $mySqlCmd) { 
    if( $do -notin @("read", "update") )
    {
        return $false
    }
    $SqlConn = New-Object System.Data.SqlClient.SqlConnection
    $SqlConn.ConnectionString = "Data Source=$DBserver;Initial Catalog=$Database;user id=$DBuserName;pwd=$DBpassword"

    # Try to open the connection, wait up to 30 seconds
    $timeout = 30
    $timer = [System.Diagnostics.Stopwatch]::StartNew()

    while ($SqlConn.State -ne 'Open' -and $timer.Elapsed.TotalSeconds -lt $timeout) {
        try {
            # Open connection
            $SqlConn.Open()
            Start-Sleep -Seconds 1
        } catch {
            # If the connection fails to open, catch the exception and continue waiting.
            # process_log "Error opening connection: $_"
        }
    }

    $timer.Stop()

    # Check connection status
    if ($SqlConn.State -ne 'Open') {
        return "Unconnected_"
    }

    $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
    $SqlCmd.connection = $SqlConn
    $sqlCmd.CommandText = $mySqlCmd

    if ($do -eq "read") {
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $sqlCmd
        $dataset = New-Object System.Data.DataSet
        $adapter.Fill($dataSet)
        $SqlConn.close()
        return $dataSet
    } 
    if ($do -eq "update") {
        $SqlCmd.executenonquery()
        $SqlConn.close()
        return $null
    }
}        

Function FTP($ftpurl,$do,$filename) { 
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

Function process_log($log) {
        $timestamp = Get-Date -Format "[yyyy-MM-dd HH:mm:ss] "
        $timestamp+$log | Add-Content $logfile
    }
Function result_log($log) {
        Add-Content -Path $outputfile -Value (Get-Date -Format "[yyyy-MM-dd HH:mm:ss] ")
        $timestamp+$log | Add-Content $outputfile
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


    # Server.json
    $configPath = "c:\TestManager\Server.json"
    $config = Get-Content -Raw -Path $configPath | ConvertFrom-Json
    $ftpServer  = $config.ftpServer
    $username   = $config.username
    $password   = $config.password
    $Database   = $config.Database
    $DBserver   = $config.DBserver 
    $DBuserName = $config.DBuserName
    $DBpassword = $config.DBpassword
    
    # TR_Result.json
    $TRPath = "c:\TestManager\TR_Result.json"
    $TRconfig = Get-Content -Raw -Path $TRPath | ConvertFrom-Json
    $TCM_ID     = $TRconfig.TCM_ID 
    $TR_ID      = $TRconfig.TR_ID 
    $TestResult = $TRconfig.TestResult
    $TestStatus = $TRconfig.TestStatus
    $Text_Log_File_Path = $TRconfig.Text_Log_File_Path
    $Test_TimeOut       = $TRconfig.Test_TimeOut
