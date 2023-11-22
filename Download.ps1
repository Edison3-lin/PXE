################## Build Log function ##########################
function process_log($log)
{
   $timestamp = Get-Date -Format "[yyyy-MM-dd HH:mm:ss] "
   $timestamp+$log | Add-Content $logfile
}
function result_log($log)
{
    Add-Content -Path $outputfile -Value (Get-Date -Format "[yyyy-MM-dd HH:mm:ss] ")
    $timestamp+$log | Add-Content $outputfile
}

function Down_Common($f)
{

    # Write-Host "Server: $ftpServer"
    # Write-Host "Username: $username"
    # Write-Host "Password: $password"
    
    process_log  "   === $f attached ==="
    $attDir = $f.split('.')[0]
    $ftpDirectory = "/Test_Item/$attDir/"
# process_log $attDir
# return    
    $commonFilePath = ".\ItemDownload\"
    if (-not (Test-Path -Path $commonFilePath -PathType Container)) {
        New-Item -Path $commonFilePath -ItemType Directory
    }
    
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
            process_log  "Download.... $fileName"
            $webClient.DownloadFile("$ftpServer$ftpDirectory$fileName", "$commonFilePath$fileName")

        }
        catch {
            process_log "!!!<$fileName>: $($_.Exception.Message)"
        }
    }
    $ftpReader.Close()
    $ftpStream.Close()
    $ftpResponse.Close()
}

$file = Get-Item $PSCommandPath
$Directory = Split-Path -Path $PSCommandPath -Parent
$Directory += '\MyLog'
$baseName = $file.BaseName
$logfile = $Directory+'\'+$baseName+"_process.log"
$tempfile = $Directory+'\temp.log'
$outputfile = $Directory+'\'+$baseName+'_result.log'

$configPath = ".\Server.json"
$config = Get-Content -Raw -Path $configPath | ConvertFrom-Json
$ftpServer = $config.ftpServer
$username = $config.username
$password = $config.password

# Specify the remote file to download
# $remoteFile = @("abt1.ps1","abt2.ps1","abt3.ps1")

# Specify the local destination for the downloaded file
$localPath = $Directory = Split-Path -Path $PSCommandPath -Parent
$localPath += "\ItemDownload"
if (-not (Test-Path -Path $localPath -PathType Container)) {
    New-Item -Path $localPath -ItemType Directory
}
# Create a WebClient object and set credentials
$webClient = New-Object System.Net.WebClient
$webClient.Credentials = New-Object System.Net.NetworkCredential($username, $password)

process_log "Download.. $remoteFile"
# $remoteFile = image_installation_application_default.dll

try {
    Down_Common($remoteFile)
}
catch {
    process_log "Directory not exist? <$f>: $($_.Exception.Message)"
}

process_log  "======Download finished======"
# Release WebClient
$webClient.Dispose()
return 0
