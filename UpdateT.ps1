Set-StrictMode -Version Latest
function Invoke-Administrator([String] $FilePath, [String[]] $ArgumentList = '') 
{
  $Current = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
  $Administrator = [Security.Principal.WindowsBuiltInRole]::Administrator
  if (-not $Current.IsInRole($Administrator)) 
  {
    $PowerShellPath = (Get-Process -Id $PID).Path
    $Command = "" + $FilePath + "$ArgumentList" + ""
    Start-Process $PowerShellPath "-NoProfile -ExecutionPolicy Bypass -File $Command" -Verb RunAs
    exit
  } 
  else 
  {
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy ByPass
  }
}

Invoke-Administrator $PSCommandPath

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

function Down_Common()
{
    $ftpDirectory = "/TestManager/"
    $commonFilePath = ".\"
    
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


# Specify the local destination for the downloaded file
$localPath = $Directory = Split-Path -Path $PSCommandPath -Parent
$localPath += "\"
if (-not (Test-Path -Path $localPath -PathType Container)) {
    New-Item -Path $localPath -ItemType Directory
}
# Create a WebClient object and set credentials
$webClient = New-Object System.Net.WebClient
$webClient.Credentials = New-Object System.Net.NetworkCredential($username, $password)

process_log "Download.. TestManager"
# $remoteFile = image_installation_application_default.dll

try {
    Down_Common
}
catch {
    process_log "Directory not exist?"
}

process_log  "======Download finished======"
# Release WebClient
$webClient.Dispose()
return 0
