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

function Get_Version()
{
    # 获取目录下的所有文件
    $files = Get-ChildItem -Path ".\" -File
    foreach ($file in $files) {
        if($file.Name -like "TM????.exe")
        {
            $f = $file.Name.Substring(2, 4) 
            return $f
        }
    }
    return $null    
}
function Down_Common()
{
    $ftpDirectory = "/TestManager/"
    $commonFilePath = ".\"

    # 获取目录下的所有文件
    $files = Get-ChildItem -Path $commonFilePath -File

    # 删除所有文件
    foreach ($file in $files) {
        if($file.Name -ne "UpdateT.ps1")
        {
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
    if( [int]$dir[-2] -gt [int]$version)
    {
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

$version = Get_Version

try {
    Down_Common
}
catch {
    Write-Host "Directory not exist?"
}

Write-Host  "======Download finished======"
# Release WebClient
$webClient.Dispose()
