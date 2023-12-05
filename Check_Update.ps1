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

    $ftpRequest = [System.Net.FtpWebRequest]::Create("$ftpServer$ftpDirectory")
    $ftpRequest.Credentials = New-Object System.Net.NetworkCredential($username, $password)
    $ftpRequest.Method = [System.Net.WebRequestMethods+Ftp]::ListDirectory

    # Get FTP directoryListing
    $ftpResponse = $ftpRequest.GetResponse()
    $ftpStream = $ftpResponse.GetResponseStream()
    $ftpReader = New-Object System.IO.StreamReader($ftpStream)
    $directoryListing = $ftpReader.ReadToEnd()
    $dir = $directoryListing -split "`r`n"
    $ftpReader.Close()
    $ftpStream.Close()
    $ftpResponse.Close()

    return ([int]$dir[-2] -gt [int]$version)
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
    $CheckUpdate =  Down_Common
}
catch {
    Write-Host "Directory not exist?"
    $webClient.Dispose()
    return $false
}

# Release WebClient
$webClient.Dispose()

return $CheckUpdate
