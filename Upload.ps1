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

$file = Get-Item $PSCommandPath
$Directory = Split-Path -Path $PSCommandPath -Parent
$baseName = $file.BaseName
$logfile = $Directory+'\'+$baseName+"_process.log"
$tempfile = $Directory+'\temp.log'
$outputfile = $Directory+'\'+$baseName+'_result.log'

# Server.json
$configPath = ".\Server.json"
$config = Get-Content -Raw -Path $configPath | ConvertFrom-Json
$ftpServer = $config.ftpServer
$username = $config.username
$password = $config.password

# TR_Result.json
$TRPath = ".\TR_Result.json"
$TRconfig = Get-Content -Raw -Path $TRPath | ConvertFrom-Json
$TCM_ID = $TRconfig.TCM_ID 
$TR_ID = $TRconfig.TR_ID 
$Text_Log_File_Path = $TRconfig.Text_Log_File_Path

# local and FTP directory
$UUID = Get-WmiObject Win32_ComputerSystemProduct | Select-Object -ExpandProperty UUID 
# $localFilePath = "c:\TestManager\ResultUpload\"
$remoteFilePath = "/Test_Log/$UUID/$TCM_ID/$TR_ID/"

# Create a WebClient object and set credentials
$webClient = New-Object System.Net.WebClient
$webClient.Credentials = New-Object System.Net.NetworkCredential($username, $password)

# $files = Get-ChildItem -Path $localFilePath

foreach ($log in $Text_Log_File_Path) {
    $f = Get-Item $log
    $file = $f.Name
    $sourceFilePath = $log
    $destinationFilePath = $remoteFilePath+$file
    try {
        # WebClient upload files
        process_log "$sourceFilePath to $ftpServer$destinationFilePath"

        $webClient.UploadFile("$ftpServer$destinationFilePath", $sourceFilePath)
    }
    catch {
        process_log "!!!<$f>: $($_.Exception.Message)"
    }
    process_log "<$f> upload to $localFilePath"
}

process_log  "======Upload finished======"
# Release WebClient
$webClient.Dispose()
return 0
