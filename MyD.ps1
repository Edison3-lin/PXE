 
#==============================全局变量区=========================#
[bool]$isDownloaded=$False
 
#==============================下载函数===========================#
Function Download([String]$url, [String]$fullFileName)
{
    if([String]::IsNullOrEmpty($url) -or [String]::IsNullOrEmpty($fullFileName))
    {
        return $false;
    }
    try
    {
        $client = New-Object System.Net.WebClient 
        $client.UseDefaultCredentials = $True
 
        #监视WebClient 的下载完成事件 
         Register-ObjectEvent -InputObject $client -EventName DownloadFileCompleted `
        -SourceIdentifier Web.DownloadFileCompleted -Action {   
            #下载完成，结束下载
            $Global:isDownloaded = $True
        }
        #监视WebClient 的进度事件
        Register-ObjectEvent -InputObject $client -EventName DownloadProgressChanged `
        -SourceIdentifier Web.DownloadProgressChanged -Action {
            #将下载的进度信息记录到全局的Data对象中
            $Global:Data = $event
        }
 
        $Global:isDownloaded =$False
 
        #监视PowerShell退出事件
        Register-EngineEvent -SourceIdentifier ([System.Management.Automation.PSEngineEvent]::Exiting) -Action {
            #PowerShell 结束事件
            Get-EventSubscriber | Unregister-Event
            Get-Job | Remove-Job -Force
           }
           
         #启用定时器，设置1秒一次输出下载进度
        $timer = New-Object timers.timer
        # 1 second interval
        $timer.Interval = 1000
        #Create the event subscription
        Register-ObjectEvent -InputObject $timer -EventName Elapsed -SourceIdentifier Timer.Output -Action {
            $percent = $Global:Data.SourceArgs.ProgressPercentage
            $totalBytes = $Global:Data.SourceArgs.TotalBytesToReceive
            $receivedBytes = $Global:Data.SourceArgs.BytesReceived
           
            If ($percent -ne $null) {
                 #这里你可以选择将进度显示到命令行 也可以选择将进度写到文件，具体看自己需求
                 #我这里选择将进度输出到命令行
                    Write-Host "当前下载进度:$percent  已下载:$receivedBytes 总大小:$totalBytes"
                    
            }
           
        }
        $timer.Enabled = $True
 
        #使用异步方式下载文件
         $client.DownloadFileAsync($url, $fullFileName)
          While (-Not $isDownloaded)
           {
                #等待下载线程结束
                Start-Sleep -m 100
           }
 
         $timer.Enabled = $False
         
        #清除监视
        Get-EventSubscriber | Unregister-Event
        Get-Job | Remove-Job -Force
        #关闭下载线程
        $client.Dispose()
        Remove-Variable client
      
         Write-Host "Finish "
    }
    catch
    {
       
        return $false;  
    }
    return $true;
}

$ftpServer = "ftp://127.0.0.1:21"
$username = "sit001"
$password = "sit1234"

# 創建 NetworkCredential 對象
$credentials = New-Object System.Net.NetworkCredential($username, $password)

# 創建 WebClient 實例，並設置 Credentials
$webClient = New-Object System.Net.WebClient
$webClient.Credentials = $credentials

Download -url "ftp://127.0.0.1:21/Test_Item/Test_Collection/Microsoft.PowerShell_profile.ps1" -fullFileName ".\ItemDownload\Microsoft.PowerShell_profile.ps1"