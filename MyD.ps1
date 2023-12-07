# FTP 相關設定
$ftpUrl = "ftp://127.0.0.1:21/TEST_ITEM/common_bios_pxeboot_default/Abst64_unsign.exe"
$localPath = "c:\TestManager\ItemDownload\Abst64_unsign.exe"

$username = "sit001"
$password = "sit1234"

# 創建 NetworkCredential 對象
$credentials = New-Object System.Net.NetworkCredential($username, $password)

# 創建 WebClient 實例，並設置 Credentials
$webClient = New-Object System.Net.WebClient
$webClient.Credentials = $credentials

# 註冊 DownloadFileCompleted 事件
Register-ObjectEvent -InputObject $webClient -EventName DownloadFileCompleted -Action {
    # 下載完成事件處理
    if ($EventArgs.Error -eq $null) {
        Write-Host "下載完成"
    } else {
        Write-Host "下載失敗: $($EventArgs.Error.Message)"
    }

    # 解除事件綁定
    Unregister-Event -SourceIdentifier $eventSubscriberName
} -OutVariable eventSubscriberName

# 使用 DownloadFile 開始下載
$webClient.DownloadFileAsync((New-Object System.Uri $ftpUrl), $localPath)

# 等待下載完成
while ($webClient.IsBusy) {
    Start-Sleep -Seconds 1
}

# 清理 WebClient
$webClient.Dispose()
