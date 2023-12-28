function CheckWindowsSMode {
    $sModeStatus = Get-WindowsOptionalFeature -Online | Where-Object { $_.FeatureName -eq "Microsoft-Windows-Subsystem-For-Linux" }
    
    if ($sModeStatus -eq $null) {
        Write-Host "Windows is not in S Mode."
    } else {
        Write-Host "Windows is in S Mode."
    }
}

CheckWindowsSMode

# $a = (Get-WmiObject -Class Win32_OperatingSystem).OperatingSystemSKU
# write-host $a

# 下列清單列出可能的 SKU 值。
# PRODUCT_UNDEFINED （0）

# 未定義
# PRODUCT_ULTIMATE （1）
# Ultimate Edition，例如 Windows Vista Ultimate。

# PRODUCT_HOME_BASIC （2）
# Home Basic Edition

# PRODUCT_HOME_PREMIUM （3）
# Home 進階版 Edition

# PRODUCT_ENTERPRISE （4）
# 企業版

# PRODUCT_BUSINESS （6）
# Business Edition

# PRODUCT_STANDARD_SERVER （7）
# Windows Server Standard Edition （桌面體驗安裝）

# PRODUCT_DATACENTER_SERVER （8）
# Windows Server Datacenter Edition （桌面體驗安裝）

# PRODUCT_SMALLBUSINESS_SERVER （9）
# Small Business Server Edition

# PRODUCT_ENTERPRISE_SERVER （10）
# Enterprise Server Edition

# PRODUCT_STARTER （11）
# Starter Edition

# PRODUCT_DATACENTER_SERVER_CORE （12）
# Datacenter Server Core Edition

# PRODUCT_STANDARD_SERVER_CORE （13）
# Standard Server Core Edition

# PRODUCT_ENTERPRISE_SERVER_CORE （14）
# Enterprise Server Core Edition

# PRODUCT_WEB_SERVER （17）
# 網頁伺服器版本

# PRODUCT_HOME_SERVER （19）
# Home Server Edition

# PRODUCT_STORAGE_EXPRESS_SERVER （20）
# 儲存體 Express Server Edition

# PRODUCT_STORAGE_STANDARD_SERVER （21）
# Windows 儲存體 Server Standard Edition （桌面體驗安裝）

# PRODUCT_STORAGE_WORKGROUP_SERVER （22）
# Windows 儲存體 Server Workgroup Edition （桌面體驗安裝）

# PRODUCT_STORAGE_ENTERPRISE_SERVER （23）
# 儲存體 Enterprise Server Edition

# PRODUCT_SERVER_FOR_SMALLBUSINESS （24）
# 適用於 Small Business Edition 的伺服器

# PRODUCT_SMALLBUSINESS_SERVER_PREMIUM （25）
# Small Business Server 進階版 Edition

# PRODUCT_ENTERPRISE_N （27）
# Windows Enterprise Edition

# PRODUCT_ULTIMATE_N （28）
# Windows Ultimate Edition

# PRODUCT_WEB_SERVER_CORE （29）
# Windows Server Web Server Edition （Server Core 安裝）

# PRODUCT_STANDARD_SERVER_V （36）
# 不含 Hyper-V 的 Windows Server Standard Edition

# PRODUCT_DATACENTER_SERVER_V （37）
# 不含 Hyper-V 的 Windows Server Datacenter Edition （完整安裝）

# PRODUCT_ENTERPRISE_SERVER_V （38）
# 不含 Hyper-V 的 Windows Server Enterprise Edition （完整安裝）

# PRODUCT_DATACENTER_SERVER_CORE_V （39）
# 不含 Hyper-V 的 Windows Server Datacenter Edition （Server Core 安裝）

# PRODUCT_STANDARD_SERVER_CORE_V （40）
# 不含 Hyper-V 的 Windows Server Standard Edition （Server Core 安裝）

# PRODUCT_ENTERPRISE_SERVER_CORE_V （41）
# 不含 Hyper-V 的 Windows Server Enterprise Edition （Server Core 安裝）

# PRODUCT_HYPERV （42）
# Microsoft Hyper-V Server

# PRODUCT_STORAGE_EXPRESS_SERVER_CORE （43）
# 儲存體 Server Express Edition （Server Core 安裝）

# PRODUCT_STORAGE_STANDARD_SERVER_CORE （44）
# 儲存體 Server Standard Edition （Server Core 安裝）

# PRODUCT_STORAGE_WORKGROUP_SERVER_CORE （45）
# 儲存體 Server Workgroup Edition （Server Core 安裝）

# PRODUCT_STORAGE_ENTERPRISE_SERVER_CORE （46）
# 儲存體 Server Enterprise Edition （Server Core 安裝）

# PRODUCT_PROFESSIONAL （48）
# Windows 專業版

# PRODUCT_SB_SOLUTION_SERVER （50）
# Windows Server Essentials （桌面體驗安裝）

# PRODUCT_SMALLBUSINESS_SERVER_PREMIUM_CORE （63）
# Small Business Server 進階版 （Server Core 安裝）

# PRODUCT_CLUSTER_SERVER_V （64）
# 不含 Hyper-V 的 Windows 計算叢集伺服器

# PRODUCT_CORE_ARM （97）
# Windows RT

# PRODUCT_CORE （101）
# Windows Home

# PRODUCT_PROFESSIONAL_WMC （103）
# Windows Professional with Media Center

# PRODUCT_MOBILE_CORE （104）
# Windows Mobile

# PRODUCT_IOTUAP （123）
# Windows IoT （物聯網） 核心

# PRODUCT_DATACENTER_NANO_SERVER （143）
# Windows Server Datacenter Edition （Nano Server 安裝）

# PRODUCT_STANDARD_NANO_SERVER （144）
# Windows Server Standard Edition （Nano Server 安裝）

# PRODUCT_DATACENTER_WS_SERVER_CORE （147）
# Windows Server Datacenter Edition （Server Core 安裝）

# PRODUCT_STANDARD_WS_SERVER_CORE （148）
# Windows Server Standard Edition （Server Core 安裝）

# PRODUCT_ENTERPRISE_FOR_VIRTUAL_DESKTOPS （175）
# 適用于虛擬桌面的 Windows 企業版 （Azure 虛擬桌面）

# PRODUCT_DATACENTER_SERVER_AZURE_EDITION （407）
# Windows Server Datacenter：Azure Edition
