# # 执行 systeminfo 命令并捕获输出
# $systemInfoOutput = systeminfo

# # 创建一个空的哈希表来存储解析后的键值对
# $parsedSystemInfo = @{}

# # 遍历 systeminfo 的输出
# foreach ($line in $systemInfoOutput) {
#     # 检查行是否包含键值对（即是否包含冒号）
#     if ($line -match ':') {
#         # 将行分割成键和值
#         $splitLine = $line -split ':', 2

#         # 清理和修剪键和值
#         $key = $splitLine[0].Trim()
#         $value = $splitLine[1].Trim()

#         # 将键值对添加到哈希表中
#         $parsedSystemInfo[$key] = $value
#     }
# }

# # 将哈希表转换为 JSON
# $jsonSystemInfo = $parsedSystemInfo | ConvertTo-Json

# # 输出 JSON
# $jsonSystemInfo


# # 指定输出文件的路径
# $outputFilePath = "C:\TestManager\edison.json"

# # 将 JSON 数据写入文件
# $jsonSystemInfo | Out-File -FilePath $outputFilePath


# # 定义用于访问注册表的路径
# $registryPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
# $registryValueName = "AppsUseLightTheme"

# # 尝试获取注册表项的值
# $appsUseLightTheme = Get-ItemPropertyValue -Path $registryPath -Name $registryValueName -ErrorAction SilentlyContinue

# # 0=Dark; 1=light
# Write-Host $appsUseLightTheme


# # 創建兩個陣列
# $array1 = 1..5   # 包含數字 1 到 5
# $array2 = 6..10  # 包含數字 6 到 10

# # 創建一個二維陣列
# $twoDimensionalArray = @($array1, $array2)

# # 顯示整個二維陣列
# # $twoDimensionalArray

# # 訪問並顯示二維陣列中的特定元素
# # 顯示第一個陣列的第三個元素
# $twoDimensionalArray[0][2]

# # 顯示第二個陣列的第一個元素
# $twoDimensionalArray[1][0]


# 創建一個列表
$list1 = New-Object System.Collections.Generic.List[Object]
$list2 = New-Object System.Collections.Generic.List[Object]

# # 添加一些元素到列表
$a = 10
$b = 11
$c = 12
$list1.Add($a)
$list1.Add($b)
$list1.Add($c)

# # 在索引 1 的位置插入一個新元素 'newElement'
# $list1.Insert(1, 'newElement')
# $list1 = 1..5
# $list2 = 7..9
# 顯示更新後的列表
$null = $list1.Remove(11)
    if($a -in $list1) {
        Write-Host 'xxxxxxxx'
    }    
    else {
        Write-Host 'yyyyyyyyyyyy'
    }
