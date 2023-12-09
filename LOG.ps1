Function process_log($log) {
        $timestamp = Get-Date -Format "[yyyy-MM-dd HH:mm:ss] "
        $timestamp+$log | Add-Content $logfile
    }
Function result_log($log) {
        Add-Content -Path $outputfile -Value (Get-Date -Format "[yyyy-MM-dd HH:mm:ss] ")
        $timestamp+$log | Add-Content $outputfile
    }