Function DATABASE($do, $mySqlCmd) { 
        if( $do -notin @("read", "update") )
        {
            Write-Host "xxxxxx"
            return "xxxxxx"
        }
        $SqlConn = New-Object System.Data.SqlClient.SqlConnection
        $SqlConn.ConnectionString = "Data Source=$DBserver;Initial Catalog=$Database;user id=$DBuserName;pwd=$DBpassword"

        # Try to open the connection, wait up to 30 seconds
        $timeout = 30
        $timer = [System.Diagnostics.Stopwatch]::StartNew()

        while ($SqlConn.State -ne 'Open' -and $timer.Elapsed.TotalSeconds -lt $timeout) {
            try {
                # Open connection
                $SqlConn.Open()
                Start-Sleep -Seconds 1
            } catch {
                # If the connection fails to open, catch the exception and continue waiting.
                # process_log "Error opening connection: $_"
            }
        }

        $timer.Stop()

        # Check connection status
        if ($SqlConn.State -ne 'Open') {
            return "Unconnected_"
        }

        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $SqlCmd.connection = $SqlConn
        $sqlCmd.CommandText = $mySqlCmd

        if ($do -eq "read") {
            $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $sqlCmd
            $dataset = New-Object System.Data.DataSet
            $adapter.Fill($dataSet)
            $SqlConn.close()
            return $dataSet
        } 
        if ($do -eq "update") {
            $SqlCmd.executenonquery()
            $SqlConn.close()
            return $null
        }
    }        
