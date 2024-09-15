Function Get-IISSmtpServerStatus {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [string[]]
        $ComputerName,

        [Parameter()]
        [ValidateSet(
            'Queue',
            'Drop',
            'BadMail',
            'Pickup',
            'LogFile'
        )]
        [string[]]
        $SelectedFolder,

        [Parameter()]
        [switch]
        $AggregateByHost
    )

    begin {
        # Function to count items in the given directory
        Function GetItemCount {
            [CmdletBinding()]
            param (
                [Parameter(Mandatory)]
                [string]
                $Directory
            )
            switch ((Test-Path $Directory)) {
                $true {
                    $items = @(Get-ChildItem -LiteralPath $Directory -File)
                    if ($items) {
                        $size = ($items | Measure-Object Length -Sum).Sum
                    }
                    else {
                        $size = 0
                    }

                    [PSCustomObject]@{
                        Directory = $Directory
                        ItemCount = $items.Count
                        Size      = $size
                    }
                }
                $false {
                    [PSCustomObject]@{
                        Directory = $Directory
                        ItemCount = 0
                        Size      = 0
                    }
                }
            }
        }

        # Helper function to add item count results
        Function AddItemCountResult {
            param (
                [ref]$result,
                [string]$type,
                [string]$directory,
                [string]$computername,
                [bool]$islocalhost
            )

            switch ($islocalhost) {
                $true { $items = GetItemCount $directory -ErrorAction Stop }
                $false { $items = GetItemCount (GetNetworkPath $directory -computerName $computername) -ErrorAction Stop }
                Default {}
            }

            $result.Value.Items += [PSCustomObject]@{
                Type        = $type
                LocalPath   = $directory
                NetworkPath = (GetNetworkPath $directory -computerName $computername)
                TotalCount  = $items.ItemCount
                TotalSize   = $items.Size
            }
        }

        # Function to get network path for directories
        Function GetNetworkPath {
            param (
                [string]$computerName,
                [string]$directory
            )
            if (!($computerName -eq 'localhost')) {
                "\\$computerName\$($directory -replace ':','$')"
            }
            else {
                $directory
            }
        }

        $virtual_smtp_server_status_collection = [System.Collections.Generic.List[object]]@()
    }

    process {
        foreach ($computer_name in $ComputerName) {
            # Get the Virtual SMTP Servers
            $smtp_server = @(Get-IISSmtpServer $computer_name)

            # Skip if no Virtual SMTP Servers were retrieved
            if (!$smtp_server) {
                Continue
            }

            $results = @()
            foreach ($server in $smtp_server) {
                try {
                    $result = [ordered]@{
                        PSTypeName               = 'IIS_Smtp_Server_Status'
                        ComputerName             = $server.ComputerName
                        VirtualServerName        = $server.VirtualServerName
                        VirtualServerDisplayName = $server.VirtualServerDisplayName
                        VirtualServerState       = $server.VirtualServerState
                        SmtpServiceState         = $server.SmtpServiceState
                        Items                    = @()
                    }

                    if ($SelectedFolder) {
                        foreach ($type in $SelectedFolder) {
                            switch ($type) {
                                'Queue' { AddItemCountResult -result ([ref]$result) -type 'Queue' -directory $server.QueueDirectory -islocalhost $server.IsLocalHost -computername $server.ComputerName }
                                'Pickup' { AddItemCountResult -result ([ref]$result) -type 'Pickup' -directory $server.PickupDirectory -islocalhost $server.IsLocalHost -computername $server.ComputerName }
                                'BadMail' { AddItemCountResult -result ([ref]$result) -type 'BadMail' -directory $server.BadMailDirectory -islocalhost $server.IsLocalHost -computername $server.ComputerName }
                                'Drop' { AddItemCountResult -result ([ref]$result) -type 'Drop' -directory $server.DropDirectory -islocalhost $server.IsLocalHost -computername $server.ComputerName }
                                'LogFile' { AddItemCountResult -result ([ref]$result) -type 'LogFile' -directory $server.LogFileDirectory -islocalhost $server.IsLocalHost -computername $server.ComputerName }
                            }
                        }
                    }
                    else {
                        AddItemCountResult -result ([ref]$result) -type 'Queue' -directory $server.QueueDirectory -islocalhost $server.IsLocalHost -computername $server.ComputerName
                        AddItemCountResult -result ([ref]$result) -type 'Pickup' -directory $server.PickupDirectory -islocalhost $server.IsLocalHost -computername $server.ComputerName
                        AddItemCountResult -result ([ref]$result) -type 'BadMail' -directory $server.BadMailDirectory -islocalhost $server.IsLocalHost -computername $server.ComputerName
                        AddItemCountResult -result ([ref]$result) -type 'Drop' -directory $server.DropDirectory -islocalhost $server.IsLocalHost -computername $server.ComputerName
                        AddItemCountResult -result ([ref]$result) -type 'LogFile' -directory $server.LogFileDirectory -islocalhost $server.IsLocalHost -computername $server.ComputerName
                    }

                    $results += [PSCustomObject]$result
                }
                catch {
                    Write-Error $_.Exception.Message
                }
            }
            $virtual_smtp_server_status_collection.AddRange(@($results))
        }
    }

    end {
        if ($AggregateByHost) {

            $smtp_instance_aggregated_by_computername = [System.Collections.Generic.List[object]]@()
            foreach ($unique_computer_name in ($virtual_smtp_server_status_collection.ComputerName | Sort-Object | Select-Object -Unique)) {
                $virtual_server = @($virtual_smtp_server_status_collection | Where-Object { $_.ComputerName -eq $unique_computer_name })
                $smtp_instance_aggregated_by_computername.Add([PSCustomObject]@{
                        ComputerName                = $unique_computer_name
                        SmtpServiceState            = $($virtual_server[0].SmtpServiceState)
                        VirtualSMTPServerCollection = @($virtual_server | Select-Object -ExcludeProperty ComputerName, SmtpServiceState)
                    })
            }

            return $smtp_instance_aggregated_by_computername
        }
        else {
            return $virtual_smtp_server_status_collection
        }

    }


}
