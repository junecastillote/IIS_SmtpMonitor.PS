Function Get-IISSmtpItemCount {
    [CmdletBinding()]
    param (
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]
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
        $ItemType
    )

    # Function to count items in the given directory
    Function GetItemCount {
        [CmdletBinding()]
        param (
            [Parameter(Mandatory)]
            [string]
            $Directory
        )
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

    # Remove ItemType from the PSBoundParameters
    if ($ItemType) {
        $null = $PSBoundParameters.Remove('ItemType')
    }

    # Get the Virtual SMTP Servers
    $smtp_server = @(Get-IISSmtpServer @PSBoundParameters)

    # Stop script if no Virtual SMTP Servers were retrieved
    if (!$smtp_server) {
        return $null
    }

    $results = @()
    foreach ($server in $smtp_server) {
        try {
            $result = [ordered]@{
                ComputerName             = $server.ComputerName
                VirtualServerName        = $server.VirtualServerName
                VirtualServerDisplayName = $server.VirtualServerDisplayName
                VirtualServerState       = $server.VirtualServerState
                SmtpServiceState         = $server.SmtpServiceState
                Items                    = @()
            }

            if ($ItemType) {
                foreach ($type in $ItemType) {
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

    $results
}
