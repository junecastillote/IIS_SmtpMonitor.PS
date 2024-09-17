Function Get-IISSmtpServer {
    [CmdletBinding()]
    param (
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]
        $ComputerName
    )

    [bool]$is_localhost = $false

    if (!$ComputerName) {
        $ComputerName = ($env:COMPUTERNAME).ToUpper()
        [bool]$is_localhost = $true
    }
    # Command expression to retrieve the SMTP server instances
    $command = '$(
    $smtp_service = [adsi]"IIS://localhost/SMTPSVC"
    if ($smtp_service.Name) {
        $smtp_server = @($smtp_service.psbase.Children | Where-Object { $_.Class -eq "IIsSmtpServer" })
    };
    if ($smtp_server.Count -gt 0) {
        foreach ($server in $smtp_server) {
            $server | Add-Member -MemberType NoteProperty -Name Path -Value $server.Path -Force
            $server | Add-Member -MemberType NoteProperty -Name Name -Value ($server.Path -replace "IIS://localhost/", "") -Force
        }
    }
    $smtp_server
    )'

    switch -regex ($ComputerName) {
        # local machine
        "^(localhost|\.|$($env:COMPUTERNAME))$" {
            $ComputerName = ($env:COMPUTERNAME).ToUpper()
            $is_localhost = $true

            try {
                $smtpSvc = (Get-Service smtpsvc -ErrorAction Stop)
                $metabase_file = "$env:SystemRoot\system32\inetsrv\metabase.xml"
            }
            catch {
                SayError "[$($ComputerName)] The SMTP Server is not installed on this computer. $($_)"
                return $null
            }


            try {
                $smtp_server = @(
                    Invoke-Expression $command -ErrorAction Stop
                )
            }
            catch {
                SayError $_.Exception.Message
                return $null
            }
        }

        default {
            # remote machine
            $is_localhost = $false

            try {
                # Get the SMTPSVC service object on the remote machine.
                $smtpSvc = (Invoke-Command -ComputerName $ComputerName -ScriptBlock { Get-Service smtpsvc -ErrorAction Stop } -ErrorAction Stop)
            }
            catch {
                SayError "[$($ComputerName)] The SMTP Server is not installed on this computer."
                return $null
            }

            try {
                # Get the system32 path on the remote machine.
                $system_root = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
                    "$($env:SystemRoot)\system32"
                } -ErrorAction Stop

                # Compose the metabase.xml full UNC file path.
                $metabase_file = "\\$($ComputerName)\$($system_root -replace ':','$')\inetsrv\metabase.xml"
            }
            catch {
                SayError "[$($ComputerName)] $($_.Exception.Message)"
                return $null
            }

            try {
                # Get the SMTP Virtual Server instances on the remote machine.
                $smtp_server = @(Invoke-Command -ComputerName $ComputerName -ScriptBlock {
                        Invoke-Expression $using:command -ErrorAction Stop
                    } -ErrorAction Stop)
            }
            catch {
                SayError "[$($ComputerName)] $($_.Exception.Message)"
                return $null
            }
        }
    }

    if (!(Test-Path $metabase_file)) {
        SayError "The metabase file path [$($metabase_file)] on computer [$($ComputerName)] is invalid or missing."
        return $null
    }

    try {
        [xml]$iis_metabase = Get-Content $metabase_file -ErrorAction Stop
    }
    catch {
        SayError $_.Exception.Message
        return $null
    }

    foreach ($server in $smtp_server) {
        $server_from_metabase = $iis_metabase.configuration.MBProperty.IIsSmtpServer | Where-Object { $server.Name -eq ($_.Location -replace '/lm/', '') }
        if ($server_from_metabase.RelayIpList) {
            $relay_ip_list = [System.Collections.Generic.List[string]]@()
            $octet_strings = ($server_from_metabase.RelayIpList.Substring(160) -split '(.{8})' | Where-Object { $_ -ne '' })
            foreach ($octet in $octet_strings) {
                $relay_ip_list.Add(($octet -split '(.{2})' | Where-Object { $_ -ne '' } | ForEach-Object { [convert]::ToInt32($_, 16) }) -join ".")
            }
            $server.RelayIpList = $relay_ip_list
        }
        else {
            $server.RelayIpList = @()
        }

        $server | Add-Member -MemberType NoteProperty -Name VirtualServerName -Value (($server.Name).ToUpper()) -Force
        $server | Add-Member -MemberType NoteProperty -Name VirtualServerDisplayName -Value ($server.ServerComment)[0] -Force
        if (!$server.ServerComment) {
            $server.ServerComment = "[SMTP Virtual Server #$(($server.Name -split '/')[-1])]"
            $server.VirtualServerDisplayName = $server.ServerComment
        }

        $server | Add-Member -MemberType NoteProperty -Name VirtualServerState -Value $null -Force
        $server | Add-Member -MemberType NoteProperty -Name ComputerName -Value $ComputerName -Force
        $server | Add-Member -MemberType NoteProperty -Name IsLocalHost -Value $is_localhost -Force
        $server | Add-Member -MemberType NoteProperty -Name SmtpServiceState -Value $smtpsvc.Status -Force

        ## If LogFileDirectory property exists
        if ($server.LogFileDirectory) {
            $server.LogFileDirectory = "$($server.LogFileDirectory)\$($server.VirtualServerName -replace '/','')"
        }

        ## If LogFileDirectory property does not exist.
        if (!($server | Get-Member -Name LogFileDirectory)) {
            $server | Add-Member -MemberType NoteProperty -Name LogFileDirectory -Value "$($iis_metabase.configuration.MBProperty.IIsSmtpService.LogFileDirectory)\$($server.VirtualServerName -replace '/','')" -Force
        }

        # VirtualServerState friendly value lookup table.
        $server_state_table = @{
            1 = 'Starting'
            2 = 'Started'
            3 = 'Stopping'
            4 = 'Stopped'
            5 = 'Pausing'
            6 = 'Paused'
            7 = 'Continuing'
        }

        $server.VirtualServerState = $(
            if ($server.ServerState) {
                # Lookup the virtual smtp server status
                $server_state_table[(($server.ServerState)[0])]
            }
            else {
                # This scenario applies whent the SMTPSVC service is not running.
                'Stopped'
            }
        )
    }

    $smtp_server
}