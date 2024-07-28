Function Get-IISSmtpServer {
    [CmdletBinding()]
    param (
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]
        $ComputerName
    )

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

    if ($ComputerName) {
        $ComputerName = $ComputerName.ToUpper()
        try {
            $system_root = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
                "$($env:SystemRoot)\system32"
            } -ErrorAction Stop
            $metabase_file = "\\$($ComputerName)\$($system_root -replace ':','$')\inetsrv\metabase.xml"

            $smtp_server = @(Invoke-Command -ComputerName IISSMTP01 -ScriptBlock {
                    Invoke-Expression $using:command
                } -ErrorAction Stop)
        }
        catch {
            SayError "[$($ComputerName)] $($_.Exception.Message)"
            return $null
        }
    }

    $is_localhost = $false

    if (!$ComputerName -or $ComputerName -eq 'Localhost' -or $ComputerName -eq '.' -or $ComputerName -eq $env:COMPUTERNAME) {
        $metabase_file = "$env:SystemRoot\system32\inetsrv\metabase.xml"
        $ComputerName = ($env:COMPUTERNAME).ToUpper()
        $is_localhost = $true
        $smtp_server = @(
            Invoke-Expression $command
        )
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

    switch ($is_localhost) {
        $true { $service_state = (Get-Service SmtpSvc).Status }
        $false {
            $service_state = $(
                Invoke-Command -ComputerName $ComputerName -ScriptBlock {
                        (Get-Service SmtpSvc).Status
                }
            )
        }
        Default {}
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

        $server | Add-Member -MemberType NoteProperty -Name VirtualServerName -Value (($server.Name).ToUpper())
        $server | Add-Member -MemberType NoteProperty -Name VirtualServerDisplayName -Value ($server.ServerComment)[0]
        if (!$server.ServerComment) {
            $server.ServerComment = "[SMTP Virtual Server #$(($server.Location -split '/')[-1])]"
            $server.VirtualServerDisplayName = $server.ServerComment
        }

        $server | Add-Member -MemberType NoteProperty -Name VirtualServerState -Value $null
        $server | Add-Member -MemberType NoteProperty -Name ComputerName -Value $ComputerName
        $server | Add-Member -MemberType NoteProperty -Name IsLocalHost -Value $is_localhost
        $server | Add-Member -MemberType NoteProperty -Name SmtpServiceState -Value $service_state

        if ($server.LogFileDirectory) {
            $server.LogFileDirectory = "$($server.LogFileDirectory)\$($server.VirtualServerName -replace '/','')"
        }

        if (!($server | Get-Member -Name LogFileDirectory)) {
            $server | Add-Member -MemberType NoteProperty -Name LogFileDirectory -Value "$($iis_metabase.configuration.MBProperty.IIsSmtpService.LogFileDirectory)\$($server.VirtualServerName -replace '/','')"
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

        $server.VirtualServerState = $server_state_table[(($server.ServerState)[0])]
    }

    $smtp_server
}