Function Get-IISSmtpServer {
    [CmdletBinding()]
    param (
        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'Name')]
        [Parameter(ParameterSetName = 'Instance')]
        [ValidateNotNullOrEmpty()]
        [string]
        $ComputerName,

        [Parameter(
            Mandatory,
            ParameterSetName = 'Name'
        )]
        [ValidateNotNullOrEmpty()]
        [string]
        $ServerName,

        [Parameter(
            Mandatory,
            ParameterSetName = 'Instance'
        )]
        [ValidateNotNullOrEmpty()]
        [int]
        $ServerInstance,

        [Parameter()]
        [switch]
        $Status
    )



    # Compose Get-CimInstance parameters

    # $adsi_path = "IIS://localhost/SMTPSVC"

    if ($ComputerName) {
        try {
            $system_root = ((Get-CimInstance -ComputerName $ComputerName -ClassName Win32_OperatingSystem -ErrorAction Stop).SystemDirectory -replace ':', '$')
            $metabase_file = "\\$($ComputerName)\$($system_root)\inetsrv\metabase.xml"
        }
        catch {
            SayError $_.Exception.Message
            SayError "Could not automatically determine the metabase path on the remote remopte computer [$($ComputerName)]. Will assume the default path [$("$env:SystemRoot\system32\inetsrv\metabase.xml")] instead."
            $metabase_file = "\\$($ComputerName)\C$\Windows\system32\inetsrv\metabase.xml"
        }
    }

    if (!$ComputerName) {
        $metabase_file = "$env:SystemRoot\system32\inetsrv\metabase.xml"
        $ComputerName = $env:COMPUTERNAME
        $local_host = $true
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

    switch ($PSCmdlet.ParameterSetName) {
        'Name' {
            [System.Object]$smtp_server = $iis_metabase.configuration.MBProperty.IIsSmtpServer | Where-Object {
                $_.Location -eq "/LM/$($ServerName)"
            }
        }
        'Instance' {
            [System.Object]$smtp_server = $iis_metabase.configuration.MBProperty.IIsSmtpServer | Where-Object {
                $_.Location -eq "/LM/SmtpSvc/$($ServerInstance)"
            }
        }

        Default {
            [System.Object]$smtp_server = $iis_metabase.configuration.MBProperty.IIsSmtpServer
        }
    }

    foreach ($server in $smtp_server) {
        if ($server.RelayIpList) {
            $relay_ip_list = [System.Collections.Generic.List[string]]@()
            $octet_strings = ($server.RelayIpList.Substring(160) -split '(.{8})' | Where-Object { $_ -ne '' })
            foreach ($octet in $octet_strings) {
                $relay_ip_list.Add(($octet -split '(.{2})' | Where-Object { $_ -ne '' } | ForEach-Object { [convert]::ToInt32($_, 16) }) -join ".")
            }
            $server | Add-Member -MemberType NoteProperty -Name RelayIps -Value $relay_ip_list
        }
        else {
            $server | Add-Member -MemberType NoteProperty -Name RelayIpList -Value $relay_ip_list
        }

        $server | Add-Member -MemberType NoteProperty -Name ServerName -Value ($server.Location -replace '/lm/', '')
        $server | Add-Member -MemberType NoteProperty -Name ServerDisplayName -Value $server.ServerComment
        if (!$server.ServerComment) {
            $server.ServerComment = "[SMTP Virtual Server #$(($server.Location -split '/')[-1])]"
            $server.ServerDisplayName = $server.ServerComment
        }

        $server | Add-Member -MemberType NoteProperty -Name ServerState -Value $null
        $server | Add-Member -MemberType NoteProperty -Name ComputerName -Value $ComputerName

        if ($Status) {
            # ServerState friendly value lookup table.
            $server_state_table = @{
                1 = 'Starting'
                2 = 'Started'
                3 = 'Stopping'
                4 = 'Stopped'
                5 = 'Pausing'
                6 = 'Paused'
                7 = 'Continuing'
            }

            if ($local_host) {
                $server_state = $server_state_table[$(([adsi]"IIS://localhost/$($server.ServerName)").ServerState)]
            }
            else {
                $server_state = $server_state_table[
                $(
                    Invoke-Command -ComputerName $ComputerName -ScriptBlock {
                        (([adsi]"IIS://localhost/$($using:server.ServerName)").ServerState)
                    }
                )
                ]
            }

            $server.ServerState = $server_state
        }
    }

    $smtp_server
    # $smtp_server | ForEach-Object {
    #     $relay_ip_list = [System.Collections.Generic.List[string]]@()
    #     $octet_strings = ($_.RelayIpList -split '(.{8})' | Where-Object { $_ -ne '' })

    #     foreach ($octet in $octet_strings) {
    #         $relay_ip_list.Add(($octet -split '(.{2})' | Where-Object { $_ -ne '' } | ForEach-Object { [convert]::ToInt32($_, 16) }) -join ".")
    #     }
    # }
}