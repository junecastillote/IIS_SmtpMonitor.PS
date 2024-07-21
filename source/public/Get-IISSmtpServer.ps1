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
        $ServerInstance
    )

    if (!(IsFeatureInstalled)) {
        return $null
    }

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

    # Compose Get-CimInstance parameters
    $param_collection = @{
        NameSpace = 'root/MicrosoftIISv2'
        ClassName = 'IIsSmtpServer'
    }

    if ($ComputerName) {
        $param_collection.Add(
            'ComputerName', $ComputerName
        )
    }

    if ($PSCmdlet.ParameterSetName -eq 'Name') {
        $param_collection.Add(
            'Filter', "name = '$($ServerName)'"
        )
    }

    if ($PSCmdlet.ParameterSetName -eq 'Instance') {
        $param_collection.Add(
            'Filter', "name = 'SmtpSvc/$($ServerInstance)'"
        )
    }

    try {
        $smtp_server_collection = @(Get-CimInstance @param_collection -ErrorAction Stop)

        if (!$smtp_server_collection) {
            return $null
        }

        $property_names = ($smtp_server_collection | Get-Member -MemberType Property).Name

        foreach ($current_smtp_server in $smtp_server_collection) {
            [PSCustomObject]$(
                $result_object = @{}

                foreach ($property in $property_names) {
                    $result_object.Add($property, ($current_smtp_server.$property))
                }

                # Fill in the PSComputerName with the localhost name if empty.
                if (!$current_smtp_server.PSComputerName) {
                    $result_object.PSComputerName = $env:COMPUTERNAME
                }

                # Replace the ServerState int value with the friendly name.
                $result_object.ServerState = $server_state_table[($current_smtp_server.ServerState)]
                $result_object
            )
        }
    }
    catch {
        SayError $_.Exception.Message
        return $null
    }
}