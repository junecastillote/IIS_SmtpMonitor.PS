Function Get-IISSmtpServerDirectory {
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

    # Compose Get-CimInstance parameters
    $param_collection = @{}

    if ($ComputerName) {
        $param_collection.Add(
            'ComputerName', $ComputerName
        )
    }

    if ($PSCmdlet.ParameterSetName -eq 'Name') {
        $param_collection.Add(
            'ServerName', $($ServerName)
        )
    }

    if ($PSCmdlet.ParameterSetName -eq 'Instance') {
        $param_collection.Add(
            'ServerInstance', $($ServerInstance)
        )
    }

    Get-IISSmtpServerSetting @param_collection -ErrorAction Stop |
    Select-Object @{
        n = 'ComputerName'
        e = { $_.PSComputerName }
    },
    @{
        n = 'ServerName'
        e = { $_.Name }
    }, *Directory*
}