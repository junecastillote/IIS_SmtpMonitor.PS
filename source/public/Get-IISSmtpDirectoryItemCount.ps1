Function Get-IISSmtpDirectoryItemCount {
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


        # [Parameter(Mandatory)]
        # [Parameter(ParameterSetName = 'Default')]
        # [Parameter(ParameterSetName = 'Name')]
        # [Parameter(ParameterSetName = 'Instance')]
        [ValidateSet(
            'DropDirectory',
            'QueueDirectory',
            'PickupDirectory',
            'BadMailDirectory'
        )]
        [String]
        $Directory
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

    # $properties = (@('ComputerName', 'ServerName'))
    # if ($Directory) {
    #     $properties += $Directory
    # }
    $dir_collection = Get-IISSmtpServerDirectory @param_collection -ErrorAction Stop
    foreach ($dir in $dir_collection) {

    }

}