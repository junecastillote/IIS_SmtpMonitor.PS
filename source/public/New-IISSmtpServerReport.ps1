Function New-IISSmtpServerStatusReport {
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
        [string]
        $OutputDirectory = $($env:temp),

        [Parameter(Mandatory)]
        [string]
        $OrganizationName,

        [Parameter()]
        [string]$Title,

        [Parameter()]
        [string]
        $HtmlReportFileName,

        [Parameter()]
        [ValidateRange(0, ([int]::MaxValue))]
        [int]
        $PickupThreshold,

        [Parameter()]
        [ValidateRange(0, ([int]::MaxValue))]
        [int]
        $BadMailThreshold,

        [Parameter()]
        [ValidateRange(0, ([int]::MaxValue))]
        [int]
        $DropThreshold,

        [Parameter()]
        [ValidateRange(0, ([int]::MaxValue))]
        [int]
        $LogFileThreshold
    )

    begin {

        $now = [datetime]::Now

        $module_info = Get-Module $($MyInvocation.MyCommand.ModuleName)\
        $html_template = "$($module_Info.ModuleBase)\source\private\email_template.html"
        $report_html_file = "$($OutputDirectory)\$($OrganizationName)_IISSmtpMonitor.PS_$(($now).ToString('yyyy-MM-dd_HH-mm-ss'))_report.html"
        $teams_card_file = "$($OutputDirectory)\$($OrganizationName)_IISSmtpMonitor.PS_$(($now).ToString('yyyy-MM-dd_HH-mm-ss'))_report.json"

        $virtual_smtp_server_status_collection = [System.Collections.Generic.List[System.Object]]@()

        if (!$Title) {
            $report_title = "[$($OrganizationName)] IIS SMTP Server Status Report"
        }
        else {
            $report_title = $Title
        }
    }
    process {
        foreach ($computer_name in $ComputerName) {
            $virtual_smtp_server_status = @(Get-IISSmtpServerStatus -ComputerName $computer_name)
            if ($virtual_smtp_server_status) {
                $virtual_smtp_server_status_collection.AddRange($virtual_smtp_server_status)
            }
        }
    }
    end {
        if ($virtual_smtp_server_status_collection.Count -lt 1) {
            Continue
        }

        foreach ($virtual_smtp_server_status in $virtual_smtp_server_status_collection) {
            ## HTML report

        }
    }
}