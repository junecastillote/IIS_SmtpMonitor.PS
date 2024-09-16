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
        $QueueThreshold,

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

        $module_info = Get-Module $($MyInvocation.MyCommand.ModuleName)
        $html_template = Get-Content "$($module_Info.ModuleBase)\source\private\html_template.html" -Raw
        $html_template = $html_template.Replace(
            'vOrganizationName',
            $OrganizationName
        ).Replace(
            'vReportDate',
            $now
        )
        $html_smtp_server_section = Get-Content "$($module_Info.ModuleBase)\source\private\smtp_server_section.html" -Raw
        $html_smtp_instance_section = Get-Content "$($module_Info.ModuleBase)\source\private\smtp_instance_section.html" -Raw

        # $report_html_file = "$($OutputDirectory)\$($OrganizationName)_IISSmtpMonitor.PS_$(($now).ToString('yyyy-MM-dd_HH-mm-ss'))_report.html"
        # $teams_card_file = "$($OutputDirectory)\$($OrganizationName)_IISSmtpMonitor.PS_$(($now).ToString('yyyy-MM-dd_HH-mm-ss'))_report.json"

        $virtual_smtp_server_status_collection = [System.Collections.Generic.List[System.Object]]@()

        # if (!$Title) {
        #     $report_title = "[$($OrganizationName)] IIS SMTP Server Status Report"
        # }
        # else {
        #     $report_title = $Title
        # }

        if (!$SelectedFolder) {
            $SelectedFolder = @(
                'Queue',
                'Drop',
                'BadMail',
                'Pickup',
                'LogFile'
            )
        }
    }
    process {
        foreach ($computer_name in $ComputerName) {
            $virtual_smtp_server_status = @(Get-IISSmtpServerStatus -ComputerName $computer_name -AggregateByHost)
            if ($virtual_smtp_server_status) {
                $virtual_smtp_server_status_collection.AddRange($virtual_smtp_server_status)
            }
        }
    }
    end {
        if ($virtual_smtp_server_status_collection.Count -lt 1) {
            Continue
        }

        $report_item_html = [System.Collections.Generic.List[string]]@()
        foreach ($server_item in $virtual_smtp_server_status_collection) {
            $report_item_html.Add('<table style="border-collapse: collapse;">')

            #Region SMTP Server Section
            $current_server_section = $html_smtp_server_section
            $current_server_section = $current_server_section.Replace(
                'vComputerName', $server_item.ComputerName
            )

            if ($server_item.SmtpServiceState -ne 'Running') {
                $current_server_section = $current_server_section.Replace(
                    ';">vSmtpServiceState',
                    '; color: red; font-weight: bold;">' + $server_item.SmtpServiceState
                )
            }
            else {
                $current_server_section = $current_server_section.Replace(
                    'vSmtpServiceState',
                    $server_item.SmtpServiceState
                )
            }
            $report_item_html.Add($current_server_section)
            #EndRegion SMTP Server Section

            #Region SMTP Instance Section
            foreach ($instance_item in $server_item.VirtualSMTPServerCollection) {

                $current_instance_section = $html_smtp_instance_section

                if ($instance_item.VirtualServerState -ne 'Started') {
                    $current_instance_section = $current_instance_section.Replace(
                        ';">vVirtualServerState',
                        '; color: red; font-weight: bold;">' + $instance_item.VirtualServerState
                    )
                }
                else {
                    $current_instance_section = $current_instance_section.Replace(
                        'vVirtualServerState',
                        $instance_item.VirtualServerState
                    )
                }

                $queue = ($instance_item.Items | Where-Object { $_.Type -eq 'Queue' })
                $pickup = ($instance_item.Items | Where-Object { $_.Type -eq 'Pickup' })
                $badmail = ($instance_item.Items | Where-Object { $_.Type -eq 'BadMail' })
                $drop = ($instance_item.Items | Where-Object { $_.Type -eq 'Drop' })
                $logfile = ($instance_item.Items | Where-Object { $_.Type -eq 'LogFile' })

                $current_instance_section = $current_instance_section.Replace(
                    'vPathQueue', $queue.LocalPath
                ).Replace(
                    'vPathPickup', $pickup.LocalPath
                ).Replace(
                    'vPathBadMail', $badmail.LocalPath
                ).Replace(
                    'vPathDrop', $drop.LocalPath
                ).Replace(
                    'vPathLogFile', $logfile.LocalPath
                ).Replace(
                    'vVirtualServerDisplayName', $instance_item.VirtualServerDisplayName
                )

                if ($QueueThreshold -and $queue.TotalCount -gt $QueueThreshold ) {
                    $current_instance_section = $current_instance_section.Replace(
                        ';">vCountQueue',
                        '; color: red; font-weight: bold;">' + $queue.TotalCount
                    )
                }
                else {
                    $current_instance_section = $current_instance_section.Replace(
                        'vCountQueue',
                        $queue.TotalCount
                    )
                }

                if ($PickupThreshold -and $pickup.TotalCount -gt $PickupThreshold ) {
                    $current_instance_section = $current_instance_section.Replace(
                        ';">vCountPickup',
                        '; color: red; font-weight: bold;">' + $pickup.TotalCount
                    )
                }
                else {
                    $current_instance_section = $current_instance_section.Replace(
                        'vCountPickup',
                        $pickup.TotalCount
                    )
                }

                if ($BadMailThreshold -and $badmail.TotalCount -gt $BadMailThreshold ) {
                    $current_instance_section = $current_instance_section.Replace(
                        ';">vCountBadMail',
                        '; color: red; font-weight: bold;">' + $badmail.TotalCount
                    )
                }
                else {
                    $current_instance_section = $current_instance_section.Replace(
                        'vCountBadMail',
                        $badmail.TotalCount
                    )
                }

                if ($DropThreshold -and $drop.TotalCount -gt $DropThreshold ) {
                    $current_instance_section = $current_instance_section.Replace(
                        ';">vCountDrop',
                        '; color: red; font-weight: bold;">' + $drop.TotalCount
                    )
                }
                else {
                    $current_instance_section = $current_instance_section.Replace(
                        'vCountDrop',
                        $drop.TotalCount
                    )
                }

                if ($LogFileThreshold -and $logfile.TotalCount -gt $LogFileThreshold ) {
                    $current_instance_section = $current_instance_section.Replace(
                        ';">vCountLogFile',
                        '; color: red; font-weight: bold;">' + $logfile.TotalCount
                    )
                }
                else {
                    $current_instance_section = $current_instance_section.Replace(
                        'vCountLogFile',
                        $logfile.TotalCount
                    )
                }

                $report_item_html.Add($current_instance_section)
            }
            #EndRegion SMTP Instance Section
            $report_item_html.Add('</table><hr>')
        }
        $html_template = $html_template.Replace(
            '<!-- DATA -->',
            ($report_item_html -join "`n")
        )

        $html_template
    }
}