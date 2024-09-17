Function New-IISSmtpServerStatusReport {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory, ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [string[]]
        $ComputerName,

        # [Parameter()]
        # [ValidateSet(
        #     'Queue',
        #     'Drop',
        #     'BadMail',
        #     'Pickup',
        #     'LogFile'
        # )]
        # [string[]]
        # $SelectedFolder,

        [Parameter()]
        [string]
        $OutputDirectory = $($env:temp),

        [Parameter(Mandatory)]
        [string]
        $OrganizationName,

        [Parameter()]
        [string]$Title,

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
        $html_smtp_issue_section = Get-Content "$($module_Info.ModuleBase)\source\private\smtp_issue_section.html" -Raw

        $issue_collection = [System.Collections.Generic.List[string]]@()

        # $report_html_file = "$($OutputDirectory)\$($OrganizationName)_IISSmtpMonitor.PS_$(($now).ToString('yyyy-MM-dd_HH-mm-ss'))_report.html"
        # $teams_card_file = "$($OutputDirectory)\$($OrganizationName)_IISSmtpMonitor.PS_$(($now).ToString('yyyy-MM-dd_HH-mm-ss'))_report.json"

        $virtual_smtp_server_status_collection = [System.Collections.Generic.List[System.Object]]@()

        # if (!$Title) {
        #     $report_title = "[$($OrganizationName)] IIS SMTP Server Status Report"
        # }
        # else {
        #     $report_title = $Title
        # }

        # if (!$SelectedFolder) {
        #     $SelectedFolder = @(
        #         'Queue',
        #         'Drop',
        #         'BadMail',
        #         'Pickup',
        #         'LogFile'
        #     )
        # }
    }
    process {
        foreach ($computer_name in $ComputerName) {
            # $virtual_smtp_server_status = @(Get-IISSmtpServerStatus -ComputerName $computer_name -AggregateByHost -SelectedFolder $SelectedFolder)
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

        $report_html = [System.Collections.Generic.List[string]]@()

        foreach ($server_item in $virtual_smtp_server_status_collection) {
            $report_html.Add('<table style="border-collapse: collapse;">')

            #Region SMTP Server Section
            $current_server_section = $html_smtp_server_section
            $current_server_section = $current_server_section.Replace(
                'vComputerName', "SERVER: $($server_item.ComputerName)"
            )

            if ($server_item.SmtpServiceState -ne 'Running') {
                $current_server_section = $current_server_section.Replace(
                    ';">vSmtpServiceState',
                    '; color: red; font-weight: bold;">' + $server_item.SmtpServiceState
                )
                $issue_collection.Add("$($server_item.ComputerName):The SMTP service is <b>[$($server_item.SmtpServiceState)]</b>.")
            }
            else {
                $current_server_section = $current_server_section.Replace(
                    'vSmtpServiceState',
                    $server_item.SmtpServiceState
                )
            }
            $report_html.Add($current_server_section)
            #EndRegion SMTP Server Section

            #Region SMTP Instance Section
            foreach ($instance_item in $server_item.VirtualSMTPServerCollection) {

                $current_instance_section = $html_smtp_instance_section

                if ($instance_item.VirtualServerState -ne 'Started') {
                    $current_instance_section = $current_instance_section.Replace(
                        ';">vVirtualServerState',
                        '; color: red; font-weight: bold;">' + $instance_item.VirtualServerState
                    )
                    $issue_collection.Add("$($server_item.ComputerName):$($instance_item.VirtualServerDisplayName) - The SMTP virtual server instance is <b>[$($instance_item.VirtualServerState)]</b>.")
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
                        '; color: red; font-weight: bold;">' + ($queue.TotalCount).ToString("N0")
                    )
                    $issue_collection.Add("$($server_item.ComputerName):$($instance_item.VirtualServerDisplayName) - <b>Queue</b> count is <b>$($queue.TotalCount)</b>. Threshold is $($QueueThreshold).")
                }
                else {
                    $current_instance_section = $current_instance_section.Replace(
                        'vCountQueue',
                        ($queue.TotalCount).ToString("N0")
                    )
                }

                if ($PickupThreshold -and $pickup.TotalCount -gt $PickupThreshold ) {
                    $current_instance_section = $current_instance_section.Replace(
                        ';">vCountPickup',
                        '; color: red; font-weight: bold;">' + ($pickup.TotalCount).ToString("N0")
                    )
                    $issue_collection.Add("$($server_item.ComputerName):$($instance_item.VirtualServerDisplayName) - <b>Pickup</b> count is <b>$($pickup.TotalCount)</b>. Threshold is $($PickupThreshold).")
                }
                else {
                    $current_instance_section = $current_instance_section.Replace(
                        'vCountPickup',
                        ($pickup.TotalCount).ToString("N0")
                    )
                }

                if ($BadMailThreshold -and $badmail.TotalCount -gt $BadMailThreshold ) {
                    $current_instance_section = $current_instance_section.Replace(
                        ';">vCountBadMail',
                        '; color: red; font-weight: bold;">' + ($badmail.TotalCount).ToString("N0")
                    )
                    $issue_collection.Add("$($server_item.ComputerName):$($instance_item.VirtualServerDisplayName) - <b>BadMail</b> count is <b>$($badmail.TotalCount)</b>. Threshold is $($BadMailThreshold).")
                }
                else {
                    $current_instance_section = $current_instance_section.Replace(
                        'vCountBadMail',
                        ($badmail.TotalCount).ToString("N0")
                    )
                }

                if ($DropThreshold -and $drop.TotalCount -gt $DropThreshold ) {
                    $current_instance_section = $current_instance_section.Replace(
                        ';">vCountDrop',
                        '; color: red; font-weight: bold;">' + ($drop.TotalCount).ToString("N0")
                    )
                    $issue_collection.Add("$($server_item.ComputerName):$($instance_item.VirtualServerDisplayName) - <b>Drop</b> count is <b>$($drop.TotalCount)</b>. Threshold is $($DropThreshold).")
                }
                else {
                    $current_instance_section = $current_instance_section.Replace(
                        'vCountDrop',
                        ($drop.TotalCount).ToString("N0")
                    )
                }

                if ($LogFileThreshold -and $logfile.TotalCount -gt $LogFileThreshold ) {
                    $current_instance_section = $current_instance_section.Replace(
                        ';">vCountLogFile',
                        '; color: red; font-weight: bold;">' + ($logfile.TotalCount).ToString("N0")
                    )
                    $issue_collection.Add("$($server_item.ComputerName):$($instance_item.VirtualServerDisplayName) - <b>LogFile</b> count is <b>$($logfile.TotalCount)</b>. Threshold is $($LogFileThreshold).")
                }
                else {
                    $current_instance_section = $current_instance_section.Replace(
                        'vCountLogFile',
                        ($logfile.TotalCount).ToString("N0")
                    )
                }

                $report_html.Add($current_instance_section)
            }
            #EndRegion SMTP Instance Section
            $report_html.Add('</table><hr>')
        }
        $html_template = $html_template.Replace(
            '<!-- DATA -->',
            ($report_html -join "`n")
        )

        if ($issue_collection.Count -gt 1) {
            $issue_section = [System.Collections.Generic.List[string]]@()
            $issue_section.Add('<table style="border-collapse: collapse;">')
            # $issue_section.Add('<tr><th style="border: 1px solid #dddddd; padding: 5px; text-align: left; font-size: larger;" colspan="2">ISSUE LIST</th></tr>')
            $issue_section.Add('<tr><th style="border: none; padding: 5px; text-align: left; font-size: larger;" colspan="2">ISSUE LIST</th></tr>')

            foreach ($issue in $issue_collection) {
                # $issue | Out-Default
                $current_issue = $html_smtp_issue_section
                $current_issue = $current_issue.Replace(
                    'vTarget', "$($issue.Split(':')[0])"
                ).Replace(
                    'vIssue', "$($issue.Split(':')[-1])"
                )

                $issue_section.Add($current_issue)
            }
            $issue_section.Add('</table><hr>')

            $html_template = $html_template.Replace(
                '<!-- ISSUE -->',
            ($issue_section -join "`n")
            )
        }

        $html_template
    }
}