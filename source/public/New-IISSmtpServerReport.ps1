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
        $LogFileThreshold,

        [Parameter()]
        [switch]
        $OpenHtmlReport
    )

    begin {

        $now = [datetime]::Now

        $module_info = Get-Module $($MyInvocation.MyCommand.ModuleName)
        $html_report = Get-Content "$($module_Info.ModuleBase)\source\private\html_template.html" -Raw

        $virtual_smtp_server_status_collection = [System.Collections.Generic.List[System.Object]]@()

        if (!$Title) {
            $report_title = "[$($OrganizationName)] IIS SMTP Server Status Report"
        }
        else {
            $report_title = $Title
        }


        $html_report = $html_report.Replace(
            'vOrganizationName',
            $OrganizationName
        ).Replace(
            'vReportDate',
            $now
        ).Replace(
            'vTitle',
            $report_title
        ).Replace(
            'vModuleInfo',
            '<a href="' + "$($module_info.ProjectUri.ToString())" + '" target="_blank">' + "$($module_info.Name) v$($module_info.Version.ToString())" + '</a>'
        )
        $html_smtp_server_section = Get-Content "$($module_Info.ModuleBase)\source\private\smtp_server_section.html" -Raw
        $html_smtp_instance_section = Get-Content "$($module_Info.ModuleBase)\source\private\smtp_instance_section.html" -Raw
        $html_smtp_issue_section = Get-Content "$($module_Info.ModuleBase)\source\private\smtp_issue_section.html" -Raw

        $issue_collection = [System.Collections.Generic.List[string]]@()

        $report_html_file = "$($OutputDirectory)\$($OrganizationName)_IISSmtpMonitor.PS_$(($now).ToString('yyyy-MM-dd_HH-mm-ss'))_report.html"
        # $teams_card_file = "$($OutputDirectory)\$($OrganizationName)_IISSmtpMonitor.PS_$(($now).ToString('yyyy-MM-dd_HH-mm-ss'))_report.json"



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

        $html_body = [System.Collections.Generic.List[string]]@()

        foreach ($server_item in $virtual_smtp_server_status_collection) {
            $html_body.Add('<table style="border-collapse: collapse;">')

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
                $issue_collection.Add("$($server_item.ComputerName):SMTP service is $($server_item.SmtpServiceState).")
            }
            else {
                $current_server_section = $current_server_section.Replace(
                    ';">vSmtpServiceState',
                    '; color: green; font-weight: bold;">' + $server_item.SmtpServiceState
                )
            }
            $html_body.Add($current_server_section)
            #EndRegion SMTP Server Section

            #Region SMTP Instance Section
            foreach ($instance_item in $server_item.VirtualSMTPServerCollection) {

                $current_instance_section = $html_smtp_instance_section

                if ($instance_item.VirtualServerState -ne 'Started') {
                    $current_instance_section = $current_instance_section.Replace(
                        ';">vVirtualServerState',
                        '; color: red; font-weight: bold;">' + $instance_item.VirtualServerState
                    )
                    $issue_collection.Add("$($server_item.ComputerName):$($instance_item.VirtualServerDisplayName) - Virtual server instance is $($instance_item.VirtualServerState).")
                }
                else {
                    $current_instance_section = $current_instance_section.Replace(
                        ';">vVirtualServerState',
                        '; color: green; font-weight: bold;">' + $instance_item.VirtualServerState
                    )
                }

                $queue = ($instance_item.Folders | Where-Object { $_.Type -eq 'Queue' })
                $pickup = ($instance_item.Folders | Where-Object { $_.Type -eq 'Pickup' })
                $badmail = ($instance_item.Folders | Where-Object { $_.Type -eq 'BadMail' })
                $drop = ($instance_item.Folders | Where-Object { $_.Type -eq 'Drop' })
                $logfile = ($instance_item.Folders | Where-Object { $_.Type -eq 'LogFile' })

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
                ).Replace(
                    'vSizeQueue', ([System.Math]::Round(($queue.TotalSize / 1MB), 2)).ToString("N2")
                ).Replace(
                    'vSizePickup', ([System.Math]::Round(($pickup.TotalSize / 1MB), 2)).ToString("N2")
                ).Replace(
                    'vSizeBadMail', ([System.Math]::Round(($badmail.TotalSize / 1MB), 2)).ToString("N2")
                ).Replace(
                    'vSizeDrop', ([System.Math]::Round(($drop.TotalSize / 1MB), 2)).ToString("N2")
                ).Replace(
                    'vSizeLogFile', ([System.Math]::Round(($logfile.TotalSize / 1MB), 2)).ToString("N2")
                )

                # Queue
                if ($QueueThreshold -and $queue.TotalCount -gt $QueueThreshold ) {
                    $current_instance_section = $current_instance_section.Replace(
                        ';">vCountQueue',
                        '; color: red; font-weight: bold;">' + ($queue.TotalCount).ToString("N0")
                    )
                    $issue_collection.Add("$($server_item.ComputerName):$($instance_item.VirtualServerDisplayName) - Queue count is $($queue.TotalCount.ToString("N0")). Threshold is $($QueueThreshold.ToString("N0")).")
                }
                elseif (!$QueueThreshold) {
                    $current_instance_section = $current_instance_section.Replace(
                        'vCountQueue',
                        ($queue.TotalCount).ToString("N0")
                    )
                }
                else {
                    $current_instance_section = $current_instance_section.Replace(
                        ';">vCountQueue',
                        '; color: green; font-weight: bold;">' + ($queue.TotalCount).ToString("N0")
                    )
                }

                # Pickup
                if ($PickupThreshold -and $pickup.TotalCount -gt $PickupThreshold ) {
                    $current_instance_section = $current_instance_section.Replace(
                        ';">vCountPickup',
                        '; color: red; font-weight: bold;">' + ($pickup.TotalCount).ToString("N0")
                    )
                    $issue_collection.Add("$($server_item.ComputerName):$($instance_item.VirtualServerDisplayName) - Pickup count is $($pickup.TotalCount.ToString("N0")). Threshold is $($PickupThreshold.ToString("N0")).")
                }
                elseif (!$PickupThreshold) {
                    $current_instance_section = $current_instance_section.Replace(
                        'vCountPickup',
                        ($pickup.TotalCount).ToString("N0")
                    )
                }
                else {
                    $current_instance_section = $current_instance_section.Replace(
                        ';">vCountPickup',
                        '; color: green; font-weight: bold;">' + ($pickup.TotalCount).ToString("N0")
                    )
                }

                # BadMail
                if ($BadMailThreshold -and $badmail.TotalCount -gt $BadMailThreshold ) {
                    $current_instance_section = $current_instance_section.Replace(
                        ';">vCountBadMail',
                        '; color: red; font-weight: bold;">' + ($badmail.TotalCount).ToString("N0")
                    )
                    $issue_collection.Add("$($server_item.ComputerName):$($instance_item.VirtualServerDisplayName) - BadMail count is $($badmail.TotalCount.ToString("N0")). Threshold is $($BadMailThreshold.ToString("N0")).")
                }
                elseif (!$BadMailThreshold) {
                    $current_instance_section = $current_instance_section.Replace(
                        'vCountBadMail',
                        ($badmail.TotalCount).ToString("N0")
                    )
                }
                else {
                    $current_instance_section = $current_instance_section.Replace(
                        ';">vCountBadMail',
                        '; color: green; font-weight: bold;">' + ($badmail.TotalCount).ToString("N0")
                    )
                }

                # Drop
                if ($DropThreshold -and $drop.TotalCount -gt $DropThreshold ) {
                    $current_instance_section = $current_instance_section.Replace(
                        ';">vCountDrop',
                        '; color: red; font-weight: bold;">' + ($drop.TotalCount).ToString("N0")
                    )
                    $issue_collection.Add("$($server_item.ComputerName):$($instance_item.VirtualServerDisplayName) - Drop count is $($drop.TotalCount.ToString("N0")). Threshold is $($DropThreshold.ToString("N0")).")
                }
                elseif (!$DropThreshold) {
                    $current_instance_section = $current_instance_section.Replace(
                        'vCountDrop',
                        ($drop.TotalCount).ToString("N0")
                    )
                }
                else {
                    $current_instance_section = $current_instance_section.Replace(
                        ';">vCountDrop',
                        '; color: green; font-weight: bold;">' + ($drop.TotalCount).ToString("N0")
                    )
                }

                # LogFile
                if ($LogFileThreshold -and $logfile.TotalCount -gt $LogFileThreshold ) {
                    $current_instance_section = $current_instance_section.Replace(
                        ';">vCountLogFile',
                        '; color: red; font-weight: bold;">' + ($logfile.TotalCount).ToString("N0")
                    )
                    $issue_collection.Add("$($server_item.ComputerName):$($instance_item.VirtualServerDisplayName) - LogFile count is $($logfile.TotalCount.ToString("N0")). Threshold is $($LogFileThreshold.ToString("N0")).")
                }
                elseif (!$LogFileThreshold) {
                    $current_instance_section = $current_instance_section.Replace(
                        'vCountLogFile',
                        ($logfile.TotalCount).ToString("N0")
                    )
                }
                else {

                    $current_instance_section = $current_instance_section.Replace(
                        ';">vCountLogFile',
                        '; color: green; font-weight: bold;">' + ($logfile.TotalCount).ToString("N0")
                    )
                }

                $html_body.Add($current_instance_section)
            }
            $html_body.Add('</table><hr>')
            #EndRegion SMTP Instance Section
        }

        $html_report = $html_report.Replace(
            '<!-- DATA -->',
            ($html_body -join "`n")
        )

        if ($issue_collection.Count -gt 1) {
            $issue_section = [System.Collections.Generic.List[string]]@()
            $issue_section.Add('<table style="border-collapse: collapse;">')
            $issue_section.Add('<tr><th style="border: none; padding: 5px; text-align: left; font-size: larger;" colspan="2">ISSUE LIST</th></tr>')

            foreach ($issue in $issue_collection) {
                $current_issue = $html_smtp_issue_section
                $current_issue = $current_issue.Replace(
                    'vTarget', "$($issue.Split(':')[0])"
                ).Replace(
                    'vIssue', "$($issue.Split(':')[-1])"
                )

                $issue_section.Add($current_issue)
            }
            $issue_section.Add('</table><hr>')

            $html_report = $html_report.Replace(
                '<!-- ISSUE -->',
                ($issue_section -join "`n")
            )
        }

        $html_report | Out-File $report_html_file -Force

        $result = [PSCustomObject]@{
            PSTypeName          = 'IISSmtpServerReport'
            ReportGeneratedDate = $now
            OrganizationName    = $OrganizationName
            Title               = $report_title
            Issues              = ($issue_collection -join "`n")
            HtmlFileName        = (Resolve-Path $report_html_file).Path
            HtmlContent         = $html_report
            TeamsCardFileName   = ''
            TeamsCardContent    = ''
        }

        $visible_properties = [string[]]@('Title', 'ReportGeneratedDate', 'HtmlFileName', 'TeamsCardFileName')
        [Management.Automation.PSMemberInfo[]]$default_properties = [System.Management.Automation.PSPropertySet]::new('DefaultDisplayPropertySet', $visible_properties)
        $result | Add-Member -MemberType MemberSet -Name PSStandardMembers -Value $default_properties

        if ($OpenHtmlReport) {
            Invoke-Item $report.HtmlFileName
        }

        return $result
    }
}