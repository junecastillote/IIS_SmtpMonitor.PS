Import-Module .\IIS_SmtpMonitor.PS.psd1 -Force

# =================================================================
# Create a new IIS SMTP server report
# =================================================================
$report_params = @{
    # IIS Smtp server hostname(s) to check.
    ComputerName     = @('SMTP1', 'SMTP2')
    # ComputerName     = @('SMTP1')

    # Organization name to appear in the report. Mandatory.
    OrganizationName = 'PoshLab'

    # Title it optional.
    # Default = "[OrganizationName] IIS SMTP Server Status Report"
    Title            = ''

    # The output directory for the resulting report. Optional.
    # Default = $env:temp
    OutputDirectory  = ''

    # Specify each threshold. 0 means no threshold.
    QueueThreshold   = 50
    PickupThreshold  = 50
    BadMailThreshold = 50
    DropThreshold    = 50
    LogFileThreshold = 50

    # Specify $true or $false, whether to open the HTML report at the end.
    # Do not enable if running the script unattended (automation).
    OpenHtmlReport   = $false
}

# Create the report
$report = New-IISSmtpServerStatusReport @report_params
# =================================================================

# =================================================================
# Send the email report
# =================================================================
$mail_params = @{
    InputObject          = $report

    # The email sender address.
    MailFrom             = 'IIS SMTP Server Monitor <smtp_monitor.no_reply@mg.poshlab.xyz>'

    # TO email recipients. Mandatory.
    MailTo               = @('smtp_admin@poshlab.xyz')

    # CC and BCC email recipients. Optional.
    MailCc               = @()
    MailBcc              = @()

    # SMTP relay server name or IP.
    SmtpServer           = 'SMTP1'

    # SMTP relay server port. Optional.
    # Default = 25.
    SmtpServerPort       = 25

    # SMTP relay server credential if required.
    SmtpServerCredential = $null

    # Enable TLS/SSL if required.
    SmtpSSLEnabled       = $false

    # Specify $true to send the report only when there are issues detected (alert mode).
    # Specify $false to send the report anyway (report mode).
    SendOnIssueOnly      = $true
}

# Send the report.
Send-IISSmtpReportToEmail @mail_params
# =================================================================

# =================================================================
# Send to Teams.
# =================================================================
$teams_webhook_url = @('')
if ($teams_webhook_url) {
    $report | Send-IISSmtpReportToTeams -TeamsWebhookUrl $teams_webhook_url
}
# =================================================================