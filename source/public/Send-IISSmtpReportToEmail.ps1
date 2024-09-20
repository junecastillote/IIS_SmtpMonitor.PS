Function Send-IISSmtpReportToEmail {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,
            ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [PSTypeNameAttribute('IISSmtpServerReport')]
        $InputObject,

        [Parameter(Mandatory)]
        [mailaddress]
        $MailFrom,

        [Parameter(Mandatory)]
        [mailaddress[]]
        $MailTo,

        [Parameter()]
        [mailaddress[]]
        $MailCc,

        [Parameter()]
        [mailaddress[]]
        $MailBcc,

        [Parameter(Mandatory)]
        [string]
        $SmtpServer,

        [Parameter()]
        [switch]
        $SmtpSSLEnabled,

        [Parameter()]
        [string]
        $SmtpServerPort = 25,

        [Parameter()]
        [pscredential]
        $SmtpServerCredential,

        [Parameter()]
        [switch]
        $SendOnIssueOnly
    )
    begin {
        $star_divider = ('*' * 70)

        $mail_prop = @{}

        $mail_prop.Add('From', $MailFrom)
        $mail_prop.Add('To', $MailTo)
        $mail_prop.Add('SmtpServer', $SmtpServer)
        $mail_prop.Add('Port', $SmtpServerPort)

        if ($SmtpServerCredential) {
            $mail_prop.Add('Credential', $SmtpServerCredential)
        }
        if ($SmtpSSLEnabled) {
            $mail_prop.Add('UseSSL', $SmtpServerCredential)
        }
        if ($MailCc) {
            $mail_prop.Add('Cc', $MailCc)
        }
        if ($MailBcc) {
            $mail_prop.Add('Bcc', $MailBcc)
        }

    }
    process {
        foreach ($item in $InputObject) {
            if ($SendOnIssueOnly -and !$item.Issues) {
                SayInfo "No issues to report. Email not sent."
                continue
            }

            if ($item.Issues) {
                $mail_prop.Add('Priority', 'High')
                $item.Title = "ALERT! $($item.Title)"
            }

            try {
                SayInfo "Sending email report."
                Send-MailMessage @mail_prop -Subject $item.Title -Body $item.HtmlContent -BodyAsHtml -ErrorAction Stop -WarningAction SilentlyContinue
                # SayInfo "Done."
            }
            catch {
                SayError "Failed to send email report. `n$star_divider`n$_$star_divider"
            }
        }
    }
    end {

    }
}
