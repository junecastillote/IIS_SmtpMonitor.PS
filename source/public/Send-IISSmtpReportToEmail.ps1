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
        if ($SendOnIssueOnly -and !$InputObject.Issues) {
            SayInfo "No issues to report. Email not sent."
            Continue
        }

        if ($InputObject.Issues) {
            $mail_prop.Add('Priority', 'High')
            $InputObject.Title = "ALERT! $($InputObject.Title)"
        }

        try {
            Send-MailMessage @mail_prop -Subject $InputObject.Title -Body $InputObject.HtmlContent -BodyAsHtml -ErrorAction Stop
        }
        catch {
            SayError "Failed to send email report. `n$star_divider`n$_$star_divider"
        }
    }
    end {

    }
}
