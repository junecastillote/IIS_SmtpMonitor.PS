Function Send-IISSmtpReportToTeams {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory,
            ValueFromPipeline)]
        [ValidateNotNullOrEmpty()]
        [PSTypeNameAttribute('IISSmtpServerReport')]
        $InputObject,

        [Parameter(Mandatory)]
        [string[]]
        $TeamsWebhookUrl
    )
    begin {

        $star_divider = ('*' * 70)
        $url_counter = 1
    }
    process {
        foreach ($item in $InputObject) {
            if (!$item.Issues) {
                SayInfo "No issues to report. Teams chat not sent."
                Continue
            }
            foreach ($url in $TeamsWebhookUrl) {
                SayInfo "Posting alert to Teams with URL [#$($url_counter)]"
                $Params = @{
                    "URI"         = $url
                    "Method"      = 'POST'
                    "Body"        = $item.TeamsCardContent
                    "ContentType" = 'application/json'
                }
                try {
                    Invoke-RestMethod @Params -ErrorAction Stop
                }
                catch {
                    SayError "Failed to post to channel. `n$star_divider`n$_$star_divider"
                }
                $url_counter++
            }
        }
    }
    end {

    }
}
