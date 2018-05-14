Function Send-TervisMailMessage {
    param(
        $To,
        $Cc,
        $Bcc,
        $Subject,
        $Body,
        $From,
        $Attachments,
        [Switch]$BodyAsHTML,
        [Switch]$UseAuthentication
    )
    $Parameters = $PSBoundParameters | ConvertFrom-PSBoundParameters -ExcludeProperty UseAuthentication -AsHashTable
    if ($UseAuthentication) {
        $Credential = Get-PasswordstatePassword -ID 3971 -AsCredential
        Send-MailMessage @Parameters -SmtpServer "smtp.office365.com" -Port "587" -Credential $Credential -UseSsl
    } else {
        Send-MailMessage @Parameters -SmtpServer tervis-com.mail.protection.outlook.com -UseSsl
    }
}