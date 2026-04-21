# Mail functions
function Send-ErrorNotification {
    param(
        [string]$Subject,
        [string]$Body
    )
    try {
        $smtp = New-Object Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
        $smtp.EnableSsl = $EnableSsl
        $smtp.UseDefaultCredentials = $false                 # must be set before Credentials
        $smtp.Credentials = New-Object System.Net.NetworkCredential($SmtpUser, $SmtpPassword)
        # For older .NET frameworks, specifying the target name may help
        $smtp.TargetName = "STARTTLS/smtp.yandex.ru"
        $smtp.Timeout = 10000                                 # 10 seconds timeout

        $mail = New-Object Net.Mail.MailMessage($MailFrom, $MailTo, $Subject, $Body)
        $mail.BodyEncoding = [System.Text.Encoding]::UTF8
        $mail.SubjectEncoding = [System.Text.Encoding]::UTF8

        $smtp.Send($mail)
        Write-Host "✅ Email sent successfully to $MailTo" -ForegroundColor Green
        return $true
    } catch {
        Write-Warning "❌ Failed to send email notification: $_"
        return $false
    }
}