#This script will email the ad audit trail at the end of the month and save it to the ad group audit folder also

#Copy events.log to the $outfile location and append date, then remove file
$Source = "C:\scripts\events.log"
$OutFile = "\\nas\_audit\AD_audit- $(get-date -f yyyy-MM-dd).txt"
Copy-Item $Source $OutFile

$body = "This is a backup of the AD audits performed over the month.  The file is backed up and emailed for archival purposes.  The current month's events are in c:\scripts\events.log on server.internal.lan"
Send-MailMessage -to "audits@email.com" -From "Internal Mailer <sysbot@internal.lan>" -Subject "Monthly AD Audit" -SmtpServer relay.internal.lan -Attachments $OutFile -Body $body

#Now clean up the old events.log and start over
Remove-Item C:\scripts\events.log
