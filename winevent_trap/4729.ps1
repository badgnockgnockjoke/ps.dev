Param(
      		$MemberName, 
      		$TargetUserName,
			$SubjectUserName
 )
'[{0:yyyyMMdd--hh:mm:ss}] {1}' -f [datetime]::Now,'Group membership change - AD user removed from group' | Out-file c:\xiss\events.log -append
$p=@{
    UserChanged=$MemberName
	PersonMakingChanges=$SubjectUserName
	GroupName=$TargetUserName
}
$parms=New-Object PsObject -Property $P
$parms |Format-List |Out-String | Out-file c:\xiss\events.log -append

$subject='[{0:yyyyMMdd--hh:mm:ss}] {1}' -f [datetime]::Now,'AD Group membership change - AD user removed from group'
$body=$parms |Format-List |Out-String 

Send-MailMessage -to "your@email.com" -From "alerts@internal.lan" -Subject $subject -SmtpServer relay.internal.lan.local -Body $body