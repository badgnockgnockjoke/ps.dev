#######################################################
#
#   DMARC failure report (build 20231109)
#   Please install Outlook 2016 or higher, checks Exchange/O365 
#   Script will unarchive zip/tgz dmarc reports, search 
#   for 'fail', and email the results.  Script will clean
#   temporary files if desired.
#
#######################################################



$delFile = 1     ##change this to 0 and it will leave the original archives in place
$debugPrint = 1  ##do you want to print debugging information?

$OutlookEmail="Name of Mailbox"
$OutlookEmailInbox ="Inbox"
$OutlookEmailDestination = "dmarc"

$filepath = "c:\_dmarcfail"  ##folder to save attachments to
$date = Get-Date
$date = $date.ToString('yyyyMMddhhmm')
$fLogName = "dmarcfailscript-$date.txt"
$analyzepath = (join-path $filepath $date)
$OutLog = "c:\_dmarcfail\$fLogName" ##output log location

New-Item -ItemType Directory -Path  $analyzepath -ErrorAction Inquire
New-Item -Path $OutLog -ItemType File -Value "Initializing log..."

if ($debugPrint -eq 1) { Write-Output "Starting variables filepath, analyzepath, and OutLog:  $filepath | $analyzepath | $OutLog" | Add-Content $OutLog }
Write-Output "============================================"  | Add-Content $OutLog
Write-Output "Log for checking DMARC reports for failures"  | Add-Content $OutLog
Write-Output "DMARC failure report (build 20231109)"   | Add-Content $OutLog
Write-Output "============================================"  | Add-Content $OutLog

##define new object and assign Outlook 15 COM object
#Add-Type -assembly "Microsoft.Office.Interop.Outlook"
$Outlook = New-Object -ComObject Outlook.Application
$mapi = $Outlook.GetNamespace("MAPI")
$fldr = $mapi.Stores[$OutlookEmail].GetRootFolder().Folders($OutlookEmailInbox)
$destFldr = $fldr.Folders($OutlookEmailDestination)
$AllEmails = $fldr.Items.Restrict("[Unread]=True")
$c1 = $fldr.items.Count
$c2 = $AllEmails.Count

Write-Output "Inbox count / Unread : $c1 / $c2" | Add-Content $OutLog
if ($debugPrint -eq 1) { Write-Output "Filtering the messages..."  | Add-Content $OutLog}

$emails = $AllEmails | Where-Object {
    $_.SenderEmailAddress -like "*dmarc*" -or
    $_.Sender -like "*dmarc*" -or   
    $_.Subject -like "*dmarc*" -or
    $_.Body -like "*dmarc*"
}

$tempEmails = @()

##if emails count equal 0 then exit
if ($emails.Count -eq 0) {
    Write-Output "No emails to process"
    if ($delFile -eq 1) {
        Write-Output "Deleting old $OutLog"
        remove-item $OutLog -Force
    }
    Write-Output "There are no emails to process, exiting..."
    exit
}


Write-Output "Saving all DMARC reports to disk..."  | Add-Content $OutLog        
foreach ($email in $emails) {   
    $email.Attachments | ForEach-Object { 
        $fikeNameString = $_.FileName             
        Write-Output "Saving file: " (Join-Path $analyzepath $fikeNameString)  | Add-Content $OutLog
        if ($fikeNameString -like '*zip*' -or $fikeNameString -like '*gz*' -or $fikeNameString -like '*tgz*') {
            Write-Output "Archive attachment found, saving to disk"    | Add-Content $OutLog                     
            $_.saveasfile((Join-Path $analyzepath $fikeNameString)) 
        }                            
    }        
    $tempEmails += $email 
}

$c3 = $tempEmails.Count
#Write-Output ""  | Add-Content $OutLog        
$tempEmails.Move($destFldr)

##Find all and extract from supported archives - picked common ones
Write-Output "Extracting major known archive types (zip/gz/tgz) in $analyzepath" | Add-Content $OutLog
Get-ChildItem -Path (Join-Path $analyzepath *.zip) | ForEach-Object { C:\xiss\7z.exe e -bb0 -bd  $_.FullName "-o$($_.Directory)" }
Get-ChildItem -Path (Join-Path $analyzepath *.gz) | ForEach-Object { C:\xiss\7z.exe e -bb0 -bd  $_.FullName "-o$($_.Directory)" }
Get-ChildItem -Path (Join-Path $analyzepath *.tgz) | ForEach-Object { C:\xiss\7z.exe e -bb0 -bd $_.FullName "-o$($_.Directory)" }

$listFailFile = @()
##List and print names of files that contained failures
Write-Output "Parsing for the word 'fail'."  | Add-Content $OutLog
Get-ChildItem -Path $analyzepath | Select-String "fail" | %{ 
        Write-Output "Found in file $_.FileName" | Add-Content $OutLog
        $listFailFile += $_.FullName
    }

#Get-ChildItem -Path $analyzepath -Filter "*.xml"| Select-String "pass" | % {  Copy-Item -Path $filepath -Destination $(join-path $filepath $_.Name) }

##Moving erasure logic here to erase the archives and only parse the XML
##I tried to use * and it ended up wiping out all the files and the anaylze folder
if ($delFile -eq 1) {
    Write-Output "Erasing temporary files."  | Add-Content $OutLog    
    Get-ChildItem -Path (Join-Path $analyzepath *.zip) | Remove-Item
    Get-ChildItem -Path (Join-Path $analyzepath *.gz) | Remove-Item
    Get-ChildItem -Path (Join-Path $analyzepath *.tgz) | Remove-Item
}

Write-Output "...and we are done!  If there are items in the $filepath folder then an email will be sent."  | Add-Content $OutLog

##email infrastructure with results if there are files in the analyze path
#$parms = Get-ChildItem -Path (Join-Path $analyzepath *) | Format-List 

#if ($parms.Length -gt 0) {    
    $subject = "***ATTENTION REQUIRED*** Please check DMARC logs in $analyzepath - confirm failures"
    $body = "List of files that failed:  $failFileList \D Please see attachment" 
    Send-MailMessage -to "capncrunch@bananarama.com" -From "serious@security.cc" -Subject $subject -SmtpServer relay.server.nameor.IP -Body $body -Attachments $OutLog
    if ($delFile -eq 1) {
        Write-Output "Deleting old $OutLog"
        remove-item $OutLog -Force 
    }
#}

##null the objects to free up the memory

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($mapi) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook) | Out-Null

$Outlook = $null | Out-Null
$mapi = $null | Out-Null

[GC]::Collect()
[GC]::WaitForPendingFinalizers()

