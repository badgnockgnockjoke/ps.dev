Import-Module ActiveDirectory
$VerbosePreference = "continue"
$list = (Get-ADComputer -LDAPFilter "(&(objectcategory=computer)(OperatingSystem=*server*))").Name
Write-Verbose  -Message "Trying to query $($list.count) servers found in AD"
$logfilepath = "F:\ADList-Tasks.csv"
$ErrorActionPreference = "SilentlyContinue"

foreach ($computername in $list)
{
    #$path = "\\" + $computername + "\c$\Windows\System32\Tasks"
    #$tasks = Get-ChildItem -Path $path -File
    $tasks = Get-ScheduledTask -CimSession $computername | Get-ScheduledTaskInfo
    Get-DnsServerForwarder -ComputerName $computername.ToString -Verbose
    $tempname = $computername.substring(3)
    Write-Host $computername "    " $tempname "is the name of the system"
    if ($tempname -like 'co' )
 {
    Set-Service -ComputerName $computername -Name "ManageEngine UEMS - Agent" -StartupType Automatic -Status Running
    Set-Service -ComputerName $computername -Name "ManageEngine Unified Endpoint Security - Agent" -StartupType Automatic -Status Running
    
}
    if ($tasks)
    {
        Write-Verbose -Message "I found $($tasks.count) tasks for $computername"
    }

    foreach ($item in $tasks)
    {
   #     Write-Verbose -Message "Writing the log file with values for $computername"           
        #Add-content -path $logfilepath -Value "$computername,$item,$check"
        ##| Sort-Object -property LastRunTime | ConvertTo-Csv -Delimiter ","
        Write-Output $item    
    }

}


#Get-ScheduledTask - | Get-ScheduledTaskInfo | Sort-Object -Property LastRunTime | Format-Table -AutoSize