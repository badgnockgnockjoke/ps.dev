Import-Module ActiveDirectory
$VerbosePreference = "continue"
$list = (Get-ADComputer -LDAPFilter "(&(objectcategory=computer)(OperatingSystem=*server*))").Name
Write-Verbose  -Message "Trying to query $($list.count) servers found in AD"
$logfilepath = "F:\ADList-Tasks.csv"
$ErrorActionPreference = "SilentlyContinue"

foreach ($computername in $list)
{
    $path = "\\" + $computername + "\c$\Windows\System32\Tasks"
    $tasks = Get-ChildItem -Path $path -File

    if ($tasks)
    {
        Write-Verbose -Message "I found $($tasks.count) tasks for $computername"
    }

    foreach ($item in $tasks)
    {
        $AbsolutePath = $path + "\" + $item.Name
        $task = [xml] (Get-Content $AbsolutePath)
        [STRING]$check = $task.Task.Principals.Principal.UserId

        if ($task.Task.Principals.Principal.UserId)
        {
          Write-Verbose -Message "Writing the log file with values for $computername"           
          Add-content -path $logfilepath -Value "$computername,$item,$check"
        }

    }
}


#Get-ScheduledTask - | Get-ScheduledTaskInfo | Sort-Object -Property LastRunTime | Format-Table -AutoSize