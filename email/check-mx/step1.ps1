param (
    [string]$InputFile,
    [string]$OutputFile
)

Try {
    #step one - MX query only - if you try anything with formatting it will say no and you end up looking at full queries.  

    #Follow instructions from notes.  The input file should just be a list of domains.
    #It doesn't matter if there are duplicates, but PowerShell has a problem interpreting some
    #of the emails that are not alphanumeric - symbols make it do stuff you don't want it to do
    $domains = @()
    $domains = Get-Content $InputFile | Sort-Object -Unique

    ForEach ($domain in $domains) { 
        Resolve-DnsName -Name $domain -Type MX | Where-Object { $_.Type -eq 'Mx' } | Export-Csv -Path $OutputFile -NoTypeInformation -Append
    }
} catch { Write-Host "Something went wrong."}