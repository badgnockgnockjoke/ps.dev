param (
    [string]$InputFile,
    [string]$OutputFile
)

try {
    #step 2 - Now that I've realized there is no spoon....
    Import-csv -Path $InputFile | Select Name,NameExchange | ForEach-Object { 
        if ($_ -like "*microsoft.com*") { $_  } 
    } | Export-Csv -Path $OutputFile
} catch { write-host "Something wrong" }