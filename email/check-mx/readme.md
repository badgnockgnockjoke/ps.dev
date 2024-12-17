The  scripts were designed to output a CSV file with a list.  It will be a list of domains that have an MX record matching the search string.  In the script, where Resolve-Dnsname is called it will output a list.  However, if you pipe that into a Export-Csv or try to Select (header name), I don't think knows of certain things and gives you something  completely different.  I threw this together to determine which client domains contained a reference to a specific host.  The original use-case was it being called in conjunction with the DMARC validation script.

TL;DR:
Checks for string (domain name) from output of MX record query.  The list of domains is one text file and output is a CSV with all MX records of the matched domain.

Steps:
1. Make a list of domain names to check DNS MX records of.
2. Open PowerShell and from the prompt run .\step1.ps1 -InputFile (name_of_file.txt) -OutputFile (intermediate.csv)
3. Run .\step2.ps1 -InputFile (intermediate.csv) -OutputFile (final_list.csv)
4. Rejoin because the logic is so simple, yet saved you an hour and you now have time for cat videos.


Required files: 
	- step1.ps1 
	- step2.ps1
	- domains.txt