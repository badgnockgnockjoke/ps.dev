Quick PowerShell scripts to fire off emails and store the running tally in a text file.  

Use: 
- Save contents in "C:\scripts" and modify *.ps1 with the email addresses to/from
- Open Task Scheduler
- Drill down and create a folder "ADAudit"
- Click "Import Task" in the right pane, select the xml file associated with the EventID.
- Schedule the monthlyaudit.ps1 script to automate sending monthly reports of AD activity according to the events being trapped

I sat down to figure this out because of companies that charge thousands of dollars a year for a pretty GUI.  This is the basic framework, extrapolate from here.  
