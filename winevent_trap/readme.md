Quick PowerShell scripts to fire off emails when certain EventIDs come up, and store the running tally in a text file.  Modify the .ps1 files to redirect the output elsewhere.

Use: 
- Save contents in "C:\scripts" and modify *.ps1 with the email addresses to/from
- Open Task Scheduler
- Drill down and create a folder "ADAudit"
- Click "Import Task" in the right pane, select the xml file associated with the EventID (look inside the TaskScheduler folder)
- Schedule the BackupADAudit.ps1 script to automate sending monthly reports of AD activity according to the events being trapped


This basic framework was implemented to save money for a client.  They didn't need all of the predefined User reports, however, if you throw in some ODBC logging, you would easily be able to create a front-end to generate reports.  
