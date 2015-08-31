## MailMan
### Powershell MS Outlook Enumeration and Internal Phishing tool

###Synopsis

Mailman is a tool that can be used to browse/search a client's Outlook folders as well as send phishing emails internally. 
Mailman should be usefule in situations where lateral movement is restricted or phishing with a legitimate account is needed to further
access. 

###Quick Usage Guide

#####Get-OutlookInstance -DisablePrompt 
This function creates an Outlook ComObject or attaches to a running instance of Outlook. Use the DisablePrompt switch
to set/create the necessary registry keys to disable the programmatic access prompt for Outlook.
Specify an AdminUser and AdminPass if the current user does not have permission to edit the registry. 

#####Get-SMTPAddress -FullName "William Striker"
This function returns the Primary SMTP address of a user from the Global Address List based on there Full Name.

#####Invoke-Spam -Targets "testuser@testing.com" -Subject "Wub Wub Wub" -Body "Hey! This is a test email" -Attachment 

This function will send an email to the specified target/s. The TargetList parameter can be used to read in targets from a file. HTML tags maybe used in the Body parameter to embed a URL or whatever suits your needs. 

#####Invoke-MailSearch -DefaultFolder "Inbox" -Keyword "password" -MaxSearch 400 -MaxResults 50 -MaxThreads 15

This function will conduct a multithreaded search through specified Outlook Default folder for emails that contain the keyword. 




  
