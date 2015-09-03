## MailRaider
### Powershell MS Outlook Enumeration and Internal Phishing tool

###Synopsis

Mailman is a tool that can be used to browse/search a client's Outlook folders as well as send phishing emails internally using their Outlook client. Mailman should be usefule in situations where lateral movement is restricted or phishing with a legitimate account is needed to obtain further access. 

###Quick Usage Guide

#####Disable-SecuritySettings -AdminUser "LOCALHOST\Admin" -AdminPass "IamAdmin#123" 
This function sets/creates the required registry keys in order to disamble the Outlook programmatic access 
prompt. Please note that if Anti-Virus is not installed and updated on the host, changing these keys will not make a difference. 

#####View-Email -FolderName "Inbox" -Index 25
This function selects the specified folder and then displays the EmailItem at the selected index. This is useful for viewing
individual e-mails, one at a time. 

#####Get-SMTPAddress -FullName "William Striker"
This function returns the Primary SMTP address of a user from the Global Address List based on there Full Name.

#####Invoke-SendEmail -Targets "testuser@testing.com" -Subject "Wub Wub Wub" -Body "Hey! This is a test email" -Attachment 

This function will send an email to the specified target/s. The TargetList parameter can be used to read in targets from a file. HTML tags maybe used in the Body parameter to embed a URL or whatever suits your needs. 

#####Invoke-MailSearch -DefaultFolder "Inbox" -Keyword "password" -MaxSearch 400 -MaxResults 50 -MaxThreads 15

This function will conduct a multithreaded search through specified Outlook Default folder for emails that contain the keyword. 




  
