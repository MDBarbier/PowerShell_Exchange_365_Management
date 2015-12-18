# This script sets up a connection to a Exchange 365 hosted server after requesting outlook credentials from the user.

# User must be an exchange 365 admin.

 

 

 

Function help {

 

    Write-host "`n"

 

    Write-Host "The available commands are:"

 

    Write-host "`n"

 

    Write-host "go - this starts the connection to e365 and will prompt you for exchange credentials"

    Write-host "getUsageReport - this outputs the last time that every user logged on. Use it with redirection (>) to output to a text file."

    Write-Host "getCalPermissions - view someones calendar permissions"

    Write-host "addCalPermissions - add on a person to a user's calendar if they don't currently have any access at all"

    Write-host "remCalPermissions - remove someones permissions to view a calendar"

    Write-host "amendCalPermissions - amend the level of permisisons a user has to a calendar" 

    Write-host "getContactPermissions - view someones contacts permissions"

    Write-host "amendContactPermissions - amend the level of access someone has to a user's contact if they already have some sort of access"

    Write-host "addContactPermissions - add a new user to have contact access who currently has no access"

    Write-host "remContactPermissions - remove a user from having access to the person's contacts"

 

    Write-host "`n"

 

}

 

Function endSession {

 

    Get-PSSession | Remove-PSSession

}

 

Function GO {

 

    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection

    Import-PSSession $Session

 

    Write-host "`n"

 

    Write-host "If no errors have been displayed then you are now connected to the Exchange 365 hosted server."

 

    Write-host "Use function specified cmdlet endSession to remove the connection to o365 exchange server"

 

}

 

Function getUsageReport {

 

Write-host `n

Write-host "Use getUsageReport > filepath.txt to output to a file of your choice"

Write-warning "This will take a while..."

 

 

Get-Mailbox | ForEach-Object {Get-MailboxStatistics -Identity $($_.UserPrincipalName) | Select DisplayName, LastLogonTime}

 

 

}

 

Function getCalPermissions {

 

    Write-host "Please enter the email address of the person whose calendar permissions you want to see: "

    $user = Read-host

    $userString = $user + ":\Calendar"

    #Write-host $userString

    Get-MailboxFolderPermission $userString

}

 

Function addCalPermissions {

 

    Write-Host "NOTE this is to add new permissions for someone who has none currently, if you want to amend existing permissions you should use the amendCalPermissions cmdlet instead!"

    Write-host `n

    Write-host "Please enter the email address of the person whose calendar permissions you want to amend: "

    Write-host `n

    Write-host "For info on permission levels view http://blogs.technet.com/b/benw/archive/2015/01/21/the-mystery-of-calendar-permissions-explained.aspx"

    Write-Host `n

 

    $user = Read-host

    $userString = $user + ":\Calendar"

    Write-host "Please enter the email address of the person who you want to give permissions: "

    $recipient = Read-host

    Write-host "Please enter level of permissions you want to assign?"

    $permissionLevel = Read-Host

 

    add-MailboxFolderPermission -Identity $userString -User $recipient -AccessRights $permissionLevel

}

 

Function remCalPermissions {

 

    Write-Host "NOTE this is to remove permissions for, it will remove that user completely. If you want to change the level of permissions, use amendCalPermissions."

    Write-host `n

    Write-host "Please enter the email address of the person whose calendar permissions you want to amend: "

    Write-host `n

    Write-host "For info on permission levels view http://blogs.technet.com/b/benw/archive/2015/01/21/the-mystery-of-calendar-permissions-explained.aspx"

    Write-Host `n

 

    $user = Read-host

    $userString = $user + ":\Calendar"

    Write-host "Please enter the email address of the person who you want to remove permission: "

    $recipient = Read-host

    

    remove-MailboxFolderPermission -Identity $userString -User $recipient

}

 

Function amendCalPermissions {

 

    Write-Host "NOTE this is to amend existing permissions, if you want to grant someone permissions who has none currently you should use the addCalPermissions cmdlet instead!"

    Write-host `n

    Write-host "Please enter the email address of the person whose calendar permissions you want to amend: "

    Write-host `n

    Write-host "For info on permission levels view http://blogs.technet.com/b/benw/archive/2015/01/21/the-mystery-of-calendar-permissions-explained.aspx"

    Write-Host `n

   

    Write-host "Please enter the email address of the person whose calendar permissions you want to amend: "

    $user = Read-host

    $userString = $user + ":\Calendar"

    Write-host "Please enter the email address of the person who you want to give permissions: "

    $recipient = Read-host

    Write-host "Please enter level of permissions you want to assign?"

    $permissionLevel = Read-Host

 

    Set-MailboxFolderPermission -Identity $userString -User $recipient -AccessRights $permissionLevel

}

 

Function getContactsPermissions {

 

    Write-host "Please enter the email address of the person whose permissions you want to see: "

    $user = Read-host

    $userString = $user + ":\Contacts"

    #Write-host $userString

    Get-MailboxFolderPermission $userString

}

 

Function addContactsPermissions {

 

    Write-Host "NOTE this is to add new permissions for someone who has none currently, if you want to amend existing permissions you should use the amendContactPermissions cmdlet instead!"

    Write-host `n

    Write-host "Please enter the email address of the person whose calendar permissions you want to amend: "

    Write-host `n

    Write-host "For info on permission levels view http://blogs.technet.com/b/benw/archive/2015/01/21/the-mystery-of-calendar-permissions-explained.aspx"

    Write-Host `n

 

    $user = Read-host

    $userString = $user + ":\Contacts"

    Write-host "Please enter the email address of the person who you want to give permissions: "

    $recipient = Read-host

    Write-host "Please enter level of permissions you want to assign?"

    $permissionLevel = Read-Host

 

    add-MailboxFolderPermission -Identity $userString -User $recipient -AccessRights $permissionLevel

}

 

Function remContactsPermissions {

 

    Write-Host "NOTE this is to remove permissions for, it will remove that user completely. If you want to change the level of permissions, use amendContactsPermissions."

    Write-host `n

    Write-host "Please enter the email address of the person whose calendar permissions you want to amend: "

    Write-host `n

    Write-host "For info on permission levels view http://blogs.technet.com/b/benw/archive/2015/01/21/the-mystery-of-calendar-permissions-explained.aspx"

    Write-Host `n

 

    $user = Read-host

    $userString = $user + ":\Contacts"

    Write-host "Please enter the email address of the person who you want to remove permission: "

    $recipient = Read-host

    

    remove-MailboxFolderPermission -Identity $userString -User $recipient

}

 

Function amendContactsPermissions {

 

    Write-Host "NOTE this is to amend existing permissions, if you want to grant someone permissions who has none currently you should use the addContactsPermissions cmdlet instead!"

    Write-host `n

    Write-host "Please enter the email address of the person whose permissions you want to amend: "

    Write-host `n

    Write-host "For info on permission levels view http://blogs.technet.com/b/benw/archive/2015/01/21/the-mystery-of-calendar-permissions-explained.aspx"

    Write-Host `n

   

    Write-host "Please enter the email address of the person whose calendar permissions you want to amend: "

    $user = Read-host

    $userString = $user + ":\Contacts"

    Write-host "Please enter the email address of the person who you want to give permissions: "

    $recipient = Read-host

    Write-host "Please enter level of permissions you want to assign?"

    $permissionLevel = Read-Host

 

    Set-MailboxFolderPermission -Identity $userString -User $recipient -AccessRights $permissionLevel

}

 

 

Function welcome{

Write-host `n

Write-host "Welcome to Matt's Exchange365 management PowerShell script."

Write-host `n

Write-warning "NOTE: When you load this script up you have to use the format '. . \<script name>.ps1' as this loads the functions,"

Write-warning "otherwise you will not be able to access them"

Write-host `n

Write-host "Enter GO to connect to Exchange 365 (you will be prompted for exchange credentials)"

Write-host `n

Write-host "Use the command 'help' for details of the commands avilable "

Write-host `n

Write-host "Enter 'get-help .\exchange_365_management.ps1' for help file."

Write-Host `n

}

 

welcome

 

<#

        .SYNOPSIS

        This script will allow you to connect to exchange 365. It has certain common functions built in.

       

        .DESCRIPTION

        Upon calling the function "GO" you will be prompted for your credentials and then the connection will be set up.

        To end the connection, which you should always do rather than just closing the PowerShell window, call "endSession".

        To get a list of the functions available, load the script up (. .\exchange_365_management.ps1) and enter "help".

 

        NOTE: when loading the script you must CD to the directory and then use ". .\exchange_365_management.ps1" to ensure functions are loaded.

       

 

    #>

 