<#
Version 1.0

This script onboards a new user into M365 and GoTo.
#>

# Import class modules in the same folder. These imports must come first in the script.
using module .\Class-GotoWizard.psm1

# Dot sourcing
. "$PSScriptRoot\GlobalFunctions.ps1"

#This directive will throw an error if not running PowerShell core (a.k.a PowerShell v6+)
#Requires -PSEdition Core

# main
Show-Introduction
Use-Module "Microsoft.Graph.Users"
Use-Module "ExchangeOnlineManagement"
Use-Module "Microsoft.Powershell.SecretManagement"
Use-Module "PSAuthClient" # Docs for this module found here https://github.com/alflokken/PSAuthClient
TryConnect-MgGraph -Scopes @("User.ReadWrite.All", "Group.ReadWrite.All", "Organization.Read.All")
TryConnect-ExchangeOnline
$upn = Prompt-UPN
$user = Get-M365User $upn
if ($null -ne $user)
{
    $manager = Get-UserManager $user 
    $licenses = Get-UserLicenses $user 
    $groups = Get-UserGroups $user
    $adminRoles = Get-UserAdminRoles $user
    Show-UserProperties -User $user -Manager $manager -Licenses $licenses -Groups $groups -AdminRoles $adminRoles
}

$script:licenseStepCompleted = $false
$script:groupStepCompleted = $false
$script:mailboxStepCompleted = $false
$script:gotoStepCompleted = $false

$mainMenuSelection = Prompt-MainMenu

switch ($mainMenuSelection)
{
    1
    {
        if ($null -ne $user)
        {
            $manager = Get-UserManager $user 
            $licenses = Get-UserLicenses $user 
            $groups = Get-UserGroups $user
            $adminRoles = Get-UserAdminRoles $user
            Show-UserProperties -User $user -Manager $manager -Licenses $licenses -Groups $groups -AdminRoles $adminRoles
        }
    }
    2
    {
        Write-Host "You selected option 2! (Grant licenses)" -ForegroundColor $infoColor
        Start-M365LicenseWizard $user
        $script:licenseStepCompleted = $true
    }
    3
    {
        Write-Host "You selected option 3! (Assign groups)" -ForegroundColor $infoColor
        $script:groupStepCompleted = $true
        do
        {
            $groupEmail = Prompt-BRSEmail -EmailType "group"
            $group = Get-M365Group $groupEmail
        }
        while ($null -eq $group)

        Assign-M365Group -User $user -Group $group -ExistingGroups $groups
    }
    4
    {
        Write-Host "You selected option 4! (Grant shared mailboxes)" -ForegroundColor $infoColor
        $script:mailboxStepCompleted = $true
        do
        {
            $mailboxEmail = Prompt-BrsEmail -EmailType "mailbox"
            $mailbox = Get-SharedMailbox $mailboxEmail
        }
        while ($null -eq $mailbox)

        Grant-MailboxAccess -User $user -Mailbox $mailbox
    }
    5
    {
        Write-Host "You selected option 5! (Setup GoTo account)" -ForegroundColor $infoColor
        $script:gotoStepCompleted = $true
        $gotoWizard = [GotoWizard]::New($upn)
        $gotoWizard.Start()
    }
    6
    {
        Write-Host "You selected option 6! (Finish)" -ForegroundColor $infoColor
    }
}

$mainMenuSelection = Prompt-MainMenu

Read-Host "Press Enter to exit"
