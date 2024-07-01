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

do
{
    $upn = Prompt-BrsEmail "Enter user UPN"
    $isValidEmail = Test-ValidBrsEmail $upn
    if (-not($isValidEmail)) { continue }
    $user = Get-M365User -UPN $upn -Detailed -WarningAction "SilentlyContinue"

    if ($null -eq $user)
    {
        Write-Host "User does not exist yet." -ForegroundColor $infoColor
        $createUser = Prompt-YesOrNo "Create this user in M365?"
        if ($createUser)
        {
            $user = Start-M365UserWizard $upn
            # Get more details about the user.
            $user = Invoke-GetWithRetry { Get-M365User -UPN $user.UserPrincipalName -Detailed }
        }
    }
}
while ($null -eq $user)

$script:grantLicensesCompleted = $false
$script:assignGroupsCompleted = $false
$script:grantMailboxesCompleted = $false
$script:gotoSetupCompleted = $false

$userProps = Get-UserProperties $user
Show-UserProperties -BasicProps $userProps.basicProps -Licenses $userProps.Licenses -Groups $userProps.Groups -AdminRoles $userProps.AdminRoles

$keepGoing = $true
while ($keepGoing)
{
    $mainMenuSelection = Prompt-MainMenu
    switch ($mainMenuSelection)
    {
        1 # Show M365 user info
        {
            $userProps = Get-UserProperties $user
            Show-UserProperties -BasicProps $userProps.basicProps -Licenses $userProps.Licenses -Groups $userProps.Groups -AdminRoles $userProps.AdminRoles
            break
        }
        2 # Grant licenses
        {
            Start-M365LicenseWizard $user
            break
        }
        3 # Assign groups
        {
            Start-M365GroupWizard $user
            break
        }
        4 # Grant shared mailboxes
        {
            Start-MailboxWizard $user
            break
        }
        5 # Setup GoTo account
        {
            $gotoWizard = [GotoWizard]::New($upn)
            $gotoWizard.Start()
            break
        }
        6 # Finish
        {
           $keepGoing = $false
           break
        }
    }
}

Read-Host "Press Enter to exit"