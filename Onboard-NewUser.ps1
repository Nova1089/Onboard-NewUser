<#
Version 1.0

This script sets up a new user in M365.
#>

# functions
function Initialize-ColorScheme
{
    Set-Variable -Name "successColor" -Value "Green" -Scope "Script" -Option "Constant"
    Set-Variable -Name "infoColor" -Value "DarkCyan" -Scope "Script" -Option "Constant"
    Set-Variable -Name "warningColor" -Value "Yellow" -Scope "Script" -Option "Constant"
    Set-Variable -Name "failColor" -Value "Red" -Scope "Script" -Option "Constant"
}

function Show-Introduction
{
    Write-Host "This script sets up a new user in M365." -ForegroundColor $infoColor
    Read-Host "Press Enter to continue"
}

function Use-Module($moduleName)
{    
    $keepGoing = -not(Test-ModuleInstalled $moduleName)
    while ($keepGoing)
    {
        Prompt-InstallModule $moduleName
        Test-SessionPrivileges
        Install-Module $moduleName

        if ((Test-ModuleInstalled $moduleName) -eq $true)
        {
            Write-Host "Importing module..." -ForegroundColor $infoColor
            Import-Module $moduleName
            $keepGoing = $false
        }
    }
}

function Test-ModuleInstalled($moduleName)
{    
    $module = Get-Module -Name $moduleName -ListAvailable
    return ($null -ne $module)
}

function Prompt-InstallModule($moduleName)
{
    do 
    {
        Write-Host "$moduleName module is required." -ForegroundColor $infoColor
        $confirmInstall = Read-Host -Prompt "Would you like to install the module? (y/n)"
    }
    while ($confirmInstall -inotmatch "^\s*y\s*$") # regex matches a y but allows spaces
}

function Test-SessionPrivileges
{
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentSessionIsAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if ($currentSessionIsAdmin -ne $true)
    {
        Write-Host ("Please run script with admin privileges.`n" +
            "1. Open Powershell as admin.`n" +
            "2. CD into script directory.`n" +
            "3. Run .\scriptname`n") -ForegroundColor $failColor
        Read-Host "Press Enter to exit"
        exit
    }
}

function TryConnect-MgGraph($scopes)
{
    $connected = Test-ConnectedToMgGraph
    while (-not($connected))
    {
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor $infoColor

        if ($null -ne $scopes)
        {
            Connect-MgGraph -Scopes $scopes -ErrorAction SilentlyContinue | Out-Null
        }
        else
        {
            Connect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        }

        $connected = Test-ConnectedToMgGraph
        if (-not($connected))
        {
            Read-Host "Failed to connect to Microsoft Graph. Press Enter to try again"
        }
        else
        {
            Write-Host "Successfully connected!" -ForegroundColor $successColor
        }
    }    
}

function Test-ConnectedToMgGraph
{
    return $null -ne (Get-MgContext)
}

function TryConnect-ExchangeOnline
{
    $connectionStatus = Get-ConnectionInformation -ErrorAction SilentlyContinue

    while ($null -eq $connectionStatus)
    {
        Write-Host "Connecting to Exchange Online..."
        Connect-ExchangeOnline -ErrorAction SilentlyContinue
        $connectionStatus = Get-ConnectionInformation

        if ($null -eq $connectionStatus)
        {
            Read-Host -Prompt "Failed to connect to Exchange Online. Press Enter to try again"
        }
    }
}

function Prompt-UPN
{
    $upn = Read-Host "Enter the UPN for the user (PreferredFirstName.LastName@blueravensolar.com)"
    $isValidEmail = Validate-BrsEmail $upn
    while (-not($isValidEmail))
    {
        $upn = Read-Host "Enter the UPN for the user"
        $isValidEmail = Validate-BrsEmail $upn
    }

    return $upn.Trim()
}

function Validate-BrsEmail($email)
{
    $isValidEmail = $email -imatch '^\s*[\w\.-]+\.[\w\.-]+(@blueravensolar\.com)\s*$'
    
    if (-not($isValidEmail))
    {
        Write-Warning ("Email is invalid: $email `n" +
            "    Expected format is PreferredFirstName.LastName@blueravensolar.com `n")
    }

    return $isValidEmail
}

function Get-M365User($upn)
{
    if ($null -eq $upn) { throw "Can't get M365 user. UPN is null." }
    
    $user = (Get-MgUser -UserID $upn -Property @("CreatedDateTime", 
                                                "DisplayName", 
                                                "UserPrincipalName",   
                                                "JobTitle", 
                                                "Department", 
                                                "UsageLocation", 
                                                "LicenseDetails",
                                                "Id") -ErrorAction "SilentlyContinue")
    if ($null -eq $user)
    {
        Write-Host "User does not exist yet." -ForegroundColor $infoColor
    }
    return $user
}

function Get-UserManager($user)
{
    $managerId = Get-MgUserManager -UserId $user.UserPrincipalName -ErrorAction "SilentlyContinue" | Select-Object -ExpandProperty "ID"
    if ($null -eq $managerId) { return $null }
    return Get-MgUser -UserId $managerId | Select-Object -ExpandProperty "DisplayName"
}

function Get-UserLicenses($user)
{
    if ($null -eq $script:lookupTable)
    {
        $script:lookupTable = @{
            "3b555118-da6a-4418-894f-7df1e2096870" = "Microsoft 365 Business Basic"
            "cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46" = "Microsoft 365 Business Premium"
            "f245ecc8-75af-4f8e-b61f-27d8114de5f3" = "Microsoft 365 Business Standard"
            "18181a46-0d4e-45cd-891e-60aabd171b4e" = "Office 365 E1"
            "6fd2c87f-b296-42f0-b197-1e91e994b900" = "Office 365 E3"
            "c7df2760-2c81-4ef7-b578-5b5392b571df" = "Office 365 E5"
            "05e9a617-0261-4cee-bb44-138d3ef5d965" = "Microsoft 365 E3"
            "06ebc4ee-1bb5-47dd-8120-11324bc54e06" = "Microsoft 365 E5"
            "4b9405b0-7788-4568-add1-99614e613b69" = "Exchange Online (Plan 1)"
            "19ec0d23-8335-4cbd-94ac-6050e30712fa" = "Exchange Online (Plan 2)"
            "a403ebcc-fae0-4ca2-8c8c-7a907fd6c235" = "Microsoft Fabric (Free)"
            "f30db892-07e9-47e9-837c-80727f46fd3d" = "Microsoft Power Automate Free"
            "5b631642-bd26-49fe-bd20-1daaa972ef80" = "Microsoft PowerApps for Developer"
            "dcb1a3ae-b33f-4487-846a-a640262fadf4" = "Microsoft Power Apps Plan 2 Trial"
            "3dd6cf57-d688-4eed-ba52-9e40b5468c3e" = "Microsoft Defender for Office 365 (Plan 2)"
        }
    }

    $licenseDetails = Get-MGUserLicenseDetail -UserId $user.UserPrincipalName
    $licenses = New-Object System.Collections.Generic.List[String]

    foreach ($license in $licenseDetails)
    {
        $licenseName = $script:lookupTable[$license.SkuId]
        $licenses.Add($licenseName)
    }

    return Write-Output $licenses -NoEnumerate
}

function Get-UserGroups($user)
{
    return Get-MgUserMemberOfAsGroup -UserId $user.UserPrincipalName | Select-Object -ExpandProperty "DisplayName"
}

function Get-UserAdminRoles($user)
{
    return Get-MgUserMemberOfAsDirectoryRole -UserId $user.UserPrincipalName | Select-Object -ExpandProperty "DisplayName"
}

function Display-UserProperties($user, $manager, $licenses, $groups, $adminRoles)
{
    Write-Host "User found!" -ForegroundColor $successColor

    $basicProps = [PSCustomObject]@{
        "Created Date/Time" = $user.CreatedDateTime
        "Display Name"      = $user.DisplayName
        "UPN"               = $user.UserPrincipalName
        "Title"             = $user.JobTitle
        "Department"        = $user.Department
        "Manager"           = $manager
        "Usage Location"    = $user.UsageLocation
    }
    $basicProps | Out-Host

    Show-Separator "Licenses"
    $licenses | Out-Host

    Show-Separator "Groups"
    $groups | Out-Host
    
    Show-Separator "Admin Roles"
    $adminRoles | Out-Host
    
    Write-Host "`n"
}

function Show-Separator($title, [ConsoleColor]$color = "DarkCyan", [switch]$noLineBreaks)
{
    if ($title)
    {
        $separator = " $title "
    }
    else
    {
        $separator = ""
    }

    # Truncate if it's too long.
    If (($separator.length - 6) -gt ((Get-host).UI.RawUI.BufferSize.Width))
    {
        $separator = $separator.Remove((Get-host).UI.RawUI.BufferSize.Width - 5)
    }

    # Pad with dashes.
    $separator = "--$($separator.PadRight(((Get-host).UI.RawUI.BufferSize.Width)-3,"-"))"

    if (-not($noLineBreaks))
    {        
        # Add line breaks.
        $separator = "`n$separator`n"
    }

    Write-Host $separator -ForegroundColor $color
}

function Prompt-Menu 
{
    if ($licenseStepCompleted) { $licenseBox = "[X]"} else { $licenseBox = "[ ]" }
    if ($groupStepCompleted) { $groupBox = "[X]" } else { $groupBox = "[ ]" }
    if ($mailboxStepCompleted) { $mailboxBox = "[X]" } else { $mailboxBox = "[ ]" }

    Read-Host ("What would you like to do next?`n`n" +
                "[1] Display User Info`n" +
                "[2] $licenseBox Grant licenses`n" +
                "[3] $groupBox Assign groups`n" +
                "[4] $mailboxBox Grant shared mailboxes`n" +
                "[5] Finish`n")
}

function Prompt-BrsEmail
{
    param
    (
        [ValidateSet("group", "mailbox", IgnoreCase = $false)]
        $emailType
    )

    do
    {
        $email = Read-Host "Enter $emailType email (you may omit the @blueravensolar.com)"
    }
    while ($null -eq $email)

    $email = $email.Trim()
    $isStandardFormat = $email -imatch '^\S+@blueravensolar.com$'

    if (-not($isStandardFormat))
    {
        $email += '@blueravensolar.com'
    }

    return $email
}

function Get-M365Group($email)
{
    $group = Get-MgGroup -Filter "mail eq '$email'" -ErrorAction "SilentlyContinue"
    if ($group)
    {
        Write-Host "Found Group!" -ForegroundColor $successColor
        $group | Select-Object -Property @("DisplayName", "Mail", "Description") | Out-Host        
    }
    else
    {
        Write-Host "Group not found." -ForegroundColor $warningColor
    }
    return $group
}

function Assign-M365Group($user, $group, $existingGroups)
{
    if ($existingGroups -contains $group.DisplayName)
    {
        Write-Host "$($user.DisplayName) is already a member of that group." -ForegroundColor $warningColor
        return
    }
    
    try
    {
        New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $user.Id -ErrorAction "Stop" | Out-Null
    }
    catch
    {
        $errorRecord = $_
        Write-Host "There was an issue assigning the group." -ForegroundColor $failColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $failColor
    }    
}

function Get-SharedMailbox($email)
{
    if ($null -eq $email) { throw "Can't get shared mailbox. Email is null." }

    $mailbox = Get-EXOMailbox -Identity $email -ErrorAction "SilentlyContinue"
    if (($mailbox) -and ($mailbox.RecipientTypeDetails -eq 'SharedMailbox'))
    {
        Write-Host "Mailbox found!" -ForegroundColor $successColor
        $mailbox | Select-Object -Property @("DisplayName", "UserPrincipalName", @{ label = "Type"; expression = {$_.RecipientTypeDetails} }) | Out-Host
        return $mailbox
    }
    elseif ($mailbox)
    {
        Write-Host "Mailbox was found but it's not a shared mailbox. Type is '$($mailbox.RecipientTypeDetails)'." -ForegroundColor $warningColor
    }
    else
    {
        Write-Host "Mailbox not found." -ForegroundColor $warningColor
    }
    return $null
}

function Grant-MailboxAccess($user, $mailbox)
{
   $accessType = Prompt-MailboxAccessType

   try
   {
        switch ($accessType)
        {
            1
            {
                Add-MailboxPermission -Identity $mailbox.PrimarySmtpAddress -User $user.UserPrincipalName -AccessRights "FullAccess" -Confirm:$false -WarningAction "SilentlyContinue" -ErrorAction "Stop" | Out-Null
            }
            2
            {
                Add-RecipientPermission -Identity $mailbox.PrimarySmtpAddress -Trustee $user.UserPrincipalName -AccessRights "SendAs" -Confirm:$false -WarningAction "SilentlyContinue" -ErrorAction "Stop" | Out-Null
            }
            3
            {
                Add-MailboxPermission -Identity $mailbox.PrimarySmtpAddress -User $user.UserPrincipalName -AccessRights "FullAccess" -Confirm:$false -WarningAction "SilentlyContinue" -ErrorAction "Stop" | Out-Null
                Add-RecipientPermission -Identity $mailbox.PrimarySmtpAddress -Trustee $user.UserPrincipalName -AccessRights "SendAs" -Confirm:$false -WarningAction "SilentlyContinue" -ErrorAction "Stop" | Out-Null
            }
        }
        Write-Host "Successfully granted access! (If they didn't already have the access.)" -ForegroundColor $successColor
   }
   catch
   {
        $errorRecord = $_
        Write-Host "There was an issue granting mailbox access. Please try again." -ForegroundColor $failColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $failColor
   }    
}

function Prompt-MailboxAccessType
{
    do
    {
        $accessType = Read-Host  ("Access type?`n`n" +
            "[1] Read & Manage`n" +
            "[2] Send As`n" +
            "[3] Both`n")

        $accessType = $accessType.Trim()
        $isValidResponse = $accessType -imatch '^\s*[123]\s*$' # regex matches 1, 2, or 3 but allows spaces.

        if (-not($isValidResponse))
        {
            Write-Host "Please enter a number 1-3." -ForegroundColor $warningColor
        }
    }
    while (-not($isValidResponse))

    return $accessType
}

function Get-GoToAccessToken($clientId, $clientSecret)
{
    # Function handles GoTo connect OAUTH2 authorization code grant flow. (Obtains auth code then uses that to get a temp access token.)
    # https://developer.goto.com/guides/Authentication/03_HOW_accessToken/
    # https://developer.goto.com/Authentication/#section/Authorization-Flows

    $authUri = "https://authentication.logmeininc.com/oauth/authorize"
    $accessTokenUri = "https://authentication.logmeininc.com/oauth/token"
    $redirectUri = "https://localhost"
    
    if ($null -eq $clientId) { $clientId = Read-Host "Enter client ID" }
    if ($null -eq $clientSecret) { $clientSecret = Read-Host "Enter client secret" }

    $authCode = Invoke-OAuth2AuthorizationEndpoint -Uri $authUri -Client_id $clientId -Redirect_uri $redirectUri
    $accessToken = $accessToken = Invoke-OAuth2TokenEndpoint @authCode -Uri $accessTokenUri -Client_secret $clientSecret -Client_auth_method "client_secret_basic"

    return $accessToken
}

# main
Initialize-ColorScheme
Show-Introduction
Use-Module "Microsoft.Graph.Users"
Use-Module "ExchangeOnlineManagement"
Use-Module "PSAuthClient" # Docs for this module found here https://github.com/alflokken/PSAuthClient
TryConnect-MgGraph -Scopes @("User.ReadWrite.All", "Group.ReadWrite.All")
TryConnect-ExchangeOnline
$upn = Prompt-UPN
$user = Get-M365User $upn
if ($null -ne $user)
{
    $manager = Get-UserManager $user 
    $licenses = Get-UserLicenses $user 
    $groups = Get-UserGroups $user
    $adminRoles = Get-UserAdminRoles $user
    Display-UserProperties -User $user -Manager $manager -Licenses $licenses -Groups $groups -AdminRoles $adminRoles
}

$script:licenseStepCompleted = $false
$script:groupStepCompleted = $false
$script:mailboxStepCompleted = $false

$menuSelection = Prompt-Menu

switch ($menuSelection)
{
    1
    {
        if ($null -ne $user)
        {
            $manager = Get-UserManager $user 
            $licenses = Get-UserLicenses $user 
            $groups = Get-UserGroups $user
            $adminRoles = Get-UserAdminRoles $user
            Display-UserProperties -User $user -Manager $manager -Licenses $licenses -Groups $groups -AdminRoles $adminRoles
        }
    }
    2
    {
        Write-Host "You selected option 2! (Grant licenses)" -ForegroundColor $infoColor
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
        Write-Host "You selected option 5!" -ForegroundColor $infoColor
    }
}

$menuSelection = Prompt-Menu

Read-Host "Press Enter to exit"
