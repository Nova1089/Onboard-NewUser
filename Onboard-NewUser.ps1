<#
Version 1.0

This script creates or modifies a new user in M365 and GoTo.
#>

#This directive will throw an error if not running PowerShell core (PowerShell v6+).
#Requires -PSEdition Core

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
    Write-Host "This script creates or modifies a new user in M365 and GoTo." -ForegroundColor $infoColor
    Read-Host "Press Enter to continue"
}

function Use-Module($moduleName)
{    
    $keepGoing = -not(Confirm-ModuleInstalled $moduleName)
    while ($keepGoing)
    {
        Prompt-InstallModule $moduleName
        Confirm-AdminPrivilege
        Install-Module $moduleName

        if ((Confirm-ModuleInstalled $moduleName) -eq $true)
        {
            Write-Host "Importing module..." -ForegroundColor $infoColor
            Import-Module $moduleName
            $keepGoing = $false
        }
    }
}

function Confirm-ModuleInstalled($moduleName)
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

function Confirm-AdminPrivilege
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
    $connected = Confirm-ConnectedToMgGraph
    while (-not($connected))
    {
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor $infoColor

        if ($null -ne $scopes)
        {
            Connect-MgGraph -Scopes $scopes -ErrorAction "SilentlyContinue" | Out-Null
        }
        else
        {
            Connect-MgGraph -ErrorAction "SilentlyContinue" | Out-Null
        }

        $connected = Confirm-ConnectedToMgGraph
        if (-not($connected))
        {
            Read-Host "Failed to connect to Microsoft Graph. Press Enter to try again"
        }
    }
}

function Confirm-ConnectedToMgGraph
{
    return $null -ne (Get-MgContext)
}

function TryConnect-ExchangeOnline
{
    $connectionStatus = Get-ConnectionInformation -ErrorAction "SilentlyContinue"

    while ($null -eq $connectionStatus)
    {
        Write-Host "Connecting to Exchange Online..." -ForegroundColor $infoColor
        Connect-ExchangeOnline -ErrorAction "SilentlyContinue" -ShowBanner:$false 
        $connectionStatus = Get-ConnectionInformation
        if ($null -eq $connectionStatus)
        {
            Read-Host -Prompt "Failed to connect to Exchange Online. Press Enter to try again"
        }
    }
}

function Prompt-BrsEmail($message)
{
    do
    {
        $email = Read-Host "`n$message (you may omit the @blueravensolar.com)"
    }
    while (($null -eq $email) -or ("" -eq $email))

    $email = $email.Trim()
    $hasDomain = $email -imatch '^\S*@blueravensolar.com$'

    if (-not($hasDomain))
    {
        $email += '@blueravensolar.com'
    }

    return $email
}

function Confirm-ValidBrsEmail($email)
{
    $isBrsEmail = $email -imatch '^\S+@blueravensolar.com$'
    if (-not($isBrsEmail))
    {
        Write-Host "Invalid BRS email." -ForegroundColor $warningColor
        return $false
    }

    # Regex matches word.word@blueravensolar.com
    $isStandard = $email -imatch '^[\w-]+\.[\w-]+(@blueravensolar\.com)$'    
    if (-not($isStandard))
    {
        Write-Host "Email is not standard (First.Last@blueravensolar.com)" -ForegroundColor $warningColor
        $continue = Prompt-YesOrNo "Are you sure you want to use this email?"
        if (-not($continue)) { return $false }
    }

    return $true
}

function Prompt-YesOrNo($question)
{
    Write-Host "$question`n[Y] Yes  [N] No"

    do
    {
        $response = Read-Host
        $validResponse = $response -imatch '^\s*[yn]\s*$' # regex matches y or n but allows spaces
        if (-not($validResponse)) 
        {
            Write-Host "Please enter y or n." -ForegroundColor $warningColor
        }
    }
    while (-not($validResponse))

    if ($response -imatch '^\s*y\s*$') # regex matches a y but allows spaces
    {
        return $true
    }
    return $false
}

function Get-M365User
{
    # Needs cmdlet binding to gain WarningAction param.
    [CmdletBinding()]
    Param($upn, [switch]$detailed)
    
    if ($null -eq $upn) { throw "Can't get M365 user. UPN is null." }

    try
    {
        if ($detailed)
        {
            $user = (Get-MgUser -UserID $upn -Property @(
                    "CreatedDateTime", 
                    "DisplayName", 
                    "UserPrincipalName",   
                    "JobTitle", 
                    "Department", 
                    "UsageLocation", 
                    "LicenseDetails",
                    "Id") -ErrorAction "Stop")            
        }
        else
        {
            $user = Get-MgUser -UserID $upn -ErrorAction "Stop"
        }
        Write-Host "User found!" -ForegroundColor $successColor
    }
    catch
    {
        $errorRecord = $_
        if ($errorRecord.Exception.Message -ilike "*[Request_ResourceNotFound]*")
        {
            Write-Warning "User not found." # Use Write-Warning instead of Write-Host so we can silence it when we choose.
            return
        }
        Write-Host "There was an issue getting the user." -ForegroundColor $warningColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $warningColor
    }
    return $user
}

function Start-M365UserCreationWizard($upn)
{
    $emailParts = Get-EmailParts $upn
    $correct = Prompt-YesOrNo "This preferred first name?: $($emailParts.FirstName)"
    if (-not($correct)) { $emailParts.FirstName = Read-Host "Enter preferred first name" }

    $correct = Prompt-YesOrNo "This last name?: $($emailParts.LastName)"
    if (-not($correct)) { $emailParts.LastName = Read-Host "Enter last name" }

    $correct = Prompt-YesOrNo "This display name?: $($emailParts.DisplayName)"
    if (-not($correct))
    {
        $emailParts.DisplayName = Read-Host "Enter display name" 
        $emailParts.MailNickName = Read-Host "Enter mail nickname (first.last)"
    }

    $jobTitle = Read-Host "Enter job title"
    $department = Read-Host "Enter department"
    do
    {
        $managerUpn = Prompt-BrsEmail "Enter manager UPN"
        $manager = Get-M365User -UPN $managerUpn
    }
    while ($null -eq $manager)

    $manager | Select-Object -Property @("DisplayName", "UserPrincipalName") | Out-Host

    $params = [ordered]@{
        "UPN"           = $upn
        "MailNickName"  = $emailParts.MailNickName
        "DisplayName"   = $emailParts.DisplayName
        "FirstName"     = $emailParts.FirstName
        "LastName"      = $emailParts.LastName
        "JobTitle"      = $jobTitle
        "Department"    = $department
        "Manager"       = $manager.displayName
        "UsageLocation" = "US" # We always set this to US, even for those out-of-country.        
    }

    Write-Host "Create user with these parameters?"
    $params | Out-Host
    $continue = Prompt-YesOrNo
    if (-not($continue)) { return }

    $params.Remove("Manager") # Manager was just in params to display in host. Remove before passing to New-M365User.

    do
    {
        $user = New-M365User @params
        if ($null -eq $user)
        {
            $tryAgain = Prompt-YesOrNo "Try again?"
            if (-not($tryAgain))
            {
                Read-Host "Press Enter to exit"
                exit
            }
        }
    }
    while ($null -eq $user)
    
    Set-UserManager -User $user -Manager $manager
    return $user
}

function Get-EmailParts($email)
{    
    $email = $email.Trim()    
    $mailNickName = $email.Split('@')[0] # Get the part of the email before the @ sign.
    $nameParts = $mailNickName.Split('.')
    $firstName = Capitalize-FirstLetter $nameParts[0]
    if ($nameParts[1]) { $lastName = Capitalize-FirstLetter $nameParts[1] }    
    $displayName = "$firstName $lastName"
    
    return @{
        "FirstName"    = $firstName
        "LastName"     = $lastName
        "DisplayName"  = $displayName
        "MailNickName" = $mailNickName
    }
}

function Capitalize-FirstLetter($string)
{
    if ($null -eq $string) { return }
    $string = $string.Trim()
    return $string.substring(0, 1).ToUpper() + $string.substring(1)
}

function New-M365User($upn, $mailNickName, $displayName, $firstName, $lastName, $jobTitle, $department, $usageLocation)
{
    $tempPassword = New-TempPassword
    Write-Host "Temp Password: $tempPassword" -ForegroundColor $infoColor
    $passwordProfile = @{
        "password"                             = $tempPassword
        "forceChangePasswordNextSignInWithMfa" = $true
    }

    $params = @{
        "UserPrincipalName" = $upn
        "MailNickName"      = $mailNickName
        "DisplayName"       = $displayName
        "GivenName"         = $firstName
        "Surname"           = $lastName
        "JobTitle"          = $jobTitle
        "Department"        = $department
        "UsageLocation"     = $usageLocation
        "PasswordProfile"   = $passwordProfile
        "AccountEnabled"    = $true
    }
    
    try
    {
        $user = New-MgUser @params -ErrorAction "Stop"
    }
    catch
    {
        $errorRecord = $_
        Write-Host "There was an issue creating user." -ForegroundColor $warningColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $warningColor
    }

    if ($user) { $script:logger.LogChange("M365 user created with UPN: $upn. Temp PW: $tempPassword") }

    return $user
}

function New-TempPassword
{
    $words = @("red", "orange", "yellow", "green", "blue", "purple", "silver", "gold", "flower", "mushroom", "lake", "river",
        "mountain", "valley", "jungle", "cavern", "rain", "thunder", "lightning", "storm", "fire", "lion", "wolf", "bear", "hawk",
        "dragon", "goblin", "fairy", "wizard", "sun", "moon", "emerald", "ruby", "saphire", "diamond", "treasure", "journey", "voyage",
        "adventure", "quest", "song", "dance", "painting", "magic", "castle", "dungeon", "tower", "sword", "torch", "potion")
    $specialChars = @('!', '@', '#', '$', '%', '^', '&', '*', '-', '+', '=', '?')

    $word1 = $words | Get-Random
    $coinFlip = Get-Random -Maximum 2 # max exclusive
    if ($coinFlip -eq 1) { $word1 = $word1.ToUpper() }
    
    $word2 = $words | Get-Random
    $coinFlip = Get-Random -Maximum 2 # max exclusive
    if ($coinFlip -eq 1) { $word2 = $word2.ToUpper() }

    $word3 = $words | Get-Random
    $coinFlip = Get-Random -Maximum 2 # max exclusive
    if ($coinFlip -eq 1) { $word3 = $word3.ToUpper() }

    $specialChar = $specialChars | Get-Random
    $num = Get-Random -Maximum 100 # max exclusive
    return $word1 + '/' + $word2 + '/' + $word3 + '/' + $specialChar + $num
}

function Set-UserManager($user, $manager)
{
    try
    {
        Set-MgUserManagerByRef -UserId $user.Id -OdataId "https://graph.microsoft.com/v1.0/users/$($manager.Id)" -ErrorAction "Stop" | Out-Null
        Write-Host "Manager assigned!" -ForegroundColor $successColor
    }
    catch
    {
        $errorRecord = $_
        Write-Host "There was an issue assigning manager." -ForegroundColor $warningColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $warningColor
        $script:logger.LogWarning("There was an issue assigning manager in M365: $($manager.UserPrincipalName)")
    }
}

function Invoke-GetWithRetry([ScriptBlock]$scriptBlock, $initialDelayInSeconds = 2, $maxRetries = 4)
{
    # API may not have the info we're trying to get yet. This will automatically retry a set amount of times.

    $retryCount = 0
    $delay = $initialDelayInSeconds
    do
    {
        # The call operator (&). Invokes a script block in a new script scope.
        # https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_operators?view=powershell-7.4#call-operator-
        $response = & $scriptBlock

        if ($null -eq $response)
        {
            if ($retryCount -ge 2)
            { 
                Write-Host "$scriptBlock returned null. Retrying in $delay seconds..." -ForegroundColor $warningColor
                Start-SleepTimer -Seconds $delay
            }
            else
            {
                Start-Sleep -Seconds $delay
            }            
            $delay *= 2
            $retryCount++
        }
    }
    while (($null -eq $response) -and ($retryCount -lt $maxRetries))

    if ($retryCount -ge $maxRetries) { Write-Host "Timed out trying to get a response." -ForegroundColor $warningColor }

    return $response
}

function Start-SleepTimer($seconds)
{
    for ($i = 0; $i -lt $seconds; $i++)
    {
        Write-Progress -Activity "Waiting..." -Status "$i / $seconds seconds"
        Start-Sleep -Seconds 1
    }
}

function Get-UserProperties($user, [switch]$retryGetManager)
{
    if ($retryGetManager)
    {
        $manager = Invoke-GetWithRetry { Get-UserManager -User $user }
    }
    else
    {
        $manager = Get-UserManager -User $user
    }
    
    $basicProps = [PSCustomObject]@{
        "Created Date/Time" = $user.CreatedDateTime.ToLocalTime()
        "Display Name"      = $user.DisplayName
        "UPN"               = $user.UserPrincipalName
        "Title"             = $user.JobTitle
        "Department"        = $user.Department
        "Manager"           = $manager.displayName
        "Usage Location"    = $user.UsageLocation
    }

    return @{
        "BasicProps" = $basicProps
        "Licenses"   = Get-UserLicenses $user
        "Groups"     = Get-UserGroups $user
        "AdminRoles" = Get-UserAdminRoles $user
    }
}

function Get-UserManager($user)
{
    try
    {
        $manager = Get-MgUserManager -UserId $user.UserPrincipalName -ErrorAction "Stop"
    }
    catch
    {
        $errorRecord = $_
        if ($errorRecord.Exception.Message -ilike "*[Request_ResourceNotFound]*")
        {
            return
        }        
        Write-Host "There was an issue getting user's manager." -ForegroundColor $warningColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $warningColor
        return
    }
    
    if ($null -eq $manager) { return }
    # Returns a dictionary. The additional properties hold all the relevant data.
    return $manager.AdditionalProperties
}

function Get-UserLicenses($user)
{
    if ($null -eq $script:licenseLookupTable)
    {
        # License SKU IDs listed here: https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference
        $script:licenseLookupTable = @{
            "8f0c5670-4e56-4892-b06d-91c085d7004f" = "App Connect IW"
            "4b9405b0-7788-4568-add1-99614e613b69" = "Exchange Online (Plan 1)"
            "19ec0d23-8335-4cbd-94ac-6050e30712fa" = "Exchange Online (Plan 2)"
            "efccb6f7-5641-4e0e-bd10-b4976e1bf68e" = "Enterprise Mobility + Security E3"
            "b05e124f-c7cc-45a0-a6aa-8cf78c946968" = "Enterprise Mobility + Security E5"
            "c2273bd0-dff7-4215-9ef5-2c7bcfb06425" = "Microsoft 365 Apps for Enterprise"
            "3b555118-da6a-4418-894f-7df1e2096870" = "Microsoft 365 Business Basic"
            "cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46" = "Microsoft 365 Business Premium"
            "f245ecc8-75af-4f8e-b61f-27d8114de5f3" = "Microsoft 365 Business Standard"
            "05e9a617-0261-4cee-bb44-138d3ef5d965" = "Microsoft 365 E3"
            "06ebc4ee-1bb5-47dd-8120-11324bc54e06" = "Microsoft 365 E5"
            "44575883-256e-4a79-9da4-ebe9acabe2b2" = "Microsoft 365 F1"
            "66b55226-6b4f-492c-910c-a3b7a3c9d993" = "Microsoft 365 F3"
            "4ef96642-f096-40de-a3e9-d83fb2f90211" = "Microsoft Defender for Office 365 (Plan 1)"
            "3dd6cf57-d688-4eed-ba52-9e40b5468c3e" = "Microsoft Defender for Office 365 (Plan 2)"
            "a403ebcc-fae0-4ca2-8c8c-7a907fd6c235" = "Microsoft Fabric (Free)"
            "dcb1a3ae-b33f-4487-846a-a640262fadf4" = "Microsoft Power Apps Plan 2 Trial"
            "f30db892-07e9-47e9-837c-80727f46fd3d" = "Microsoft Power Automate Free"
            "5b631642-bd26-49fe-bd20-1daaa972ef80" = "Microsoft PowerApps for Developer"
            "1f2f344a-700d-42c9-9427-5cea1d5d7ba6" = "Microsoft Stream"
            "3ab6abff-666f-4424-bfb7-f0bc274ec7bc" = "Microsoft Teams Essentials"
            "36a0f3b3-adb5-49ea-bf66-762134cf063a" = "Microsoft Teams Premium"
            "4cde982a-ede4-4409-9ae6-b003453c8ea6" = "Microsoft Teams Rooms Pro"
            "18181a46-0d4e-45cd-891e-60aabd171b4e" = "Office 365 E1"
            "6fd2c87f-b296-42f0-b197-1e91e994b900" = "Office 365 E3"
            "c7df2760-2c81-4ef7-b578-5b5392b571df" = "Office 365 E5"
            "7b26f5ab-a763-4c00-a1ac-f6c4b5506945" = "Power BI Premium P1"
            "f8a1db68-be16-40ed-86d5-cb42ce701560" = "Power BI Pro"
            "6470687e-a428-4b7a-bef2-8a291ad947c9" = "Windows Store for Business"
        }
    }

    try
    {
        $licenseDetails = Get-MGUserLicenseDetail -UserId $user.UserPrincipalName -ErrorAction "Stop"
    }
    catch
    {
        $errorRecord = $_
        Write-Host "There was an issue getting user's licenses." -ForegroundColor $warningColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $warningColor
        return
    }
    
    $licenses = [System.Collections.Generic.List[object]]::new(5)
    foreach ($license in $licenseDetails)
    {
        $licenseName = $script:licenseLookupTable[$license.SkuId]
        $licenses.Add( [PSCustomObject]@{"Name" = $licenseName; "SkuId" = $license.SkuId } )
    }

    return Write-Output $licenses -NoEnumerate
}

function Get-UserGroups($user)
{
    try
    {
        $groups = Get-MgUserMemberOfAsGroup -UserId $user.UserPrincipalName -ErrorAction "Stop"
    }
    catch
    {
        $errorRecord = $_
        Write-Host "There was an issue getting user's groups." -ForegroundColor $warningColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $warningColor
    }
    return $groups
}

function Get-UserAdminRoles($user)
{
    try
    {
        $adminRoles = Get-MgUserMemberOfAsDirectoryRole -UserId $user.UserPrincipalName -ErrorAction "Stop"
    }
    catch
    {
        $errorRecord = $_
        Write-Host "There was an issue getting user's admin roles." -ForegroundColor $warningColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $warningColor
    }
    return $adminRoles
}

function Show-UserProperties($basicProps, $licenses, $groups, $adminRoles)
{
    Show-Separator "M365 User"
    if ($basicProps) { $basicProps | Out-Host }    

    Show-Separator "Licenses"
    if ($licenses) { $licenses | Select-Object -ExpandProperty "Name" | Sort-Object | Out-Host }    

    Show-Separator "Groups"
    if ($groups) { $groups | Select-Object -ExpandProperty "DisplayName" | Sort-Object | Out-Host }
    
    Show-Separator "Admin Roles"
    if ($adminRoles) { $adminRoles | Select-Object -ExpandProperty "DisplayName" | Sort-Object | Out-Host }
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
    $hostWidthInChars = (Get-host).UI.RawUI.BufferSize.Width

    # Truncate title if it's too long.
    if (($separator.length) -gt $hostWidthInChars)
    {
        $separator = $separator.Remove($hostWidthInChars - 5)
        $separator += " "
    }

    # Pad with dashes.
    $separator = "--$($separator.PadRight($hostWidthInChars - 2, "-"))"

    if (-not($noLineBreaks))
    {        
        # Add line breaks.
        $separator = "`n$separator`n"
    }

    Write-Host $separator -ForegroundColor $color
}

function Prompt-MainMenu
{
    $selection = Read-Host ("`nWhat next?`n" +
        "[1] Show M365 User Info`n" +
        "[2] $(New-Checkbox $script:grantLicensesCompleted) Manage licenses`n" +
        "[3] $(New-Checkbox $script:assignGroupsCompleted) Manage groups`n" +
        "[4] $(New-Checkbox $script:grantMailboxesCompleted) Manage shared mailboxes`n" +
        "[5] $(New-Checkbox $script:gotoSetupCompleted) Setup GoTo account`n" +
        "[6] Finish`n")

    do
    {
        $selection = $selection.Trim()
        $validSelection = $selection -imatch '^[1-6]$' # regex matches 1-6
        if (-not($validSelection))
        {
            Write-Host "Please enter 1-6." -ForegroundColor $warningColor
            $selection = Read-Host
        }
    }
    while (-not($validSelection))

    return [int]$selection
}

function New-Checkbox($checked)
{
    if ($checked)
    {
        return "[X]"
    }
    return "[ ]"
}

function Start-M365LicenseWizard($user)
{
    $keepGoing = $true
    while ($keepGoing)
    {
        $selection = Prompt-LicenseMenu
        switch ($selection)
        {
            1 # View assigned licenses
            {
                Show-Separator "Licenses"
                # Capture this in a var first because Get-UserLicenses does not work in the pipeline. (It doesn't enumerate its output!)
                $licenses = Get-UserLicenses $user
                if ($null -eq $licenses) { break }
                $licenses | Select-Object -ExpandProperty "Name" | Sort-Object | Out-Host
                break
            }
            2 # Grant license
            {
                $availableLicenses = Get-AvailableLicenses
                if ($null -eq $availableLicenses) { break }
                $license = Prompt-LicenseToGrant $availableLicenses
                if ($null -eq $license) { break }
                $hasLicense = Confirm-HasLicense -User $user -License $license
                if ($hasLicense) 
                { 
                    Write-Host "User already has that license." -ForegroundColor $warningColor
                    break
                }
                $success = Grant-License -User $user -License $license
                if ($success) { $script:grantLicensesCompleted = $true }      
                break
            }
            3 # Revoke license
            {
                $assignedLicenses = Get-UserLicenses $user
                if ($null -eq $assignedLicenses) { break }
                $license = Prompt-LicenseToRevoke $assignedLicenses
                if ($null -eq $license) { break }
                $hasLicense = Confirm-HasLicense -User $user -License $license
                if (-not($hasLicense)) 
                { 
                    Write-Host "User doesn't have that license." -ForegroundColor $warningColor
                    break
                }
                Revoke-License -User $user -License $license
                break
            }
            4 # Finish
            {
                $keepGoing = $false
                break
            }
        }
    }
}

function Prompt-LicenseMenu
{
    $selection = Read-Host ("`nChoose an option:`n" +
    "[1] View assigned licenses`n" +                        
    "[2] Grant license`n" +
    "[3] Revoke license`n" +
    "[4] Finish with licenses`n")

    do
    {        
        $selection = $selection.Trim()
        $validSelection = $selection -imatch '^[1-4]$' # regex matches 1-4
        if (-not($validSelection))
        {
            Write-Host "Please enter 1-4." -ForegroundColor $warningColor
            $selection = Read-Host
        }
    }
    while (-not($validSelection))

    return [int]$selection
}

function Get-AvailableLicenses
{
    $uri = "https://graph.microsoft.com/v1.0/subscribedSkus"
    try
    {
        $licenses = Invoke-MgGraphRequest -Method "Get" -Uri $uri -ErrorAction "Stop"
    }    
    catch
    {
        $errorRecord = $_
        Write-Host "There was an issue getting available licenses." -ForegroundColor $warningColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $warningColor
        return
    }

    $licenseTable = [System.Collections.Generic.List[object]]::new(30)
    foreach ($license in $licenses.value)
    {
        $name = $script:licenseLookupTable[$license.skuId]
        if ($null -eq $name ) { $name = $license.skuPartNumber }
        $amountPurchased = $license.prepaidUnits.enabled
        $amountAvailable = $amountPurchased - $license.consumedUnits

        $licenseInfo = [PSCustomObject]@{
            "Name"      = $name
            "Available" = $amountAvailable
            "Purchased" = $amountPurchased
            "SkuId"     = $license.skuId        
        }
        $licenseTable.Add($licenseInfo)
    }

    return Write-Output $licenseTable -NoEnumerate
}

function Prompt-LicenseToGrant($availableLicenses)
{
    if ($null -eq $availableLicenses -or $availableLicenses.Count -eq 0)
    {
        Write-Host "There are no available licenses to grant." -ForegroundColor $warningColor
        return
    }
    
    # Display available licenses with an option number next to each.
    $option = 0    
    $availableLicenses | Sort-Object -Property "Name" | ForEach-Object { $option++; $_ | Add-Member -NotePropertyName "Option" -NotePropertyValue $option }
    $availableLicenses | Sort-Object -Property "Option" | Format-Table -Property @("Option", "Name", "Available", "Purchased") | Out-Host
    $selection = (Read-Host "Select an option (1-$option)") -As [int]

    do
    {        
        # Check that selection is a number between 1 and option count. (Avoids use of regex because that's not great for matching multi-digit number ranges.)
        $validSelection = ($selection -is [int]) -and (($selection -ge 1) -and ($selection -le $option))
        if (-not($validSelection)) 
        {
            Write-Host "Please enter 1-$option." -ForegroundColor $warningColor
            $tryAgain = Prompt-YesOrNo "Try again?"
            if (-not($tryAgain)) { return }
            $selection = (Read-Host "Select an option (1-$option)") -As [int]
        }
    }
    while (-not($validSelection))

    foreach ($license in $availableLicenses)
    {
        if ($license.option -eq [int]$selection)
        {
            return $license
        }
    }
}

function Confirm-HasLicense($user, $license)
{
    $grantedLicenses = Get-UserLicenses -User $user
    if ($null -eq $grantedLicenses) { return $false }
    foreach ($grantedLicense in $grantedLicenses)
    {
        if ($grantedLicense.SkuId -eq $license.SkuId) 
        {
            return $true
        }
    }
    return $false
}

function Grant-License($user, $license)
{
    try
    {
        Set-MgUserLicense -UserId $user.Id -AddLicenses @{SkuId = $license.SkuId } -RemoveLicenses @() -ErrorAction "Stop" | Out-Null
        Write-Host "License granted: $($license.name)" -ForegroundColor $successColor
        $script:logger.LogChange("Granted M365 license: $($license.name)")
        $success = $true
    }
    catch
    {
        $errorRecord = $_
        Write-Host "There was an issue granting the license." -ForegroundColor $warningColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $warningColor
        $script:logger.LogWarning("There was an issue granting M365 license: $($license.name)")
        $success = $false
    }
    return $success
}

function Prompt-LicenseToRevoke($assignedLicenses)
{    
    if (($null -eq $assignedLicenses) -or ($assignedLicenses.Count -eq 0))
    {
        Write-Host "User has no licenses to revoke." -ForegroundColor $warningColor
        return
    }
    
    # Display assigned licenses with an option number next to each.
    $option = 0
    $optionList = $assignedLicenses | Sort-Object "Name" | ForEach-Object { 
        $option++
        [PSCustomObject]@{
            "Option" = $option
            "Name"   = $_.Name
            "SkuId"  = $_.SkuId
        }
    }
    $optionList | Sort-Object -Property "Option" | Format-Table -Property @("Option", "Name") | Out-Host

    $selection = (Read-Host "Select an option (1-$option)") -As [int]
    do
    {
        # Check that selection is a number between 1 and option count. (Avoids use of regex because that's not great for matching multi-digit number ranges.)
        $validSelection = ($selection -is [int]) -and (($selection -ge 1) -and ($selection -le $option))
        if (-not($validSelection)) 
        {
            Write-Host "Please enter 1-$option." -ForegroundColor $warningColor
            $tryAgain = Prompt-YesOrNo "Try again?"
            if (-not($tryAgain)) { return }
            $selection = (Read-Host "Select an option (1-$option)") -As [int]
        }
    }
    while (-not($validSelection))

    return $optionList[$selection - 1]
}

function Revoke-License($user, $license)
{
    try
    {
        Set-MgUserLicense -UserId $user.Id -AddLicenses @() -RemoveLicenses @($license.SkuId) -ErrorAction "Stop" | Out-Null
        Write-Host "License revoked: $($license.name)" -ForegroundColor $successColor
        $script:logger.LogChange("Revoked M365 license: $($license.name)")
    }
    catch
    {
        $errorRecord = $_
        Write-Host "There was an issue revoking the license." -ForegroundColor $warningColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $warningColor
        $script:logger.LogWarning("There was an issue revoking M365 license: $($license.name)")
    }
}

function Start-M365GroupWizard($user)
{
    $keepGoing = $true
    while ($keepGoing)
    {
        $selection = Prompt-GroupMenu

        # We give this switch statement a label so we can break out of it from nested loops.
        :outerSwitch switch ($selection)
        {
            1 # View assigned groups
            {
                Show-Separator "Groups"
                $assignedGroups = Get-UserGroups -User $user
                if ($null -eq $assignedGroups) { break }
                $assignedGroups | Select-Object -ExpandProperty "DisplayName" | Sort-Object | Out-Host
                break
            }
            2 # Assign group
            {
                do
                {
                    $groupEmail = Prompt-BrsEmail "Enter group email"
                    $group = Get-M365Group $groupEmail
                    if ($null -eq $group)
                    {
                        $tryAgain = Prompt-YesOrNo "Try again?"
                        if (-not($tryAgain)) { break outerSwitch }
                        continue
                    }

                    $isAlreadyMember = Confirm-IsMemberOfGroup -User $user -Group $group
                    if ($isAlreadyMember)
                    {
                        Write-Host "$($user.DisplayName) is already a member of the group: $($group.DisplayName)." -ForegroundColor $warningColor
                        $tryAgain = Prompt-YesOrNo "Try again?"
                        if (-not($tryAgain)) { break outerSwitch }
                        continue
                    }
                }
                while (($null -eq $group) -or ($isAlreadyMember))

                $success = Assign-M365Group -User $user -Group $group
                if ($success) { $script:assignGroupsCompleted = $true }                
                break
            }
            3 # Remove group
            {
                $assignedGroups = Get-UserGroups -User $user
                if ($null -eq $assignedGroups) { break }
                $group = Prompt-GroupToUnassign $assignedGroups
                if ($null -eq $group) { break }
                $isAlreadyMember = Confirm-IsMemberOfGroup -User $user -Group $group
                if (-not($isAlreadyMember))
                {
                    Write-Host "User isn't a member of that group." -ForegroundColor $warningColor
                    break
                }
                Unassign-M365Group -User $user -Group $group
                break                
            }
            4 # Finish with groups
            {
                $keepGoing = $false
                break
            }
        }
    }
}

function Prompt-GroupMenu
{
    $selection = Read-Host ("`nChoose an option:`n" +
    "[1] View assigned groups`n" +                        
    "[2] Assign group`n" +
    "[3] Remove group`n" +
    "[4] Finish with groups`n")

    do
    {        
        $selection = $selection.Trim()
        $validSelection = $selection -imatch '^[1-4]$' # regex matches 1-4
        if (-not($validSelection))
        {
            Write-Host "Please enter 1-4." -ForegroundColor $warningColor
            $selection = Read-Host
        }
    }
    while (-not($validSelection))

    return [int]$selection
}

function Get-M365Group($email)
{
    try
    {
        $group = Get-MgGroup -Filter "mail eq '$email'" -ErrorAction "Stop"
    }
    catch
    {
        $errorRecord = $_    
        Write-Host "There was an issue getting the group." -ForegroundColor $warningColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $warningColor
        return
    }

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

function Confirm-IsMemberOfGroup($user, $group)
{
    $currentAssignedGroups = Get-UserGroups -User $user
    if ($null -eq $currentAssignedGroups) { return $false}
    foreach ($assignedGroup in $currentAssignedGroups)
    {
        if ($assignedGroup.Id -eq $group.Id)
        {            
            return $true
        }
    }
    return $false
}

function Assign-M365Group($user, $group)
{  
    try
    {
        New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $user.Id -ErrorAction "Stop" | Out-Null
        Write-Host "Group assigned: $($group.DisplayName)" -ForegroundColor $successColor
        $script:logger.LogChange("Assigned M365 group: $($group.DisplayName)")
        $success = $true
    }
    catch
    {
        $errorRecord = $_
        Write-Host "There was an issue assigning the group." -ForegroundColor $warningColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $warningColor
        $script:logger.LogWarning("There was an issue assigning M365 group: $($group.DisplayName)")
        $success = $false
    }
    return $success
}

function Prompt-GroupToUnassign($assignedGroups)
{
    if (($null -eq $assignedGroups) -or ($assignedGroups.Count -eq 0))
    {
        Write-Host "User has no groups to unassign." -ForegroundColor $warningColor
        return
    }

    # Display assigned groups with an option number next to each.
    $option = 0
    $optionList = $assignedGroups | Sort-Object "DisplayName" | ForEach-Object {
        $option++
        [PSCustomObject]@{
            "Option" = $option
            "DisplayName" = $_.DisplayName
            "Mail" = $_.Mail
            "Description" = $_.Description
            "Id" = $_.Id
        }
    }
    $optionList | Sort-Object -Property "Option" | Format-Table -Property @("Option", "DisplayName", "Mail", "Description") | Out-Host

    $selection = (Read-Host "Select an option (1-$option)") -As [int]
    do
    {
        # Check that selection is a number between 1 and option count. (Avoids use of regex because that's not great for matching multi-digit number ranges.)
        $validSelection = ($selection -is [int]) -and (($selection -ge 1) -and ($selection -le $option))
        if (-not($validSelection)) 
        {
            Write-Host "Please enter 1-$option." -ForegroundColor $warningColor
            $tryAgain = Prompt-YesOrNo "Try again?"
            if (-not($tryAgain)) { return }
            $selection = (Read-Host "Select an option (1-$option)") -As [int]
        }
    }
    while (-not($validSelection))

    return $optionList[$selection - 1]
}

function Unassign-M365Group($user, $group)
{
    try
    {
        Remove-MgGroupMemberByRef -GroupId $group.Id -DirectoryObjectId $user.Id -ErrorAction "Stop" | Out-Null
        Write-Host "Group unassigned: $($group.DisplayName)" -ForegroundColor $successColor
        $script:logger.LogChange("Unassigned M365 group: $($group.DisplayName)")
    }
    catch
    {
        $errorRecord = $_
        Write-Host "There was an issue unassigning the group." -ForegroundColor $warningColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $warningColor
        $script:logger.LogWarning("There was an issue unassigning M365 group: $($group.DisplayName)")
    }
}

function Start-MailboxWizard($user)
{
    $keepGoing = $true
    while ($keepGoing)
    {
        $selection = Prompt-MailboxMenu

        # We give this switch statement a label so we can break out of it from nested loops.
        :outerSwitch switch ($selection)
        {
            1 # View assigned mailboxes
            {
                Write-Host ("Sorry, at this time there is no fast way to get all assigned mailboxes for a user.`n" +
                    "For that you may run this script instead, but it takes a little while.`n" +
                    "https://help.blueravensolar.com/a/solutions/articles/19000077594") -ForegroundColor $infoColor
                break
            }
            2 # Grant access to mailbox
            {
                do
                {
                    $mailboxUpn = Prompt-BrsEmail "Enter mailbox email"
                    $mailbox = Get-SharedMailbox $mailboxUpn
                    if ($null -eq $mailbox)
                    {
                        $tryAgain = Prompt-YesOrNo "Try again?"
                        if (-not($tryAgain)) { break outerSwitch }
                        continue
                    }
                }
                while ($null -eq $mailbox)

                $success = Grant-MailboxAccess -User $user -Mailbox $mailbox
                if ($success) { $script:grantMailboxesCompleted = $true }                
                break
            }
            3 # Revoke access to mailbox
            {
                do
                {
                    $mailboxUpn = Prompt-BrsEmail "Enter mailbox email"
                    $mailbox = Get-SharedMailbox $mailboxUpn
                    if ($null -eq $mailbox)
                    {
                        $tryAgain = Prompt-YesOrNo "Try again?"
                        if (-not($tryAgain)) { break outerSwitch }
                        continue
                    }
                }
                while ($null -eq $mailbox)

                Revoke-MailboxAccess -User $user -Mailbox $mailbox
                break
            }
            4 # Finish with mailboxes
            {
                $keepGoing = $false
                break
            }
        }
    }
}

function Prompt-MailboxMenu
{
    $selection = Read-Host ("`nChoose an option:`n" +
    "[1] View assigned mailboxes`n" +                        
    "[2] Grant access to mailbox`n" +
    "[3] Revoke access to mailbox`n" +
    "[4] Finish with mailboxes`n")

    do
    {
        $selection = $selection.Trim()
        $validSelection = $selection -imatch '^[1-4]$' # regex matches 1-4
        if (-not($validSelection))
        {
            Write-Host "Please enter 1-4." -ForegroundColor $warningColor
            $selection = Read-Host
        }
    }
    while (-not($validSelection))

    return [int]$selection
}

function Get-SharedMailbox($upn)
{
    if ($null -eq $upn) { throw "Can't get shared mailbox. UPN is null." }

    try
    {
        $mailbox = Get-EXOMailbox -Identity $upn -ErrorAction "Stop"
    }
    catch
    {
        $errorRecord = $_
        if ($errorRecord.Exception.Message -ilike "*HttpStatusCode=404*")
        {
            Write-Host "Mailbox not found." -ForegroundColor $warningColor
            return
        }
        Write-Host "There was an issue getting the mailbox." -ForegroundColor $warningColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $warningColor
        return
    }
    
    if (($mailbox) -and ($mailbox.RecipientTypeDetails -eq 'SharedMailbox'))
    {
        Write-Host "Mailbox found!" -ForegroundColor $successColor
        $mailbox | Select-Object -Property @("DisplayName", "UserPrincipalName", @{ label = "Type"; expression = { $_.RecipientTypeDetails } }) | Out-Host        
    }
    elseif ($mailbox)
    {
        Write-Host "Mailbox was found but it's not a shared mailbox. Type is '$($mailbox.RecipientTypeDetails)'." -ForegroundColor $warningColor
        return
    }

    return $mailbox
}

function Grant-MailboxAccess($user, $mailbox)
{
    Write-Host "Access type to grant?" 
    $accessType = Prompt-MailboxAccessType
    try
    {
        switch ($accessType)
        {
            1 # Read and Manage
            {
                Add-MailboxPermission -Identity $mailbox.UserPrincipalName -User $user.UserPrincipalName -AccessRights "FullAccess" -Confirm:$false -WarningAction "SilentlyContinue" -ErrorAction "Stop" | Out-Null
                Write-Host "Granted read and manage access to mailbox! (If they didn't already have it.)" -ForegroundColor $successColor
                $script:logger.LogChange("Granted read and manage access to mailbox: $($mailbox.UserPrincipalName)")
                break
            }
            2 # Send As
            {
                Add-RecipientPermission -Identity $mailbox.UserPrincipalName -Trustee $user.UserPrincipalName -AccessRights "SendAs" -Confirm:$false -WarningAction "SilentlyContinue" -ErrorAction "Stop" | Out-Null
                Write-Host "Granted send as access to mailbox! (If they didn't already have it.)" -ForegroundColor $successColor
                $script:logger.LogChange("Granted send as access to mailbox: $($mailbox.UserPrincipalName)")
                break
            }
            3 # Both
            {
                Add-MailboxPermission -Identity $mailbox.UserPrincipalName -User $user.UserPrincipalName -AccessRights "FullAccess" -Confirm:$false -WarningAction "SilentlyContinue" -ErrorAction "Stop" | Out-Null
                Add-RecipientPermission -Identity $mailbox.UserPrincipalName -Trustee $user.UserPrincipalName -AccessRights "SendAs" -Confirm:$false -WarningAction "SilentlyContinue" -ErrorAction "Stop" | Out-Null
                Write-Host "Granted read and manage + send as access to mailbox! (If they didn't already have it.)" -ForegroundColor $successColor
                $script:logger.LogChange("Granted read and manage + send as access to mailbox: $($mailbox.UserPrincipalName)")
                break
            }
            4 # Go back
            {
                return
            }
        }
        $success = $true
    }
    catch
    {
        $errorRecord = $_
        Write-Host "There was an issue granting mailbox access." -ForegroundColor $warningColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $warningColor
        $script:logger.LogWarning("There was an issue granting access to mailbox: $($mailbox.UserPrincipalName)")
        $success = $false
    }
    return $success
}

function Prompt-MailboxAccessType
{
    $accessType = Read-Host  ("[1] Read & Manage`n" +
        "[2] Send As`n" +
        "[3] Both`n" +
        "[4] Go back`n")

    do
    {
        $accessType = $accessType.Trim()
        $isValidResponse = $accessType -imatch '^[1-4]$' # regex matches 1-4
        if (-not($isValidResponse))
        {
            Write-Host "Please enter 1-4." -ForegroundColor $warningColor
            $accessType = Read-Host
        }
    }
    while (-not($isValidResponse))

    return [int]$accessType
}

function Revoke-MailboxAccess($user, $mailbox)
{
    Write-Host "Access type to revoke?"
    $accessType = Prompt-MailboxAccessType
    try
    {
        switch ($accessType)
        {
            1 # Read and Manage
            {
                Remove-MailboxPermission -Identity $mailbox.UserPrincipalName -User $user.UserPrincipalName -AccessRights "FullAccess" -Confirm:$false -WarningAction "SilentlyContinue" -ErrorAction "Stop"
                Write-Host "Revoked read and manage access to mailbox! (Assuming they had it.)" -ForegroundColor $successColor
                $script:logger.LogChange("Revoked read and manage access to mailbox: $($mailbox.UserPrincipalName)")
                break
            }
            2 # Send As
            {
                Remove-RecipientPermission -Identity $mailbox.UserPrincipalName -Trustee $user.UserPrincipalName -AccessRights "SendAs" -Confirm:$false -WarningAction "SilentlyContinue" -ErrorAction "Stop"
                Write-Host "Revoked send as access to mailbox! (Assuming they had it.)" -ForegroundColor $successColor
                $script:logger.LogChange("Revoked send as access to mailbox: $($mailbox.UserPrincipalName)")
                break
            }
            3 # Both
            {
                Remove-MailboxPermission -Identity $mailbox.UserPrincipalName -User $user.UserPrincipalName -AccessRights "FullAccess" -Confirm:$false -WarningAction "SilentlyContinue" -ErrorAction "Stop"
                Remove-RecipientPermission -Identity $mailbox.UserPrincipalName -Trustee $user.UserPrincipalName -AccessRights "SendAs" -Confirm:$false -WarningAction "SilentlyContinue" -ErrorAction "Stop"
                Write-Host "Revoked read and manage + send as access to mailbox! (Assuming they had it.)" -ForegroundColor $successColor
                $script:logger.LogChange("Revoked read and manage + send as access to mailbox: $($mailbox.UserPrincipalName)")
                break
            }
            4 # Go back
            {
                return
            }
        }
    }
    catch
    {
        $errorRecord = $_
        Write-Host "There was an issue revoking mailbox access." -ForegroundColor $warningColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $warningColor
        $script:logger.LogWarning("There was an issue revoking access to mailbox: $($mailbox.UserPrincipalName)")
    }
}

function SafelyInvoke-RestMethod($method, $uri, $headers, $body)
{
    try
    {
        $response = Invoke-RestMethod -Method $method -Uri $uri -Headers $headers -Body $body -ErrorVariable "responseError"
        $success = $true
    }
    catch
    {
        Write-Host $responseError[0].Message -ForegroundColor $warningColor
        $success = $false
    }

    if ($response)
    {
        return $response
    }
    return $success
}

function UriEncode-QueryParam($queryParam)
{ 
    return [uri]::EscapeDataString($queryParam)
}

function ConvertTo-Base64($text)
{
    return [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($text))
}

# classes
class Logger
{
    # Note: Don't try to implement logger as a singleton. 
    # For one you can't truly make private constructors & members in PowerShell.
    # For two the static instance will persist across the global terminal scope/session, even after exiting the script.

    # Constructors
    Logger()
    { 
        $this.Logs = [System.Collections.Generic.List[object]]::new(10)
    }
    
    # fields
    hidden [System.Collections.Generic.List[object]] $Logs

    # methods
    [void] LogChange($message)
    {
        $logEntry = @{
            Timestamp = Get-Date
            Level     = 'Change'
            Message   = $message
        }
        $this.Logs.Add($logEntry)
    }

    [void] LogWarning($message)
    {
        $logEntry = @{
            Timestamp = Get-Date
            Level     = 'Warning'
            Message   = $message
        }
        $this.Logs.Add($logEntry)
    }

    [void] LogError($message)
    {
        $logEntry = @{
            Timestamp = Get-Date
            Level     = 'Error'
            Message   = $message
        }
        $this.Logs.Add($logEntry)
    }

    [void] ShowLogs()
    {
        Show-Separator "Logs"
        $this.Logs = @(,$this.Logs) | Sort-Object "Timestamp" # Wrapping this.Logs in array is important so that pipeline works properly when Logs.Count is 1.
        foreach ($log in $this.Logs)
        {
            # Pipe to Get-Date for a simplified timestamp.
            $message = "[$($log.Timestamp | Get-Date -Format 'yyyy-MM-dd hh:mm tt')] $($log.Message)"
            switch ($log.Level)
            {
                'Change' 
                {
                    Write-Host $message -ForegroundColor $script:successColor
                    break
                }
                'Warning' 
                { 
                    Write-Host $message -ForegroundColor $script:warningColor 
                    break
                }
                'Error' 
                { 
                    Write-Host $message -ForegroundColor $script:failColor
                    break
                }
            }
        }
    }
}

class GotoWizard
{
    # fields
    [string] $upn
    [object] $gotoSecret
    [string] $clientId
    [string] $clientSecret
    [string] $accountKey
    [string] $accessToken
    [object] $gotoUser
    [System.Collections.Specialized.OrderedDictionary] $allCustomRoles
    [bool] $assignRoleCompleted
    [bool] $assignCallerIdCompleted

    # constructors
    GotoWizard($upn)
    {
        Write-Host "Connecting to GoTo..." -ForegroundColor $script:infoColor
        $this.upn = $upn
        $this.gotoSecret = $this.GetApiSecret()
        if ($null -eq $this.gotoSecret) { return }

        $this.clientId = $this.gotoSecret.ClientID
        $this.clientSecret = $this.gotoSecret.ClientSecret
        $this.accountKey = $this.gotoSecret.AccountKey
        $this.accessToken = $this.GetAccessToken()
        if ($null -eq $this.accessToken) { return }

        $this.gotoUser = $this.GetUser()
        $this.allCustomRoles = $this.GetAllCustomRoles()
    }

    # Default constructor defined for Pester testing purposes.
    GotoWizard() {}

    # methods
    [void] Start()
    {       
        if ($null -eq $this.gotoSecret)
        {
            Write-Host "Could not start GoTo wizard because GoTo secret was not obtained." -ForegroundColor $script:warningColor
            return
        }

        if ($null -eq $this.accessToken)
        {
            Write-Host "Could not start GoTo wizard because GoTo access token was not obtained." -ForegroundColor $script:warningColor
            return
        }
        
        if ($this.gotoUser)
        {
            Write-Host "Found GoTo user!" -ForegroundColor $script:successColor
        }
        else
        {
            Write-Host "Goto user not found." -ForegroundColor $script:infoColor

            $shouldCreate = Prompt-YesOrNo "Create user?"
            if (-not($shouldCreate)) { return }

            $emailParts = Get-EmailParts $this.upn
            $correct = Prompt-YesOrNo "This first name?: $($emailParts.FirstName)"
            if (-not($correct)) { $emailParts.FirstName = Read-Host "Enter first name" }

            $correct = Prompt-YesOrNo "This last name?: $($emailParts.LastName)"
            if (-not($correct)) { $emailParts.LastName = Read-Host "Enter last name" }
            
            $this.gotouser = $this.CreateUser($this.upn, $emailParts.FirstName, $emailParts.LastName)
        }

        $keepGoing = $true
        while ($keepGoing)
        {
            $selection = $this.PromptMenu()

            :outerSwitch switch ($selection)
            {
                1 # Show user info
                {
                    $this.ShowUserInfo()
                    break
                }
                2 # Assign role
                {
                    $roleSelection = $this.PromptRoleToAssign()
                    $this.AssignUserRole($roleSelection)
                    break      
                }
                3 # Assign outbound caller ID
                {
                    $this.DisplayLinkToOutboundCallerId()
                    break
                }
                4 # Finish
                {
                    if ($this.gotoUser) { $script:gotoSetupCompleted = $true }                    
                    $keepGoing = $false
                    break
                }
            }
        }
    }

    [object] GetApiSecret()
    {
        $secret = $null
        $keepGoing = $true        
        do
        {
            try
            {
                $secret = Get-Secret -Name "YZJrirO-73fEk6aZO5QgZg" -AsPlainText
                $keepGoing = $false
            }
            catch
            {
                $errorRecord = $_
                $exceptionType = $errorRecord.Exception.GetType().FullName
                # We must check for exception type this way instead of catching it directly since the type is unavailable when the class definition loads.
                if ($exceptionType -eq "Microsoft.PowerShell.SecretManagement.PasswordRequiredException")
                {
                    Write-Host "You entered an incorrect password for your secret store." -ForegroundColor $script:warningColor
                }
                else
                {
                    Write-Host "There was an issue getting GoTo API secrets." -ForegroundColor $script:warningColor
                    Write-Host $errorRecord.Exception.Message -ForegroundColor $script:warningColor

                }
                $tryAgain = Prompt-YesOrNo "Try again?"
                if (-not($tryAgain)) { $keepGoing = $false }
            }
        }
        while ($keepGoing)

        return $secret
    }

    [string] GetAccessToken()
    {
        $token = $this.TryRefreshToken()
        if ($token) { return $token }
        return $this.GetAccessTokenByAuthCode()
    }

    [string] TryRefreshToken()
    {
        try
        {
            $refreshToken = Get-Secret -Name "gtrt" -AsPlainText -ErrorAction "Stop"
        }
        catch
        {
            return $null
        }
        
        $method = "Post"
        $uri = "https://authentication.logmeininc.com/oauth/token"
        $headers = @{
            "Authorization" = "Basic $(ConvertTo-Base64 "$($this.clientId):$($this.clientSecret)")"
        }
        $body = @{
            "grant_type"    = "refresh_token"
            "refresh_token" = $refreshToken
        }

        try
        {
            $response = Invoke-RestMethod -Method $method -Uri $uri -Headers $headers -Body $body
        }
        catch
        {
            return $null
        }

        if ($response)
        {
            if ($response.refresh_token)
            {
                Set-Secret -Name "gtrt" -Secret $response.refresh_token -Vault "LocalStore"
            }            
            return $response.access_token
        }
        return $null
    }

    [string] GetAccessTokenByAuthCode()
    {
        # Function handles GoTo connect OAUTH2 authorization code grant flow. (Obtains auth code then uses that to get a temp access token.)
        # https://developer.goto.com/guides/Authentication/03_HOW_accessToken/
        # https://developer.goto.com/Authentication/#section/Authorization-Flows

        $authUri = "https://authentication.logmeininc.com/oauth/authorize"
        $accessTokenUri = "https://authentication.logmeininc.com/oauth/token"
        $redirectUri = "https://localhost"
    
        $authCode = Invoke-OAuth2AuthorizationEndpoint -Uri $authUri -Client_id $this.clientId -Redirect_uri $redirectUri
        $token = Invoke-OAuth2TokenEndpoint @authCode -Uri $accessTokenUri -Client_secret $this.clientSecret -Client_auth_method "client_secret_basic"
        if ($token.refresh_token) 
        { 
            Set-Secret -Name "gtrt" -Secret $token.refresh_token -Vault "LocalStore"
        }
        return $token.access_token
    }

    [object] CreateUser($upn, $firstName, $lastName)
    {       
        $method = "Post"
        $uri = "https://api.getgo.com/admin/rest/v1/accounts/$($this.accountKey)/users"
        $headers = @{
            "Authorization" = "Bearer $($this.accessToken)"
            "Content-Type" = "application/json"
        }        
        $body = @{
            "users" = @( [PSCustomObject]@{
                "email" = $upn
                "firstName" = $firstName
                "lastName" = $lastName
            })
            "licenseKeys" = @( 7902142105473202120 ) # License key for GoTo Connect Voice.
        } | ConvertTo-Json

        $response = SafelyInvoke-RestMethod -Method $method -Uri $uri -Headers $headers -Body $body
        if ($response)
        {
            Write-Host "GoTo user created!" -ForegroundColor $script:successColor
            $script:logger.LogChange("GoTo user created with username: $upn")
        }
        else
        {
            $script:logger.LogWarning("There was an issue creating GoTo user with username: $upn")
        }
        return $response
    }

    [int] PromptMenu()
    {
        $selection = Read-Host ("`nChoose an option:`n" +
            "[1] Show GoTo user info`n" +
            "[2] $(New-Checkbox $this.assignRoleCompleted) Assign role`n" +
            "[3] $(New-Checkbox $this.assignCallerIdCompleted) Assign outbound caller ID`n" +
            "[4] Finish GoTo setup`n")

        do
        {
            $selection = $selection.Trim()
            $validSelection = $selection -imatch '^[1-4]$' # regex matches 1-4
            if (-not($validSelection))
            {
                Write-Host "Please enter 1-4." -ForegroundColor $script:warningColor
                $selection = Read-Host
            }
        }
        while (-not($validSelection))

        return [int]$selection
    }

    [object[]] GetUser()
    {
        $method = "Get"
        # Docs: https://developer.goto.com/admin/#operation/Get%20Users
        $uri = "https://api.getgo.com/admin/rest/v1/accounts/$($this.accountKey)/users"
        $headers = @{
            "Authorization" = "Bearer $($this.accessToken)"
        }
        # eq is the "equals" operator. https://developer.goto.com/admin/#section/Resource-Filtering 
        $emailQuery = "email eq `"$($this.upn)`""

        # Here we apply the query params directly to the URI instead of passing them to the body param of Invoke-RestMethod.
        # That way would apply URI encoding to the query params, but encode spaces with + signs instead of %20. The GoTo API doesn't accept this.
        $emailQuery = UriEncode-QueryParam $emailQuery
        $uri = $uri + "?filter=$emailQuery"

        $response = SafelyInvoke-RestMethod -Method $method -Uri $uri -Headers $headers
        if ($response.results) { return $response.results[0] }
        return $null
    }

    [void] ShowUserInfo()
    {
        $this.gotoUser = Invoke-GetWithRetry { $this.GetUser() }
        $role = $this.GetUserRole()        
        # I'd also show their external caller ID, but this is not accessible from the public API.
        $this.gotoUser | Add-Member -NotePropertyName "role" -NotePropertyValue $role -PassThru | Format-List -Property @("email", "firstName", "lastName", "role") | Out-Host
    }

    [string] GetUserRole()
    {
        $this.gotoUser = Invoke-GetWithRetry { $this.GetUser() }
        
        # Check for a built-in (system) admin role.
        if ($this.gotoUser.adminRoles)
        {
            if ($this.gotoUser.adminRoles -eq "SUPER_USER")
            {
                return "Super Admin"
            }

            if ($this.gotoUser.adminRoles -eq "MANAGE_ACCOUNT")
            {
                return "Admin (Configure PBX)"
            }
        }
        
        # Check for a custom role.
        if ($this.gotoUser.roleSets)
        {
            return $this.allCustomRoles[$this.gotoUser.roleSets[0]]
        }

        # If they have no built-in (system) admin role or custom role, their role is simply "Member".
        return "Member"
    }

    [System.Collections.Specialized.OrderedDictionary] GetAllCustomRoles()
    {
        # Gets custom roles (not system roles)

        $method = "Get"
        $uri = "https://api.getgo.com/admin/rest/v1/accounts/$($this.accountKey)/rolesets"
        $headers = @{
            "Authorization" = "Bearer $($this.accessToken)"
        }

        $response = SafelyInvoke-RestMethod -Method $method -Uri $uri -Headers $headers
        $roles = [ordered]@{} # ordered dictionary
        foreach ($role in $response.results)
        {
            $roles.Add($role.id, $role.name)
        }
        return $roles
    }

    [int] PromptRoleToAssign()
    {
        $menuText = ("`nSelect a role to assign:`n" +
            "[1] Super Admin`n" +
            "[2] Admin (Configure PBX)`n" +
            "[3] Member`n")

        $optionsCount = 3
        foreach ($roleName in $this.allCustomRoles.Values)
        {
            $optionsCount++
            $menuText = $menuText + "[$optionsCount] $roleName`n"            
        }
        $selection = Read-Host $menuText

        do
        {
            $selection = $selection.Trim()
            $validSelection = $selection -imatch "^[1-$optionsCount]$" # regex matches 1-optionsCount
            if (-not($validSelection))
            {
                Write-Host "Please enter 1-$optionsCount." -ForegroundColor $script:warningColor
                $selection = Read-Host
            }
        }
        while (-not($validSelection))

        return [int]$selection
    }

    [void] AssignUserRole($roleSelection)
    {
        # Need to refresh the user's info here because the method of assigning a new role depends on their current one.
        $this.gotoUser = Invoke-GetWithRetry { $this.GetUser() }
        
        $method = "Put"
        $uri = ""
        $headers = @{
            "Authorization" = "Bearer $($this.accessToken)"
            "Content-Type"  = "application/json"
        }
        $body = @{}
        $newRoleName = ""

        switch ($roleSelection)
        {
            1 # roleSelection is Super Admin
            {
                $uri = "https://api.getgo.com/admin/rest/v1/accounts/$($this.accountKey)/users/$($this.gotoUser.key)"
                $body = @{
                    "adminRoles" = @("SUPER_USER")
                }
                $newRoleName = "Super Admin"
                break
            }
            2 # roleSelection is Admin (Configure PBX)
            {
                $uri = "https://api.getgo.com/admin/rest/v1/accounts/$($this.accountKey)/users/$($this.gotoUser.key)"
                $body = @{
                    "adminRoles" = @("MANAGE_ACCOUNT")
                }
                $newRoleName = "Admin (Configure PBX)"
                break
            }
            3 # roleSelection is member
            {
                if ($this.gotoUser.adminRoles) # user has a built-in (system) admin role that must be removed
                {
                    $uri = "https://api.getgo.com/admin/rest/v1/accounts/$($this.accountKey)/users/$($this.gotoUser.key)"
                    $body = @{ 
                        "adminRoles" = @() # set admin role to blank array
                    }
                }
                elseif ($this.gotoUser.roleSets) # user has a custom role that must be removed
                {
                    $method = "Delete"
                    $uri = "https://api.getgo.com/admin/rest/v1/accounts/$($this.accountKey)/users/rolesets"
                    $emailQuery = "email eq `"$($this.gotoUser.email)`""
                    $emailQuery = UriEncode-QueryParam $emailQuery
                    $uri = $uri + "?filter=$emailQuery"
                }
                else
                {
                    Write-Host "User is already a member!" -ForegroundColor $script:successColor
                    return
                }
                $newRoleName = "Member"
                break
            }
            { $_ -ge 4 } # roleSelection is >= 4 and therefore a custom role
            {
                $roleId = @($this.allCustomRoles.keys)[$roleSelection - 4] # enumerates the keys and accesses them by index
                $uri = "https://api.getgo.com/admin/rest/v1/accounts/$($this.accountKey)/rolesets/$roleId/users"
                $emailQuery = "email eq `"$($this.gotoUser.email)`""
                $emailQuery = UriEncode-QueryParam $emailQuery
                $uri = $uri + "?filter=$emailQuery"
                $newRoleName = $this.allCustomRoles[$roleId]
                break
            }
        }

        $response = SafelyInvoke-RestMethod -Method $method -Uri $uri -Headers $headers -Body ($body | ConvertTo-Json)
        if ($response)
        { 
            Write-Host "Assigned role: $newRoleName`n" -ForegroundColor $script:successColor
            $script:logger.LogChange("Assigned GoTo role: $newRoleName")
            $this.assignRoleCompleted = $true
        }
        else
        {
            $script:logger.LogWarning("There was an issue assigning GoTo role: $newRoleName")
        }
    }

    [void] DisplayLinkToOutboundCallerId()
    {
        $url = "https://my.jive.com/pbx/brs/extensions/lines/$(($this.GetLine()).id)/general?source=root.nav.pbx.extensions.lines.list"
        Write-Host "At this time the GoTo API can't change the outbound caller ID. You'll need to change it here:`n$url" -ForegroundColor $script:infoColor
        $shouldLaunch = Prompt-YesOrNo "Launch browser to this link?"
        if ($shouldLaunch)
        {
            try 
            {                
                Start-Process $url # When passed a URL Start-Process will launch system default browser.
            }
            catch
            {
                $errorRecord = $_
                Write-Host "There was an issue launching the url." -ForegroundColor $script:warningColor
                Write-Host $errorRecord.Exception.Message -ForegroundColor $script:warningColor
            }      
        }
        $this.assignCallerIdCompleted = $true
    }

    [object] GetLine()
    {
        $method = "Get"
        $uri = "https://api.goto.com/users/v1/users/$($this.gotoUser.key)/lines"
        $headers = @{
            "Authorization" = "Bearer $($this.accessToken)"
        }

        $response = SafelyInvoke-RestMethod -Method $method -Uri $uri -Headers $headers
        return $response.items
    }
}

# main
Initialize-ColorScheme
Show-Introduction
Use-Module "Microsoft.Graph.Users"
Use-Module "ExchangeOnlineManagement"
Use-Module "Microsoft.Powershell.SecretManagement"
Use-Module "PSAuthClient" # Docs for this module found here https://github.com/alflokken/PSAuthClient
TryConnect-MgGraph -Scopes @("User.ReadWrite.All", "Group.ReadWrite.All", "Organization.Read.All")
TryConnect-ExchangeOnline
Set-Variable -Name "logger" -Value ([Logger]::New()) -Scope "Script" -Option "Constant"

$keepGoing = $true
do
{
    $upn = Prompt-BrsEmail "Enter user UPN"
    $isValidEmail = Confirm-ValidBrsEmail $upn
    if (-not($isValidEmail)) { continue }

    $user = Get-M365User -UPN $upn -Detailed -WarningAction "SilentlyContinue"
    if ($null -eq $user)
    {
        Write-Host "User does not exist yet." -ForegroundColor $infoColor
        $createUser = Prompt-YesOrNo "Create user with this UPN?: $upn"
        if ($createUser)
        {
            $user = Start-M365UserCreationWizard $upn
            if ($null -eq $user) { continue }
            # Get more details about the user.
            $user = Invoke-GetWithRetry { Get-M365User -UPN $user.UserPrincipalName -Detailed }
            if ($user)
            { 
                $userProps = Get-UserProperties -User $user -RetryGetManager
                $keepGoing = $false 
            }            
        }
    }
    else
    {
        $userProps = Get-UserProperties $user
        $keepGoing = $false
    }
}
while ($keepGoing)

Show-UserProperties -BasicProps $userProps.basicProps -Licenses $userProps.Licenses -Groups $userProps.Groups -AdminRoles $userProps.AdminRoles

$script:grantLicensesCompleted = $false
$script:assignGroupsCompleted = $false
$script:grantMailboxesCompleted = $false
$script:gotoSetupCompleted = $false

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
            if ($null -eq $gotoWizard)
            {
                $gotoWizard = [GotoWizard]::New($upn) 
            }

            if ($gotoWizard) { $gotoWizard.Start() }            
            break
        }
        6 # Finish
        {
           $keepGoing = $false
           break
        }
    }
}

$userProps = Get-UserProperties $user
Show-UserProperties -BasicProps $userProps.basicProps -Licenses $userProps.Licenses -Groups $userProps.Groups -AdminRoles $userProps.AdminRoles
$script:logger.ShowLogs()

Read-Host "Press Enter to exit"