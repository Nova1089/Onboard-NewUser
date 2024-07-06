<#
Version 1.0

This script onboards or modifies a new user in M365 and GoTo.
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
    Write-Host "This script onboards or modifies a new user in M365 and GoTo." -ForegroundColor $infoColor
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

function Test-ValidBrsEmail($email)
{
    $isBrsEmail = $email -imatch '^\S+@blueravensolar.com$'
    if (-not($isBrsEmail))
    {
        Write-Host "Email can't have spaces and needs to end in @blueravensolar.com." -ForegroundColor $warningColor
        return $false
    }

    $isStandard = $email -imatch '^[\w-]+\.[\w-]+(@blueravensolar\.com)$'    
    if (-not($isStandard))
    {
        Write-Host "Email is not standard (First.Last@blueravensolar.com)" -ForegroundColor $warningColor
        $continue = Prompt-YesOrNo "Are you sure you want to use this email?"
        if (-not($continue)) { return $false }
    }

    return $true
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
            Write-Warning "User not found."
            return
        }
        Write-Host "There was an issue getting the user." -ForegroundColor $warningColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $warningColor
    }
    return $user
}

function Start-M365UserWizard($upn)
{
    $emailParts = Get-EmailParts $upn
    $correct = Prompt-YesOrNo "This preferred first name?: $($emailParts.FirstName)"
    if (-not($correct)) { $emailParts.FirstName = Read-Host "Enter preferred first name" }

    $correct = Prompt-YesOrNo "This last name?: $($emailParts.LastName)"
    if (-not($correct)) { $emailParts.LastName = Read-Host "Enter last name" }

    $correct = Prompt-YesOrNo "This display name?: $($emailParts.DisplayName)"
    if (-not($correct)) { $emailParts.DisplayName = Read-Host "Enter display name" }

    $jobTitle = Read-Host "Enter job title"
    $department = Read-Host "Enter department"
    do
    {
        $managerUpn = Prompt-BrsEmail "Enter manager UPN"
        $manager = Get-M365User -UPN $managerUpn
    }
    while ($null -eq $manager)

    $params = @{
        "UPN"           = $upn
        "MailNickName"  = $emailParts.MailNickName
        "DisplayName"   = $emailParts.DisplayName
        "FirstName"     = $emailParts.FirstName
        "LastName"      = $emailParts.LastName
        "JobTitle"      = $jobTitle
        "Department"    = $department
        "UsageLocation" = "US" # We always set this to US, even for those out-of-country.
    }

    $user = New-M365User @params
    Set-UserManager -User $user -Manager $manager
    return $user
}

function New-M365User($upn, $mailNickName, $displayName, $firstName, $lastName, $jobTitle, $department, $usageLocation)
{
    $tempPassword = Get-TempPassword
    Write-Host "Temp Password: $tempPassword" -ForegroundColor $infoColor
    $passwordProfile = @{
        "password"                             = Get-TempPassword
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
        $user = New-Mguser @params
    }
    catch
    {
        $errorRecord = $_
        Write-Host "There was an issue creating user." -ForegroundColor $warningColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $warningColor
    }
    return $user
}

function Get-TempPassword
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

function Get-EmailParts($email)
{    
    $email = $email.Trim()
    # Get the part of the email before the @ sign.
    $mailNickName = $email.Split('@')[0]
    $nameParts = $mailNickName.Split('.')
    $displayName = ""
    foreach ($name in $nameParts)
    {
        $displayName += (Capitalize-FirstLetter $name) + " "
    }
    
    return @{
        "FirstName"    = Capitalize-FirstLetter $nameParts[0]
        "LastName"     = Capitalize-FirstLetter $nameParts[1]
        "DisplayName"  = $displayName
        "MailNickName" = $mailNickName
    }
}

function Capitalize-FirstLetter($string)
{
    return $string.substring(0, 1).ToUpper() + $string.substring(1)
}

function Set-UserManager($user, $manager)
{
    try
    {
        Set-MgUserManagerByRef -UserId $user.Id -OdataId "https://graph.microsoft.com/v1.0/users/$($manager.Id)" -ErrorAction "Stop" | Out-Null
    }
    catch
    {
        $errorRecord = $_
        Write-Host "There was an issue assigning manager." -ForegroundColor $warningColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $warningColor
    }
    
}

function Get-UserProperties($user)
{
    $manager = Invoke-GetWithRetry { Get-UserManager -User $user }
    $basicProps = [PSCustomObject]@{
        "Created Date/Time (UTC)" = $user.CreatedDateTime
        "Display Name"            = $user.DisplayName
        "UPN"                     = $user.UserPrincipalName
        "Title"                   = $user.JobTitle
        "Department"              = $user.Department
        "Manager"                 = $manager.displayName
        "Usage Location"          = $user.UsageLocation
    }

    return @{
        "BasicProps" = $basicProps
        "Licenses"   = Get-UserLicenses $user
        "Groups"     = Get-UserGroups $user
        "AdminRoles" = Get-UserAdminRoles $user
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
                Write-Warning "$scriptBlock returned null. Retrying in $delay seconds..."
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

    if ($retryCount -ge $maxRetries) { Write-Warning "Timed out trying to get a response." }

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
    # Returns a dictionary.
    return $manager.AdditionalProperties
}

function Get-UserLicenses($user)
{
    if ($null -eq $script:licenseLookupTable)
    {
        $script:licenseLookupTable = @{
            "8f0c5670-4e56-4892-b06d-91c085d7004f" = "App Connect IW"
            "4b9405b0-7788-4568-add1-99614e613b69" = "Exchange Online (Plan 1)"
            "19ec0d23-8335-4cbd-94ac-6050e30712fa" = "Exchange Online (Plan 2)"
            "3b555118-da6a-4418-894f-7df1e2096870" = "Microsoft 365 Business Basic"
            "cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46" = "Microsoft 365 Business Premium"
            "f245ecc8-75af-4f8e-b61f-27d8114de5f3" = "Microsoft 365 Business Standard"
            "05e9a617-0261-4cee-bb44-138d3ef5d965" = "Microsoft 365 E3"
            "06ebc4ee-1bb5-47dd-8120-11324bc54e06" = "Microsoft 365 E5"
            "4ef96642-f096-40de-a3e9-d83fb2f90211" = "Microsoft Defender for Office 365 (Plan 1)"
            "3dd6cf57-d688-4eed-ba52-9e40b5468c3e" = "Microsoft Defender for Office 365 (Plan 2)"
            "a403ebcc-fae0-4ca2-8c8c-7a907fd6c235" = "Microsoft Fabric (Free)"
            "dcb1a3ae-b33f-4487-846a-a640262fadf4" = "Microsoft Power Apps Plan 2 Trial"
            "f30db892-07e9-47e9-837c-80727f46fd3d" = "Microsoft Power Automate Free"
            "5b631642-bd26-49fe-bd20-1daaa972ef80" = "Microsoft PowerApps for Developer"
            "1f2f344a-700d-42c9-9427-5cea1d5d7ba6" = "Microsoft Stream"
            "3ab6abff-666f-4424-bfb7-f0bc274ec7bc" = "Microsoft Teams Essentials"
            "4cde982a-ede4-4409-9ae6-b003453c8ea6" = "Microsoft Teams Rooms Pro"
            "18181a46-0d4e-45cd-891e-60aabd171b4e" = "Office 365 E1"
            "6fd2c87f-b296-42f0-b197-1e91e994b900" = "Office 365 E3"
            "c7df2760-2c81-4ef7-b578-5b5392b571df" = "Office 365 E5"
            "7b26f5ab-a763-4c00-a1ac-f6c4b5506945" = "Power BI Premium P1"
            "6470687e-a428-4b7a-bef2-8a291ad947c9" = "Windows Store for Business"
        }
    }

    try
    {
        $licenseDetails = Get-MGUserLicenseDetail -UserId $user.UserPrincipalName
    }
    catch
    {
        $errorRecord = $_
        Write-Host "There was an issue getting user's licenses." -ForegroundColor $warningColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $warningColor
        return
    }
    
    $licenses = New-Object System.Collections.Generic.List[object]
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
    $basicProps | Out-Host

    Show-Separator "Licenses"
    $licenses | Select-Object -ExpandProperty "Name" | Sort-Object | Out-Host

    Show-Separator "Groups"
    $groups | Select-Object -ExpandProperty "DisplayName" | Sort-Object | Out-Host
    
    Show-Separator "Admin Roles"
    $adminRoles | Select-Object -ExpandProperty "DisplayName" | Sort-Object | Out-Host
    
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

function Prompt-MainMenu
{
    $selection = Read-Host ("`nWhat next?`n" +
        "[1] Show M365 User Info`n" +
        "[2] $(New-Checkbox($script:grantLicensesCompleted)) Manage licenses`n" +
        "[3] $(New-Checkbox($script:assignGroupsCompleted)) Manage groups`n" +
        "[4] $(New-Checkbox($script:grantMailboxesCompleted)) Manage shared mailboxes`n" +
        "[5] $(New-Checkbox($script:gotoSetupCompleted)) Setup GoTo account`n" +
        "[6] Finish`n")

    do
    {
        $isValidSelection = $selection -imatch '^\s*[1-6]\s*$' # regex matches 1-6 but allows spaces
        if (-not($isValidSelection))
        {
            Write-Host "Please enter a number 1-6." -ForegroundColor $warningColor
            $selection = Read-Host
        }
    }
    while (-not($isValidSelection))

    return $selection.Trim()
}

function Prompt-BrsEmail($message)
{
    do
    {
        $email = Read-Host "`n$message (you may omit the @blueravensolar.com)"
    }
    while ($null -eq $email)

    $email = $email.Trim()
    $hasDomain = $email -imatch '^\S+@blueravensolar.com$'

    if (-not($hasDomain))
    {
        $email += '@blueravensolar.com'
    }

    return $email
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
                $licenses | Select-Object -ExpandProperty "Name" | Sort-Object | Out-Host
                break
            }
            2 # Grant license
            {
                $availableLicenses = Get-AvailableLicenses
                $license = Prompt-LicenseToGrant $availableLicenses
                Grant-License -User $user -License $license
                $script:grantLicensesCompleted = $true
                break
            }
            3 # Revoke license
            {
                $assignedLicenses = Get-UserLicenses $user
                $license = Prompt-LicenseToRevoke $assignedLicenses
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
    do
    {
        $response = Read-Host ("`nChoose an option:`n" +
            "[1] View assigned licenses`n" +                        
            "[2] Grant license`n" +
            "[3] Revoke license`n" +
            "[4] Finish with licenses`n")
        
        $validResponse = $response -imatch '^\s*[1-4]\s*$' # regex matches 1-4 but allows spaces
        if (-not($validResponse))
        {
            Write-Host "Please enter 1-4." -ForegroundColor $warningColor
        }
    }
    while (-not($validResponse))

    return [int]$response
}

function Get-AvailableLicenses
{
    $uri = "https://graph.microsoft.com/v1.0/subscribedSkus"
    $licenses = Invoke-MgGraphRequest -Method "Get" -Uri $uri

    $licenseTable = New-Object System.Collections.Generic.List[object]
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
    $option = 0
    $availableLicenses | Sort-Object -Property "Name" | ForEach-Object { $option++; $_ | Add-Member -NotePropertyName "Option" -NotePropertyValue $option }
    $availableLicenses | Sort-Object -Property "Option" | Format-Table -Property @("Option", "Name", "Available", "Purchased") | Out-Host

    do
    {
        $response = (Read-Host "Select an option (1 - $option)") -As [int]
        # check that response is a number between 1 and option count (avoids use of regex because that's not great for matching multi-digit number ranges)
        $validResponse = ($response -is [int]) -and (($response -ge 1) -and ($response -le $option))
        if (-not($validResponse)) 
        {
            Write-Host "Please enter a number 1-$option." -ForegroundColor $warningColor
        }
    }
    while (-not($validResponse))

    foreach ($license in $availableLicenses)
    {
        if ($license.option -eq [int]$response)
        {
            return $license
        }
    }
    return $null
}

function Grant-License($user, $license)
{
    try
    {
        Set-MgUserLicense -UserId $user.Id -AddLicenses @{SkuId = $license.SkuId } -RemoveLicenses @() -ErrorAction "Stop" | Out-Null
        Write-Host "License granted: $($license.name)" -ForegroundColor $successColor
    }
    catch
    {
        $errorRecord = $_
        Write-Host "There was an issue granting the license." -ForegroundColor $warningColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $warningColor
    }  
}

function Prompt-LicenseToRevoke($assignedLicenses)
{    
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

    do
    {
        $response = (Read-Host "Select an option (1 - $option)") -As [int]
        # check that response is a number between 1 and option count (avoids use of regex because that's not great for matching multi-digit number ranges)
        $validResponse = ($response -is [int]) -and (($response -ge 1) -and ($response -le $option))
        if (-not($validResponse)) 
        {
            Write-Host "Please enter a number 1-$option." -ForegroundColor $warningColor
        }
    }
    while (-not($validResponse))

    foreach ($license in $optionList)
    {
        if ($license.option -eq [int]$response)
        {
            return $license
        }
    }
    return $null
}

function Revoke-License($user, $license)
{
    try
    {
        Set-MgUserLicense -UserId $user.Id -AddLicenses @() -RemoveLicenses @($license.SkuId) -ErrorAction "Stop" | Out-Null
        Write-Host "License revoked: $($license.name)" -ForegroundColor $successColor
    }
    catch
    {
        $errorRecord = $_
        Write-Host "There was an issue revoking the license." -ForegroundColor $warningColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $warningColor
    } 
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
            Write-Warning "Please enter y or n."
        }
    }
    while (-not($validResponse))

    if ($response -imatch '^\s*y\s*$') # regex matches a y but allows spaces
    {
        return $true
    }
    return $false
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
                Get-UserGroups -User $user | Select-Object -ExpandProperty "DisplayName" | Sort-Object | Out-Host
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

                    $isAlreadyMember = Test-IsMemberOfGroup -User $user -Group $group
                    if ($isAlreadyMember)
                    {
                        Write-Host "$($user.DisplayName) is already a member of the group: $($group.DisplayName)." -ForegroundColor $warningColor
                        $tryAgain = Prompt-YesOrNo "Try again?"
                        if (-not($tryAgain)) { break outerSwitch }
                        continue
                    }
                }
                while (($null -eq $group) -or ($isAlreadyMember))

                Assign-M365Group -User $user -Group $group
                $script:assignGroupsCompleted = $true
                break      
            }
            3 # Remove group
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

                    $isAlreadyMember = Test-IsMemberOfGroup -User $user -Group $group
                    if (-not($isAlreadyMember))
                    {
                        Write-Host "$($user.DisplayName) is not a member of the group: $($group.DisplayName)." -ForegroundColor $warningColor
                        $tryAgain = Prompt-YesOrNo "Try again?"
                        if (-not($tryAgain)) { break outerSwitch }
                        continue
                    }
                }
                while (($null -eq $group) -or (-not($isAlreadyMember)))

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
    do
    {
        $response = Read-Host ("`nChoose an option:`n" +
            "[1] View assigned groups`n" +                        
            "[2] Assign group`n" +
            "[3] Remove group`n" +
            "[4] Finish with groups`n")
        
        $validResponse = $response -imatch '^\s*[1-4]\s*$' # regex matches 1-4 but allows spaces
        if (-not($validResponse))
        {
            Write-Host "Please enter 1-4." -ForegroundColor $warningColor
        }
    }
    while (-not($validResponse))

    return [int]$response
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
        if ($errorRecord.Exception.Message -ilike "*[Request_ResourceNotFound]*")
        {
            Write-Host "Group not found." -ForegroundColor $warningColor
            return
        }        
        Write-Host "There was an issue getting the group." -ForegroundColor $warningColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $warningColor
    }

    if ($group)
    {
        Write-Host "Found Group!" -ForegroundColor $successColor
        $group | Select-Object -Property @("DisplayName", "Mail", "Description") | Out-Host        
    }

    return $group
}

function Test-IsMemberOfGroup($user, $group)
{
    $currentAssignedGroups = Get-UserGroups -User $user
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
        Write-Host "Group assigned: $($group.Mail)" -ForegroundColor $successColor
    }
    catch
    {
        $errorRecord = $_
        Write-Host "There was an issue assigning the group." -ForegroundColor $warningColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $warningColor
    }    
}

function Unassign-M365Group($user, $group)
{
    try
    {
        Remove-MgGroupMemberByRef -GroupId $group.Id -DirectoryObjectId $user.Id -ErrorAction "Stop" | Out-Null
        Write-Host "Group removed: $($group.Mail)" -ForegroundColor $successColor
    }
    catch
    {
        $errorRecord = $_
        Write-Host "There was an issue unassigning the group." -ForegroundColor $warningColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $warningColor
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
                    "For that you may run this script instead, but it takes a little while...`n" +
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

                Grant-MailboxAccess -User $user -Mailbox $mailbox
                $script:grantMailboxesCompleted = $true
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
    do
    {
        $response = Read-Host ("`nChoose an option:`n" +
            "[1] View assigned mailboxes`n" +                        
            "[2] Grant access to mailbox`n" +
            "[3] Revoke access to mailbox`n" +
            "[4] Finish with mailboxes`n")
        
        $validResponse = $response -imatch '^\s*[1-4]\s*$' # regex matches 1-4 but allows spaces
        if (-not($validResponse))
        {
            Write-Host "Please enter 1-4." -ForegroundColor $warningColor
        }
    }
    while (-not($validResponse))

    return [int]$response
}

function Get-SharedMailbox($upn)
{
    if ($null -eq $upn) { throw "Can't get shared mailbox. UPN is null." }

    $mailbox = Get-EXOMailbox -Identity $upn -ErrorAction "SilentlyContinue"
    if (($mailbox) -and ($mailbox.RecipientTypeDetails -eq 'SharedMailbox'))
    {
        Write-Host "Mailbox found!" -ForegroundColor $successColor
        $mailbox | Select-Object -Property @("DisplayName", "UserPrincipalName", @{ label = "Type"; expression = { $_.RecipientTypeDetails } }) | Out-Host
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
    Write-Host "Access type to grant?`n" 
    $accessType = Prompt-MailboxAccessType
    try
    {
        switch ($accessType)
        {
            1 # Read and Manage
            {
                Add-MailboxPermission -Identity $mailbox.UserPrincipalName -User $user.UserPrincipalName -AccessRights "FullAccess" -Confirm:$false -WarningAction "SilentlyContinue" -ErrorAction "Stop" | Out-Null
                break
            }
            2 # Send As
            {
                Add-RecipientPermission -Identity $mailbox.UserPrincipalName -Trustee $user.UserPrincipalName -AccessRights "SendAs" -Confirm:$false -WarningAction "SilentlyContinue" -ErrorAction "Stop" | Out-Null
                break
            }
            3 # Both
            {
                Add-MailboxPermission -Identity $mailbox.UserPrincipalName -User $user.UserPrincipalName -AccessRights "FullAccess" -Confirm:$false -WarningAction "SilentlyContinue" -ErrorAction "Stop" | Out-Null
                Add-RecipientPermission -Identity $mailbox.UserPrincipalName -Trustee $user.UserPrincipalName -AccessRights "SendAs" -Confirm:$false -WarningAction "SilentlyContinue" -ErrorAction "Stop" | Out-Null
                break
            }
        }
        Write-Host "Granted access to mailbox! (If they didn't already have the access.)" -ForegroundColor $successColor
    }
    catch
    {
        $errorRecord = $_
        Write-Host "There was an issue granting mailbox access. Please try again." -ForegroundColor $warningColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $warningColor
    }    
}

function Revoke-MailboxAccess($user, $mailbox)
{
    Write-Host "Access type to revoke?`n"
    $accessType = Prompt-MailboxAccessType
    try
    {
        switch ($accessType)
        {
            1 # Read and Manage
            {
                Remove-MailboxPermission -Identity $mailbox.UserPrincipalName -User $user.UserPrincipalName -AccessRights "FullAccess" -Confirm:$false -WarningAction "SilentlyContinue" -ErrorAction "Stop"
                break
            }
            2 # Send As
            {
                Remove-RecipientPermission -Identity $mailbox.UserPrincipalName -Trustee $user.UserPrincipalName -AccessRights "SendAs" -Confirm:$false -WarningAction "SilentlyContinue" -ErrorAction "Stop"
                break
            }
            3 # Both
            {
                Remove-MailboxPermission -Identity $mailbox.UserPrincipalName -User $user.UserPrincipalName -AccessRights "FullAccess" -Confirm:$false -WarningAction "SilentlyContinue" -ErrorAction "Stop"
                Remove-RecipientPermission -Identity $mailbox.UserPrincipalName -Trustee $user.UserPrincipalName -AccessRights "SendAs" -Confirm:$false -WarningAction "SilentlyContinue" -ErrorAction "Stop"
                break
            }
        }
        Write-Host "Revoked access to mailbox! (Assuming they had access.)" -ForegroundColor $successColor
    }
    catch
    {
        $errorRecord = $_
        Write-Host "There was an issue revoking mailbox access. Please try again." -ForegroundColor $warningColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $warningColor
    }
}

function Prompt-MailboxAccessType
{
    do
    {
        $accessType = Read-Host  ("[1] Read & Manage`n" +
            "[2] Send As`n" +
            "[3] Both`n")
        $accessType = $accessType.Trim()
        $isValidResponse = $accessType -imatch '^[1-3]$' # regex matches 1-3

        if (-not($isValidResponse))
        {
            Write-Host "Please enter a number 1-3." -ForegroundColor $warningColor
        }
    }
    while (-not($isValidResponse))

    return [int]$accessType
}

function UriEncode-QueryParam($queryParam)
{ 
    return [uri]::EscapeDataString($queryParam)
}

function SafelyInvoke-RestMethod($method, $uri, $headers, $body)
{
    try
    {
        $response = Invoke-RestMethod -Method $method -Uri $uri -Headers $headers -Body $body -ErrorVariable "responseError"
    }
    catch
    {
        Write-Host $responseError[0].Message -ForegroundColor $warningColor
        return $false # Wasn't successful.
    }

    if ($response)
    {
        return $response
    }
    return $true # Was successful.
}

function New-Checkbox($checked)
{
    if ($checked)
    {
        return "[X]"
    }
    return "[ ]"
}

class Logger
{
    # Implementing logger as a singleton. Meaning there should only be one instance in the program!

    # Can't make the constructor private, so we'll just hide it.
    hidden Logger()
    { 
        $this.logs = New-Object System.Collections.Generic.List[object]
    }

    # Can't make this field private or readonly, but we can hide it.
    hidden static [Logger] $instance

    # This is how you're supposed to instantiate the logger.
    static [Logger] GetInstance()
    {
        if ($null -eq [Logger]::instance)
        {
            [Logger]::instance = [Logger]::new()
        }
        return [Logger]::instance
    }    
    
    # fields
    hidden [System.Collections.Generic.List[object]] $logs

    # methods
    [void] LogChange($message)
    {
        $logEntry = [PSCustomObject]@{
            Timestamp = Get-Date
            Level     = 'Change'
            Message   = $message
        }
        $this.logs.Add($logEntry)
    }

    [void] LogWarning($message)
    {
        $logEntry = [PSCustomObject]@{
            Timestamp = Get-Date
            Level     = 'Warning'
            Message   = $message
        }
        $this.logs.Add($logEntry)
    }

    [void] LogError($message)
    {
        $logEntry = [PSCustomObject]@{
            Timestamp = Get-Date
            Level     = 'Error'
            Message   = $message
        }
        $this.logs.Add($logEntry)
    }

    [void] ShowLogs()
    {
        $this.Logs = $this.Logs | Sort-Object TimeStamp
        foreach ($log in $this.Logs)
        {
            # Pipe to Get-Date for a simplified timestamp.
            $message = "[$($log.Timestamp | Get-Date -Format 'yyyy-mm-dd a\t hh:mm tt')] $($log.Message)"
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

# classes
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
    [bool] $roleStepCompleted
    [bool] $cidStepCompleted

    # constructors
    GotoWizard($upn)
    {
        $this.upn = $upn
        $this.gotoSecret = $this.GetApiSecret()
        $this.clientId = $this.gotoSecret.ClientID
        $this.clientSecret = $this.gotoSecret.ClientSecret
        $this.accountKey = $this.gotoSecret.AccountKey
        $this.accessToken = $this.GetAccessToken()
        $this.gotoUser = $this.GetUser()
        $this.allCustomRoles = $this.GetAllCustomRoles()
    }

    # Default constructor defined for Pester testing purposes.
    GotoWizard() {}

    # methods
    [void] Start()
    {
        if ($null -eq $this.gotoUser)
        {
            Write-Host "Goto user not found." -ForegroundColor $script:warningColor
            return
        }
        Write-Host "Found GoTo user!`n" -ForegroundColor $script:successColor

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
                    $script:gotoSetupCompleted = $true
                    $keepGoing = $false
                    break
                }
            }
        }
    }

    [object] GetApiSecret()
    {
        return Get-Secret -Name "YZJrirO-73fEk6aZO5QgZg" -AsPlainText
    }

    [string] GetAccessToken()
    {
        # Function handles GoTo connect OAUTH2 authorization code grant flow. (Obtains auth code then uses that to get a temp access token.)
        # https://developer.goto.com/guides/Authentication/03_HOW_accessToken/
        # https://developer.goto.com/Authentication/#section/Authorization-Flows

        $authUri = "https://authentication.logmeininc.com/oauth/authorize"
        $accessTokenUri = "https://authentication.logmeininc.com/oauth/token"
        $redirectUri = "https://localhost"
    
        $authCode = Invoke-OAuth2AuthorizationEndpoint -Uri $authUri -Client_id $this.clientId -Redirect_uri $redirectUri
        $token = Invoke-OAuth2TokenEndpoint @authCode -Uri $accessTokenUri -Client_secret $this.clientSecret -Client_auth_method "client_secret_basic"
        return $token.access_token
    }

    [int] PromptMenu()
    {
        $selection = Read-Host ("Choose an option:`n" +
            "[1] Show GoTo user info`n" +
            "[2] $(New-Checkbox $this.roleStepCompleted) Assign role`n" +
            "[3] $(New-Checkbox $this.cidStepCompleted) Assign outbound caller ID`n" +
            "[4] Finish GoTo setup`n")

        while ($selection -notmatch '^\s*[1-4]\s*$') # regex matches 1-4 but allows spaces
        {
            Write-Host "Please enter 1-4." -ForegroundColor $script:warningColor
            $selection = Read-Host
        }

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
        # That other way will apply URI encoding to the query params, but encode spaces with + signs instead of %20. 
        # The GoTo API doesn't accept this.
        $emailQuery = UriEncode-QueryParam $emailQuery
        $uri = $uri + "?filter=$emailQuery"

        $response = SafelyInvoke-RestMethod -Method $method -Uri $uri -Headers $headers
        return $response.results[0]
    }

    [void] ShowUserInfo()
    {
        $role = $this.GetUserRole()
        # I'd also show their external caller ID, but this is not accessible from the public API.
        $this.gotoUser | Add-Member -NotePropertyName "role" -NotePropertyValue $role -PassThru | Format-List -Property @("email", "firstName", "lastName", "role") | Out-Host
    }

    [string] GetUserRole()
    {
        $this.gotoUser = $this.GetUser()
        
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
        $menuText = ("Select a role to assign:`n" +
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

        while ($selection -notmatch "^\s*[1-$optionsCount]\s*$") # regex matches 1 - optionsCount but allows spaces
        {
            Write-Host "Please enter 1-$optionsCount." -ForegroundColor $script:warningColor
            $selection = Read-Host
        }

        return [int]$selection.Trim()
    }

    [void] AssignUserRole($roleSelection)
    {
        # Need to refresh the user's info here because the method of assigning a new role depends on their current one.
        $this.gotoUser = $this.GetUser()
        
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
        if ($response) { Write-Host "Assigned role: $newRoleName`n" -ForegroundColor $script:successColor }        
    }

    [void] DisplayLinkToOutboundCallerId()
    {
        $line = $this.GetLine()
        Write-Host "At this time the GoTo API can't change the outbound caller ID. You'll need to change it here:" -ForegroundColor $script:infoColor
        Write-Host "https://my.jive.com/pbx/brs/extensions/lines/$($line.id)/general?source=root.nav.pbx.extensions.lines.list" -ForegroundColor $script:infoColor
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

do
{
    $upn = Prompt-BrsEmail "Enter user UPN"
    $isValidEmail = Test-ValidBrsEmail $upn
    if (-not($isValidEmail)) { continue }
    $user = Get-M365User -UPN $upn -Detailed -WarningAction "SilentlyContinue"

    if ($null -eq $user)
    {
        Write-Host "User does not exist yet. Create them with this UPN?: $upn" -ForegroundColor $infoColor
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

# Initialize logger. Singleton that should have one unchanging instance.
Set-Variable -Name "logger" -Value ([Logger]::GetInstance()) -Scope "Script" -Option "Constant"

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