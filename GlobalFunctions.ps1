#  functions
function Initialize-ColorScheme
{
    Set-Variable -Name "successColor" -Value "Green" -Scope "Script"
    Set-Variable -Name "infoColor" -Value "DarkCyan" -Scope "Script"
    Set-Variable -Name "warningColor" -Value "Yellow" -Scope "Script"
    Set-Variable -Name "failColor" -Value "Red" -Scope "Script"
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
    $upn = Read-Host "Enter the UPN for the user (PreferredFirst.Last@blueravensolar.com)"
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
    $isBrsEmail = $email -match '^\s*\S+@blueravensolar.com\s*$'

    if (-not($isBrsEmail))
    {
        Write-Warning "Email needs to end in @blueravensolar.com"
        return $false
    }

    $isStandard = $email -imatch '^\s*[\w-]+\.[\w-]+(@blueravensolar\.com)\s*$'    
    
    if (-not($isStandard))
    {
        Write-Warning "Email is not standard (PreferredFirstName.LastName@blueravensolar.com)"
        $continue = Prompt-YesOrNo "Are you sure you want to use this email?"
        if (-not($continue)) { return $false }
    }

    return $true
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

    $licenseDetails = Get-MGUserLicenseDetail -UserId $user.UserPrincipalName
    $licenses = New-Object System.Collections.Generic.List[object]

    foreach ($license in $licenseDetails)
    {
        $licenseName = $script:licenseLookupTable[$license.SkuId]
        $licenses.Add( @{"Name" = $licenseName; "SkuId" = $license.SkuId} )
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

function Show-UserProperties($user, $manager, $licenses, $groups, $adminRoles)
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
    $licenses | Select-Object -ExpandProperty "Name" | Sort-Object | Out-Host

    Show-Separator "Groups"
    $groups | Sort-Object | Out-Host
    
    Show-Separator "Admin Roles"
    $adminRoles | Sort-Object | Out-Host
    
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
    "[2] $(New-Checkbox($script:grantLicensesCompleted)) Grant licenses`n" +
    "[3] $(New-Checkbox($script:assignGroupsCompleted)) Assign groups`n" +
    "[4] $(New-Checkbox($script:mailboxStepCompleted)) Grant shared mailboxes`n" +
    "[5] $(New-Checkbox($script:gotoStepCompleted)) Setup GoTo Account`n" +
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
    while(-not($isValidSelection))

    return $selection.Trim()
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
                Write-Host "Current assigned licenses:" -ForegroundColor $infoColor
                Get-UserLicenses $user | Select-Object -ExpandProperty "Name" | Sort-Object | Out-Host
            }
            2 # Grant license
            {
                $availableLicenses = Get-AvailableLicenses
                $license = Prompt-LicenseToGrant $availableLicenses
                Grant-License -User $user -License $license
                $script:grantLicensesCompleted = $true
            }
            3 # Revoke license
            {
                $assignedLicenses = Get-UserLicenses -User $user
                $license = Prompt-LicenseToRevoke $assignedLicenses
                Revoke-License -User $user -License $license
            }
            4 # Finish
            {
                $keepGoing = $false
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
            "Name" = $name
            "Available" = $amountAvailable
            "Purchased" = $amountPurchased
            "SkuId" = $license.skuId        
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
        Write-Host "There was an issue granting the license." -ForegroundColor $failColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $failColor
    }  
}

function Prompt-LicenseToRevoke($assignedLicenses)
{    
    $option = 0
    $optionList = $assignedLicenses | Sort-Object "Name" | ForEach-Object { 
        $option++
        [PSCustomObject]@{
            "Option" = $option
            "Name" = $_.Name
            "SkuId" = $_.SkuId
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
        Write-Host "There was an issue revoking the license." -ForegroundColor $failColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $failColor
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

Start-M365GroupWizard
{    
    $keepGoing = $true
    while ($keepGoing)
    {
        $selection = Prompt-GroupMenu

        switch ($selection)
        {
            1 # View assigned groups
            {

            }
            2 # Assign group
            {
                do
                {
                    $groupEmail = Prompt-BRSEmail -EmailType "group"
                    $group = Get-M365Group $groupEmail
                }
                while ($null -eq $group)

                Assign-M365Group -User $user -Group $group -ExistingGroups $groups
                $script:assignGroupsCompleted = $true
            }
            3 # Remove group
            {

            }
            4 # Finish with groups
            {

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
        $mailbox | Select-Object -Property @("DisplayName", "UserPrincipalName", @{ label="Type"; expression={$_.RecipientTypeDetails} }) | Out-Host
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
        Write-Host "There was an issue granting mailbox access. Please try again." -ForegroundColor $warningColor
        Write-Host $errorRecord.Exception.Message -ForegroundColor $warningColor
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
        $isValidResponse = $accessType -imatch '^[1-3]$' # regex matches 1-3

        if (-not($isValidResponse))
        {
            Write-Host "Please enter a number 1-3." -ForegroundColor $warningColor
        }
    }
    while (-not($isValidResponse))

    return $accessType
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
        return
    }

    return $response
}

function New-Checkbox($checked)
{
    if ($checked)
    {
        return "[X]"
    }
    return "[ ]"
}

# Initialize script scoped variables
Initialize-ColorScheme