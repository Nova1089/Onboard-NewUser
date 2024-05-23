<#
Version 1.0

Script synopsis...
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
    Write-Host "This script does some stuff..." -ForegroundColor $infoColor
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

function Get-M365User($upn)
{
    return (Get-MgUser -UserID $upn -Property @("CreatedDateTime", "DisplayName", "UserPrincipalName", "JobTitle", "Department", "UsageLocation", "LicenseDetails") -ErrorAction "SilentlyContinue")
}

function Display-BasicUserProperties($user)
{
    Write-Host "User found!" -ForegroundColor $successColor
    $managerId = Get-MgUserManager -UserId $user.UserPrincipalName | Select-Object -ExpandProperty "ID"
    $managerDisplayName = Get-MgUser -UserId $managerId | Select-Object -ExpandProperty "DisplayName"
    $licenses = Get-UserLicenses $user.UserPrincipalName
    $userProps = [PSCustomObject]@{
        "Created Date/Time" = $user.CreatedDateTime
        "Display Name" = $user.DisplayName
        "UPN" = $user.UserPrincipalName
        "Title" = $user.JobTitle
        "Department" = $user.Department
        "Manager"      = $managerDisplayName
        "Usage Location" = $user.UsageLocation
        "Licenses" = $licenses
    }
    $userProps | Out-Host
}

function Get-UserLicenses($upn)
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

    $licenseDetails = Get-MGUserLicenseDetail -UserId $upn
    $licenses = New-Object System.Collections.Generic.List[String]

    foreach ($license in $licenseDetails)
    {
        $licenseName = $script:lookupTable[$license.SkuId]
        $licenses.Add($licenseName)
    }

    return $licenses
}

# main
Initialize-ColorScheme
Show-Introduction
Use-Module -ModuleName "Microsoft.Graph.Users"
TryConnect-MgGraph -Scopes User.ReadWrite.All
$upn = Read-Host "Please enter the UPN for the new user (PreferredFirstName.LastName@blueravensolar.com)"
$user = Get-M365User $upn
if ($null -ne $user)
{
    Display-BasicUserProperties $user
}
