# Dot sourcing
. "$PSScriptRoot\GlobalFunctions.ps1"

class GotoWizard
{
    # properties
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
    GotoWizard(){}

    # methods
    Start()
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

    ShowUserInfo()
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

    AssignUserRole($roleSelection)
    {
        # Need to refresh the user's info here because the method of assigning a new role depends on their current one.
        $this.gotoUser = $this.GetUser()
        
        $method = "Put"
        $uri = ""
        $headers = @{
            "Authorization" = "Bearer $($this.accessToken)"
            "Content-Type" = "application/json"
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
            {$_ -ge 4} # roleSelection is >= 4 and therefore a custom role
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

    DisplayLinkToOutboundCallerId()
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