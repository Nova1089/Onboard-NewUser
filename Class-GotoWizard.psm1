# Import modules in the same folder. These imports must come first in the script.
using module .\GlobalFunctions.psm1

# Initialize script scoped variables
Initialize-ColorScheme

class GotoWizard
{
    # properties
    [string] $upn
    [string] $gotoSecret
    [string] $clientId
    [string] $clientSecret
    [string] $accountKey
    [string] $accessToken
    [object] $user
    [System.Collections.Specialized.OrderedDictionary] $allCustomRoles
    [bool] $roleStepCompleted
    [bool] $cidStepCompleted

    # constructors
    GotoWizard($upn)
    {
        $this.upn = $upn
        $this.gotoSecret = GetApiSecret
        $this.clientId = $this.gotoSecret.ClientId
        $this.clientSecret = $this.gotoSecret.ClientSecret
        $this.accountKey = $this.gotoSecret.AccountKey
        $this.accessToken = $this.GetAccessToken()
        $this.user = $this.GetUser()
        $this.allCustomRoles = $this.GetAllCustomRoles()
    }

    # Default constructor defined for Pester testing purposes.
    GotoWizard(){}

    # methods
    Start()
    {
        if ($this.gotoUser)
        {
            Write-Host "Found Goto user!" -ForegroundColor $script:successColor
            ShowUserInfo
            $menuSelection = Prompt-GotoMenu
            switch ($menuSelection)
            {
                1 # show user info
                {
                    ShowUserInfo
                }
                2 # assign role
                {
                    $roleSelection = PromptRoleToAssign
                    AssignUserRole($roleSelection)
                }
                3 # assign outbound caller ID
                {
                    $this.DisplayLinkToOutboundCallerId()
                }
                4 # finish
                {
                    # break out of the goto wizard
                }
            }
        }
        else
        {
            Write-Host "Could not find goto user :(" -ForegroundColor $script:warningColor
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
        return $response.results
    }

    ShowUserInfo()
    {
        $role = $this.GetUserRole()
        # I'd also show their external caller ID, but this is not accessible from the public API.
        $this.user | Add-Member -NotePropertyName "role" -NotePropertyValue $role -PassThru | Format-List -Property @("email", "firstName", "lastName", "role") | Out-Host
    }

    [string] GetUserRole()
    {
        # Check for a built-in (system) admin role.
        if ($this.user.adminRoles)
        {
            if ($this.user.adminRoles -eq "SUPER_USER")
            {
                return "Super Admin"
            }

            if ($this.user.adminRoles -eq "MANAGE_ACCOUNT")
            {
                return "Admin (Configure PBX)"
            }
        }
        
        # Check for a custom role.
        if ($this.user.roleSets)
        {
            return $this.allCustomRoles[$this.user.roleSets]
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

    [int] PromptMenu()
    {
        $selection = Read-Host ("What next?`n`n" +
            "[1] Show Goto User Info`n" +
            "[2] $(New-Checkbox $this.roleStepCompleted) Assign role`n" +
            "[3] $(New-Checkbox $this.cidStepCompleted) Assign outbound caller ID`n" +
            "[4] Finish Goto setup`n")

        while ($selection -notmatch '^\s*[1-4]\s*$') # regex matches 1-4 but allows spaces
        {
            $selection = Read-Host "Please enter 1-4"
        }

        return [int]$selection.Trim()
    }

    [int] PromptRoleToAssign()
    {
        $menuText = ("Select a role to assign.`n`n" +
            "[1] Super Admin`n" +
            "[2] Admin (Configure PBX)`n" +
            "[3] Member`n")

        $optionsCount = 3
        foreach ($roleName in $this.allCustomRoles.Values)
        {
            $menuText = $menuText + "[$optionsCount] $roleName`n"
            $optionsCount++
        }

        $selection = Read-Host $menuText

        while ($selection -notmatch "^\s*[1-$optionsCount]\s*$") # regex matches 1 - optionsCount but allows spaces
        {
            $selection = Read-Host "Please enter 1-$optionsCount"
        }

        return [int]$selection.Trim()
    }

    AssignUserRole($roleSelection)
    {
        $method = "Put"
        $uri = ""
        $headers = @{
            "Authorization" = "Bearer $($this.accessToken)"
        }
        $body = @{}

        switch ($roleSelection)
        {
            1 # selection is Super Admin
            {
                $uri = "https://api.getgo.com/admin/rest/v1/accounts/$($this.accountKey)/users/$($this.user.key)"
                $body = @{
                    "adminRoles" = @("SUPER_USER")
                }
            }
            2 # selection is Admin (Configure PBX)
            {
                $uri = "https://api.getgo.com/admin/rest/v1/accounts/$($this.accountKey)/users/$($this.user.key)"
                $body = @{
                    "adminRoles" = @("MANAGE_ACCOUNT")
                }
            }
            3 # selection is member
            {
                if ($this.user.adminRoles) # user has a built-in (system) admin role that must be removed
                {
                    $uri = "https://api.getgo.com/admin/rest/v1/accounts/$($this.accountKey)/users/$($this.user.key)"
                    $body = @{ 
                        "adminRoles" = @() # set admin role to blank array
                    }
                }
                elseif ($this.user.roleSets) # user has a custom role that must be removed
                {
                    $method = "Delete"
                    $uri = "https://api.getgo.com/admin/rest/v1/accounts/$($this.accountKey)/users/rolesets"
                    $emailQuery = "email eq `"$($this.user.email)`""
                    $emailQuery = UriEncode-QueryParam $emailQuery
                    $uri = $uri + "?filter=$emailQuery"
                }
            }
            {$_ -ge 4} # roleSelection is >= 4 and therefore a custom role
            {
                $roleId = @($this.allCustomRoles.keys)[$roleSelection - 4] # enumerates the keys and accesses them by index
                $uri = "https://api.getgo.com/admin/rest/v1/accounts/$($this.accountKey)/rolesets/$roleId/users"
                $emailQuery = "email eq `"$($this.user.email)`""
                $emailQuery = UriEncode-QueryParam $emailQuery
                $uri = $uri + "?filter=$emailQuery"
            }
        }

        SafelyInvoke-RestMethod -Method $method -Uri $uri -Headers $headers -Body ($body | ConvertTo-Json)
    }

    DisplayLinkToOutboundCallerId()
    {
        $line = $this.GetLine()
        Write-Host "At this time the GoTo API can't change the outbound caller ID. You'll need to change it here:" -ForegroundColor "DarkCyan"
        Write-Host "https://my.jive.com/pbx/brs/extensions/lines/$($line.id)/general?source=root.nav.pbx.extensions.lines.list" -ForegroundColor "DarkCyan"
    }

    [object] GetLine()
    {
        $method = "Get"
        $uri = "https://api.goto.com/users/v1/users/$($this.user.key)/lines"
        $headers = @{
            "Authorization" = "Bearer $($this.accessToken)"
        }

        $response = SafelyInvoke-RestMethod -Method $method -Uri $uri -Headers $headers
        return $response.items
    }
}