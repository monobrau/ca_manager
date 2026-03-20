<#
.SYNOPSIS
    Creates or updates the shared Entra app used by CA Manager, Entra Secret Rotate, and Exchange Online Analyzer (default: River Run Security Investigator).
.DESCRIPTION
    Default -DisplayName matches Exchange Online Analyzer's New-GraphInboxRulesApp.ps1 so all tools use one registration and the same EOA/ESR WCM credential entries.

    If an app with -DisplayName already exists, compares required Microsoft Graph delegated (Scope) and application (Role) permissions,
    public client redirect URIs, and sign-in audience. When anything is missing, you can update the existing registration, replace it
    (delete and recreate), or quit. When everything matches, you can add a new client secret and optionally save to WCM.

    After creation, set M365_GRAPH_TOOLKIT_CLIENT_ID to the Application (client) ID on each machine that runs the tools.

    Requires an admin sign-in with Application.ReadWrite.All and AppRoleAssignment.ReadWrite.All for bootstrap.
.PARAMETER SaveToWCM
    After create/update/replace, save ClientId and new secret under EOA-GraphApp- and ESR-GraphApp- in Credential Manager.
.PARAMETER MultiTenant
    signInAudience AzureADMultipleOrgs. Default AzureADMyOrg. Checked on existing apps; update path can align audience when you choose Update.
.PARAMETER DisplayName
    App registration display name in Entra (default: River Run Security Investigator, same as XOA).
.PARAMETER UpdateExisting
    Non-interactive: if the app exists and permissions or redirects are incomplete, patch the app and grant missing consent without prompting.
.PARAMETER ReplaceExisting
    Non-interactive: if the app exists, delete it and create a new one (destructive).
.PARAMETER Force
    With -ReplaceExisting, skip confirmation. With interactive Replace choice, still prompts unless -Force.
.PARAMETER NewSecret
    After a successful Update or when the app was already complete, create a new client secret (and SaveToWCM when -SaveToWCM).
.EXAMPLE
    .\New-UnifiedGraphToolkitApp.ps1 -SaveToWCM
.EXAMPLE
    .\New-UnifiedGraphToolkitApp.ps1 -UpdateExisting -SaveToWCM
.NOTES
    Application role IDs: Microsoft Graph 00000003-0000-0000-c000-000000000000 (graphpermissions.merill.net).
#>

#Requires -Version 5.1
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [switch]$SaveToWCM,
    [switch]$MultiTenant,
    # Same Entra app registration as Exchange Online Analyzer (New-GraphInboxRulesApp.ps1); shared via EOA/ESR WCM keys.
    [string]$DisplayName = 'River Run Security Investigator',
    [switch]$UpdateExisting,
    [switch]$ReplaceExisting,
    [switch]$Force,
    [switch]$NewSecret
)

$ErrorActionPreference = 'Stop'
$graphResourceId = '00000003-0000-0000-c000-000000000000'

foreach ($m in @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Applications')) {
    if (-not (Get-Module -ListAvailable -Name $m)) {
        Write-Error "Install-Module $m -Scope CurrentUser (required to create the app and secret)."
    }
    Import-Module $m -ErrorAction Stop
}

$appRoles = @(
    @{ id = '01c0a623-fc9b-48e9-b794-0756f8e8f067'; name = 'Policy.ReadWrite.ConditionalAccess' }
    @{ id = '810c84a8-4a9e-49e6-bf7d-12d183f40d01'; name = 'Mail.Read' }
    @{ id = '40f97065-369a-49f4-947c-6a255697ae91'; name = 'MailboxSettings.Read' }
    @{ id = 'df021288-bdef-4463-88db-98f22de89214'; name = 'User.Read.All' }
    @{ id = 'b0afded3-3588-46d8-8b3d-9842eff778da'; name = 'AuditLog.Read.All' }
    @{ id = '7ab1d382-f21e-4acd-a863-ba3e13f7da61'; name = 'Directory.Read.All' }
    @{ id = '498476e4-1e14-4a69-9742-9339357b62d5'; name = 'Organization.Read.All' }
    @{ id = '246dd0d5-5bd0-4def-940b-0421030a5b68'; name = 'Policy.Read.All' }
    @{ id = '9a5d68dd-52b0-4cc2-bd40-abcf44ac3a30'; name = 'Application.Read.All' }
    @{ id = '1bfefb4e-e0b5-418b-a88f-73c46d2cc8e9'; name = 'Application.ReadWrite.All' }
    @{ id = '06b708a9-e830-4db3-a914-8e69da51d44f'; name = 'AppRoleAssignment.ReadWrite.All' }
    @{ id = '230c1aed-a721-4c5d-9cb4-a90514e508ef'; name = 'Reports.Read.All' }
    @{ id = '332a536c-c7ef-4017-ab91-336970924f0d'; name = 'Sites.Read.All' }
    @{ id = 'bf394140-e372-4bf9-a898-299cfc7564e5'; name = 'SecurityEvents.Read.All' }
    @{ id = '38d9df27-64da-44fd-b7c5-a6fbac20248f'; name = 'UserAuthenticationMethod.Read.All' }
)

$delegatedPermissionValues = @(
    'Policy.Read.All'
    'Policy.ReadWrite.ConditionalAccess'
    'User.Read.All'
    'Group.Read.All'
    'Organization.Read.All'
    'Application.Read.All'
    'Application.ReadWrite.All'
    'AppRoleAssignment.ReadWrite.All'
)

$bootstrapScopes = @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All')
$expectedRedirectUris = @(
    'https://login.microsoftonline.com/common/oauth2/nativeclient'
    'http://localhost'
)

if ($UpdateExisting -and $ReplaceExisting) {
    Write-Error "Specify only one of -UpdateExisting or -ReplaceExisting."
}

$modRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$credMod = Join-Path $modRoot 'Modules\GraphToolkitCredential.psm1'

function Get-NormalizedResourceAccessSets {
    param($RequiredResourceAccess, [string]$ResourceAppId)
    $scopeIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $roleIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    if (-not $RequiredResourceAccess) { return @{ ScopeIds = $scopeIds; RoleIds = $roleIds } }
    $graphBlock = @($RequiredResourceAccess | Where-Object {
            $rid = if ($_.ResourceAppId) { $_.ResourceAppId } else { $_.resourceAppId }
            $rid -eq $ResourceAppId
        })[0]
    $resAcc = $null
    if ($graphBlock) {
        $resAcc = $graphBlock.ResourceAccess
        if (-not $resAcc) { $resAcc = $graphBlock.resourceAccess }
    }
    if (-not $graphBlock -or -not $resAcc) { return @{ ScopeIds = $scopeIds; RoleIds = $roleIds } }
    foreach ($ra in @($resAcc)) {
        $rid = if ($null -ne $ra.Id) { [string]$ra.Id } else { [string]$ra.id }
        $typ = if ($ra.Type) { [string]$ra.Type } else { [string]$ra.type }
        if (-not $rid) { continue }
        if ($typ -match '^[Ss]cope$') { [void]$scopeIds.Add($rid) }
        elseif ($typ -match '^[Rr]ole$') { [void]$roleIds.Add($rid) }
    }
    return @{ ScopeIds = $scopeIds; RoleIds = $roleIds }
}

function Get-ToolkitPermissionReport {
    param(
        $Application,
        [System.Collections.Generic.HashSet[string]]$ExpectedScopeIds,
        [System.Collections.Generic.HashSet[string]]$ExpectedRoleIds,
        [string]$GraphResourceId,
        [string[]]$ExpectedRedirects,
        [string]$ExpectedAudience
    )
    $rra = $Application.RequiredResourceAccess
    if (-not $rra) { $rra = $Application.requiredResourceAccess }
    $sets = Get-NormalizedResourceAccessSets -RequiredResourceAccess $rra -ResourceAppId $GraphResourceId
    $missingScopes = [System.Collections.ArrayList]::new()
    foreach ($x in $ExpectedScopeIds) { if (-not $sets.ScopeIds.Contains($x)) { [void]$missingScopes.Add($x) } }
    $missingRoles = [System.Collections.ArrayList]::new()
    foreach ($x in $ExpectedRoleIds) { if (-not $sets.RoleIds.Contains($x)) { [void]$missingRoles.Add($x) } }
    $extraScopes = [System.Collections.ArrayList]::new()
    foreach ($x in $sets.ScopeIds) { if (-not $ExpectedScopeIds.Contains($x)) { [void]$extraScopes.Add($x) } }
    $extraRoles = [System.Collections.ArrayList]::new()
    foreach ($x in $sets.RoleIds) { if (-not $ExpectedRoleIds.Contains($x)) { [void]$extraRoles.Add($x) } }

    $pc = $Application.PublicClient
    if (-not $pc) { $pc = $Application.publicClient }
    $redirects = @()
    if ($pc) {
        $uris = $pc.RedirectUris
        if (-not $uris) { $uris = $pc.redirectUris }
        if ($uris) { $redirects = @($uris) }
    }
    $missingRedirects = @($ExpectedRedirects | Where-Object { $r = $_; -not ($redirects | Where-Object { $_ -eq $r }) })

    $audience = $Application.SignInAudience
    if (-not $audience) { $audience = $Application.signInAudience }
    $audienceOk = ($audience -eq $ExpectedAudience)

    $permOk = ($missingScopes.Count -eq 0 -and $missingRoles.Count -eq 0)
    $redirectOk = ($missingRedirects.Count -eq 0)

    return [pscustomobject]@{
        MissingScopes      = $missingScopes
        MissingRoles       = $missingRoles
        ExtraScopes        = $extraScopes
        ExtraRoles         = $extraRoles
        PermissionsOk      = $permOk
        RedirectsOk        = $redirectOk
        MissingRedirects   = $missingRedirects
        CurrentRedirects   = $redirects
        AudienceOk         = $audienceOk
        CurrentAudience    = $audience
        ExpectedAudience   = $ExpectedAudience
        IsFullyAligned     = ($permOk -and $redirectOk -and $audienceOk)
    }
}

function Write-ToolkitPermissionReport {
    param($Report, $AppRoleLookup, $ScopeIdToValue)
    Write-Host "`n--- Permission check (Microsoft Graph) ---" -ForegroundColor Cyan
    if ($Report.PermissionsOk) {
        Write-Host "  Delegated (Scope) IDs: all required entries present." -ForegroundColor Green
    } else {
        if ($Report.MissingScopes.Count -gt 0) {
            Write-Host "  Missing delegated permission IDs:" -ForegroundColor Yellow
            foreach ($id in $Report.MissingScopes) {
                $lbl = if ($ScopeIdToValue.ContainsKey($id)) { $ScopeIdToValue[$id] } else { $id }
                Write-Host "    - $lbl ($id)" -ForegroundColor Yellow
            }
        }
        if ($Report.MissingRoles.Count -gt 0) {
            Write-Host "  Missing application permission IDs:" -ForegroundColor Yellow
            foreach ($id in $Report.MissingRoles) {
                $lbl = if ($AppRoleLookup.ContainsKey($id)) { $AppRoleLookup[$id] } else { $id }
                Write-Host "    - $lbl ($id)" -ForegroundColor Yellow
            }
        }
    }
    if ($Report.ExtraScopes.Count -gt 0 -or $Report.ExtraRoles.Count -gt 0) {
        Write-Host "  Note: App has extra Graph Scope/Role IDs not in the toolkit manifest (left unchanged on Update)." -ForegroundColor DarkGray
    }
    Write-Host "`n--- Public client / audience ---" -ForegroundColor Cyan
    if ($Report.RedirectsOk) {
        Write-Host "  Redirect URIs: OK" -ForegroundColor Green
    } else {
        Write-Host "  Missing redirect URIs:" -ForegroundColor Yellow
        $Report.MissingRedirects | ForEach-Object { Write-Host "    - $_" -ForegroundColor Yellow }
    }
    if ($Report.AudienceOk) {
        Write-Host "  Sign-in audience: OK ($($Report.CurrentAudience))" -ForegroundColor Green
    } else {
        Write-Host "  Sign-in audience: expected '$($Report.ExpectedAudience)', current '$($Report.CurrentAudience)'" -ForegroundColor Yellow
    }
}

function Grant-ToolkitApplicationRoleAssignments {
    param(
        [string]$ServicePrincipalId,
        [string]$GraphSpId,
        [array]$AppRolesDefinition
    )
    $existing = @()
    try {
        $existing = @(Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ServicePrincipalId -All -ErrorAction Stop)
    } catch {
        Write-Warning "Could not list app role assignments: $($_.Exception.Message)"
        $existing = @()
    }
    $assignedForGraph = @{}
    foreach ($a in $existing) {
        if ($a.ResourceId -eq $GraphSpId -and $a.AppRoleId) {
            $assignedForGraph[[string]$a.AppRoleId] = $true
        }
    }
    foreach ($role in $AppRolesDefinition) {
        if ($assignedForGraph.ContainsKey([string]$role.id)) { continue }
        try {
            New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $ServicePrincipalId -PrincipalId $ServicePrincipalId -AppRoleId $role.id -ResourceId $GraphSpId -ErrorAction Stop | Out-Null
            Write-Host "  [app granted] $($role.name)" -ForegroundColor Green
        } catch {
            Write-Warning "  [app] $($role.name): $($_.Exception.Message)"
        }
    }
}

function Grant-ToolkitDelegatedOAuth2Grant {
    param(
        [string]$ServicePrincipalId,
        [string]$GraphSpId,
        [string[]]$DelegatedPermissionValues
    )
    $scopeString = ($DelegatedPermissionValues -join ' ').Trim()
    try {
        $filterEnc = [uri]::EscapeDataString("clientId eq '$ServicePrincipalId' and resourceId eq '$GraphSpId'")
        $existingGrants = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants?`$filter=$filterEnc" -OutputType PSObject
        if ($existingGrants.value -and @($existingGrants.value).Count -gt 0) {
            $g = @($existingGrants.value)[0]
            $merged = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            $existingScope = if ($g.scope) { [string]$g.scope } else { '' }
            foreach ($p in ($existingScope -split '\s+')) { if ($p) { [void]$merged.Add($p) } }
            foreach ($p in $DelegatedPermissionValues) { [void]$merged.Add($p) }
            $newScope = ($merged | Sort-Object) -join ' '
            Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/oauth2PermissionGrants/$($g.id)" -Body (@{ scope = $newScope } | ConvertTo-Json) -ContentType 'application/json' -ErrorAction Stop | Out-Null
            Write-Host "  Delegated consent (oauth2PermissionGrant) updated." -ForegroundColor Green
        } else {
            $grantBody = @{
                clientId    = $ServicePrincipalId
                consentType = 'AllPrincipals'
                resourceId  = $GraphSpId
                scope       = $scopeString
            }
            Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/oauth2PermissionGrants' -Body ($grantBody | ConvertTo-Json) -ContentType 'application/json' -ErrorAction Stop | Out-Null
            Write-Host "  Delegated consent (oauth2PermissionGrant) created." -ForegroundColor Green
        }
    } catch {
        Write-Warning "oauth2PermissionGrant: $($_.Exception.Message)"
    }
}

function Merge-ToolkitRequiredResourceAccess {
    param(
        $ExistingApplication,
        [string]$GraphResourceId,
        [array]$ToolkitResourceAccessEntries
    )
    $rra = $ExistingApplication.RequiredResourceAccess
    if (-not $rra) { $rra = $ExistingApplication.requiredResourceAccess }
    $other = @()
    if ($rra) {
        $other = @($rra | Where-Object {
                $rid = if ($_.ResourceAppId) { $_.ResourceAppId } else { $_.resourceAppId }
                $rid -ne $GraphResourceId
            })
    }
    $graphBlock = @{
        resourceAppId  = $GraphResourceId
        resourceAccess = @($ToolkitResourceAccessEntries)
    }
    return @($other) + @($graphBlock)
}

function Invoke-RemoveToolkitApplication {
    param([string]$ApplicationObjectId, [string]$AppClientId)
    try {
        $sps = @(Get-MgServicePrincipal -Filter "appId eq '$AppClientId'" -ErrorAction SilentlyContinue)
        foreach ($x in $sps) {
            Remove-MgServicePrincipal -ServicePrincipalId $x.Id -ErrorAction SilentlyContinue
        }
    } catch { }
    Remove-MgApplication -ApplicationId $ApplicationObjectId -ErrorAction Stop
}

function Get-SecretPlainText {
    param($SecretObject)
    if ($null -ne $SecretObject.SecretText) { return $SecretObject.SecretText }
    if ($null -ne $SecretObject.secretText) { return $SecretObject.secretText }
    return $null
}

# --- Main ---

Write-Host "`n=== Unified Graph Toolkit app ===" -ForegroundColor Cyan
Write-Host "Connecting to Graph (bootstrap scopes)..." -ForegroundColor Yellow
Connect-MgGraph -Scopes $bootstrapScopes -NoWelcome -ErrorAction Stop | Out-Null

$tenantId = (Get-MgContext).TenantId
Write-Host "TenantId: $tenantId" -ForegroundColor Gray

Write-Host "Loading Microsoft Graph service principal..." -ForegroundColor Yellow
$graphSpUri = "https://graph.microsoft.com/v1.0/servicePrincipals(appId='$graphResourceId')?`$select=id,oauth2PermissionScopes"
$graphSpObj = Invoke-MgGraphRequest -Method GET -Uri $graphSpUri -OutputType PSObject
$graphSpId = $graphSpObj.id
$publishedScopes = @($graphSpObj.oauth2PermissionScopes)
if ($publishedScopes.Count -eq 0) {
    Write-Error "Could not read oauth2PermissionScopes from Microsoft Graph service principal."
}

$scopeIdToValue = @{}
foreach ($s in $publishedScopes) {
    $sid = if ($null -ne $s.id) { [string]$s.id } else { [string]$s.Id }
    $val = if ($null -ne $s.value) { [string]$s.value } else { [string]$s.Value }
    if ($sid) { $scopeIdToValue[$sid] = $val }
}

$raList = New-Object System.Collections.ArrayList
$seen = @{}
foreach ($val in $delegatedPermissionValues) {
    $scope = @($publishedScopes | Where-Object { $_.value -eq $val -or $_.Value -eq $val })[0]
    if (-not $scope) {
        Write-Warning "Delegated permission '$val' not found on Graph SP - skip."
        continue
    }
    $sid = if ($null -ne $scope.id) { $scope.id } else { $scope.Id }
    $key = "Scope:$sid"
    if (-not $seen.ContainsKey($key)) {
        $seen[$key] = $true
        [void]$raList.Add(@{ id = [string]$sid; type = 'Scope' })
    }
}
foreach ($role in $appRoles) {
    $key = "Role:$($role.id)"
    if (-not $seen.ContainsKey($key)) {
        $seen[$key] = $true
        [void]$raList.Add(@{ id = $role.id; type = 'Role' })
    }
}

$expectedScopeIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
$expectedRoleIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($item in @($raList.ToArray())) {
    if ($item.type -match '^[Ss]cope$') { [void]$expectedScopeIds.Add([string]$item.id) }
    elseif ($item.type -match '^[Rr]ole$') { [void]$expectedRoleIds.Add([string]$item.id) }
}

$appRoleLookup = @{}
foreach ($r in $appRoles) { $appRoleLookup[[string]$r.id] = $r.name }

$audience = if ($MultiTenant) { 'AzureADMultipleOrgs' } else { 'AzureADMyOrg' }

$escapedName = $DisplayName.Replace("'", "''")
$filter = "displayName eq '$escapedName'"
$existingApps = @(Get-MgApplication -Filter $filter -ConsistencyLevel eventual -ErrorAction SilentlyContinue)

if ($existingApps.Count -gt 1) {
    Write-Error "Multiple app registrations named '$DisplayName' ($($existingApps.Count)). Rename or delete duplicates in Entra, then re-run."
}

$existingApp = if ($existingApps.Count -eq 1) { $existingApps[0] } else { $null }

if ($existingApp) {
    $sel = 'id,appId,displayName,requiredResourceAccess,publicClient,signInAudience'
    $fullApp = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/applications/$($existingApp.Id)?`$select=$sel" -OutputType PSObject
    $disp = if ($fullApp.displayName) { $fullApp.displayName } else { $fullApp.DisplayName }
    $objId = if ($fullApp.id) { $fullApp.id } else { $fullApp.Id }
    $cliId = if ($fullApp.appId) { $fullApp.appId } else { $fullApp.AppId }
    $report = Get-ToolkitPermissionReport -Application $fullApp -ExpectedScopeIds $expectedScopeIds -ExpectedRoleIds $expectedRoleIds `
        -GraphResourceId $graphResourceId -ExpectedRedirects $expectedRedirectUris -ExpectedAudience $audience
    Write-Host "`nFound existing app: $disp" -ForegroundColor Cyan
    Write-Host "  Object ID: $objId  Client ID: $cliId" -ForegroundColor Gray
    Write-ToolkitPermissionReport -Report $report -AppRoleLookup $appRoleLookup -ScopeIdToValue $scopeIdToValue

    if ($report.IsFullyAligned) {
        Write-Host "`nThis app already has all toolkit permissions, redirect URIs, and expected sign-in audience." -ForegroundColor Green
        if (-not $NewSecret -and -not $SaveToWCM) {
            Write-Host "Use -NewSecret to create a client secret, or -SaveToWCM with -NewSecret to rotate and store in WCM." -ForegroundColor Gray
        }
        if ($NewSecret -or $SaveToWCM) {
            if (-not $NewSecret -and $SaveToWCM) {
                Write-Warning "-SaveToWCM requires a new secret. Adding -NewSecret."
                $NewSecret = $true
            }
            if ($PSCmdlet.ShouldProcess($DisplayName, 'Create new client secret')) {
                $pwdCred = @{ displayName = 'GraphToolkit'; endDateTime = (Get-Date).AddMonths(24).ToUniversalTime().ToString('o') }
                $secret = Add-MgApplicationPassword -ApplicationId $fullApp.Id -PasswordCredential $pwdCred -ErrorAction Stop
                $shownSecret = Get-SecretPlainText -SecretObject $secret
                Write-Host "`nClient secret (save now - shown once):" -ForegroundColor Yellow
                Write-Host "  $shownSecret" -ForegroundColor White
                if ($SaveToWCM) {
                    try {
                        Import-Module $credMod -Force -ErrorAction Stop
                        $tdn = $null
                        try {
                            $org = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization' -ErrorAction SilentlyContinue
                            if ($org.value -and $org.value[0].displayName) { $tdn = $org.value[0].displayName }
                        } catch { }
                        Save-GraphToolkitAppToWCM -TenantId $tenantId -ClientId $cliId -ClientSecret $shownSecret -TenantDisplayName $tdn
                        Write-Host "Saved to WCM (EOA + ESR prefixes)." -ForegroundColor Green
                    } catch {
                        Write-Warning "WCM save failed: $($_.Exception.Message)"
                    }
                }
                Write-Host "`nM365_GRAPH_TOOLKIT_CLIENT_ID = $cliId" -ForegroundColor Cyan
            }
        }
        exit 0
    }

    # Not fully aligned
    $action = $null
    if ($ReplaceExisting) {
        $action = 'R'
    } elseif ($UpdateExisting) {
        $action = 'U'
    } else {
        Write-Host "`nChoose an action:" -ForegroundColor Yellow
        Write-Host "  [U] Update  - merge toolkit permissions into this app, fix redirects/audience, grant missing admin consent" -ForegroundColor White
        Write-Host "  [R] Replace - DELETE this app and create a new one (client ID changes; update env vars and WCM)" -ForegroundColor White
        Write-Host "  [Q] Quit    - no changes" -ForegroundColor White
        do {
            $action = Read-Host "Enter U, R, or Q"
            $action = if ($action) { $action.Trim().ToUpperInvariant().Substring(0, 1) } else { '' }
        } while ($action -notin 'U', 'R', 'Q')
    }

    if ($action -eq 'Q') {
        Write-Host "No changes made." -ForegroundColor Gray
        exit 0
    }

    if ($action -eq 'R') {
        if (-not $Force -and -not $ReplaceExisting) {
            $confirm = Read-Host "Type DELETE to confirm removal of app '$DisplayName' ($cliId)"
            if ($confirm -ne 'DELETE') {
                Write-Host "Cancelled." -ForegroundColor Gray
                exit 0
            }
        }
        if (-not $PSCmdlet.ShouldProcess($DisplayName, 'Delete app registration')) { exit 0 }
        Write-Host "Removing existing app..." -ForegroundColor Yellow
        Invoke-RemoveToolkitApplication -ApplicationObjectId $objId -AppClientId $cliId
        $existingApp = $null
    }

    if ($action -eq 'U') {
        if (-not $PSCmdlet.ShouldProcess($DisplayName, 'Update app registration')) { exit 0 }
        Write-Host "`nUpdating app registration..." -ForegroundColor Yellow
        $mergedRra = Merge-ToolkitRequiredResourceAccess -ExistingApplication $fullApp -GraphResourceId $graphResourceId -ToolkitResourceAccessEntries @($raList.ToArray())
        $newRedirects = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($u in @($fullApp.PublicClient.RedirectUris)) { if ($u) { [void]$newRedirects.Add($u) } }
        foreach ($u in $expectedRedirectUris) { [void]$newRedirects.Add($u) }
        $patchBody = @{
            requiredResourceAccess = @($mergedRra | ForEach-Object {
                    @{
                        resourceAppId  = $_.ResourceAppId
                        resourceAccess = @($_.ResourceAccess | ForEach-Object {
                                @{ id = [string]$_.Id; type = [string]$_.Type }
                            })
                    }
                })
            signInAudience         = $audience
            publicClient           = @{ redirectUris = @($newRedirects.ToArray()) }
        }
        $patchJson = $patchBody | ConvertTo-Json -Depth 14
        Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/applications/$objId" -Body $patchJson -ContentType 'application/json' -ErrorAction Stop | Out-Null
        Write-Host "  Application manifest patched." -ForegroundColor Green

        $sp = @(Get-MgServicePrincipal -Filter "appId eq '$cliId'" -ErrorAction Stop | Select-Object -First 1)
        if ($sp.Count -eq 0) {
            Write-Host "  Creating service principal..." -ForegroundColor Yellow
            $sp = @(New-MgServicePrincipal -AppId $cliId -ErrorAction Stop)
        }
        $spId = $sp[0].Id
        Grant-ToolkitApplicationRoleAssignments -ServicePrincipalId $spId -GraphSpId $graphSpId -AppRolesDefinition $appRoles
        Grant-ToolkitDelegatedOAuth2Grant -ServicePrincipalId $spId -GraphSpId $graphSpId -DelegatedPermissionValues $delegatedPermissionValues

        Write-Host "`nUpdate complete. Verify 'Grant admin consent' in Entra if any permission still shows as not consented." -ForegroundColor Cyan
        $clientId = $cliId
        $objectId = $objId

        if ($NewSecret -or $SaveToWCM) {
            if ($SaveToWCM -and -not $NewSecret) { $NewSecret = $true }
            if ($PSCmdlet.ShouldProcess($DisplayName, 'Create new client secret after update')) {
                $pwdCred = @{ displayName = 'GraphToolkit'; endDateTime = (Get-Date).AddMonths(24).ToUniversalTime().ToString('o') }
                $secret = Add-MgApplicationPassword -ApplicationId $objectId -PasswordCredential $pwdCred -ErrorAction Stop
                $shownSecret = Get-SecretPlainText -SecretObject $secret
                Write-Host "`nClient secret (save now - shown once):" -ForegroundColor Yellow
                Write-Host "  $shownSecret" -ForegroundColor White
                if ($SaveToWCM) {
                    try {
                        Import-Module $credMod -Force -ErrorAction Stop
                        $tdn = $null
                        try {
                            $org = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization' -ErrorAction SilentlyContinue
                            if ($org.value -and $org.value[0].displayName) { $tdn = $org.value[0].displayName }
                        } catch { }
                        Save-GraphToolkitAppToWCM -TenantId $tenantId -ClientId $clientId -ClientSecret $shownSecret -TenantDisplayName $tdn
                        Write-Host "Saved to WCM (EOA + ESR prefixes)." -ForegroundColor Green
                    } catch { Write-Warning "WCM save failed: $($_.Exception.Message)" }
                }
            }
        }
        Write-Host "`nM365_GRAPH_TOOLKIT_CLIENT_ID = $clientId" -ForegroundColor Cyan
        exit 0
    }
}

# --- Create new app (no existing, or replaced) ---
if (-not $PSCmdlet.ShouldProcess($DisplayName, 'Create app registration')) {
    Write-Host "WhatIf: would create app with $($raList.Count) resourceAccess entries." -ForegroundColor Gray
    exit 0
}

Write-Host "Creating application '$DisplayName'..." -ForegroundColor Yellow
$appCreateBody = @{
    displayName            = $DisplayName
    signInAudience         = $audience
    publicClient           = @{ redirectUris = @($expectedRedirectUris) }
    requiredResourceAccess = @(
        @{
            resourceAppId  = $graphResourceId
            resourceAccess = @($raList.ToArray())
        }
    )
}
$createJson = $appCreateBody | ConvertTo-Json -Depth 14
$app = Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/applications' -Body $createJson -ContentType 'application/json' -OutputType PSObject

$objectId = $app.id
$clientId = $app.appId
Write-Host "  Object ID: $objectId" -ForegroundColor Gray
Write-Host "  Client ID: $clientId" -ForegroundColor Green

Write-Host "Creating service principal..." -ForegroundColor Yellow
$sp = New-MgServicePrincipal -AppId $clientId -ErrorAction Stop
Write-Host "  Service principal ID: $($sp.Id)" -ForegroundColor Gray

Grant-ToolkitApplicationRoleAssignments -ServicePrincipalId $sp.Id -GraphSpId $graphSpId -AppRolesDefinition $appRoles
Grant-ToolkitDelegatedOAuth2Grant -ServicePrincipalId $sp.Id -GraphSpId $graphSpId -DelegatedPermissionValues $delegatedPermissionValues

Write-Host "Creating client secret..." -ForegroundColor Yellow
$pwdCred = @{
    displayName = 'GraphToolkit'
    endDateTime = (Get-Date).AddMonths(24).ToUniversalTime().ToString('o')
}
$secret = Add-MgApplicationPassword -ApplicationId $objectId -PasswordCredential $pwdCred -ErrorAction Stop
Write-Host "  Secret expires: $($secret.EndDateTime)" -ForegroundColor Gray

$tenantDisplayName = $null
try {
    $org = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization' -ErrorAction Stop
    if ($org.value -and $org.value[0].displayName) { $tenantDisplayName = $org.value[0].displayName }
} catch { }

if ($SaveToWCM) {
    Write-Host "Saving credentials to Windows Credential Manager (EOA + ESR prefixes)..." -ForegroundColor Yellow
    try {
        Import-Module $credMod -Force -ErrorAction Stop
        $secTxt = Get-SecretPlainText -SecretObject $secret
        Save-GraphToolkitAppToWCM -TenantId $tenantId -ClientId $clientId -ClientSecret $secTxt -TenantDisplayName $tenantDisplayName
        Write-Host "  Saved." -ForegroundColor Green
    } catch {
        Write-Warning "WCM save failed: $($_.Exception.Message). Install-Module CredentialManager -Scope CurrentUser"
    }
}

Write-Host "`n=== Done ===" -ForegroundColor Cyan
Write-Host "Set on each workstation:" -ForegroundColor Yellow
Write-Host "  M365_GRAPH_TOOLKIT_CLIENT_ID = $clientId" -ForegroundColor White
Write-Host "`nClient secret (save now - shown once):" -ForegroundColor Yellow
Write-Host "  $(Get-SecretPlainText -SecretObject $secret)" -ForegroundColor White
Write-Host "`nEntra portal: App registrations > $DisplayName > verify API permissions." -ForegroundColor Gray
Write-Host ""
