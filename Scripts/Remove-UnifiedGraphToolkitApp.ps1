<#
.SYNOPSIS
    Removes all Entra app registrations matching the toolkit display name (same idea as Remove-GraphInboxRulesApp.ps1 for XOA).
.DESCRIPTION
    Deletes each matching application and its service principal(s). Optional -RemoveWCM clears EOA/ESR Graph app credentials
    for the signed-in tenant on this machine.

    Default display name is River Run Security Investigator (shared with Exchange Online Analyzer); deleting removes that app for all tools.
.PARAMETER DisplayName
    Exact app registration display name to match (default: River Run Security Investigator, same as New-UnifiedGraphToolkitApp.ps1 / XOA).
.PARAMETER TenantId
    If set, Connect-MgGraph uses this tenant for admin sign-in.
.PARAMETER Force
    Skip the console confirmation prompt (use when the caller already confirmed, e.g. CA Manager GUI).
.PARAMETER RemoveWCM
    After successful deletes, remove EOA-GraphApp-{tenantId} and ESR-GraphApp-{tenantId} from Windows Credential Manager.
#>
#Requires -Version 5.1
param(
    [string]$DisplayName = 'River Run Security Investigator',
    [string]$TenantId = $null,
    [switch]$Force,
    [switch]$RemoveWCM
)

$ErrorActionPreference = 'Stop'
foreach ($m in @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Applications')) {
    if (-not (Get-Module -ListAvailable -Name $m)) {
        Write-Error "Install-Module $m -Scope CurrentUser"
    }
    Import-Module $m -ErrorAction Stop
}

$scopes = @('Application.ReadWrite.All', 'AppRoleAssignment.ReadWrite.All')
Write-Host "`n=== Remove unified Graph toolkit app(s) ===" -ForegroundColor Cyan
Write-Host "Connecting with $($scopes -join ', ')..." -ForegroundColor Yellow
try {
    if ($TenantId) {
        Connect-MgGraph -Scopes $scopes -NoWelcome -TenantId $TenantId.Trim() -ErrorAction Stop | Out-Null
        Write-Host "Tenant: $TenantId" -ForegroundColor Gray
    } else {
        Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop | Out-Null
    }
} catch {
    Write-Error "Graph connect failed: $_"
}

$escapedName = $DisplayName.Replace("'", "''")
$filter = "displayName eq '$escapedName'"
$existingApps = @(Get-MgApplication -Filter $filter -ConsistencyLevel eventual -ErrorAction SilentlyContinue)

if ($existingApps.Count -eq 0) {
    Write-Host "`nNo app registrations named '$DisplayName' found." -ForegroundColor Gray
    exit 0
}

Write-Host "`nFound $($existingApps.Count) app(s) named '$DisplayName':" -ForegroundColor Yellow
foreach ($a in $existingApps) {
    Write-Host "  Object Id: $($a.Id)  Client Id: $($a.AppId)" -ForegroundColor Gray
}

if (-not $Force) {
    Write-Host "`nRemove ALL of these registrations? (y/n): " -ForegroundColor Yellow -NoNewline
    $reply = Read-Host
    if ($reply -ne 'y' -and $reply -ne 'Y') {
        Write-Host "Cancelled." -ForegroundColor Gray
        exit 0
    }
}

$ctxTenantId = (Get-MgContext).TenantId
foreach ($a in $existingApps) {
    $objId = $a.Id
    $cliId = $a.AppId
    Write-Host "Removing app $cliId ..." -ForegroundColor Yellow
    try {
        $sps = @(Get-MgServicePrincipal -Filter "appId eq '$cliId'" -ErrorAction SilentlyContinue)
        foreach ($x in $sps) {
            Remove-MgServicePrincipal -ServicePrincipalId $x.Id -ErrorAction SilentlyContinue
        }
    } catch { }
    Remove-MgApplication -ApplicationId $objId -ErrorAction Stop
    Write-Host "  Removed." -ForegroundColor Green
}

Write-Host "`nDone. Removed $($existingApps.Count) app registration(s)." -ForegroundColor Cyan

if ($RemoveWCM) {
    $modRoot = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
    $credMod = Join-Path $modRoot 'Modules\GraphAppCredential.psm1'
    if (Test-Path $credMod) {
        Import-Module $credMod -Force -ErrorAction Stop
        foreach ($pfx in @('EOA', 'ESR')) {
            Remove-GraphAppCredentialFromWCM -TenantId $ctxTenantId -Prefix $pfx -ErrorAction SilentlyContinue
        }
        Write-Host "Removed WCM Graph app credentials for tenant $ctxTenantId (EOA + ESR)." -ForegroundColor Green
    } else {
        Write-Warning "GraphAppCredential.psm1 not found; skipped WCM cleanup."
    }
}
