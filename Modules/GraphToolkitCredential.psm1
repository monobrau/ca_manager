<#
.SYNOPSIS
    Save unified Graph app credentials to Windows Credential Manager under both EOA and ESR key prefixes
    so Entra Secret Rotate, Exchange Online Analyzer, and bulk workers find the same ClientId/secret.
.NOTES
    Target format matches GraphAppCredential.psm1: UserName = "TenantId|ClientId", Password = ClientSecret.
#>

$script:PrefixEOA = 'EOA-GraphApp-'
$script:PrefixESR = 'ESR-GraphApp-'

function Save-GraphToolkitAppToWCM {
    param(
        [Parameter(Mandatory = $true)][string]$TenantId,
        [Parameter(Mandatory = $true)][string]$ClientId,
        [Parameter(Mandatory = $true)][string]$ClientSecret,
        [Parameter(Mandatory = $false)][string]$TenantDisplayName
    )
    $userName = "${TenantId}|${ClientId}"
    foreach ($prefix in @($script:PrefixEOA, $script:PrefixESR)) {
        $target = "$prefix$TenantId"
        if (Get-Module -ListAvailable -Name CredentialManager) {
            try {
                Import-Module CredentialManager -ErrorAction Stop
                $cred = New-Object PSCredential $userName, (ConvertTo-SecureString $ClientSecret -AsPlainText -Force)
                New-StoredCredential -Target $target -Credentials $cred -ErrorAction Stop | Out-Null
            } catch {
                Start-Process -FilePath "cmdkey.exe" -ArgumentList "/generic:$target", "/user:$userName", "/pass:$ClientSecret" -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue | Out-Null
            }
        } else {
            Start-Process -FilePath "cmdkey.exe" -ArgumentList "/generic:$target", "/user:$userName", "/pass:$ClientSecret" -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue | Out-Null
        }
    }
    if ($TenantDisplayName -and -not [string]::IsNullOrWhiteSpace($TenantDisplayName)) {
        foreach ($prefix in @($script:PrefixEOA, $script:PrefixESR)) {
            $nameTarget = "${prefix}${TenantId}-DisplayName"
            try {
                if (Get-Module -ListAvailable -Name CredentialManager) {
                    Import-Module CredentialManager -ErrorAction Stop
                    $nameCred = New-Object PSCredential 'DisplayName', (ConvertTo-SecureString $TenantDisplayName -AsPlainText -Force)
                    New-StoredCredential -Target $nameTarget -Credentials $nameCred -ErrorAction Stop | Out-Null
                } else {
                    Start-Process -FilePath "cmdkey.exe" -ArgumentList "/generic:$nameTarget", "/user:DisplayName", "/pass:$TenantDisplayName" -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue | Out-Null
                }
            } catch { }
        }
    }
}

Export-ModuleMember -Function Save-GraphToolkitAppToWCM
