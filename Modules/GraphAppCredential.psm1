<#
.SYNOPSIS
    Store and retrieve Graph app credentials (app-only) in Windows Credential Manager.
.DESCRIPTION
    Target formats: EOA-GraphApp-{tenantId} (ExchangeOnlineAnalyzer), ESR-GraphApp-{tenantId} (Entra Secret Rotate).
    UserName stores "TenantId|ClientId", Password stores ClientSecret.
.NOTES
    Requires: Install-Module CredentialManager (optional; cmdkey fallback)
#>

$script:credTargetPrefixEOA = 'EOA-GraphApp-'
$script:credTargetPrefixESR = 'ESR-GraphApp-'

function _Get-WcmGraphAppTenantIdSuffixVariants {
    <#
    .SYNOPSIS
        Suffix variants after EOA-GraphApp- / ESR-GraphApp- (cmdkey targets differ by braces or GUID casing).
    #>
    param([string]$TenantId)
    $cand = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $t = $TenantId.Trim()
    if ($t) { [void]$cand.Add($t) }
    $noBrace = $t -replace '[\{\}]', ''
    if ($noBrace) {
        [void]$cand.Add($noBrace)
        [void]$cand.Add($noBrace.ToLowerInvariant())
        [void]$cand.Add($noBrace.ToUpperInvariant())
        [void]$cand.Add('{' + $noBrace + '}')
        [void]$cand.Add('{' + $noBrace.ToUpperInvariant() + '}')
    }
    return @($cand)
}

function _Get-WcmCredReadTargetVariants {
    <#
    .SYNOPSIS
        Windows stores generic credentials as LegacyGeneric:target=<name> in many cases; CredRead may need either form.
    #>
    param([string]$BaseTarget)
    $set = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    if ([string]::IsNullOrWhiteSpace($BaseTarget)) { return @() }
    $b = $BaseTarget.Trim()
    [void]$set.Add($b)
    if ($b -notlike 'LegacyGeneric:target=*') {
        [void]$set.Add('LegacyGeneric:target=' + $b)
    }
    return @($set)
}

function _Get-GraphAppCredentialFromWCMTarget {
    param([string]$Target)
    foreach ($tn in (_Get-WcmCredReadTargetVariants -BaseTarget $Target)) {
        if (Get-Module -ListAvailable -Name CredentialManager) {
            try {
                Import-Module CredentialManager -ErrorAction Stop
                $cred = Get-StoredCredential -Target $tn -ErrorAction SilentlyContinue
                if ($cred) {
                    $parts = $cred.UserName -split '\|', 2
                    if ($parts.Count -ge 2) {
                        $pw = $cred.GetNetworkCredential().Password
                        if (-not [string]::IsNullOrWhiteSpace($pw)) {
                            return [pscustomobject]@{
                                TenantId     = $parts[0]
                                ClientId     = $parts[1]
                                ClientSecret = $pw
                            }
                        }
                    }
                }
            } catch { }
        }
        try {
            $credObj = _ReadCredentialViaCredRead -Target $tn
            if (-not $credObj) { continue }
            $parts = $credObj.UserName -split '\|', 2
            if ($parts.Count -lt 2) { continue }
            if ([string]::IsNullOrWhiteSpace($credObj.CredentialBlob)) { continue }
            return [pscustomobject]@{
                TenantId     = $parts[0]
                ClientId     = $parts[1]
                ClientSecret = $credObj.CredentialBlob
            }
        } catch { }
    }
    return $null
}

function Get-GraphAppCredentialFromWCM {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TenantId,
        [Parameter(Mandatory = $false)]
        [ValidateSet('EOA', 'ESR')]
        [string]$Prefix = 'EOA'
    )
    $credPrefix = if ($Prefix -eq 'ESR') { $script:credTargetPrefixESR } else { $script:credTargetPrefixEOA }
    foreach ($tidCand in (_Get-WcmGraphAppTenantIdSuffixVariants -TenantId $TenantId)) {
        $target = "$credPrefix$tidCand"
        $got = _Get-GraphAppCredentialFromWCMTarget -Target $target
        if ($got) { return $got }
    }
    return $null
}

function Get-WCMTenantIds {
    param(
        [Parameter(Mandatory = $false)]
        [ValidateSet('EOA', 'ESR')]
        [string]$Prefix = 'EOA'
    )
    $credPrefix = if ($Prefix -eq 'ESR') { $script:credTargetPrefixESR } else { $script:credTargetPrefixEOA }
    $tenantIds = @()
    try {
        $output = cmdkey /list 2>$null
        if ($output) {
            $text = $output | Out-String
            # Do not match ...-DisplayName entries (those are not app secrets; they used to pollute the list and break connect).
            $pattern = [regex]::Escape($credPrefix) + '(?:\{)?([a-fA-F0-9]{8}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{4}-[a-fA-F0-9]{12})(?:\})?(?!-DisplayName)'
            $m = [regex]::Matches($text, $pattern)
            foreach ($match in $m) {
                if ($match.Success -and $match.Groups[1].Value) {
                    $tid = $match.Groups[1].Value
                    if ($tid -notin $tenantIds) { $tenantIds += $tid }
                }
            }
        }
    } catch {}
    return $tenantIds
}

function _Get-StoredDisplayName {
    param([string]$TenantId, [string]$Prefix = 'EOA')
    $credPrefix = if ($Prefix -eq 'ESR') { $script:credTargetPrefixESR } else { $script:credTargetPrefixEOA }
    foreach ($tidCand in (_Get-WcmGraphAppTenantIdSuffixVariants -TenantId $TenantId)) {
        $base = "${credPrefix}${tidCand}-DisplayName"
        foreach ($tn in (_Get-WcmCredReadTargetVariants -BaseTarget $base)) {
            try {
                if (Get-Module -ListAvailable -Name CredentialManager) {
                    Import-Module CredentialManager -ErrorAction Stop
                    $c = Get-StoredCredential -Target $tn -ErrorAction SilentlyContinue
                    if ($c) {
                        $pwd = $c.GetNetworkCredential().Password
                        if (-not [string]::IsNullOrWhiteSpace($pwd)) { return $pwd }
                    }
                }
                $obj = _ReadCredentialViaCredRead -Target $tn
                if ($obj -and -not [string]::IsNullOrWhiteSpace($obj.CredentialBlob)) { return $obj.CredentialBlob }
            } catch { }
        }
    }
    return $null
}

function Get-TenantDisplayNameFromWCM {
    param([Parameter(Mandatory = $true)][string]$TenantId, [string]$Prefix = 'EOA')
    $token = Get-GraphAppTokenFromWCM -TenantId $TenantId -Prefix $Prefix
    if (-not $token) { return $null }
    $headers = @{ Authorization = "Bearer $token" }
    $uri = "https://graph.microsoft.com/v1.0/organization?`$top=1&`$select=displayName,verifiedDomains"
    $max = 3
    for ($attempt = 0; $attempt -lt $max; $attempt++) {
        try {
            $resp = Invoke-RestMethod -Uri $uri -Headers $headers -Method Get -ErrorAction Stop
            $vals = $resp.value
            if ($null -eq $vals) { return $null }
            $arr = @($vals)
            if ($arr.Count -eq 0) { return $null }
            $o = $arr[0]
            $name = $null
            if ($o.PSObject.Properties['displayName'] -and $o.displayName) { $name = [string]$o.displayName }
            if (-not $name -and $o.verifiedDomains) {
                $vd = @($o.verifiedDomains | Where-Object { $_.isInitial -eq $true } | Select-Object -First 1)
                if (-not $vd) { $vd = @($o.verifiedDomains | Select-Object -First 1) }
                if ($vd -and $vd.name) { $name = [string]$vd.name }
            }
            if ($name) { return $name }
            return $null
        } catch {
            if ($attempt -lt $max - 1) { Start-Sleep -Milliseconds 750 }
        }
    }
    return $null
}

function Get-WCMTenantListWithNames {
    param([Parameter(Mandatory = $false)][ValidateSet('EOA', 'ESR')][string]$Prefix = 'EOA')
    $result = @()
    $ids = Get-WCMTenantIds -Prefix $Prefix
    $sourceLabel = if ($Prefix -eq 'ESR') { ' (ESR)' } else { '' }
    foreach ($tid in $ids) {
        $name = _Get-StoredDisplayName -TenantId $tid -Prefix $Prefix
        if (-not $name) { $name = Get-TenantDisplayNameFromWCM -TenantId $tid -Prefix $Prefix }
        if (-not $name) {
            $alt = if ($Prefix -eq 'EOA') { 'ESR' } else { 'EOA' }
            $name = Get-TenantDisplayNameFromWCM -TenantId $tid -Prefix $alt
        }
        $displayText = if ($name) { "$name$sourceLabel" } else { "$tid$sourceLabel" }
        $result += [pscustomobject]@{ TenantId = $tid; DisplayName = $name; DisplayText = $displayText; Source = $Prefix }
    }
    return $result | Sort-Object -Property DisplayText
}

function Get-GraphAppTokenFromWCM {
    param([Parameter(Mandatory = $true)][string]$TenantId, [Parameter(Mandatory = $false)][ValidateSet('EOA', 'ESR')][string]$Prefix = 'EOA')
    $cred = Get-GraphAppCredentialFromWCM -TenantId $TenantId -Prefix $Prefix
    if (-not $cred) { return $null }
    $tenantForUrl = ($cred.TenantId -replace '[\{\}]', '').Trim()
    if (-not $tenantForUrl) { $tenantForUrl = ($TenantId -replace '[\{\}]', '').Trim() }
    $tokenUrl = "https://login.microsoftonline.com/$tenantForUrl/oauth2/v2.0/token"
    $body = @{
        client_id     = $cred.ClientId
        client_secret = $cred.ClientSecret
        scope         = 'https://graph.microsoft.com/.default'
        grant_type    = 'client_credentials'
    }
    try {
        $resp = Invoke-RestMethod -Uri $tokenUrl -Method POST -Body $body -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
        return $resp.access_token
    } catch {
        return $null
    }
}

function _ReadCredentialViaCredRead {
    param([string]$Target)
    if (-not $Target) { return $null }
    $sig = @'
[DllImport("Advapi32.dll", EntryPoint = "CredReadW", CharSet = CharSet.Unicode, SetLastError = true)]
public static extern bool CredRead(string target, uint type, int reservedFlag, out IntPtr credentialPtr);

[DllImport("Advapi32.dll", EntryPoint = "CredFree", SetLastError = true)]
public static extern bool CredFree(IntPtr cred);

[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
public struct NativeCredential {
    public uint Flags;
    public uint Type;
    public IntPtr TargetName;
    public IntPtr Comment;
    public long LastWritten;
    public uint CredentialBlobSize;
    public IntPtr CredentialBlob;
    public uint Persist;
    public uint AttributeCount;
    public IntPtr Attributes;
    public IntPtr TargetAlias;
    public IntPtr UserName;
}
'@
    try {
        Add-Type -MemberDefinition $sig -Namespace 'EOACredRead' -Name 'Util' -ErrorAction Stop
    } catch {
        if ($_.Exception.Message -notmatch 'already exists') { return $null }
    }
    $ptr = [IntPtr]::Zero
    $ok = [EOACredRead.Util]::CredRead($Target, 1, 0, [ref]$ptr)
    if (-not $ok -or $ptr -eq [IntPtr]::Zero) { return $null }
    try {
        $ncred = [System.Runtime.InteropServices.Marshal]::PtrToStructure($ptr, [EOACredRead.Util+NativeCredential])
        $userName = if ($ncred.UserName -ne [IntPtr]::Zero) { [System.Runtime.InteropServices.Marshal]::PtrToStringUni($ncred.UserName) } else { $null }
        $blob = $null
        if ($ncred.CredentialBlob -ne [IntPtr]::Zero -and $ncred.CredentialBlobSize -gt 0) {
            $blob = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($ncred.CredentialBlob, [int]$ncred.CredentialBlobSize / 2)
        }
        [EOACredRead.Util]::CredFree($ptr) | Out-Null
        return [pscustomobject]@{ UserName = $userName; CredentialBlob = $blob }
    } catch {
        try { [EOACredRead.Util]::CredFree($ptr) | Out-Null } catch {}
        return $null
    }
}

function Save-GraphAppCredentialToWCM {
    param(
        [Parameter(Mandatory = $true)][string]$TenantId,
        [Parameter(Mandatory = $true)][string]$ClientId,
        [Parameter(Mandatory = $true)][string]$ClientSecret,
        [Parameter(Mandatory = $false)][string]$TenantDisplayName,
        [Parameter(Mandatory = $false)][ValidateSet('EOA', 'ESR')][string]$Prefix = 'ESR'
    )
    $credPrefix = if ($Prefix -eq 'ESR') { $script:credTargetPrefixESR } else { $script:credTargetPrefixEOA }
    $target = "$credPrefix$TenantId"
    $userName = "${TenantId}|${ClientId}"
    if (Get-Module -ListAvailable -Name CredentialManager) {
        try {
            Import-Module CredentialManager -ErrorAction Stop
            $cred = New-Object PSCredential $userName, (ConvertTo-SecureString $ClientSecret -AsPlainText -Force)
            New-StoredCredential -Target $target -Credentials $cred -ErrorAction Stop | Out-Null
        } catch {
            Start-Process -FilePath "cmdkey.exe" -ArgumentList "/generic:$target", "/user:$userName", "/pass:$ClientSecret" -Wait -PassThru -WindowStyle Hidden | Out-Null
        }
    } else {
        Start-Process -FilePath "cmdkey.exe" -ArgumentList "/generic:$target", "/user:$userName", "/pass:$ClientSecret" -Wait -PassThru -WindowStyle Hidden | Out-Null
    }
    if ($TenantDisplayName -and -not [string]::IsNullOrWhiteSpace($TenantDisplayName)) {
        $nameTarget = "${credPrefix}${TenantId}-DisplayName"
        try {
            if (Get-Module -ListAvailable -Name CredentialManager) {
                Import-Module CredentialManager -ErrorAction Stop
                $nameCred = New-Object PSCredential 'DisplayName', (ConvertTo-SecureString $TenantDisplayName -AsPlainText -Force)
                New-StoredCredential -Target $nameTarget -Credentials $nameCred -ErrorAction Stop | Out-Null
            } else {
                Start-Process -FilePath "cmdkey.exe" -ArgumentList "/generic:$nameTarget", "/user:DisplayName", "/pass:$TenantDisplayName" -Wait -PassThru -WindowStyle Hidden | Out-Null
            }
        } catch { }
    }
}

function Remove-GraphAppCredentialFromWCM {
    param([Parameter(Mandatory = $true)][string]$TenantId, [Parameter(Mandatory = $false)][ValidateSet('EOA', 'ESR')][string]$Prefix = 'ESR')
    $credPrefix = if ($Prefix -eq 'ESR') { $script:credTargetPrefixESR } else { $script:credTargetPrefixEOA }
    $target = "$credPrefix$TenantId"
    $nameTarget = "${credPrefix}${TenantId}-DisplayName"
    if (Get-Module -ListAvailable -Name CredentialManager) {
        try {
            Import-Module CredentialManager -ErrorAction Stop
            Remove-StoredCredential -Target $target -ErrorAction SilentlyContinue
            Remove-StoredCredential -Target $nameTarget -ErrorAction SilentlyContinue
            return
        } catch { }
    }
    try {
        Start-Process -FilePath "cmdkey.exe" -ArgumentList "/delete:$target" -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue
        Start-Process -FilePath "cmdkey.exe" -ArgumentList "/delete:$nameTarget" -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue
    } catch { }
}

Export-ModuleMember -Function Get-GraphAppCredentialFromWCM, Get-GraphAppTokenFromWCM, Get-WCMTenantIds, Get-TenantDisplayNameFromWCM, Get-WCMTenantListWithNames, Save-GraphAppCredentialToWCM, Remove-GraphAppCredentialFromWCM
