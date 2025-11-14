<#
.SYNOPSIS
    A GUI-based PowerShell script to manage Microsoft Entra Conditional Access policies.
.NOTES
    Author: Gemini  
    Version: 3.4 (Clean working version with all features)
    Requirements: Windows PowerShell 5.1 with Microsoft.Graph module
#>

# Check PowerShell version first (before loading anything)
# PowerShell Core is now supported
if ($PSVersionTable.PSEdition -eq "Core") {
    Write-Host "Running in PowerShell Core (pwsh) - Windows Forms support enabled" -ForegroundColor Green
    
    # Check if we're on Windows (required for Windows Forms)
    if ($IsWindows -eq $false) {
        Write-Error "This script requires Windows to run. Windows Forms is not available on non-Windows platforms."
        Write-Host "Please run this script on a Windows machine or in Windows Subsystem for Linux (WSL)." -ForegroundColor Yellow
        Read-Host "Press Enter to exit"
        exit
    }
} else {
    Write-Host "Running in Windows PowerShell" -ForegroundColor Green
}

# Load assemblies and configure immediately (must be done before ANY Windows Forms objects are created)
try {
    Write-Host "Loading Windows Forms assemblies..." -ForegroundColor Green
    
    # Load assemblies first - works in both PowerShell Core and Windows PowerShell
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing  
    
    # Only load Microsoft.VisualBasic if available (may not be in PowerShell Core)
    try {
        Add-Type -AssemblyName Microsoft.VisualBasic
        Write-Host "Microsoft.VisualBasic loaded successfully." -ForegroundColor Green
    } catch {
        Write-Host "Microsoft.VisualBasic not available, using custom InputBox implementation." -ForegroundColor Yellow
    }
    
    Write-Host "Configuring Windows Forms rendering..." -ForegroundColor Green
    
    # Try to configure rendering - if it fails, continue anyway
    try {
        [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)
        [System.Windows.Forms.Application]::EnableVisualStyles()
        Write-Host "Windows Forms configured with visual styles." -ForegroundColor Green
    } catch {
        Write-Warning "Could not set text rendering default (forms may already exist). Continuing..."
        [System.Windows.Forms.Application]::EnableVisualStyles()
        Write-Host "Windows Forms configured (basic mode)." -ForegroundColor Green
    }
    
    Write-Host "Windows Forms initialized successfully." -ForegroundColor Green
} catch {
    Write-Error "Failed to initialize Windows Forms: $_"
    Write-Host "This error can occur if Windows Forms objects were already created." -ForegroundColor Yellow
    Write-Host "Try restarting PowerShell and running the script again." -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit
}

# Check for Microsoft Graph module
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-Error "Microsoft.Graph module not installed."
    Write-Host "Please run: Install-Module Microsoft.Graph -Scope CurrentUser" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit
}

# PowerShell Core compatible InputBox function
function Show-InputBox {
    param(
        [string]$Prompt,
        [string]$Title = "Input",
        [string]$DefaultValue = ""
    )
    
    # Try to use native InputBox if Microsoft.VisualBasic is available
    try {
        if ([System.Management.Automation.PSTypeName]'Microsoft.VisualBasic.Interaction').Type) {
            return [Microsoft.VisualBasic.Interaction]::InputBox($Prompt, $Title, $DefaultValue)
        }
    } catch {
        # Fall through to custom implementation
    }
    
    # Custom InputBox implementation for PowerShell Core
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object System.Drawing.Size(400, 150)
    $form.StartPosition = "CenterParent"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Prompt
    $label.Location = New-Object System.Drawing.Point(10, 20)
    $label.Size = New-Object System.Drawing.Size(360, 20)
    $form.Controls.Add($label)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10, 45)
    $textBox.Size = New-Object System.Drawing.Size(360, 20)
    $textBox.Text = $DefaultValue
    $form.Controls.Add($textBox)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Location = New-Object System.Drawing.Point(210, 75)
    $okButton.Size = New-Object System.Drawing.Size(75, 23)
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(295, 75)
    $cancelButton.Size = New-Object System.Drawing.Size(75, 23)
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($cancelButton)

    $form.AcceptButton = $okButton
    $form.CancelButton = $cancelButton

    $result = $form.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $textBox.Text
    } else {
        return $null
    }
}

#region Helper Functions

# Validation Helpers
function Test-GraphConnectionAndSelection {
    param(
        $listView,
        [string]$itemType
    )

    if (-not $global:isConnected) {
        [System.Windows.Forms.MessageBox]::Show("Please connect to Microsoft Graph first.", "Not Connected")
        return $false
    }

    if ($listView.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a $itemType.", "No Selection")
        return $false
    }

    return $true
}

function Test-ResourceExists {
    param(
        [string]$ResourceType,
        [string]$ResourceId
    )

    try {
        switch ($ResourceType) {
            "NamedLocation" {
                $resource = Get-MgIdentityConditionalAccessNamedLocation -NamedLocationId $ResourceId -ErrorAction SilentlyContinue
            }
            "Policy" {
                $resource = Get-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $ResourceId -ErrorAction SilentlyContinue
            }
            default {
                return $false
            }
        }
        return ($null -ne $resource)
    } catch {
        return $false
    }
}

# Data Processing Helpers
function Copy-PolicyConditionArray {
    param(
        $SourceArray,
        [string]$PropertyName
    )

    if ($SourceArray -and $SourceArray.Count -gt 0) {
        $cleaned = @($SourceArray | Where-Object { $_ -ne $null -and $_ -ne "" })
        if ($cleaned.Count -gt 0) {
            return $cleaned
        }
    }
    return $null
}

function Copy-PolicyConditionObject {
    param(
        $SourceObject,
        [string[]]$PropertyNames
    )

    if (-not $SourceObject) {
        return $null
    }

    $result = @{}
    $hasValidProperty = $false

    foreach ($prop in $PropertyNames) {
        $value = $SourceObject.$prop

        if ($value -is [array] -or $value -is [System.Collections.ArrayList]) {
            $cleaned = Copy-PolicyConditionArray -SourceArray $value -PropertyName $prop
            if ($cleaned) {
                $result[$prop] = $cleaned
                $hasValidProperty = $true
            }
        }
        elseif ($null -ne $value -and $value -ne "") {
            $result[$prop] = $value
            $hasValidProperty = $true
        }
    }

    if ($hasValidProperty) {
        return $result
    }
    return $null
}

function Remove-NullValuesOptimized {
    param($obj)

    if ($null -eq $obj) {
        return $null
    }

    if ($obj -is [hashtable]) {
        $cleaned = @{}
        foreach ($key in $obj.Keys) {
            $value = Remove-NullValuesOptimized $obj[$key]
            if ($null -ne $value) {
                $cleaned[$key] = $value
            }
        }
        return ($cleaned.Count -gt 0) ? $cleaned : $null
    }

    if ($obj -is [array] -or $obj -is [System.Collections.ArrayList]) {
        $cleaned = [System.Collections.ArrayList]::new()
        foreach ($item in $obj) {
            $value = Remove-NullValuesOptimized $item
            if ($null -ne $value) {
                [void]$cleaned.Add($value)
            }
        }
        if ($cleaned.Count -gt 0) {
            # Force array return even for single items
            return ,$cleaned.ToArray()
        }
        return $null
    }

    return $obj
}

# API Operation Helpers
function Invoke-GraphOperationWithRetry {
    param(
        [string]$Method,
        [string]$Uri,
        $Body = $null,
        [int]$MaxRetries = 3,
        [int]$InitialDelayMs = 500
    )

    $attempt = 0
    $delay = $InitialDelayMs

    while ($attempt -lt $MaxRetries) {
        try {
            $params = @{
                Method = $Method
                Uri = $Uri
            }

            if ($Body) {
                if ($Body -is [string]) {
                    $params.Body = $Body
                } else {
                    $params.Body = ($Body | ConvertTo-Json -Depth 10)
                }
                $params.ContentType = "application/json"
            }

            $response = Invoke-MgGraphRequest @params
            return $response
        }
        catch {
            $attempt++
            if ($attempt -ge $MaxRetries) {
                throw
            }

            Write-Host "Retry attempt $attempt/$MaxRetries after ${delay}ms..." -ForegroundColor Yellow
            Start-Sleep -Milliseconds $delay
            $delay *= 2  # Exponential backoff
        }
    }
}

function Wait-ForGraphOperation {
    param(
        [string]$ResourceType,
        [string]$ResourceId,
        [int]$MaxAttempts = 10,
        [int]$DelayMs = 500
    )

    for ($i = 0; $i -lt $MaxAttempts; $i++) {
        if ($i -gt 0) {
            Start-Sleep -Milliseconds $DelayMs
        }

        if (Test-ResourceExists -ResourceType $ResourceType -ResourceId $ResourceId) {
            return $true
        }
    }

    return $false
}

# UI Helpers
function Show-ErrorMessage {
    param(
        [string]$Message,
        [string]$Title = "Error",
        [string]$Details = ""
    )

    $fullMessage = $Message
    if ($Details) {
        $fullMessage += "`n`n$Details"
    }

    [System.Windows.Forms.MessageBox]::Show(
        $fullMessage,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
}

function Show-SuccessMessage {
    param(
        [string]$Message,
        [string]$Title = "Success"
    )

    [System.Windows.Forms.MessageBox]::Show(
        $Message,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
}

function Show-WarningMessage {
    param(
        [string]$Message,
        [string]$Title = "Warning"
    )

    [System.Windows.Forms.MessageBox]::Show(
        $Message,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
}

#endregion Helper Functions

# Global Variables
$global:isConnected = $false
$global:tenantId = ""
$global:tenantDisplayName = ""
$script:statusLabel = $null
$script:namedLocationsListView = $null
$script:policiesListView = $null
$script:connectButton = $null
$script:disconnectButton = $null
$script:reconnectButton = $null

# Microsoft Graph Functions
function Connect-GraphAPI {
    param([string]$TenantId)
    
    $requiredScopes = @(
        "Policy.Read.All",
        "Policy.ReadWrite.ConditionalAccess", 
        "User.Read.All",
        "Group.Read.All",
        "Organization.Read.All"
    )

    try {
        if ($TenantId) {
            Connect-MgGraph -Scopes $requiredScopes -TenantId $TenantId
        } else {
            Connect-MgGraph -Scopes $requiredScopes
        }
        
        $token = Get-MgContext
        $global:isConnected = $true
        $global:tenantId = $token.TenantId
        $global:tenantDisplayName = $global:tenantId  # Default fallback
        
        # Simple synchronous attempt to get tenant name
        try {
            $org = Get-MgOrganization -Top 1 -ErrorAction SilentlyContinue
            if ($org -and $org.DisplayName) {
                $global:tenantDisplayName = $org.DisplayName
                Write-Host ("Found tenant name: " + $global:tenantDisplayName) -ForegroundColor Green
            } else {
                Write-Host "Could not retrieve tenant display name, using ID" -ForegroundColor Yellow
            }
        } catch {
            Write-Host ("Error getting tenant name: " + $_.Exception.Message) -ForegroundColor Yellow
        }
        
        Update-ConnectionUI
        return $true
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Failed to connect: $_", "Connection Error")
        return $false
    }
}

function Disconnect-GraphAPI {
    try {
        Disconnect-MgGraph
        $global:isConnected = $false
        $global:tenantId = ""
        $global:tenantDisplayName = ""
        
        # Clear the list views
        if ($script:namedLocationsListView) {
            $script:namedLocationsListView.Items.Clear()
        }
        if ($script:policiesListView) {
            $script:policiesListView.Items.Clear()
        }
        
        Update-ConnectionUI
        [System.Windows.Forms.MessageBox]::Show("Disconnected successfully!", "Disconnected")
        return $true
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error disconnecting: $_", "Disconnect Error")
        return $false
    }
}

function Show-ReconnectDialog {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Reconnect to Tenant"
    $form.Size = New-Object System.Drawing.Size(400, 200)
    $form.StartPosition = "CenterParent"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false

    # Current tenant info
    $currentLabel = New-Object System.Windows.Forms.Label
    if ($global:isConnected) {
        if ($global:tenantDisplayName -and $global:tenantDisplayName -ne $global:tenantId) {
            $currentLabel.Text = "Currently connected to: " + $global:tenantDisplayName
        } else {
            $currentLabel.Text = "Currently connected to: " + $global:tenantId
        }
    } else {
        $currentLabel.Text = "Currently not connected"
    }
    $currentLabel.Location = New-Object System.Drawing.Point(10, 20)
    $currentLabel.Size = New-Object System.Drawing.Size(370, 20)
    $form.Controls.Add($currentLabel)

    # Tenant ID input
    $tenantLabel = New-Object System.Windows.Forms.Label
    $tenantLabel.Text = "Tenant ID (leave blank for default):"
    $tenantLabel.Location = New-Object System.Drawing.Point(10, 60)
    $tenantLabel.Size = New-Object System.Drawing.Size(200, 20)
    $form.Controls.Add($tenantLabel)

    $tenantTextBox = New-Object System.Windows.Forms.TextBox
    $tenantTextBox.Location = New-Object System.Drawing.Point(10, 85)
    $tenantTextBox.Size = New-Object System.Drawing.Size(360, 20)
    $form.Controls.Add($tenantTextBox)

    # Buttons
    $connectButton = New-Object System.Windows.Forms.Button
    $connectButton.Text = "Connect"
    $connectButton.Location = New-Object System.Drawing.Point(210, 120)
    $connectButton.Size = New-Object System.Drawing.Size(75, 23)
    $connectButton.Add_Click({
        # Disconnect first if already connected
        if ($global:isConnected) {
            Disconnect-GraphAPI
        }
        
        $tenantId = $tenantTextBox.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($tenantId)) {
            $tenantId = $null
        }
        
        if (Connect-GraphAPI -TenantId $tenantId) {
            $form.Close()
            # Refresh both lists if connected successfully
            if ($global:isConnected) {
                Refresh-NamedLocationsList $script:namedLocationsListView
                Refresh-PoliciesList $script:policiesListView
            }
        }
    })
    $form.Controls.Add($connectButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(295, 120)
    $cancelButton.Size = New-Object System.Drawing.Size(75, 23)
    $cancelButton.Add_Click({ $form.Close() })
    $form.Controls.Add($cancelButton)

    $form.ShowDialog() | Out-Null
}

function Update-ConnectionUI {
    if ($script:statusLabel) {
        if ($global:isConnected) {
            if ($global:tenantDisplayName -and $global:tenantDisplayName -ne $global:tenantId) {
                $script:statusLabel.Text = "Connected to: " + $global:tenantDisplayName
            } else {
                $script:statusLabel.Text = "Connected to tenant: " + $global:tenantId
            }
            $script:statusLabel.ForeColor = [System.Drawing.Color]::Green
        } else {
            $script:statusLabel.Text = "Not connected"
            $script:statusLabel.ForeColor = [System.Drawing.Color]::Red
        }
    }
    
    if ($script:connectButton) {
        $script:connectButton.Enabled = -not $global:isConnected
    }
    
    if ($script:disconnectButton) {
        $script:disconnectButton.Enabled = $global:isConnected
    }
    
    if ($script:reconnectButton) {
        $script:reconnectButton.Enabled = $true  # Always enabled
    }
}

function Get-UserDisplayInfo {
    param([string[]]$UserIds)
    
    $userInfo = @()
    foreach ($userId in $UserIds) {
        if ($userId -eq "All") {
            $userInfo += "All Users"
        } else {
            try {
                $user = Get-MgUser -UserId $userId -ErrorAction SilentlyContinue
                if ($user) {
                    $userInfo += "$($user.DisplayName) ($($user.UserPrincipalName))"
                } else {
                    $userInfo += "Unknown User [$userId]"
                }
            } catch {
                $userInfo += "Error retrieving user [$userId]"
            }
        }
    }
    return $userInfo
}

function Resolve-UserInput {
    param([string[]]$UserInputs)
    
    $resolvedUserIds = @()
    $notFoundUsers = @()
    
    foreach ($input in $UserInputs) {
        $input = $input.Trim()
        if ([string]::IsNullOrWhiteSpace($input)) { continue }
        
        try {
            # Check if it is a GUID (User ID)
            if ($input -match "^[{]?[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}[}]?$") {
                $user = Get-MgUser -UserId $input -ErrorAction SilentlyContinue
                if ($user) {
                    $resolvedUserIds += $input
                } else {
                    $notFoundUsers += $input
                }
            } else {
                # Try as email or display name
                $user = Get-MgUser -Filter "userPrincipalName eq '$input'" -ErrorAction SilentlyContinue
                
                if (-not $user) {
                    $user = Get-MgUser -Filter "displayName eq '$input'" -ErrorAction SilentlyContinue
                }
                
                if ($user) {
                    $resolvedUserIds += $user.Id
                } else {
                    $notFoundUsers += $input
                }
            }
        } catch {
            $notFoundUsers += $input
        }
    }
    
    return @{
        ResolvedUserIds = $resolvedUserIds
        NotFoundUsers = $notFoundUsers
    }
}

# GUI Functions
function Refresh-NamedLocationsList {
    param($listView)
    
    if (-not $global:isConnected) { 
        [System.Windows.Forms.MessageBox]::Show("Please connect to Microsoft Graph first.", "Not Connected")
        return 
    }
    
    $listView.Items.Clear()
    
    try {
        Write-Host "Refreshing Named Locations list..." -ForegroundColor Yellow
        $locations = Get-MgIdentityConditionalAccessNamedLocation -All -ErrorAction Stop
        
        foreach ($location in $locations) {
            try {
                $item = New-Object System.Windows.Forms.ListViewItem($location.DisplayName)
                $item.SubItems.Add($location.Id) | Out-Null
                
                $odataType = $location.AdditionalProperties.'@odata.type'
                if ($odataType -eq '#microsoft.graph.countryNamedLocation') {
                    $item.SubItems.Add("Country") | Out-Null
                    $countries = $location.AdditionalProperties['countriesAndRegions']
                    $item.SubItems.Add(($countries -join ', ')) | Out-Null
                } elseif ($odataType -eq '#microsoft.graph.ipNamedLocation') {
                    $item.SubItems.Add("IP Range") | Out-Null
                    $item.SubItems.Add("IP Ranges") | Out-Null
                } else {
                    $item.SubItems.Add("Unknown") | Out-Null
                    $item.SubItems.Add("") | Out-Null
                }
                
                $item.Tag = $location
                $listView.Items.Add($item) | Out-Null
            } catch {
                Write-Host ("Error processing location " + $location.DisplayName + ": " + $_.Exception.Message) -ForegroundColor Red
                # Continue with other locations
            }
        }
        Write-Host ("Loaded " + $locations.Count + " Named Locations") -ForegroundColor Green
    } catch {
        Write-Host ("Error loading Named Locations: " + $_.Exception.Message) -ForegroundColor Red
        [System.Windows.Forms.MessageBox]::Show("Error loading Named Locations: $_", "Error")
    }
}

function Show-CountrySelectionDialog {
    param(
        [string[]]$PreselectedCountries = @()
    )
    
    # Complete list of ISO 3166-1 alpha-2 country codes with names
    $countryList = @(
        @{Code='AD'; Name='Andorra'},
        @{Code='AE'; Name='United Arab Emirates'},
        @{Code='AF'; Name='Afghanistan'},
        @{Code='AG'; Name='Antigua and Barbuda'},
        @{Code='AI'; Name='Anguilla'},
        @{Code='AL'; Name='Albania'},
        @{Code='AM'; Name='Armenia'},
        @{Code='AO'; Name='Angola'},
        @{Code='AQ'; Name='Antarctica'},
        @{Code='AR'; Name='Argentina'},
        @{Code='AS'; Name='American Samoa'},
        @{Code='AT'; Name='Austria'},
        @{Code='AU'; Name='Australia'},
        @{Code='AW'; Name='Aruba'},
        @{Code='AX'; Name='Åland Islands'},
        @{Code='AZ'; Name='Azerbaijan'},
        @{Code='BA'; Name='Bosnia and Herzegovina'},
        @{Code='BB'; Name='Barbados'},
        @{Code='BD'; Name='Bangladesh'},
        @{Code='BE'; Name='Belgium'},
        @{Code='BF'; Name='Burkina Faso'},
        @{Code='BG'; Name='Bulgaria'},
        @{Code='BH'; Name='Bahrain'},
        @{Code='BI'; Name='Burundi'},
        @{Code='BJ'; Name='Benin'},
        @{Code='BL'; Name='Saint Barthélemy'},
        @{Code='BM'; Name='Bermuda'},
        @{Code='BN'; Name='Brunei'},
        @{Code='BO'; Name='Bolivia'},
        @{Code='BQ'; Name='Caribbean Netherlands'},
        @{Code='BR'; Name='Brazil'},
        @{Code='BS'; Name='Bahamas'},
        @{Code='BT'; Name='Bhutan'},
        @{Code='BV'; Name='Bouvet Island'},
        @{Code='BW'; Name='Botswana'},
        @{Code='BY'; Name='Belarus'},
        @{Code='BZ'; Name='Belize'},
        @{Code='CA'; Name='Canada'},
        @{Code='CC'; Name='Cocos (Keeling) Islands'},
        @{Code='CD'; Name='Democratic Republic of the Congo'},
        @{Code='CF'; Name='Central African Republic'},
        @{Code='CG'; Name='Republic of the Congo'},
        @{Code='CH'; Name='Switzerland'},
        @{Code='CI'; Name='Côte d''Ivoire'},
        @{Code='CK'; Name='Cook Islands'},
        @{Code='CL'; Name='Chile'},
        @{Code='CM'; Name='Cameroon'},
        @{Code='CN'; Name='China'},
        @{Code='CO'; Name='Colombia'},
        @{Code='CR'; Name='Costa Rica'},
        @{Code='CU'; Name='Cuba'},
        @{Code='CV'; Name='Cape Verde'},
        @{Code='CW'; Name='Curaçao'},
        @{Code='CX'; Name='Christmas Island'},
        @{Code='CY'; Name='Cyprus'},
        @{Code='CZ'; Name='Czech Republic'},
        @{Code='DE'; Name='Germany'},
        @{Code='DJ'; Name='Djibouti'},
        @{Code='DK'; Name='Denmark'},
        @{Code='DM'; Name='Dominica'},
        @{Code='DO'; Name='Dominican Republic'},
        @{Code='DZ'; Name='Algeria'},
        @{Code='EC'; Name='Ecuador'},
        @{Code='EE'; Name='Estonia'},
        @{Code='EG'; Name='Egypt'},
        @{Code='EH'; Name='Western Sahara'},
        @{Code='ER'; Name='Eritrea'},
        @{Code='ES'; Name='Spain'},
        @{Code='ET'; Name='Ethiopia'},
        @{Code='FI'; Name='Finland'},
        @{Code='FJ'; Name='Fiji'},
        @{Code='FK'; Name='Falkland Islands'},
        @{Code='FM'; Name='Micronesia'},
        @{Code='FO'; Name='Faroe Islands'},
        @{Code='FR'; Name='France'},
        @{Code='GA'; Name='Gabon'},
        @{Code='GB'; Name='United Kingdom'},
        @{Code='GD'; Name='Grenada'},
        @{Code='GE'; Name='Georgia'},
        @{Code='GF'; Name='French Guiana'},
        @{Code='GG'; Name='Guernsey'},
        @{Code='GH'; Name='Ghana'},
        @{Code='GI'; Name='Gibraltar'},
        @{Code='GL'; Name='Greenland'},
        @{Code='GM'; Name='Gambia'},
        @{Code='GN'; Name='Guinea'},
        @{Code='GP'; Name='Guadeloupe'},
        @{Code='GQ'; Name='Equatorial Guinea'},
        @{Code='GR'; Name='Greece'},
        @{Code='GS'; Name='South Georgia and the South Sandwich Islands'},
        @{Code='GT'; Name='Guatemala'},
        @{Code='GU'; Name='Guam'},
        @{Code='GW'; Name='Guinea-Bissau'},
        @{Code='GY'; Name='Guyana'},
        @{Code='HK'; Name='Hong Kong'},
        @{Code='HM'; Name='Heard Island and McDonald Islands'},
        @{Code='HN'; Name='Honduras'},
        @{Code='HR'; Name='Croatia'},
        @{Code='HT'; Name='Haiti'},
        @{Code='HU'; Name='Hungary'},
        @{Code='ID'; Name='Indonesia'},
        @{Code='IE'; Name='Ireland'},
        @{Code='IL'; Name='Israel'},
        @{Code='IM'; Name='Isle of Man'},
        @{Code='IN'; Name='India'},
        @{Code='IO'; Name='British Indian Ocean Territory'},
        @{Code='IQ'; Name='Iraq'},
        @{Code='IR'; Name='Iran'},
        @{Code='IS'; Name='Iceland'},
        @{Code='IT'; Name='Italy'},
        @{Code='JE'; Name='Jersey'},
        @{Code='JM'; Name='Jamaica'},
        @{Code='JO'; Name='Jordan'},
        @{Code='JP'; Name='Japan'},
        @{Code='KE'; Name='Kenya'},
        @{Code='KG'; Name='Kyrgyzstan'},
        @{Code='KH'; Name='Cambodia'},
        @{Code='KI'; Name='Kiribati'},
        @{Code='KM'; Name='Comoros'},
        @{Code='KN'; Name='Saint Kitts and Nevis'},
        @{Code='KP'; Name='North Korea'},
        @{Code='KR'; Name='South Korea'},
        @{Code='KW'; Name='Kuwait'},
        @{Code='KY'; Name='Cayman Islands'},
        @{Code='KZ'; Name='Kazakhstan'},
        @{Code='LA'; Name='Laos'},
        @{Code='LB'; Name='Lebanon'},
        @{Code='LC'; Name='Saint Lucia'},
        @{Code='LI'; Name='Liechtenstein'},
        @{Code='LK'; Name='Sri Lanka'},
        @{Code='LR'; Name='Liberia'},
        @{Code='LS'; Name='Lesotho'},
        @{Code='LT'; Name='Lithuania'},
        @{Code='LU'; Name='Luxembourg'},
        @{Code='LV'; Name='Latvia'},
        @{Code='LY'; Name='Libya'},
        @{Code='MA'; Name='Morocco'},
        @{Code='MC'; Name='Monaco'},
        @{Code='MD'; Name='Moldova'},
        @{Code='ME'; Name='Montenegro'},
        @{Code='MF'; Name='Saint Martin'},
        @{Code='MG'; Name='Madagascar'},
        @{Code='MH'; Name='Marshall Islands'},
        @{Code='MK'; Name='North Macedonia'},
        @{Code='ML'; Name='Mali'},
        @{Code='MM'; Name='Myanmar'},
        @{Code='MN'; Name='Mongolia'},
        @{Code='MO'; Name='Macao'},
        @{Code='MP'; Name='Northern Mariana Islands'},
        @{Code='MQ'; Name='Martinique'},
        @{Code='MR'; Name='Mauritania'},
        @{Code='MS'; Name='Montserrat'},
        @{Code='MT'; Name='Malta'},
        @{Code='MU'; Name='Mauritius'},
        @{Code='MV'; Name='Maldives'},
        @{Code='MW'; Name='Malawi'},
        @{Code='MX'; Name='Mexico'},
        @{Code='MY'; Name='Malaysia'},
        @{Code='MZ'; Name='Mozambique'},
        @{Code='NA'; Name='Namibia'},
        @{Code='NC'; Name='New Caledonia'},
        @{Code='NE'; Name='Niger'},
        @{Code='NF'; Name='Norfolk Island'},
        @{Code='NG'; Name='Nigeria'},
        @{Code='NI'; Name='Nicaragua'},
        @{Code='NL'; Name='Netherlands'},
        @{Code='NO'; Name='Norway'},
        @{Code='NP'; Name='Nepal'},
        @{Code='NR'; Name='Nauru'},
        @{Code='NU'; Name='Niue'},
        @{Code='NZ'; Name='New Zealand'},
        @{Code='OM'; Name='Oman'},
        @{Code='PA'; Name='Panama'},
        @{Code='PE'; Name='Peru'},
        @{Code='PF'; Name='French Polynesia'},
        @{Code='PG'; Name='Papua New Guinea'},
        @{Code='PH'; Name='Philippines'},
        @{Code='PK'; Name='Pakistan'},
        @{Code='PL'; Name='Poland'},
        @{Code='PM'; Name='Saint Pierre and Miquelon'},
        @{Code='PN'; Name='Pitcairn Islands'},
        @{Code='PR'; Name='Puerto Rico'},
        @{Code='PS'; Name='Palestine'},
        @{Code='PT'; Name='Portugal'},
        @{Code='PW'; Name='Palau'},
        @{Code='PY'; Name='Paraguay'},
        @{Code='QA'; Name='Qatar'},
        @{Code='RE'; Name='Réunion'},
        @{Code='RO'; Name='Romania'},
        @{Code='RS'; Name='Serbia'},
        @{Code='RU'; Name='Russia'},
        @{Code='RW'; Name='Rwanda'},
        @{Code='SA'; Name='Saudi Arabia'},
        @{Code='SB'; Name='Solomon Islands'},
        @{Code='SC'; Name='Seychelles'},
        @{Code='SD'; Name='Sudan'},
        @{Code='SE'; Name='Sweden'},
        @{Code='SG'; Name='Singapore'},
        @{Code='SH'; Name='Saint Helena'},
        @{Code='SI'; Name='Slovenia'},
        @{Code='SJ'; Name='Svalbard and Jan Mayen'},
        @{Code='SK'; Name='Slovakia'},
        @{Code='SL'; Name='Sierra Leone'},
        @{Code='SM'; Name='San Marino'},
        @{Code='SN'; Name='Senegal'},
        @{Code='SO'; Name='Somalia'},
        @{Code='SR'; Name='Suriname'},
        @{Code='SS'; Name='South Sudan'},
        @{Code='ST'; Name='São Tomé and Príncipe'},
        @{Code='SV'; Name='El Salvador'},
        @{Code='SX'; Name='Sint Maarten'},
        @{Code='SY'; Name='Syria'},
        @{Code='SZ'; Name='Eswatini'},
        @{Code='TC'; Name='Turks and Caicos Islands'},
        @{Code='TD'; Name='Chad'},
        @{Code='TF'; Name='French Southern Territories'},
        @{Code='TG'; Name='Togo'},
        @{Code='TH'; Name='Thailand'},
        @{Code='TJ'; Name='Tajikistan'},
        @{Code='TK'; Name='Tokelau'},
        @{Code='TL'; Name='Timor-Leste'},
        @{Code='TM'; Name='Turkmenistan'},
        @{Code='TN'; Name='Tunisia'},
        @{Code='TO'; Name='Tonga'},
        @{Code='TR'; Name='Turkey'},
        @{Code='TT'; Name='Trinidad and Tobago'},
        @{Code='TV'; Name='Tuvalu'},
        @{Code='TW'; Name='Taiwan'},
        @{Code='TZ'; Name='Tanzania'},
        @{Code='UA'; Name='Ukraine'},
        @{Code='UG'; Name='Uganda'},
        @{Code='UM'; Name='United States Minor Outlying Islands'},
        @{Code='US'; Name='United States'},
        @{Code='UY'; Name='Uruguay'},
        @{Code='UZ'; Name='Uzbekistan'},
        @{Code='VA'; Name='Vatican City'},
        @{Code='VC'; Name='Saint Vincent and the Grenadines'},
        @{Code='VE'; Name='Venezuela'},
        @{Code='VG'; Name='British Virgin Islands'},
        @{Code='VI'; Name='United States Virgin Islands'},
        @{Code='VN'; Name='Vietnam'},
        @{Code='VU'; Name='Vanuatu'},
        @{Code='WF'; Name='Wallis and Futuna'},
        @{Code='WS'; Name='Samoa'},
        @{Code='YE'; Name='Yemen'},
        @{Code='YT'; Name='Mayotte'},
        @{Code='ZA'; Name='South Africa'},
        @{Code='ZM'; Name='Zambia'},
        @{Code='ZW'; Name='Zimbabwe'}
    )
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select Countries"
    $form.Size = New-Object System.Drawing.Size(600, 500)
    $form.StartPosition = "CenterParent"
    $form.FormBorderStyle = "Sizable"

    # Search box
    $searchLabel = New-Object System.Windows.Forms.Label
    $searchLabel.Text = "Search countries:"
    $searchLabel.Location = New-Object System.Drawing.Point(10, 15)
    $searchLabel.Size = New-Object System.Drawing.Size(120, 20)
    $form.Controls.Add($searchLabel)

    $searchBox = New-Object System.Windows.Forms.TextBox
    $searchBox.Location = New-Object System.Drawing.Point(130, 12)
    $searchBox.Size = New-Object System.Drawing.Size(200, 20)
    $form.Controls.Add($searchBox)

    # Selection info
    $selectionLabel = New-Object System.Windows.Forms.Label
    $selectionLabel.Text = "0 countries selected"
    $selectionLabel.Location = New-Object System.Drawing.Point(350, 15)
    $selectionLabel.Size = New-Object System.Drawing.Size(200, 20)
    $selectionLabel.ForeColor = [System.Drawing.Color]::Blue
    $form.Controls.Add($selectionLabel)

    # Quick select buttons
    $selectAllButton = New-Object System.Windows.Forms.Button
    $selectAllButton.Text = "Select All"
    $selectAllButton.Location = New-Object System.Drawing.Point(10, 45)
    $selectAllButton.Size = New-Object System.Drawing.Size(80, 25)
    $form.Controls.Add($selectAllButton)

    $selectNoneButton = New-Object System.Windows.Forms.Button
    $selectNoneButton.Text = "Select None"
    $selectNoneButton.Location = New-Object System.Drawing.Point(100, 45)
    $selectNoneButton.Size = New-Object System.Drawing.Size(80, 25)
    $form.Controls.Add($selectNoneButton)

    # Countries list with checkboxes
    $countriesListBox = New-Object System.Windows.Forms.CheckedListBox
    $countriesListBox.Location = New-Object System.Drawing.Point(10, 80)
    $countriesListBox.Size = New-Object System.Drawing.Size(560, 320)
    $countriesListBox.CheckOnClick = $true
    $form.Controls.Add($countriesListBox)

    # Populate the list
    function Refresh-CountryList {
        param([string]$SearchTerm = "")
        
        $countriesListBox.Items.Clear()
        $filteredCountries = $countryList
        
        if (-not [string]::IsNullOrWhiteSpace($SearchTerm)) {
            $filteredCountries = $countryList | Where-Object { 
                $_.Name -like "*$SearchTerm*" -or $_.Code -like "*$SearchTerm*" 
            }
        }
        
        foreach ($country in $filteredCountries | Sort-Object Name) {
            $displayText = "$($country.Code) - $($country.Name)"
            $index = $countriesListBox.Items.Add($displayText)
            
            # Check if this country was preselected
            if ($PreselectedCountries -contains $country.Code) {
                $countriesListBox.SetItemChecked($index, $true)
            }
        }
        
        Update-SelectionCount
    }

    function Update-SelectionCount {
        $selectedCount = $countriesListBox.CheckedItems.Count
        $selectionLabel.Text = "$selectedCount countries selected"
    }

    # Event handlers
    $searchBox.Add_TextChanged({
        Refresh-CountryList -SearchTerm $searchBox.Text
    })

    $selectAllButton.Add_Click({
        for ($i = 0; $i -lt $countriesListBox.Items.Count; $i++) {
            $countriesListBox.SetItemChecked($i, $true)
        }
        Update-SelectionCount
    })

    $selectNoneButton.Add_Click({
        for ($i = 0; $i -lt $countriesListBox.Items.Count; $i++) {
            $countriesListBox.SetItemChecked($i, $false)
        }
        Update-SelectionCount
    })

    $countriesListBox.Add_ItemCheck({
        # Use a timer to update count after the check state changes
        $timer = New-Object System.Windows.Forms.Timer
        $timer.Interval = 10
        $timer.Add_Tick({
            Update-SelectionCount
            $timer.Stop()
            $timer.Dispose()
        })
        $timer.Start()
    })

    # OK and Cancel buttons
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Location = New-Object System.Drawing.Point(420, 420)
    $okButton.Size = New-Object System.Drawing.Size(75, 30)
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(505, 420)
    $cancelButton.Size = New-Object System.Drawing.Size(75, 30)
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($cancelButton)

    # Initialize the list
    Refresh-CountryList

    # Show dialog and return selected countries
    $result = $form.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $selectedCountries = @()
        foreach ($item in $countriesListBox.CheckedItems) {
            # Extract country code (first 2 characters before the hyphen)
            $countryCode = $item.ToString().Substring(0, 2)
            $selectedCountries += $countryCode
        }
        return $selectedCountries
    } else {
        return $null
    }
}

function Show-CountryLocationDialog {
    param(
        $listView,
        [string]$Mode = "Create",  # "Create" or "Edit"
        $ExistingLocation = $null
    )
    
    if (-not $global:isConnected) {
        [System.Windows.Forms.MessageBox]::Show("Please connect to Microsoft Graph first.", "Not Connected")
        return
    }
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Mode + " Country Named Location"
    $form.Size = New-Object System.Drawing.Size(500, 350)
    $form.StartPosition = "CenterParent"
    $form.FormBorderStyle = "FixedDialog"

    # Name
    $nameLabel = New-Object System.Windows.Forms.Label
    $nameLabel.Text = "Display Name:"
    $nameLabel.Location = New-Object System.Drawing.Point(10, 20)
    $nameLabel.Size = New-Object System.Drawing.Size(100, 20)
    $form.Controls.Add($nameLabel)

    $nameTextBox = New-Object System.Windows.Forms.TextBox
    $nameTextBox.Location = New-Object System.Drawing.Point(120, 18)
    $nameTextBox.Size = New-Object System.Drawing.Size(350, 20)
    $form.Controls.Add($nameTextBox)

    # Countries
    $countriesLabel = New-Object System.Windows.Forms.Label
    $countriesLabel.Text = "Country Codes:"
    $countriesLabel.Location = New-Object System.Drawing.Point(10, 60)
    $countriesLabel.Size = New-Object System.Drawing.Size(100, 20)
    $form.Controls.Add($countriesLabel)

    $countriesTextBox = New-Object System.Windows.Forms.TextBox
    $countriesTextBox.Location = New-Object System.Drawing.Point(10, 85)
    $countriesTextBox.Size = New-Object System.Drawing.Size(350, 20)
    $countriesTextBox.ReadOnly = $true
    $form.Controls.Add($countriesTextBox)

    # Select Countries Button
    $selectCountriesButton = New-Object System.Windows.Forms.Button
    $selectCountriesButton.Text = "Select Countries"
    $selectCountriesButton.Location = New-Object System.Drawing.Point(370, 83)
    $selectCountriesButton.Size = New-Object System.Drawing.Size(100, 25)
    $selectCountriesButton.Add_Click({
        # Get current countries from textbox
        $currentCountries = @()
        if (-not [string]::IsNullOrWhiteSpace($countriesTextBox.Text)) {
            $currentCountries = $countriesTextBox.Text.Split(',') | ForEach-Object { $_.Trim().ToUpper() }
        }
        
        # Open country selection dialog
        $selectedCountries = Show-CountrySelectionDialog -PreselectedCountries $currentCountries
        
        if ($selectedCountries) {
            $countriesTextBox.Text = ($selectedCountries -join ', ')
        }
    })
    $form.Controls.Add($selectCountriesButton)

    # Help text for country selection
    $helpLabel = New-Object System.Windows.Forms.Label
    $helpLabel.Text = "Click 'Select Countries' to choose from a list of all countries"
    $helpLabel.Location = New-Object System.Drawing.Point(10, 110)
    $helpLabel.Size = New-Object System.Drawing.Size(460, 15)
    $helpLabel.ForeColor = [System.Drawing.Color]::Gray
    $helpLabel.Font = New-Object System.Drawing.Font($helpLabel.Font.FontFamily, 8)
    $form.Controls.Add($helpLabel)

    # Include Unknown
    $includeUnknownCheckBox = New-Object System.Windows.Forms.CheckBox
    $includeUnknownCheckBox.Text = "Include unknown/future countries"
    $includeUnknownCheckBox.Location = New-Object System.Drawing.Point(10, 135)
    $includeUnknownCheckBox.Size = New-Object System.Drawing.Size(250, 20)
    $form.Controls.Add($includeUnknownCheckBox)

    # Pre-populate if editing
    if ($ExistingLocation -and $Mode -eq "Edit") {
        $nameTextBox.Text = $ExistingLocation.DisplayName
        
        $countries = $ExistingLocation.AdditionalProperties['countriesAndRegions']
        if ($countries) {
            $countriesTextBox.Text = ($countries -join ', ')
        }
        
        $includeUnknown = $ExistingLocation.AdditionalProperties['includeUnknownCountriesAndRegions']
        if ($includeUnknown) {
            $includeUnknownCheckBox.Checked = $includeUnknown
        }
    }

    # Buttons
    $actionButton = New-Object System.Windows.Forms.Button
    $actionButton.Text = $Mode
    $actionButton.Location = New-Object System.Drawing.Point(310, 215)
    $actionButton.Size = New-Object System.Drawing.Size(75, 23)
    $actionButton.Add_Click({
        if ([string]::IsNullOrWhiteSpace($nameTextBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter a display name.", "Validation Error")
            return
        }
        if ([string]::IsNullOrWhiteSpace($countriesTextBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please select at least one country.", "Validation Error")
            return
        }

        try {
            $countryCodes = $countriesTextBox.Text.Split(',') | ForEach-Object { $_.Trim().ToUpper() }
            $success = $false
            
            if ($Mode -eq "Edit") {
                # Update existing location using PATCH
                $updateBody = @{
                    displayName = $nameTextBox.Text
                    countriesAndRegions = $countryCodes
                    includeUnknownCountriesAndRegions = $includeUnknownCheckBox.Checked
                } | ConvertTo-Json -Depth 10
                
                Write-Host ("Updating location " + $ExistingLocation.Id + " with: " + $updateBody) -ForegroundColor Cyan
                $uri = "https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations/" + $ExistingLocation.Id
                $response = Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $updateBody -ContentType "application/json"
                Write-Host ("Update response: " + ($response | ConvertTo-Json -Depth 3)) -ForegroundColor Green
                $success = $true
                [System.Windows.Forms.MessageBox]::Show("Named Location updated successfully!", "Success")
            } else {
                # Create new location using REST API
                $createParams = @{
                    "odata.type" = "#microsoft.graph.countryNamedLocation"
                    displayName = $nameTextBox.Text
                    countriesAndRegions = $countryCodes
                    includeUnknownCountriesAndRegions = $includeUnknownCheckBox.Checked
                }
                
                $jsonBody = $createParams | ConvertTo-Json -Depth 10
                Write-Host ("Creating location with: " + $jsonBody) -ForegroundColor Cyan
                
                $uri = "https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations"
                $response = Invoke-MgGraphRequest -Method POST -Uri $uri -Body $jsonBody -ContentType "application/json"
                Write-Host ("Create response: " + ($response | ConvertTo-Json -Depth 3)) -ForegroundColor Green
                $success = $true
                [System.Windows.Forms.MessageBox]::Show("Named Location created successfully!", "Success")
            }
            
            if ($success) {
                $form.Close()
                # Small delay to allow Microsoft Graph to process the changes
                Start-Sleep -Seconds 1
                Refresh-NamedLocationsList $listView
            }
        } catch {
            $errorMessage = "Error processing Named Location:`n`n"
            $errorMessage += "Error: " + $_.Exception.Message + "`n`n"
            $errorMessage += "Settings:`n"
            $errorMessage += "- Name: " + $nameTextBox.Text + "`n"
            $errorMessage += "- Countries: " + ($countryCodes -join ', ') + "`n"
            $errorMessage += "- Include Unknown: " + $includeUnknownCheckBox.Checked + "`n`n"
            
            if ($_.Exception.Message -like "*BadRequest*") {
                $errorMessage += "Common fixes:`n"
                $errorMessage += "* Use valid 2-letter country codes: US, CA, GB, DE, FR`n"
                $errorMessage += "* Remove special characters from display name`n"
                $errorMessage += "* Ensure you have the required permissions"
            }
            
            Write-Host ("ERROR: " + $errorMessage) -ForegroundColor Red
            [System.Windows.Forms.MessageBox]::Show($errorMessage, "Error Details")
        }
    })
    $form.Controls.Add($actionButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(395, 215)
    $cancelButton.Size = New-Object System.Drawing.Size(75, 23)
    $cancelButton.Add_Click({ $form.Close() })
    $form.Controls.Add($cancelButton)

    $form.ShowDialog() | Out-Null
}

function Show-CreateCountryLocationDialog {
    param($listView)
    Show-CountryLocationDialog -listView $listView -Mode "Create"
}

function Show-IpLocationDialog {
    param(
        $listView,
        [string]$Mode = "Create",  # "Create" or "Edit"
        $ExistingLocation = $null
    )

    if (-not $global:isConnected) {
        [System.Windows.Forms.MessageBox]::Show("Please connect to Microsoft Graph first.", "Not Connected")
        return
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Mode + " IP Named Location"
    $form.Size = New-Object System.Drawing.Size(550, 400)
    $form.StartPosition = "CenterParent"
    $form.FormBorderStyle = "FixedDialog"

    # Name
    $nameLabel = New-Object System.Windows.Forms.Label
    $nameLabel.Text = "Display Name:"
    $nameLabel.Location = New-Object System.Drawing.Point(10, 20)
    $nameLabel.Size = New-Object System.Drawing.Size(100, 20)
    $form.Controls.Add($nameLabel)

    $nameTextBox = New-Object System.Windows.Forms.TextBox
    $nameTextBox.Location = New-Object System.Drawing.Point(120, 18)
    $nameTextBox.Size = New-Object System.Drawing.Size(400, 20)
    $form.Controls.Add($nameTextBox)

    # IP Ranges
    $ipRangesLabel = New-Object System.Windows.Forms.Label
    $ipRangesLabel.Text = "IP Ranges:"
    $ipRangesLabel.Location = New-Object System.Drawing.Point(10, 50)
    $ipRangesLabel.Size = New-Object System.Drawing.Size(100, 20)
    $form.Controls.Add($ipRangesLabel)

    $ipRangesTextBox = New-Object System.Windows.Forms.TextBox
    $ipRangesTextBox.Location = New-Object System.Drawing.Point(10, 75)
    $ipRangesTextBox.Size = New-Object System.Drawing.Size(510, 100)
    $ipRangesTextBox.Multiline = $true
    $ipRangesTextBox.ScrollBars = "Vertical"
    $form.Controls.Add($ipRangesTextBox)

    # Help text
    $helpLabel = New-Object System.Windows.Forms.Label
    $helpLabel.Text = "Enter IP ranges (one per line). Formats: 192.168.1.1/32, 10.0.0.0/8, 2001:db8::/32"
    $helpLabel.Location = New-Object System.Drawing.Point(10, 180)
    $helpLabel.Size = New-Object System.Drawing.Size(510, 20)
    $helpLabel.ForeColor = [System.Drawing.Color]::Gray
    $helpLabel.Font = New-Object System.Drawing.Font($helpLabel.Font.FontFamily, 8)
    $form.Controls.Add($helpLabel)

    # Is Trusted
    $isTrustedCheckBox = New-Object System.Windows.Forms.CheckBox
    $isTrustedCheckBox.Text = "Trusted location (skip MFA when accessing from these IPs)"
    $isTrustedCheckBox.Location = New-Object System.Drawing.Point(10, 210)
    $isTrustedCheckBox.Size = New-Object System.Drawing.Size(400, 20)
    $form.Controls.Add($isTrustedCheckBox)

    # Pre-populate if editing
    if ($ExistingLocation -and $Mode -eq "Edit") {
        $nameTextBox.Text = $ExistingLocation.DisplayName

        $ipRanges = $ExistingLocation.AdditionalProperties['ipRanges']
        if ($ipRanges) {
            $ipRangesText = ""
            foreach ($range in $ipRanges) {
                if ($range.cidrAddress) {
                    $ipRangesText += $range.cidrAddress + "`r`n"
                }
            }
            $ipRangesTextBox.Text = $ipRangesText.TrimEnd()
        }

        $isTrusted = $ExistingLocation.AdditionalProperties['isTrusted']
        if ($isTrusted) {
            $isTrustedCheckBox.Checked = $isTrusted
        }
    }

    # Buttons
    $actionButton = New-Object System.Windows.Forms.Button
    $actionButton.Text = $Mode
    $actionButton.Location = New-Object System.Drawing.Point(360, 250)
    $actionButton.Size = New-Object System.Drawing.Size(75, 23)
    $actionButton.Add_Click({
        if ([string]::IsNullOrWhiteSpace($nameTextBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter a display name.", "Validation Error")
            return
        }
        if ([string]::IsNullOrWhiteSpace($ipRangesTextBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter at least one IP range.", "Validation Error")
            return
        }

        try {
            # Parse IP ranges
            $ipRangeLines = $ipRangesTextBox.Text -split "`n" | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
            $ipRangesArray = @()

            foreach ($line in $ipRangeLines) {
                $ipRangesArray += @{ cidrAddress = $line }
            }

            $success = $false

            if ($Mode -eq "Edit") {
                # Update existing location using PATCH
                $updateBody = @{
                    displayName = $nameTextBox.Text
                    ipRanges = $ipRangesArray
                    isTrusted = $isTrustedCheckBox.Checked
                } | ConvertTo-Json -Depth 10

                Write-Host ("Updating IP location " + $ExistingLocation.Id + " with: " + $updateBody) -ForegroundColor Cyan
                $uri = "https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations/" + $ExistingLocation.Id
                $response = Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $updateBody -ContentType "application/json"
                Write-Host ("Update response: " + ($response | ConvertTo-Json -Depth 3)) -ForegroundColor Green
                $success = $true
                Show-SuccessMessage "IP Named Location updated successfully!"
            } else {
                # Create new location using REST API
                $createParams = @{
                    "@odata.type" = "#microsoft.graph.ipNamedLocation"
                    displayName = $nameTextBox.Text
                    ipRanges = $ipRangesArray
                    isTrusted = $isTrustedCheckBox.Checked
                }

                $jsonBody = $createParams | ConvertTo-Json -Depth 10
                Write-Host ("Creating IP location with: " + $jsonBody) -ForegroundColor Cyan

                $uri = "https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations"
                $response = Invoke-MgGraphRequest -Method POST -Uri $uri -Body $jsonBody -ContentType "application/json"
                Write-Host ("Create response: " + ($response | ConvertTo-Json -Depth 3)) -ForegroundColor Green
                $success = $true
                Show-SuccessMessage "IP Named Location created successfully!"
            }

            if ($success) {
                $form.Close()
                # Wait for operation to complete
                if (Wait-ForGraphOperation -ResourceType "NamedLocation" -ResourceId $response.id -MaxAttempts 10) {
                    Refresh-NamedLocationsList $listView
                } else {
                    Start-Sleep -Seconds 1
                    Refresh-NamedLocationsList $listView
                }
            }
        } catch {
            $errorMessage = "Error processing IP Named Location:`n`n"
            $errorMessage += "Error: " + $_.Exception.Message + "`n`n"
            $errorMessage += "Settings:`n"
            $errorMessage += "- Name: " + $nameTextBox.Text + "`n"
            $errorMessage += "- IP Ranges: " + ($ipRangeLines -join ', ') + "`n"
            $errorMessage += "- Is Trusted: " + $isTrustedCheckBox.Checked + "`n`n"

            if ($_.Exception.Message -like "*BadRequest*") {
                $errorMessage += "Common fixes:`n"
                $errorMessage += "* Use valid CIDR notation: 192.168.1.0/24, 10.0.0.0/8`n"
                $errorMessage += "* For single IP use /32: 192.168.1.1/32`n"
                $errorMessage += "* IPv6 is supported: 2001:db8::/32`n"
                $errorMessage += "* Ensure you have the required permissions"
            }

            Write-Host ("ERROR: " + $errorMessage) -ForegroundColor Red
            Show-ErrorMessage $errorMessage "Error Details"
        }
    })
    $form.Controls.Add($actionButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(445, 250)
    $cancelButton.Size = New-Object System.Drawing.Size(75, 23)
    $cancelButton.Add_Click({ $form.Close() })
    $form.Controls.Add($cancelButton)

    $form.ShowDialog() | Out-Null
}

function Show-CreateIpLocationDialog {
    param($listView)
    Show-IpLocationDialog -listView $listView -Mode "Create"
}

function Edit-SelectedNamedLocation {
    param($listView)

    if (-not (Test-GraphConnectionAndSelection -listView $listView -itemType "Named Location to edit")) {
        return
    }

    $selectedItem = $listView.SelectedItems[0]
    $location = $selectedItem.Tag

    # Check the location type and open the appropriate dialog
    $odataType = $location.AdditionalProperties.'@odata.type'

    switch ($odataType) {
        '#microsoft.graph.countryNamedLocation' {
            Show-CountryLocationDialog -listView $listView -Mode "Edit" -ExistingLocation $location
        }
        '#microsoft.graph.ipNamedLocation' {
            Show-IpLocationDialog -listView $listView -Mode "Edit" -ExistingLocation $location
        }
        default {
            Show-ErrorMessage "Unknown Named Location type: $odataType" "Not Supported"
        }
    }
}

function Copy-SelectedNamedLocation {
    param($listView)

    if (-not (Test-GraphConnectionAndSelection -listView $listView -itemType "Named Location to copy")) {
        return
    }

    $selectedItem = $listView.SelectedItems[0]
    $location = $selectedItem.Tag
    $odataType = $location.AdditionalProperties.'@odata.type'

    # Prompt for new name
    $defaultName = "Copy of " + $location.DisplayName
    $newName = Show-InputBox -Prompt "Enter a name for the new Named Location:" -Title "Copy Named Location" -DefaultValue $defaultName

    if ([string]::IsNullOrWhiteSpace($newName)) {
        return
    }

    try {
        $createParams = @{
            "@odata.type" = $odataType
            displayName = $newName
        }

        # Copy settings based on location type
        switch ($odataType) {
            '#microsoft.graph.countryNamedLocation' {
                Write-Host "Copying country-based Named Location..." -ForegroundColor Cyan
                $createParams['countriesAndRegions'] = $location.AdditionalProperties['countriesAndRegions']
                $createParams['includeUnknownCountriesAndRegions'] = [bool]$location.AdditionalProperties['includeUnknownCountriesAndRegions']
            }
            '#microsoft.graph.ipNamedLocation' {
                Write-Host "Copying IP-based Named Location..." -ForegroundColor Cyan
                $createParams['ipRanges'] = $location.AdditionalProperties['ipRanges']
                $createParams['isTrusted'] = [bool]$location.AdditionalProperties['isTrusted']
            }
            default {
                Show-ErrorMessage "Unknown Named Location type: $odataType" "Not Supported"
                return
            }
        }

        # Use REST API for consistent behavior
        $uri = "https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations"
        $response = Invoke-GraphOperationWithRetry -Method POST -Uri $uri -Body $createParams

        Write-Host "Named Location copied successfully!" -ForegroundColor Green
        Show-SuccessMessage "Named Location copied successfully as '$newName'!"

        # Wait for operation to complete
        if (Wait-ForGraphOperation -ResourceType "NamedLocation" -ResourceId $response.id) {
            Refresh-NamedLocationsList $listView
        } else {
            Start-Sleep -Seconds 1
            Refresh-NamedLocationsList $listView
        }
    } catch {
        $errorMessage = "Error copying Named Location:`n`n"
        $errorMessage += "Error: " + $_.Exception.Message + "`n`n"
        $errorMessage += "Source: " + $location.DisplayName + "`n"
        $errorMessage += "Target Name: " + $newName

        Write-Host ("Copy error: " + $errorMessage) -ForegroundColor Red
        Show-ErrorMessage $errorMessage "Copy Error"
    }
}

function Rename-SelectedNamedLocation {
    param($listView)

    if (-not (Test-GraphConnectionAndSelection -listView $listView -itemType "Named Location to rename")) {
        return
    }

    $selectedItem = $listView.SelectedItems[0]
    $currentName = $selectedItem.Text
    $locationId = $selectedItem.SubItems[1].Text

    $newName = Show-InputBox -Prompt "Enter new display name:" -Title "Rename Named Location" -DefaultValue $currentName

    if ([string]::IsNullOrWhiteSpace($newName) -or $newName -eq $currentName) {
        return
    }

    try {
        # Check if the location still exists before trying to rename
        if (-not (Test-ResourceExists -ResourceType "NamedLocation" -ResourceId $locationId)) {
            Show-WarningMessage "The selected Named Location no longer exists. Refreshing list." "Not Found"
            Refresh-NamedLocationsList $listView
            return
        }

        Update-MgIdentityConditionalAccessNamedLocation -NamedLocationId $locationId -DisplayName $newName
        Show-SuccessMessage "Named Location renamed successfully!"

        # Wait for operation to complete
        if (Wait-ForGraphOperation -ResourceType "NamedLocation" -ResourceId $locationId) {
            Refresh-NamedLocationsList $listView
        } else {
            Start-Sleep -Seconds 1
            Refresh-NamedLocationsList $listView
        }
    } catch {
        if ($_.Exception.Message -like "*NotFound*" -or $_.Exception.Message -like "*404*") {
            Show-WarningMessage "The Named Location no longer exists. Refreshing list." "Not Found"
            Refresh-NamedLocationsList $listView
        } else {
            Show-ErrorMessage "Error renaming Named Location: $_"
        }
    }
}

function Remove-SelectedNamedLocation {
    param($listView)
    
    if (-not $global:isConnected) {
        [System.Windows.Forms.MessageBox]::Show("Please connect to Microsoft Graph first.", "Not Connected")
        return
    }
    
    if ($listView.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a Named Location to delete.", "No Selection")
        return
    }

    $selectedItem = $listView.SelectedItems[0]
    $locationName = $selectedItem.Text
    $locationId = $selectedItem.SubItems[1].Text

    $result = [System.Windows.Forms.MessageBox]::Show("Are you sure you want to delete '" + $locationName + "'?", "Confirm Delete", [System.Windows.Forms.MessageBoxButtons]::YesNo)
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        try {
            # Check if the location still exists before trying to delete
            $existingLocation = Get-MgIdentityConditionalAccessNamedLocation -NamedLocationId $locationId -ErrorAction SilentlyContinue
            if (-not $existingLocation) {
                [System.Windows.Forms.MessageBox]::Show("The selected Named Location no longer exists. Refreshing list.", "Already Deleted")
                Refresh-NamedLocationsList $listView
                return
            }
            
            Remove-MgIdentityConditionalAccessNamedLocation -NamedLocationId $locationId
            [System.Windows.Forms.MessageBox]::Show("Named Location deleted successfully!", "Success")
            
            # Small delay before refresh
            Start-Sleep -Seconds 1
            Refresh-NamedLocationsList $listView
        } catch {
            if ($_.Exception.Message -like "*NotFound*" -or $_.Exception.Message -like "*404*") {
                [System.Windows.Forms.MessageBox]::Show("The Named Location was already deleted. Refreshing list.", "Already Deleted")
                Refresh-NamedLocationsList $listView
            } else {
                [System.Windows.Forms.MessageBox]::Show("Error deleting Named Location: $_", "Error")
            }
        }
    }
}

function Refresh-PoliciesList {
    param($listView)
    
    if (-not $global:isConnected) { 
        [System.Windows.Forms.MessageBox]::Show("Please connect to Microsoft Graph first.", "Not Connected")
        return 
    }
    
    $listView.Items.Clear()
    
    try {
        $policies = Get-MgIdentityConditionalAccessPolicy -All
        foreach ($policy in $policies) {
            $item = New-Object System.Windows.Forms.ListViewItem($policy.DisplayName)
            $item.SubItems.Add($policy.State) | Out-Null
            
            # Get user info
            $includeUsers = $policy.Conditions.Users.IncludeUsers
            $excludeUsers = $policy.Conditions.Users.ExcludeUsers
            
            if ($includeUsers -contains "All") {
                $item.SubItems.Add("All Users") | Out-Null
            } else {
                $userInfo = Get-UserDisplayInfo -UserIds $includeUsers
                $item.SubItems.Add(($userInfo -join '; ')) | Out-Null
            }
            
            if ($excludeUsers) {
                $excludeInfo = Get-UserDisplayInfo -UserIds $excludeUsers
                $item.SubItems.Add(($excludeInfo -join '; ')) | Out-Null
            } else {
                $item.SubItems.Add("None") | Out-Null
            }
            
            $item.Tag = $policy
            $listView.Items.Add($item) | Out-Null
        }
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Error loading policies: $_", "Error")
    }
}

function Remove-SelectedPolicy {
    param($listView)
    
    if (-not $global:isConnected) {
        [System.Windows.Forms.MessageBox]::Show("Please connect to Microsoft Graph first.", "Not Connected")
        return
    }
    
    if ($listView.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a Conditional Access Policy to delete.", "No Selection")
        return
    }

    $selectedItem = $listView.SelectedItems[0]
    $policyName = $selectedItem.Text
    $policy = $selectedItem.Tag
    $policyId = $policy.Id
    $policyState = $policy.State

    # Enhanced confirmation with policy state warning
    $confirmMessage = "Are you sure you want to delete the Conditional Access Policy:`n`n"
    $confirmMessage += "'$policyName'`n`n"
    $confirmMessage += "Policy State: $policyState`n`n"
    
    if ($policyState -eq "enabled") {
        $confirmMessage += "⚠️ WARNING: This policy is currently ENABLED and may be actively protecting your organization!`n`n"
    }
    
    $confirmMessage += "This action cannot be undone!"

    $result = [System.Windows.Forms.MessageBox]::Show(
        $confirmMessage, 
        "Confirm Delete Policy", 
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        try {
            Write-Host "Deleting Conditional Access Policy: $policyName" -ForegroundColor Yellow
            
            # Check if the policy still exists before trying to delete
            $existingPolicy = Get-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policyId -ErrorAction SilentlyContinue
            if (-not $existingPolicy) {
                [System.Windows.Forms.MessageBox]::Show("The selected Conditional Access Policy no longer exists. Refreshing list.", "Already Deleted")
                Refresh-PoliciesList $listView
                return
            }
            
            # Additional confirmation for enabled policies
            if ($existingPolicy.State -eq "enabled") {
                $finalConfirm = [System.Windows.Forms.MessageBox]::Show(
                    "FINAL CONFIRMATION:`n`nYou are about to delete an ENABLED Conditional Access Policy:`n'$policyName'`n`nThis could immediately impact user access to your organization's resources.`n`nAre you absolutely certain you want to proceed?",
                    "FINAL WARNING - Delete Enabled Policy",
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Stop
                )
                
                if ($finalConfirm -eq [System.Windows.Forms.DialogResult]::No) {
                    Write-Host "Policy deletion cancelled by user." -ForegroundColor Yellow
                    return
                }
            }
            
            # Perform the deletion
            Remove-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policyId -Confirm:$false
            
            Write-Host "✅ Policy deleted successfully!" -ForegroundColor Green
            [System.Windows.Forms.MessageBox]::Show("Conditional Access Policy '$policyName' deleted successfully!", "Success")
            
            # Small delay before refresh
            Start-Sleep -Seconds 1
            Refresh-PoliciesList $listView
            
        } catch {
            Write-Host "❌ Error deleting policy: $($_.Exception.Message)" -ForegroundColor Red
            
            if ($_.Exception.Message -like "*NotFound*" -or $_.Exception.Message -like "*404*") {
                [System.Windows.Forms.MessageBox]::Show("The Conditional Access Policy was already deleted. Refreshing list.", "Already Deleted")
                Refresh-PoliciesList $listView
            } else {
                $errorMessage = "Error deleting Conditional Access Policy:`n`n"
                $errorMessage += "Policy: $policyName`n"
                $errorMessage += "Error: " + $_.Exception.Message + "`n`n"
                $errorMessage += "Common causes:`n"
                $errorMessage += "• Missing permissions (need Policy.ReadWrite.ConditionalAccess)`n"
                $errorMessage += "• Policy is being used by other services`n"
                $errorMessage += "• Network connectivity issues`n"
                $errorMessage += "• Policy was already deleted by another admin"
                
                [System.Windows.Forms.MessageBox]::Show($errorMessage, "Delete Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            }
        }
    } else {
        Write-Host "Policy deletion cancelled by user." -ForegroundColor Gray
    }
}

function Copy-SelectedPolicy {
    param($listView)

    if (-not (Test-GraphConnectionAndSelection -listView $listView -itemType "Conditional Access Policy to copy")) {
        return
    }

    $selectedItem = $listView.SelectedItems[0]
    $policy = $selectedItem.Tag
    $policyId = $policy.Id
    $sourceName = $policy.DisplayName

    # Prompt for new name
    $defaultName = "Copy of $sourceName"
    $newName = Show-InputBox -Prompt "Enter a name for the new Conditional Access Policy:" -Title "Copy Policy" -DefaultValue $defaultName

    if ([string]::IsNullOrWhiteSpace($newName)) {
        return
    }

    try {
        Write-Host "Copying Conditional Access Policy: $sourceName" -ForegroundColor Yellow

        # Get the full policy details
        $fullPolicy = Get-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policyId

        if (-not $fullPolicy) {
            Show-WarningMessage "The selected policy no longer exists. Refreshing list." "Not Found"
            Refresh-PoliciesList $listView
            return
        }

        # Start with base policy structure
        $newPolicyBody = @{
            displayName = $newName
            state = "disabled"  # Safety first
            conditions = @{}
            grantControls = @{}
        }

        # Copy all condition types using helper functions
        if ($fullPolicy.Conditions) {
            Write-Host "Copying policy conditions..." -ForegroundColor Cyan
            $conditions = @{}

            # Users conditions
            $userProps = @('IncludeUsers', 'ExcludeUsers', 'IncludeGroups', 'ExcludeGroups', 'IncludeRoles', 'ExcludeRoles')
            $userConditions = Copy-PolicyConditionObject -SourceObject $fullPolicy.Conditions.Users -PropertyNames $userProps
            if ($userConditions) {
                $conditions.users = $userConditions
                Write-Host "  - Copied user conditions" -ForegroundColor Gray
            }

            # Application conditions
            $appProps = @('IncludeApplications', 'ExcludeApplications', 'IncludeUserActions', 'IncludeAuthenticationContextClassReferences')
            $appConditions = Copy-PolicyConditionObject -SourceObject $fullPolicy.Conditions.Applications -PropertyNames $appProps
            if ($appConditions) {
                $conditions.applications = $appConditions
                Write-Host "  - Copied application conditions" -ForegroundColor Gray
            }

            # Location conditions
            $locationProps = @('IncludeLocations', 'ExcludeLocations')
            $locationConditions = Copy-PolicyConditionObject -SourceObject $fullPolicy.Conditions.Locations -PropertyNames $locationProps
            if ($locationConditions) {
                $conditions.locations = $locationConditions
                Write-Host "  - Copied location conditions" -ForegroundColor Gray
            }

            # Platform conditions
            $platformProps = @('IncludePlatforms', 'ExcludePlatforms')
            $platformConditions = Copy-PolicyConditionObject -SourceObject $fullPolicy.Conditions.Platforms -PropertyNames $platformProps
            if ($platformConditions) {
                $conditions.platforms = $platformConditions
                Write-Host "  - Copied platform conditions" -ForegroundColor Gray
            }

            # Device state conditions (legacy, may be null)
            if ($fullPolicy.Conditions.DeviceStates) {
                $deviceStateProps = @('IncludeStates', 'ExcludeStates')
                $deviceStates = Copy-PolicyConditionObject -SourceObject $fullPolicy.Conditions.DeviceStates -PropertyNames $deviceStateProps
                if ($deviceStates) {
                    $conditions.deviceStates = $deviceStates
                    Write-Host "  - Copied device state conditions" -ForegroundColor Gray
                }
            }

            # Device filter (newer)
            if ($fullPolicy.Conditions.Devices) {
                $deviceProps = @('IncludeDevices', 'ExcludeDevices', 'DeviceFilter')
                $devices = Copy-PolicyConditionObject -SourceObject $fullPolicy.Conditions.Devices -PropertyNames $deviceProps
                if ($devices) {
                    $conditions.devices = $devices
                    Write-Host "  - Copied device conditions" -ForegroundColor Gray
                }
            }

            # Client app types
            $clientAppTypes = Copy-PolicyConditionArray -SourceArray $fullPolicy.Conditions.ClientAppTypes -PropertyName 'ClientAppTypes'
            if ($clientAppTypes) {
                $conditions.clientAppTypes = $clientAppTypes
                Write-Host "  - Copied client app types" -ForegroundColor Gray
            }

            # Sign-in risk levels
            $signInRiskLevels = Copy-PolicyConditionArray -SourceArray $fullPolicy.Conditions.SignInRiskLevels -PropertyName 'SignInRiskLevels'
            if ($signInRiskLevels) {
                $conditions.signInRiskLevels = $signInRiskLevels
                Write-Host "  - Copied sign-in risk levels" -ForegroundColor Gray
            }

            # User risk levels
            $userRiskLevels = Copy-PolicyConditionArray -SourceArray $fullPolicy.Conditions.UserRiskLevels -PropertyName 'UserRiskLevels'
            if ($userRiskLevels) {
                $conditions.userRiskLevels = $userRiskLevels
                Write-Host "  - Copied user risk levels" -ForegroundColor Gray
            }

            # Service principal risk levels (newer)
            if ($fullPolicy.Conditions.ServicePrincipalRiskLevels) {
                $spRiskLevels = Copy-PolicyConditionArray -SourceArray $fullPolicy.Conditions.ServicePrincipalRiskLevels -PropertyName 'ServicePrincipalRiskLevels'
                if ($spRiskLevels) {
                    $conditions.servicePrincipalRiskLevels = $spRiskLevels
                    Write-Host "  - Copied service principal risk levels" -ForegroundColor Gray
                }
            }

            # Client applications (newer)
            if ($fullPolicy.Conditions.ClientApplications) {
                $clientAppProps = @('IncludeServicePrincipals', 'ExcludeServicePrincipals')
                $clientApps = Copy-PolicyConditionObject -SourceObject $fullPolicy.Conditions.ClientApplications -PropertyNames $clientAppProps
                if ($clientApps) {
                    $conditions.clientApplications = $clientApps
                    Write-Host "  - Copied client applications" -ForegroundColor Gray
                }
            }

            $newPolicyBody.conditions = $conditions
        }

        # Copy grant controls
        if ($fullPolicy.GrantControls) {
            Write-Host "Copying grant controls..." -ForegroundColor Cyan
            $grantControls = @{
                operator = $fullPolicy.GrantControls.Operator ?? "OR"
            }

            $builtInControls = Copy-PolicyConditionArray -SourceArray $fullPolicy.GrantControls.BuiltInControls -PropertyName 'BuiltInControls'
            if ($builtInControls) {
                $grantControls.builtInControls = $builtInControls
            }

            $customFactors = Copy-PolicyConditionArray -SourceArray $fullPolicy.GrantControls.CustomAuthenticationFactors -PropertyName 'CustomAuthenticationFactors'
            if ($customFactors) {
                $grantControls.customAuthenticationFactors = $customFactors
            }

            $termsOfUse = Copy-PolicyConditionArray -SourceArray $fullPolicy.GrantControls.TermsOfUse -PropertyName 'TermsOfUse'
            if ($termsOfUse) {
                $grantControls.termsOfUse = $termsOfUse
            }

            if ($fullPolicy.GrantControls.AuthenticationStrength) {
                $grantControls.authenticationStrength = @{ id = $fullPolicy.GrantControls.AuthenticationStrength.Id }
            }

            $newPolicyBody.grantControls = $grantControls
        }

        # Copy session controls
        if ($fullPolicy.SessionControls) {
            Write-Host "Copying session controls..." -ForegroundColor Cyan
            $sessionControls = @{}

            # Application enforced restrictions
            if ($fullPolicy.SessionControls.ApplicationEnforcedRestrictions -and $null -ne $fullPolicy.SessionControls.ApplicationEnforcedRestrictions.IsEnabled) {
                $sessionControls.applicationEnforcedRestrictions = @{
                    isEnabled = $fullPolicy.SessionControls.ApplicationEnforcedRestrictions.IsEnabled
                }
            }

            # Cloud app security
            if ($fullPolicy.SessionControls.CloudAppSecurity -and $null -ne $fullPolicy.SessionControls.CloudAppSecurity.IsEnabled) {
                $cloudAppSecurity = @{
                    isEnabled = $fullPolicy.SessionControls.CloudAppSecurity.IsEnabled
                }
                if ($fullPolicy.SessionControls.CloudAppSecurity.CloudAppSecurityType) {
                    $cloudAppSecurity.cloudAppSecurityType = $fullPolicy.SessionControls.CloudAppSecurity.CloudAppSecurityType
                }
                $sessionControls.cloudAppSecurity = $cloudAppSecurity
            }

            # Sign-in frequency
            if ($fullPolicy.SessionControls.SignInFrequency -and $null -ne $fullPolicy.SessionControls.SignInFrequency.IsEnabled) {
                $signInFreq = @{
                    isEnabled = $fullPolicy.SessionControls.SignInFrequency.IsEnabled
                }
                if ($fullPolicy.SessionControls.SignInFrequency.Type) {
                    $signInFreq.type = $fullPolicy.SessionControls.SignInFrequency.Type
                }
                if ($null -ne $fullPolicy.SessionControls.SignInFrequency.Value) {
                    $signInFreq.value = $fullPolicy.SessionControls.SignInFrequency.Value
                }
                if ($fullPolicy.SessionControls.SignInFrequency.AuthenticationType) {
                    $signInFreq.authenticationType = $fullPolicy.SessionControls.SignInFrequency.AuthenticationType
                }
                if ($fullPolicy.SessionControls.SignInFrequency.FrequencyInterval) {
                    $signInFreq.frequencyInterval = $fullPolicy.SessionControls.SignInFrequency.FrequencyInterval
                }
                $sessionControls.signInFrequency = $signInFreq
            }

            # Persistent browser
            if ($fullPolicy.SessionControls.PersistentBrowser -and $null -ne $fullPolicy.SessionControls.PersistentBrowser.IsEnabled) {
                $persistentBrowser = @{
                    isEnabled = $fullPolicy.SessionControls.PersistentBrowser.IsEnabled
                }
                if ($fullPolicy.SessionControls.PersistentBrowser.Mode) {
                    $persistentBrowser.mode = $fullPolicy.SessionControls.PersistentBrowser.Mode
                }
                $sessionControls.persistentBrowser = $persistentBrowser
            }

            # Continuous access evaluation
            if ($fullPolicy.SessionControls.ContinuousAccessEvaluation -and $null -ne $fullPolicy.SessionControls.ContinuousAccessEvaluation.Mode) {
                $sessionControls.continuousAccessEvaluation = @{
                    mode = $fullPolicy.SessionControls.ContinuousAccessEvaluation.Mode
                }
            }

            # Disable resilience defaults
            if ($null -ne $fullPolicy.SessionControls.DisableResilienceDefaults) {
                $sessionControls.disableResilienceDefaults = $fullPolicy.SessionControls.DisableResilienceDefaults
            }

            if ($sessionControls.Count -gt 0) {
                $newPolicyBody.sessionControls = $sessionControls
                Write-Host "  - Added session controls: $($sessionControls.Keys -join ', ')" -ForegroundColor Gray
            }
        }

        # Clean null values and create policy
        $cleanedPolicy = Remove-NullValuesOptimized $newPolicyBody

        Write-Host "Creating new policy: $newName (DISABLED state)" -ForegroundColor Cyan

        # Use REST API with retry
        $uri = "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies"
        $response = Invoke-GraphOperationWithRetry -Method POST -Uri $uri -Body $cleanedPolicy

        Write-Host "✅ Policy created successfully!" -ForegroundColor Green

        $successMessage = "Conditional Access Policy copied successfully as '$newName'!`n`n"
        $successMessage += "⚠️ IMPORTANT: The new policy has been created in DISABLED state for safety.`n"
        $successMessage += "Please review the settings and enable it manually when ready.`n`n"
        $successMessage += "New Policy ID: " + $response.id

        Show-SuccessMessage $successMessage

        # Wait for operation and refresh
        if (Wait-ForGraphOperation -ResourceType "Policy" -ResourceId $response.id) {
            Refresh-PoliciesList $listView
        } else {
            Start-Sleep -Seconds 2
            Refresh-PoliciesList $listView
        }

    } catch {
        Write-Host "❌ ERROR: Policy copy failed" -ForegroundColor Red
        Write-Host ("Error details: " + $_.Exception.Message) -ForegroundColor Red

        $errorDetails = "This might be due to:`n"
        $errorDetails += "• Missing permissions (need Policy.ReadWrite.ConditionalAccess)`n"
        $errorDetails += "• Referenced objects (groups, named locations) that don't exist`n"
        $errorDetails += "• Complex policy conditions that need manual recreation"

        Show-ErrorMessage ("Error copying policy '$sourceName' to '$newName':`n`n" + $_.Exception.Message) "Copy Error" $errorDetails
    }
}

function Rename-SelectedPolicy {
    param($listView)

    if (-not (Test-GraphConnectionAndSelection -listView $listView -itemType "Conditional Access Policy to rename")) {
        return
    }

    $selectedItem = $listView.SelectedItems[0]
    $currentName = $selectedItem.Text
    $policy = $selectedItem.Tag
    $policyId = $policy.Id

    $newName = Show-InputBox -Prompt "Enter new display name:" -Title "Rename Conditional Access Policy" -DefaultValue $currentName

    if ([string]::IsNullOrWhiteSpace($newName) -or $newName -eq $currentName) {
        return
    }

    try {
        # Check if the policy still exists before trying to rename
        if (-not (Test-ResourceExists -ResourceType "Policy" -ResourceId $policyId)) {
            Show-WarningMessage "The selected Conditional Access Policy no longer exists. Refreshing list." "Not Found"
            Refresh-PoliciesList $listView
            return
        }

        Update-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policyId -DisplayName $newName
        Show-SuccessMessage "Conditional Access Policy renamed successfully!"

        # Wait for operation to complete
        if (Wait-ForGraphOperation -ResourceType "Policy" -ResourceId $policyId) {
            Refresh-PoliciesList $listView
        } else {
            Start-Sleep -Seconds 1
            Refresh-PoliciesList $listView
        }
    } catch {
        if ($_.Exception.Message -like "*NotFound*" -or $_.Exception.Message -like "*404*") {
            Show-WarningMessage "The Conditional Access Policy no longer exists. Refreshing list." "Not Found"
            Refresh-PoliciesList $listView
        } else {
            Show-ErrorMessage "Error renaming Conditional Access Policy: $_"
        }
    }
}

function Show-ManageUserExceptionsDialog {
    param($listView)
    
    if (-not $global:isConnected) {
        [System.Windows.Forms.MessageBox]::Show("Please connect to Microsoft Graph first.", "Not Connected")
        return
    }
    
    if ($listView.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a Conditional Access Policy.", "No Selection")
        return
    }

    $selectedItem = $listView.SelectedItems[0]
    $policyName = $selectedItem.Text
    $policy = $selectedItem.Tag
    $policyId = $policy.Id

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Manage User Exceptions - " + $policyName
    $form.Size = New-Object System.Drawing.Size(700, 500)
    $form.StartPosition = "CenterParent"

    # Current excluded users
    $excludedLabel = New-Object System.Windows.Forms.Label
    $excludedLabel.Text = "Currently Excluded Users:"
    $excludedLabel.Location = New-Object System.Drawing.Point(10, 20)
    $excludedLabel.Size = New-Object System.Drawing.Size(200, 20)
    $form.Controls.Add($excludedLabel)

    $excludedListBox = New-Object System.Windows.Forms.ListBox
    $excludedListBox.Location = New-Object System.Drawing.Point(10, 45)
    $excludedListBox.Size = New-Object System.Drawing.Size(660, 150)
    $excludedListBox.SelectionMode = "MultiExtended"
    $form.Controls.Add($excludedListBox)

    # Add users
    $addLabel = New-Object System.Windows.Forms.Label
    $addLabel.Text = "Add Users (emails/names/IDs - one per line):"
    $addLabel.Location = New-Object System.Drawing.Point(10, 210)
    $addLabel.Size = New-Object System.Drawing.Size(300, 20)
    $form.Controls.Add($addLabel)

    $addTextBox = New-Object System.Windows.Forms.RichTextBox
    $addTextBox.Location = New-Object System.Drawing.Point(10, 235)
    $addTextBox.Size = New-Object System.Drawing.Size(450, 100)
    $form.Controls.Add($addTextBox)

    $addButton = New-Object System.Windows.Forms.Button
    $addButton.Text = "Add Users"
    $addButton.Location = New-Object System.Drawing.Point(470, 235)
    $addButton.Size = New-Object System.Drawing.Size(100, 30)
    $addButton.Add_Click({
        $userInputs = $addTextBox.Text.Split("`n") | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
        if ($userInputs.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please enter users to add.", "Input Required")
            return
        }

        try {
            $resolveResult = Resolve-UserInput -UserInputs $userInputs
            
            if ($resolveResult.NotFoundUsers.Count -gt 0) {
                $notFoundMessage = "Some users not found: " + ($resolveResult.NotFoundUsers -join ", ") + ". Continue?"
                $result = [System.Windows.Forms.MessageBox]::Show($notFoundMessage, "Users Not Found", [System.Windows.Forms.MessageBoxButtons]::YesNo)
                if ($result -eq [System.Windows.Forms.DialogResult]::No) {
                    return
                }
            }

            if ($resolveResult.ResolvedUserIds.Count -gt 0) {
                $currentPolicy = Get-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policyId
                $currentExcludeUsers = $currentPolicy.Conditions.Users.ExcludeUsers
                $newExcludeList = @()
                
                if ($currentExcludeUsers) { $newExcludeList += $currentExcludeUsers }
                
                foreach ($userId in $resolveResult.ResolvedUserIds) {
                    if ($userId -notin $newExcludeList) {
                        $newExcludeList += $userId
                    }
                }
                
                $userConditions = @{
                    IncludeUsers = $currentPolicy.Conditions.Users.IncludeUsers
                    ExcludeUsers = $newExcludeList
                }
                
                Update-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policyId -Conditions @{ Users = $userConditions }
                
                [System.Windows.Forms.MessageBox]::Show("Users added successfully!", "Success")
                $addTextBox.Clear()
                Refresh-ExcludedUsers
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error adding users: $_", "Error")
        }
    })
    $form.Controls.Add($addButton)

    # Remove button
    $removeButton = New-Object System.Windows.Forms.Button
    $removeButton.Text = "Remove Selected"
    $removeButton.Location = New-Object System.Drawing.Point(470, 275)
    $removeButton.Size = New-Object System.Drawing.Size(100, 30)
    $removeButton.Add_Click({
        if ($excludedListBox.SelectedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select users to remove.", "No Selection")
            return
        }

        $result = [System.Windows.Forms.MessageBox]::Show("Remove selected users?", "Confirm", [System.Windows.Forms.MessageBoxButtons]::YesNo)
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            try {
                $currentPolicy = Get-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policyId
                $currentExcludeUsers = $currentPolicy.Conditions.Users.ExcludeUsers
                
                # Remove selected indices
                $indicesToRemove = @()
                foreach ($index in $excludedListBox.SelectedIndices) {
                    $indicesToRemove += $index
                }
                
                $newExcludeList = @()
                for ($i = 0; $i -lt $currentExcludeUsers.Count; $i++) {
                    if ($i -notin $indicesToRemove) {
                        $newExcludeList += $currentExcludeUsers[$i]
                    }
                }
                
                $userConditions = @{
                    IncludeUsers = $currentPolicy.Conditions.Users.IncludeUsers
                    ExcludeUsers = $newExcludeList
                }
                
                Update-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policyId -Conditions @{ Users = $userConditions }
                
                [System.Windows.Forms.MessageBox]::Show("Users removed successfully!", "Success")
                Refresh-ExcludedUsers
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Error removing users: $_", "Error")
            }
        }
    })
    $form.Controls.Add($removeButton)

    # Close button
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "Close"
    $closeButton.Location = New-Object System.Drawing.Point(580, 350)
    $closeButton.Size = New-Object System.Drawing.Size(75, 23)
    $closeButton.Add_Click({ $form.Close() })
    $form.Controls.Add($closeButton)

    # Function to refresh excluded users list
    function Refresh-ExcludedUsers {
        $excludedListBox.Items.Clear()
        $currentPolicy = Get-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policyId
        $excludeUsers = $currentPolicy.Conditions.Users.ExcludeUsers
        
        if ($excludeUsers) {
            $userInfo = Get-UserDisplayInfo -UserIds $excludeUsers
            foreach ($info in $userInfo) {
                $excludedListBox.Items.Add($info) | Out-Null
            }
        }
    }

    Refresh-ExcludedUsers
    $form.ShowDialog() | Out-Null
    Refresh-PoliciesList $listView
}

function Show-ManageIncludedUsersDialog {
    param($listView)
    
    if (-not $global:isConnected) {
        [System.Windows.Forms.MessageBox]::Show("Please connect to Microsoft Graph first.", "Not Connected")
        return
    }
    
    if ($listView.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a Conditional Access Policy.", "No Selection")
        return
    }

    $selectedItem = $listView.SelectedItems[0]
    $policyName = $selectedItem.Text
    $policy = $selectedItem.Tag
    $policyId = $policy.Id

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Manage Included Users - " + $policyName
    $form.Size = New-Object System.Drawing.Size(700, 600)
    $form.StartPosition = "CenterParent"

    # Current included users
    $includedLabel = New-Object System.Windows.Forms.Label
    $includedLabel.Text = "Currently Included Users:"
    $includedLabel.Location = New-Object System.Drawing.Point(10, 20)
    $includedLabel.Size = New-Object System.Drawing.Size(200, 20)
    $form.Controls.Add($includedLabel)

    $includedListBox = New-Object System.Windows.Forms.ListBox
    $includedListBox.Location = New-Object System.Drawing.Point(10, 45)
    $includedListBox.Size = New-Object System.Drawing.Size(660, 150)
    $includedListBox.SelectionMode = "MultiExtended"
    $form.Controls.Add($includedListBox)

    # All Users checkbox
    $allUsersCheckBox = New-Object System.Windows.Forms.CheckBox
    $allUsersCheckBox.Text = "Include All Users (overrides specific user list)"
    $allUsersCheckBox.Location = New-Object System.Drawing.Point(10, 205)
    $allUsersCheckBox.Size = New-Object System.Drawing.Size(300, 20)
    $allUsersCheckBox.Add_CheckedChanged({
        if ($allUsersCheckBox.Checked) {
            $includedListBox.Enabled = $false
            $addTextBox.Enabled = $false
            $addButton.Enabled = $false
            $removeButton.Enabled = $false
        } else {
            $includedListBox.Enabled = $true
            $addTextBox.Enabled = $true
            $addButton.Enabled = $true
            $removeButton.Enabled = $true
        }
    })
    $form.Controls.Add($allUsersCheckBox)

    # Add users
    $addLabel = New-Object System.Windows.Forms.Label
    $addLabel.Text = "Add Users (emails/names/IDs - one per line):"
    $addLabel.Location = New-Object System.Drawing.Point(10, 240)
    $addLabel.Size = New-Object System.Drawing.Size(300, 20)
    $form.Controls.Add($addLabel)

    $addTextBox = New-Object System.Windows.Forms.RichTextBox
    $addTextBox.Location = New-Object System.Drawing.Point(10, 265)
    $addTextBox.Size = New-Object System.Drawing.Size(450, 100)
    $form.Controls.Add($addTextBox)

    $addButton = New-Object System.Windows.Forms.Button
    $addButton.Text = "Add Users"
    $addButton.Location = New-Object System.Drawing.Point(470, 265)
    $addButton.Size = New-Object System.Drawing.Size(100, 30)
    $addButton.Add_Click({
        $userInputs = $addTextBox.Text.Split("`n") | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
        if ($userInputs.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please enter users to add.", "Input Required")
            return
        }

        try {
            $resolveResult = Resolve-UserInput -UserInputs $userInputs
            
            if ($resolveResult.NotFoundUsers.Count -gt 0) {
                $notFoundMessage = "Some users not found: " + ($resolveResult.NotFoundUsers -join ", ") + ". Continue?"
                $result = [System.Windows.Forms.MessageBox]::Show($notFoundMessage, "Users Not Found", [System.Windows.Forms.MessageBoxButtons]::YesNo)
                if ($result -eq [System.Windows.Forms.DialogResult]::No) {
                    return
                }
            }

            if ($resolveResult.ResolvedUserIds.Count -gt 0) {
                $currentPolicy = Get-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policyId
                $currentIncludeUsers = $currentPolicy.Conditions.Users.IncludeUsers
                $newIncludeList = @()
                
                # Don't include "All" in the specific user list
                if ($currentIncludeUsers) { 
                    $newIncludeList += ($currentIncludeUsers | Where-Object { $_ -ne "All" })
                }
                
                foreach ($userId in $resolveResult.ResolvedUserIds) {
                    if ($userId -notin $newIncludeList) {
                        $newIncludeList += $userId
                    }
                }
                
                $userConditions = @{
                    IncludeUsers = $newIncludeList
                    ExcludeUsers = $currentPolicy.Conditions.Users.ExcludeUsers
                }
                
                Update-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policyId -Conditions @{ Users = $userConditions }
                
                [System.Windows.Forms.MessageBox]::Show("Users added successfully!", "Success")
                $addTextBox.Clear()
                Refresh-IncludedUsers
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error adding users: $_", "Error")
        }
    })
    $form.Controls.Add($addButton)

    # Remove button
    $removeButton = New-Object System.Windows.Forms.Button
    $removeButton.Text = "Remove Selected"
    $removeButton.Location = New-Object System.Drawing.Point(470, 305)
    $removeButton.Size = New-Object System.Drawing.Size(100, 30)
    $removeButton.Add_Click({
        if ($includedListBox.SelectedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select users to remove.", "No Selection")
            return
        }

        $result = [System.Windows.Forms.MessageBox]::Show("Remove selected users?", "Confirm", [System.Windows.Forms.MessageBoxButtons]::YesNo)
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            try {
                $currentPolicy = Get-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policyId
                $currentIncludeUsers = $currentPolicy.Conditions.Users.IncludeUsers
                
                # Remove selected indices
                $indicesToRemove = @()
                foreach ($index in $includedListBox.SelectedIndices) {
                    $indicesToRemove += $index
                }
                
                $newIncludeList = @()
                $userList = ($currentIncludeUsers | Where-Object { $_ -ne "All" })
                for ($i = 0; $i -lt $userList.Count; $i++) {
                    if ($i -notin $indicesToRemove) {
                        $newIncludeList += $userList[$i]
                    }
                }
                
                # If no users left, we need at least something or the policy will be invalid
                if ($newIncludeList.Count -eq 0) {
                    $confirmEmpty = [System.Windows.Forms.MessageBox]::Show("This will remove all specific users. The policy may become invalid without any included users or groups. Continue?", "Warning", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
                    if ($confirmEmpty -eq [System.Windows.Forms.DialogResult]::No) {
                        return
                    }
                }
                
                $userConditions = @{
                    IncludeUsers = $newIncludeList
                    ExcludeUsers = $currentPolicy.Conditions.Users.ExcludeUsers
                }
                
                Update-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policyId -Conditions @{ Users = $userConditions }
                
                [System.Windows.Forms.MessageBox]::Show("Users removed successfully!", "Success")
                Refresh-IncludedUsers
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Error removing users: $_", "Error")
            }
        }
    })
    $form.Controls.Add($removeButton)

    # Apply/Save button for All Users setting
    $applyAllButton = New-Object System.Windows.Forms.Button
    $applyAllButton.Text = "Apply All Users Setting"
    $applyAllButton.Location = New-Object System.Drawing.Point(470, 380)
    $applyAllButton.Size = New-Object System.Drawing.Size(130, 30)
    $applyAllButton.Add_Click({
        try {
            $currentPolicy = Get-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policyId
            
            if ($allUsersCheckBox.Checked) {
                # Set to "All"
                $userConditions = @{
                    IncludeUsers = @("All")
                    ExcludeUsers = $currentPolicy.Conditions.Users.ExcludeUsers
                }
            } else {
                # Remove "All" but keep existing specific users
                $currentIncludeUsers = $currentPolicy.Conditions.Users.IncludeUsers
                $specificUsers = $currentIncludeUsers | Where-Object { $_ -ne "All" }
                
                $userConditions = @{
                    IncludeUsers = $specificUsers
                    ExcludeUsers = $currentPolicy.Conditions.Users.ExcludeUsers
                }
            }
            
            Update-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policyId -Conditions @{ Users = $userConditions }
            
            [System.Windows.Forms.MessageBox]::Show("All Users setting applied successfully!", "Success")
            Refresh-IncludedUsers
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error applying All Users setting: $_", "Error")
        }
    })
    $form.Controls.Add($applyAllButton)

    # Close button
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "Close"
    $closeButton.Location = New-Object System.Drawing.Point(580, 430)
    $closeButton.Size = New-Object System.Drawing.Size(75, 23)
    $closeButton.Add_Click({ $form.Close() })
    $form.Controls.Add($closeButton)

    # Function to refresh included users list
    function Refresh-IncludedUsers {
        $includedListBox.Items.Clear()
        $currentPolicy = Get-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policyId
        $includeUsers = $currentPolicy.Conditions.Users.IncludeUsers
        
        if ($includeUsers -contains "All") {
            $allUsersCheckBox.Checked = $true
            # Disable other controls when "All" is selected
            $includedListBox.Enabled = $false
            $addTextBox.Enabled = $false
            $addButton.Enabled = $false
            $removeButton.Enabled = $false
        } else {
            $allUsersCheckBox.Checked = $false
            $includedListBox.Enabled = $true
            $addTextBox.Enabled = $true
            $addButton.Enabled = $true
            $removeButton.Enabled = $true
            
            if ($includeUsers) {
                $specificUsers = $includeUsers | Where-Object { $_ -ne "All" }
                if ($specificUsers) {
                    $userInfo = Get-UserDisplayInfo -UserIds $specificUsers
                    foreach ($info in $userInfo) {
                        $includedListBox.Items.Add($info) | Out-Null
                    }
                }
            }
        }
    }

    Refresh-IncludedUsers
    $form.ShowDialog() | Out-Null
    Refresh-PoliciesList $listView
}

function Create-MainForm {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Conditional Access Management Tool"
    $form.Size = New-Object System.Drawing.Size(1000, 700)
    $form.StartPosition = "CenterScreen"

    # Connection Buttons Panel
    $connectionPanel = New-Object System.Windows.Forms.Panel
    $connectionPanel.Location = New-Object System.Drawing.Point(10, 10)
    $connectionPanel.Size = New-Object System.Drawing.Size(970, 35)
    $form.Controls.Add($connectionPanel)

    # Connect Button
    $script:connectButton = New-Object System.Windows.Forms.Button
    $script:connectButton.Text = "Connect to Microsoft Graph"
    $script:connectButton.Location = New-Object System.Drawing.Point(0, 0)
    $script:connectButton.Size = New-Object System.Drawing.Size(180, 30)
    $script:connectButton.Add_Click({
        if (Connect-GraphAPI) {
            # Refresh both lists directly
            if ($global:isConnected) {
                Refresh-NamedLocationsList $script:namedLocationsListView
                Refresh-PoliciesList $script:policiesListView
            }
        }
    })
    $connectionPanel.Controls.Add($script:connectButton)

    # Reconnect Button
    $script:reconnectButton = New-Object System.Windows.Forms.Button
    $script:reconnectButton.Text = "Reconnect/Change Tenant"
    $script:reconnectButton.Location = New-Object System.Drawing.Point(190, 0)
    $script:reconnectButton.Size = New-Object System.Drawing.Size(160, 30)
    $script:reconnectButton.Add_Click({
        Show-ReconnectDialog
    })
    $connectionPanel.Controls.Add($script:reconnectButton)

    # Disconnect Button
    $script:disconnectButton = New-Object System.Windows.Forms.Button
    $script:disconnectButton.Text = "Disconnect"
    $script:disconnectButton.Location = New-Object System.Drawing.Point(360, 0)
    $script:disconnectButton.Size = New-Object System.Drawing.Size(100, 30)
    $script:disconnectButton.Add_Click({
        Disconnect-GraphAPI
    })
    $connectionPanel.Controls.Add($script:disconnectButton)

    # Status Label
    $script:statusLabel = New-Object System.Windows.Forms.Label
    $script:statusLabel.Text = "Not connected"
    $script:statusLabel.Location = New-Object System.Drawing.Point(470, 5)
    $script:statusLabel.Size = New-Object System.Drawing.Size(500, 20)
    $script:statusLabel.ForeColor = [System.Drawing.Color]::Red
    $connectionPanel.Controls.Add($script:statusLabel)

    # Tab Control
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Location = New-Object System.Drawing.Point(10, 55)
    $tabControl.Size = New-Object System.Drawing.Size(960, 590)
    $form.Controls.Add($tabControl)

    # Named Locations Tab
    $namedLocationsTab = New-Object System.Windows.Forms.TabPage
    $namedLocationsTab.Text = "Named Locations"
    $tabControl.TabPages.Add($namedLocationsTab)

    # Named Locations ListView
    $script:namedLocationsListView = New-Object System.Windows.Forms.ListView
    $script:namedLocationsListView.Location = New-Object System.Drawing.Point(10, 50)
    $script:namedLocationsListView.Size = New-Object System.Drawing.Size(920, 400)
    $script:namedLocationsListView.View = "Details"
    $script:namedLocationsListView.FullRowSelect = $true
    $script:namedLocationsListView.GridLines = $true
    $script:namedLocationsListView.Columns.Add("Display Name", 200) | Out-Null
    $script:namedLocationsListView.Columns.Add("ID", 250) | Out-Null
    $script:namedLocationsListView.Columns.Add("Type", 100) | Out-Null
    $script:namedLocationsListView.Columns.Add("Details", 370) | Out-Null
    $namedLocationsTab.Controls.Add($script:namedLocationsListView)

    # Named Locations Buttons
    $nlRefreshButton = New-Object System.Windows.Forms.Button
    $nlRefreshButton.Text = "Refresh"
    $nlRefreshButton.Location = New-Object System.Drawing.Point(10, 15)
    $nlRefreshButton.Size = New-Object System.Drawing.Size(80, 25)
    $nlRefreshButton.Add_Click({
        Refresh-NamedLocationsList $script:namedLocationsListView
    })
    $namedLocationsTab.Controls.Add($nlRefreshButton)

    $nlCreateButton = New-Object System.Windows.Forms.Button
    $nlCreateButton.Text = "Create Country Location"
    $nlCreateButton.Location = New-Object System.Drawing.Point(100, 15)
    $nlCreateButton.Size = New-Object System.Drawing.Size(150, 25)
    $nlCreateButton.Add_Click({
        Show-CreateCountryLocationDialog $script:namedLocationsListView
    })
    $namedLocationsTab.Controls.Add($nlCreateButton)

    $nlEditButton = New-Object System.Windows.Forms.Button
    $nlEditButton.Text = "Edit Countries"
    $nlEditButton.Location = New-Object System.Drawing.Point(260, 15)
    $nlEditButton.Size = New-Object System.Drawing.Size(100, 25)
    $nlEditButton.Add_Click({
        Edit-SelectedNamedLocation $script:namedLocationsListView
    })
    $namedLocationsTab.Controls.Add($nlEditButton)

    $nlCopyButton = New-Object System.Windows.Forms.Button
    $nlCopyButton.Text = "Copy Countries"
    $nlCopyButton.Location = New-Object System.Drawing.Point(370, 15)
    $nlCopyButton.Size = New-Object System.Drawing.Size(100, 25)
    $nlCopyButton.Add_Click({
        Copy-SelectedNamedLocation $script:namedLocationsListView
    })
    $namedLocationsTab.Controls.Add($nlCopyButton)

    $nlRenameButton = New-Object System.Windows.Forms.Button
    $nlRenameButton.Text = "Rename"
    $nlRenameButton.Location = New-Object System.Drawing.Point(480, 15)
    $nlRenameButton.Size = New-Object System.Drawing.Size(80, 25)
    $nlRenameButton.Add_Click({
        Rename-SelectedNamedLocation $script:namedLocationsListView
    })
    $namedLocationsTab.Controls.Add($nlRenameButton)

    $nlDeleteButton = New-Object System.Windows.Forms.Button
    $nlDeleteButton.Text = "Delete"
    $nlDeleteButton.Location = New-Object System.Drawing.Point(570, 15)
    $nlDeleteButton.Size = New-Object System.Drawing.Size(80, 25)
    $nlDeleteButton.Add_Click({
        Remove-SelectedNamedLocation $script:namedLocationsListView
    })
    $namedLocationsTab.Controls.Add($nlDeleteButton)

    # Conditional Access Policies Tab
    $policiesTab = New-Object System.Windows.Forms.TabPage
    $policiesTab.Text = "Conditional Access Policies"
    $tabControl.TabPages.Add($policiesTab)

    # Policies ListView
    $script:policiesListView = New-Object System.Windows.Forms.ListView
    $script:policiesListView.Location = New-Object System.Drawing.Point(10, 50)
    $script:policiesListView.Size = New-Object System.Drawing.Size(920, 400)
    $script:policiesListView.View = "Details"
    $script:policiesListView.FullRowSelect = $true
    $script:policiesListView.GridLines = $true
    $script:policiesListView.Columns.Add("Display Name", 200) | Out-Null
    $script:policiesListView.Columns.Add("State", 100) | Out-Null
    $script:policiesListView.Columns.Add("Included Users", 300) | Out-Null
    $script:policiesListView.Columns.Add("Excluded Users", 320) | Out-Null
    $policiesTab.Controls.Add($script:policiesListView)

    # Policies Buttons
    $polRefreshButton = New-Object System.Windows.Forms.Button
    $polRefreshButton.Text = "Refresh"
    $polRefreshButton.Location = New-Object System.Drawing.Point(10, 15)
    $polRefreshButton.Size = New-Object System.Drawing.Size(80, 25)
    $polRefreshButton.Add_Click({
        Refresh-PoliciesList $script:policiesListView
    })
    $policiesTab.Controls.Add($polRefreshButton)

    $polManageIncludedButton = New-Object System.Windows.Forms.Button
    $polManageIncludedButton.Text = "Manage Included Users"
    $polManageIncludedButton.Location = New-Object System.Drawing.Point(100, 15)
    $polManageIncludedButton.Size = New-Object System.Drawing.Size(140, 25)
    $polManageIncludedButton.Add_Click({
        Show-ManageIncludedUsersDialog $script:policiesListView
    })
    $policiesTab.Controls.Add($polManageIncludedButton)

    $polManageUsersButton = New-Object System.Windows.Forms.Button
    $polManageUsersButton.Text = "Manage User Exceptions"
    $polManageUsersButton.Location = New-Object System.Drawing.Point(250, 15)
    $polManageUsersButton.Size = New-Object System.Drawing.Size(150, 25)
    $polManageUsersButton.Add_Click({
        Show-ManageUserExceptionsDialog $script:policiesListView
    })
    $policiesTab.Controls.Add($polManageUsersButton)

    $polRenameButton = New-Object System.Windows.Forms.Button
    $polRenameButton.Text = "Rename Policy"
    $polRenameButton.Location = New-Object System.Drawing.Point(410, 15)
    $polRenameButton.Size = New-Object System.Drawing.Size(100, 25)
    $polRenameButton.Add_Click({
        Rename-SelectedPolicy $script:policiesListView
    })
    $policiesTab.Controls.Add($polRenameButton)

    $polCopyButton = New-Object System.Windows.Forms.Button
    $polCopyButton.Text = "Copy Policy"
    $polCopyButton.Location = New-Object System.Drawing.Point(520, 15)
    $polCopyButton.Size = New-Object System.Drawing.Size(100, 25)
    $polCopyButton.Add_Click({
        Copy-SelectedPolicy $script:policiesListView
    })
    $policiesTab.Controls.Add($polCopyButton)

    $polDeleteButton = New-Object System.Windows.Forms.Button
    $polDeleteButton.Text = "Delete Policy"
    $polDeleteButton.Location = New-Object System.Drawing.Point(630, 15)
    $polDeleteButton.Size = New-Object System.Drawing.Size(100, 25)
    $polDeleteButton.ForeColor = [System.Drawing.Color]::DarkRed
    $polDeleteButton.Add_Click({
        Remove-SelectedPolicy $script:policiesListView
    })
    $policiesTab.Controls.Add($polDeleteButton)

    # Initialize UI state
    Update-ConnectionUI

    return $form
}

# Main execution
try {
    Write-Host "Creating GUI..." -ForegroundColor Green
    Write-Host "Available Policy Management Features:" -ForegroundColor Cyan
    Write-Host "  - View and refresh policies" -ForegroundColor White
    Write-Host "  - Manage included/excluded users" -ForegroundColor White
    Write-Host "  - Rename policies" -ForegroundColor White
    Write-Host "  - Copy policies (created as DISABLED for safety)" -ForegroundColor White
    Write-Host "  - Delete policies (with confirmation warnings)" -ForegroundColor White
    Write-Host "  - Named Location management with checkbox country selector" -ForegroundColor White
    Write-Host "" -ForegroundColor White
    $mainForm = Create-MainForm
    Write-Host "Showing form..." -ForegroundColor Green
    [void]$mainForm.ShowDialog()
} catch {
    Write-Error "Error: $($_.Exception.Message)"
    Read-Host "Press Enter to exit"
}