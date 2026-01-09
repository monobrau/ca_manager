<#
.SYNOPSIS
    A GUI-based PowerShell script to manage Microsoft Entra Conditional Access policies.
.NOTES
    Version: 3.5 (Added Reset Auth, connection state tracking, help dialog, and tooltips)
    Requirements: Windows PowerShell 5.1 with Microsoft.Graph module
#>

# Suppress Write-Host output when running as EXE (no console window)
# Redirect all Write-Host to null to prevent popups
function Write-Host {
    param(
        [Parameter(ValueFromPipeline=$true)]
        [object]$Object,
        [switch]$NoNewline,
        [System.ConsoleColor]$ForegroundColor,
        [System.ConsoleColor]$BackgroundColor
    )
    # Do nothing - suppress all Write-Host output
}

# Check PowerShell version first (before loading anything)
# PowerShell Core is now supported
if ($PSVersionTable.PSEdition -eq "Core") {
    # Check if we're on Windows (required for Windows Forms)
    if ($IsWindows -eq $false) {
        [System.Windows.Forms.MessageBox]::Show("This script requires Windows to run. Windows Forms is not available on non-Windows platforms.`n`nPlease run this script on a Windows machine or in Windows Subsystem for Linux (WSL).", "Platform Not Supported", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        exit
    }
}

# Load assemblies and configure immediately (must be done before ANY Windows Forms objects are created)
try {
    # Load assemblies first - works in both PowerShell Core and Windows PowerShell
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing  
    
    # Only load Microsoft.VisualBasic if available (may not be in PowerShell Core)
    try {
        Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction SilentlyContinue
    } catch {
        # Will use custom InputBox implementation
    }
    
    # Try to configure rendering - if it fails, continue anyway
    try {
        [System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)
        [System.Windows.Forms.Application]::EnableVisualStyles()
    } catch {
        # Forms may already exist, continue anyway
        [System.Windows.Forms.Application]::EnableVisualStyles()
    }
} catch {
    [System.Windows.Forms.MessageBox]::Show("Failed to initialize Windows Forms: $_`n`nThis error can occur if Windows Forms objects were already created. Try restarting the application.", "Initialization Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}

# Hide console window for EXE (but keep it available for authentication)
# Detect if running as EXE and hide console window
try {
    $isExe = $PSCommandPath -like "*.exe" -or $MyInvocation.MyCommand.Path -like "*.exe"
    if ($isExe) {
        # Import Windows API to hide/show console window
        Add-Type -Name Window -Namespace Console -MemberDefinition '
        [DllImport("Kernel32.dll")]
        public static extern IntPtr GetConsoleWindow();
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
        '
        
        # Hide the console window initially
        $consolePtr = [Console.Window]::GetConsoleWindow()
        if ($consolePtr -ne [IntPtr]::Zero) {
            # SW_HIDE = 0
            [Console.Window]::ShowWindow($consolePtr, 0) | Out-Null
        }
    }
} catch {
    # Ignore errors - console hiding is optional
}

# Function to show console window when needed (for authentication)
function Show-ConsoleWindow {
    try {
        $consolePtr = [Console.Window]::GetConsoleWindow()
        if ($consolePtr -ne [IntPtr]::Zero) {
            # SW_SHOW = 5, SW_RESTORE = 9
            [Console.Window]::ShowWindow($consolePtr, 9) | Out-Null
        }
    } catch {
        # Ignore errors
    }
}

# Function to hide console window
function Hide-ConsoleWindow {
    try {
        $consolePtr = [Console.Window]::GetConsoleWindow()
        if ($consolePtr -ne [IntPtr]::Zero) {
            # SW_HIDE = 0
            [Console.Window]::ShowWindow($consolePtr, 0) | Out-Null
        }
    } catch {
        # Ignore errors
    }
}

# Check for Microsoft Graph module - EXE-compatible approach (silent)
# EXE environments often don't have the user's module path
# Add common module paths explicitly
$userModulePath = Join-Path $env:USERPROFILE "Documents\PowerShell\Modules"
$userModulePath2 = Join-Path $env:USERPROFILE "Documents\WindowsPowerShell\Modules"
$programModulePath = "C:\Program Files\WindowsPowerShell\Modules"
$programModulePath2 = "C:\Program Files\PowerShell\Modules"

# Add user module paths to PSModulePath if they exist and aren't already there
# Filter out OneDrive paths to avoid issues
$pathsToAdd = @()
if (Test-Path $userModulePath) {
    if ($userModulePath -notlike "*OneDrive*" -and $env:PSModulePath -notlike "*$userModulePath*") {
        $pathsToAdd += $userModulePath
    }
}
if (Test-Path $userModulePath2) {
    if ($userModulePath2 -notlike "*OneDrive*" -and $env:PSModulePath -notlike "*$userModulePath2*") {
        $pathsToAdd += $userModulePath2
    }
}
if (Test-Path $programModulePath) {
    if ($env:PSModulePath -notlike "*$programModulePath*") {
        $pathsToAdd += $programModulePath
    }
}
if (Test-Path $programModulePath2) {
    if ($env:PSModulePath -notlike "*$programModulePath2*") {
        $pathsToAdd += $programModulePath2
    }
}

# Filter out OneDrive paths from existing PSModulePath and combine with new paths
$cleanExistingPaths = ($env:PSModulePath -split ';' | Where-Object { $_ -and $_ -notlike "*OneDrive*" })
if ($pathsToAdd.Count -gt 0) {
    $env:PSModulePath = ($pathsToAdd + $cleanExistingPaths) -join ';'
} else {
    $env:PSModulePath = $cleanExistingPaths -join ';'
}

$modulesLoaded = $false
$moduleError = $null

# Import only the specific sub-modules we need (not the full Microsoft.Graph)
# This avoids hitting PowerShell's function limit (~4096 functions)
$requiredModules = @(
    "Microsoft.Graph.Authentication",  # Required for Connect-MgGraph and Get-MgContext
    "Microsoft.Graph.Identity.SignIns",
    "Microsoft.Graph.Users"
)
$loadedCount = 0
$failedModules = @()
$moduleErrors = @()

# First, check if modules are available
$availableModules = @()
foreach ($moduleName in $requiredModules) {
    $moduleAvailable = Get-Module -ListAvailable -Name $moduleName -ErrorAction SilentlyContinue
    if ($moduleAvailable) {
        $availableModules += $moduleName
    } else {
        $moduleErrors += "$moduleName (not found in module paths)"
    }
}

# If no modules are available, show error immediately
if ($availableModules.Count -eq 0) {
    $errorMsg = "Microsoft.Graph modules could not be found.`n`n"
    $errorMsg += "Checked module paths:`n"
    $errorMsg += ($env:PSModulePath -split ';' | Where-Object { $_ -and $_ -notlike "*OneDrive*" } | ForEach-Object { "  - $_" }) -join "`n"
    $errorMsg += "`n`nPlease install Microsoft.Graph module:`n`n"
    $errorMsg += "  Install-Module Microsoft.Graph -Scope CurrentUser`n`n"
    $errorMsg += "Or for all users (requires admin):`n`n"
    $errorMsg += "  Install-Module Microsoft.Graph -Scope AllUsers`n`n"
    $errorMsg += "Note: This may take several minutes as it installs multiple sub-modules.`n`n"
    $errorMsg += "After installation, restart this application."
    
    [System.Windows.Forms.MessageBox]::Show($errorMsg, "Module Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}

# Try to import available modules
foreach ($moduleName in $requiredModules) {
    try {
        # Check if module is already loaded
        $moduleAlreadyLoaded = Get-Module -Name $moduleName -ErrorAction SilentlyContinue
        if ($moduleAlreadyLoaded) {
            # Module already loaded, just verify cmdlets are available
            $loadedCount++
        } else {
            # Check if module is available before trying to import
            $moduleAvailable = Get-Module -ListAvailable -Name $moduleName -ErrorAction SilentlyContinue
            if ($moduleAvailable) {
                # Import module (PowerShell 5.1 doesn't support -SkipEditionCheck)
                Import-Module $moduleName -ErrorAction Stop
                $loadedCount++
            } else {
                $failedModules += $moduleName
                $moduleErrors += "$moduleName (not found)"
            }
        }
    } catch {
        $failedModules += $moduleName
        $moduleErrors += "$moduleName ($($_.Exception.Message))"
    }
}

if ($loadedCount -eq $requiredModules.Count) {
    $modulesLoaded = $true
} elseif ($loadedCount -gt 0) {
    # Some modules loaded - warn user but continue
    $modulesLoaded = $true
} else {
    $moduleError = "Could not load any required modules"
}

# Verify cmdlets are available
if ($modulesLoaded) {
    $testCmdlets = @("Get-MgContext", "Connect-MgGraph")
    $missingCmdlets = @()
    
    foreach ($cmdlet in $testCmdlets) {
        if (-not (Get-Command $cmdlet -ErrorAction SilentlyContinue)) {
            $missingCmdlets += $cmdlet
        }
    }
    
    if ($missingCmdlets.Count -gt 0) {
        # Required cmdlets are missing - modules not properly loaded
        $modulesLoaded = $false
        $moduleErrors += "Required cmdlets missing: $($missingCmdlets -join ', ')"
    }
}

# Show error dialog and exit if modules couldn't be loaded
if (-not $modulesLoaded) {
    $errorMsg = "Microsoft.Graph module could not be loaded.`n`n"
    
    if ($moduleErrors.Count -gt 0) {
        $errorMsg += "Failed modules:`n"
        $errorMsg += ($moduleErrors | ForEach-Object { "  - $_" }) -join "`n"
        $errorMsg += "`n`n"
    }
    
    $errorMsg += "Module search paths:`n"
    $errorMsg += ($env:PSModulePath -split ';' | Where-Object { $_ -and $_ -notlike "*OneDrive*" } | ForEach-Object { "  - $_" }) -join "`n"
    $errorMsg += "`n`nPlease install Microsoft.Graph module:`n`n"
    $errorMsg += "  Install-Module Microsoft.Graph -Scope CurrentUser`n`n"
    $errorMsg += "Or for all users (requires admin):`n`n"
    $errorMsg += "  Install-Module Microsoft.Graph -Scope AllUsers`n`n"
    $errorMsg += "Note: This may take several minutes as it installs multiple sub-modules.`n`n"
    $errorMsg += "After installation, restart this application."
    
    [System.Windows.Forms.MessageBox]::Show($errorMsg, "Module Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
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
        if (([System.Management.Automation.PSTypeName]'Microsoft.VisualBasic.Interaction').Type) {
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
    $okButton.BackColor = [System.Drawing.Color]::LightGreen
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(295, 75)
    $cancelButton.Size = New-Object System.Drawing.Size(75, 23)
    $cancelButton.BackColor = [System.Drawing.Color]::LightGray
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

# Global Variables
$global:isConnected = $false
$global:tenantId = ""
$global:tenantDisplayName = ""
$script:isConnecting = $false
$script:statusLabel = $null
$script:namedLocationsListView = $null
$script:policiesListView = $null
$script:mainForm = $null
$script:connectButton = $null
$script:disconnectButton = $null
$script:reconnectButton = $null
$script:resetAuthButton = $null

# Microsoft Graph Functions
function Reset-Authentication {
    # Forcefully clear any stuck authentication sessions
    try {
        # Try to disconnect gracefully first
        Disconnect-MgGraph -ErrorAction SilentlyContinue
    } catch {
        # Ignore errors - we're trying to clear stuck state
    }
    
    # Clear connection state
    $global:isConnected = $false
    $global:tenantId = ""
    $global:tenantDisplayName = ""
    $script:isConnecting = $false
    
    # Clear the list views
    if ($script:namedLocationsListView) {
        $script:namedLocationsListView.Items.Clear()
    }
    if ($script:policiesListView) {
        $script:policiesListView.Items.Clear()
    }
    
    Update-ConnectionUI
    [System.Windows.Forms.MessageBox]::Show("Authentication state has been reset. You can now attempt to connect again.", "Auth Reset", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}

function Connect-GraphAPI {
    param([string]$TenantId)
    
    # Prevent multiple simultaneous connection attempts
    if ($script:isConnecting) {
        [System.Windows.Forms.MessageBox]::Show("A connection attempt is already in progress. Please wait or use 'Reset Auth' if it appears stuck.", "Connection In Progress", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return $false
    }
    
    # Clear any stuck authentication state before attempting connection
    try {
        Disconnect-MgGraph -ErrorAction SilentlyContinue
    } catch {
        # Ignore errors - we're just clearing potential stuck state
    }
    
    $script:isConnecting = $true
    Update-ConnectionUI
    
    $requiredScopes = @(
        "Policy.Read.All",
        "Policy.ReadWrite.ConditionalAccess", 
        "User.Read.All",
        "Group.Read.All",
        "Organization.Read.All"
    )

    try {
        # For EXE versions, ensure the main form is visible and bring it to front
        # This helps WAM authentication find the parent window handle
        if ($script:mainForm -and -not $script:mainForm.IsDisposed) {
            $script:mainForm.WindowState = [System.Windows.Forms.FormWindowState]::Normal
            $script:mainForm.BringToFront()
            $script:mainForm.Activate()
            $script:mainForm.Focus()
            
            # Force the form to be visible and on top
            [System.Windows.Forms.Application]::DoEvents()
        }
        
        # Attempt connection with error handling
        # Use standard interactive browser authentication (WAM will be used automatically on Windows)
        # Note: The form must be visible for WAM to work properly in EXE mode
        # Suppress WAM warnings (they appear as informational messages, not critical)
        $originalWarningPreference = $WarningPreference
        $WarningPreference = 'SilentlyContinue'
        try {
            if ($TenantId) {
                Connect-MgGraph -Scopes $requiredScopes -TenantId $TenantId -ErrorAction Stop
            } else {
                Connect-MgGraph -Scopes $requiredScopes -ErrorAction Stop
            }
        } catch {
            # If window handle error occurs, try device code authentication as fallback
            # This works better when EXE has console window support
            if ($_.Exception.Message -like "*window handle*" -or $_.Exception.Message -like "*parent window*") {
                $WarningPreference = 'Continue'  # Show device code output
                
                # Show console window for device code display
                Show-ConsoleWindow
                
                $deviceCodeMsg = "Browser authentication failed. Switching to device code authentication.`n`n" +
                                "A device code will be displayed in the console window.`n`n" +
                                "Please:`n" +
                                "1. Look at the console window for the device code`n" +
                                "2. Go to https://microsoft.com/devicelogin`n" +
                                "3. Enter the displayed code`n" +
                                "4. Complete authentication in your browser`n`n" +
                                "Click OK to continue with device code authentication."
                [System.Windows.Forms.MessageBox]::Show($deviceCodeMsg, "Device Code Authentication", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
                
                # Ensure form is responsive
                [System.Windows.Forms.Application]::DoEvents()
                
                # Use device code authentication - code will display in console
                try {
                    if ($TenantId) {
                        Connect-MgGraph -Scopes $requiredScopes -TenantId $TenantId -UseDeviceAuthentication -ErrorAction Stop
                    } else {
                        Connect-MgGraph -Scopes $requiredScopes -UseDeviceAuthentication -ErrorAction Stop
                    }
                    # Hide console window again after successful authentication
                    Start-Sleep -Seconds 2
                    Hide-ConsoleWindow
                } catch {
                    $WarningPreference = $originalWarningPreference
                    # Keep console visible if authentication failed
                    throw
                }
            } else {
                $WarningPreference = $originalWarningPreference
                throw
            }
        } finally {
            $WarningPreference = $originalWarningPreference
        }
        
        # Verify connection was successful
        $token = Get-MgContext -ErrorAction Stop
        if (-not $token) {
            throw "Connection completed but no context was returned"
        }
        
        $global:isConnected = $true
        $global:tenantId = $token.TenantId
        $global:tenantDisplayName = $global:tenantId  # Default fallback
        
        # Hide console window after successful connection
        Hide-ConsoleWindow
        
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
        
        $script:isConnecting = $false
        Update-ConnectionUI
        return $true
    } catch {
        $script:isConnecting = $false
        Update-ConnectionUI
        
        $errorMsg = $_.Exception.Message
        if ($errorMsg -like "*canceled*" -or $errorMsg -like "*cancelled*" -or $errorMsg -like "*user*denied*") {
            [System.Windows.Forms.MessageBox]::Show("Connection was canceled or denied. If authentication appears stuck, use the 'Reset Auth' button to clear the session.", "Connection Canceled", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } elseif ($errorMsg -like "*timeout*" -or $errorMsg -like "*timed out*") {
            [System.Windows.Forms.MessageBox]::Show("Connection timed out. This may indicate a network issue or the authentication dialog didn't appear. Use the 'Reset Auth' button to clear the session and try again.", "Connection Timeout", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        } else {
            [System.Windows.Forms.MessageBox]::Show("Failed to connect: $errorMsg`n`nIf authentication appears stuck, use the 'Reset Auth' button to clear the session.", "Connection Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
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
    $connectButton.BackColor = [System.Drawing.Color]::LightGreen
    $connectButton.Add_Click({
        # Prevent multiple connection attempts
        if ($script:isConnecting) {
            [System.Windows.Forms.MessageBox]::Show("A connection attempt is already in progress. Please wait or use 'Reset Auth' if it appears stuck.", "Connection In Progress", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
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
    $cancelButton.BackColor = [System.Drawing.Color]::LightGray
    $cancelButton.Add_Click({ $form.Close() })
    $form.Controls.Add($cancelButton)

    $form.ShowDialog() | Out-Null
}

function Update-ConnectionUI {
    if ($script:statusLabel) {
        if ($script:isConnecting) {
            $script:statusLabel.Text = "Connecting..."
            $script:statusLabel.ForeColor = [System.Drawing.Color]::Orange
        } elseif ($global:isConnected) {
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
        $script:connectButton.Enabled = -not $global:isConnected -and -not $script:isConnecting
    }
    
    if ($script:disconnectButton) {
        $script:disconnectButton.Enabled = $global:isConnected -and -not $script:isConnecting
    }
    
    if ($script:reconnectButton) {
        $script:reconnectButton.Enabled = -not $script:isConnecting
    }
    
    if ($script:resetAuthButton) {
        $script:resetAuthButton.Enabled = $true  # Always enabled - useful when stuck
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
        $locations = Get-MgIdentityConditionalAccessNamedLocation -All -ErrorAction Stop
        
        # Get all policies to check references (cache for performance)
        $allPolicies = @()
        try {
            $allPolicies = Get-MgIdentityConditionalAccessPolicy -All -ErrorAction SilentlyContinue
        } catch {
            Write-Host "Warning: Could not load policies for reference checking: $_" -ForegroundColor Yellow
        }
        
        foreach ($location in $locations) {
            try {
                $item = New-Object System.Windows.Forms.ListViewItem($location.DisplayName)
                
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
                
                # Find policies that reference this location
                $referencingPolicies = @()
                foreach ($policy in $allPolicies) {
                    if ($policy.Conditions -and $policy.Conditions.Locations) {
                        $isReferenced = $false
                        if ($policy.Conditions.Locations.IncludeLocations) {
                            $includeLocs = $policy.Conditions.Locations.IncludeLocations
                            if ($includeLocs -is [string]) { $includeLocs = @($includeLocs) }
                            if ($includeLocs -contains $location.Id) {
                                $isReferenced = $true
                            }
                        }
                        if ($policy.Conditions.Locations.ExcludeLocations) {
                            $excludeLocs = $policy.Conditions.Locations.ExcludeLocations
                            if ($excludeLocs -is [string]) { $excludeLocs = @($excludeLocs) }
                            if ($excludeLocs -contains $location.Id) {
                                $isReferenced = $true
                            }
                        }
                        if ($isReferenced) {
                            $referencingPolicies += $policy.DisplayName
                        }
                    }
                }
                
                if ($referencingPolicies.Count -gt 0) {
                    $item.SubItems.Add(($referencingPolicies -join '; ')) | Out-Null
                } else {
                    $item.SubItems.Add("None") | Out-Null
                }
                
                $item.Tag = $location
                $listView.Items.Add($item) | Out-Null
            } catch {
                # Error processing location - continue with others
                # Continue with other locations
            }
        }
        # Loaded locations successfully
    } catch {
        # Error loading Named Locations
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
        # Check if form and controls are still valid before accessing them
        if ($null -eq $form -or $form.IsDisposed -or $null -eq $countriesListBox -or $null -eq $selectionLabel) {
            return
        }
        try {
            $selectedCount = $countriesListBox.CheckedItems.Count
            $selectionLabel.Text = "$selectedCount countries selected"
        } catch {
            # Silently ignore errors if controls are disposed
        }
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
        param($sender, $e)
        
        # Update count synchronously by calculating what it will be after the change
        # ItemCheck fires BEFORE the state changes, so we account for the pending change
        try {
            if ($null -ne $countriesListBox -and $null -ne $selectionLabel -and 
                $null -ne $form -and -not $form.IsDisposed) {
                $currentCount = $countriesListBox.CheckedItems.Count
                
                # Adjust count based on whether item is being checked or unchecked
                if ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked) {
                    # Item is being checked, so count will increase by 1
                    $newCount = $currentCount + 1
                } else {
                    # Item is being unchecked, so count will decrease by 1
                    $newCount = $currentCount - 1
                }
                
                $selectionLabel.Text = "$newCount countries selected"
            }
        } catch {
            # Silently ignore errors if controls are disposed
        }
    })

    # OK and Cancel buttons
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Location = New-Object System.Drawing.Point(420, 420)
    $okButton.Size = New-Object System.Drawing.Size(75, 30)
    $okButton.BackColor = [System.Drawing.Color]::LightGreen
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(505, 420)
    $cancelButton.Size = New-Object System.Drawing.Size(75, 30)
    $cancelButton.BackColor = [System.Drawing.Color]::LightGray
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
    $selectCountriesButton.BackColor = [System.Drawing.Color]::LightSteelBlue
    $selectCountriesButton.Add_Click({
        # Get current countries from textbox
        $currentCountries = @()
        if (-not [string]::IsNullOrWhiteSpace($countriesTextBox.Text)) {
            $currentCountries = $countriesTextBox.Text.Split(',') | 
                ForEach-Object { $_.Trim().ToUpper() } | 
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        }
        
        # Open country selection dialog
        $selectedCountries = Show-CountrySelectionDialog -PreselectedCountries $currentCountries
        
        if ($selectedCountries -and $selectedCountries.Count -gt 0) {
            $countriesTextBox.Text = ($selectedCountries -join ', ')
        } else {
            # Clear if no countries selected
            $countriesTextBox.Text = ""
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
    if ($Mode -eq "Create") {
        $actionButton.Text = "Create"
        $actionButton.BackColor = [System.Drawing.Color]::LightGreen
    } else {
        $actionButton.Text = "Apply"
        $actionButton.BackColor = [System.Drawing.Color]::Orange
    }
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
            # Parse and validate country codes
            $countryCodes = $countriesTextBox.Text.Split(',') | 
                ForEach-Object { $_.Trim().ToUpper() } | 
                Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            
            # Validate we have at least one country code
            if ($countryCodes.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("Please select at least one country.", "Validation Error")
                return
            }
            
            # Validate country codes are 2-letter ISO codes
            $invalidCodes = $countryCodes | Where-Object { $_.Length -ne 2 -or -not ($_ -match '^[A-Z]{2}$') }
            if ($invalidCodes.Count -gt 0) {
                [System.Windows.Forms.MessageBox]::Show("Invalid country codes detected: $($invalidCodes -join ', ')`n`nCountry codes must be 2-letter ISO codes (e.g., US, CA, GB).", "Validation Error")
                return
            }
            
            $success = $false
            
            if ($Mode -eq "Edit") {
                # Try using the PowerShell cmdlet first (handles formatting better)
                try {
                    $updateParams = @{
                        NamedLocationId = $ExistingLocation.Id
                        BodyParameter = @{
                            "@odata.type" = "#microsoft.graph.countryNamedLocation"
                            displayName = $nameTextBox.Text
                            countriesAndRegions = $countryCodes
                            includeUnknownCountriesAndRegions = $includeUnknownCheckBox.Checked
                        }
                    }
                    
                    Write-Host ("Updating location " + $ExistingLocation.Id + " with " + $countryCodes.Count + " countries using cmdlet") -ForegroundColor Cyan
                    Update-MgIdentityConditionalAccessNamedLocation @updateParams -ErrorAction Stop
                    $success = $true
                    [System.Windows.Forms.MessageBox]::Show("Named Location updated successfully!", "Success")
                } catch {
                    # If cmdlet fails, fall back to REST API
                    Write-Host ("Cmdlet failed, trying REST API: " + $_.Exception.Message) -ForegroundColor Yellow
                    
                    # Update existing location using PATCH REST API
                    $updateBody = @{
                        displayName = $nameTextBox.Text
                        countriesAndRegions = $countryCodes
                        includeUnknownCountriesAndRegions = $includeUnknownCheckBox.Checked
                    }
                    
                    $jsonBody = $updateBody | ConvertTo-Json -Depth 10
                    
                    Write-Host ("Updating via REST API with " + $countryCodes.Count + " countries") -ForegroundColor Cyan
                    $uri = "https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations/" + $ExistingLocation.Id
                    $response = Invoke-MgGraphRequest -Method PATCH -Uri $uri -Body $jsonBody -ContentType "application/json" -ErrorAction Stop
                    Write-Host ("Update response: " + ($response | ConvertTo-Json -Depth 3)) -ForegroundColor Green
                    $success = $true
                    [System.Windows.Forms.MessageBox]::Show("Named Location updated successfully!", "Success")
                }
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
            $errorMessage += "- Countries: " + ($countryCodes.Count) + " countries selected`n"
            $errorMessage += "- Include Unknown: " + $includeUnknownCheckBox.Checked + "`n`n"
            
            # Try to get detailed error response from Microsoft Graph
            $apiErrorDetails = $null
            if ($_.Exception -is [Microsoft.Graph.PowerShell.Runtime.GraphException]) {
                try {
                    $graphError = $_.Exception
                    if ($graphError.ErrorResponse) {
                        $apiErrorDetails = $graphError.ErrorResponse | ConvertTo-Json -Depth 5
                    }
                } catch {
                    # Try alternative method to get error details
                    try {
                        $apiErrorDetails = $_.Exception.ToString()
                    } catch {}
                }
            } elseif ($_.Exception.Response) {
                try {
                    $responseStream = $_.Exception.Response.GetResponseStream()
                    if ($responseStream) {
                        $reader = New-Object System.IO.StreamReader($responseStream)
                        $responseBody = $reader.ReadToEnd()
                        $reader.Close()
                        if ($responseBody) {
                            # Try to parse as JSON for better formatting
                            try {
                                $errorObj = $responseBody | ConvertFrom-Json
                                $apiErrorDetails = $errorObj | ConvertTo-Json -Depth 5
                            } catch {
                                $apiErrorDetails = $responseBody
                            }
                        }
                    }
                } catch {
                    # Ignore errors reading response
                }
            }
            
            if ($apiErrorDetails) {
                $errorMessage += "API Error Details:`n" + $apiErrorDetails + "`n`n"
            }
            
            if ($_.Exception.Message -like "*BadRequest*") {
                $errorMessage += "Common fixes:`n"
                $errorMessage += "* Use valid 2-letter ISO country codes: US, CA, GB, DE, FR`n"
                $errorMessage += "* Remove special characters from display name`n"
                $errorMessage += "* Ensure country codes are valid ISO 3166-1 alpha-2 codes`n"
                $errorMessage += "* Check that you have the required permissions`n"
                $errorMessage += "* If selecting many countries (" + $countryCodes.Count + "), Microsoft Graph may have limits`n"
                $errorMessage += "* Try selecting fewer countries or splitting into multiple locations"
            }
            
            Write-Host ("ERROR: " + $errorMessage) -ForegroundColor Red
            [System.Windows.Forms.MessageBox]::Show($errorMessage, "Error Details", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    $form.Controls.Add($actionButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(395, 215)
    $cancelButton.Size = New-Object System.Drawing.Size(75, 23)
    $cancelButton.BackColor = [System.Drawing.Color]::LightGray
    $cancelButton.Add_Click({ $form.Close() })
    $form.Controls.Add($cancelButton)

    $form.ShowDialog() | Out-Null
}

function Show-CreateCountryLocationDialog {
    param($listView)
    Show-CountryLocationDialog -listView $listView -Mode "Create"
}

function Edit-SelectedNamedLocation {
    param($listView)
    
    if (-not $global:isConnected) {
        [System.Windows.Forms.MessageBox]::Show("Please connect to Microsoft Graph first.", "Not Connected")
        return
    }
    
    if ($listView.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a Named Location to edit.", "No Selection")
        return
    }

    $selectedItem = $listView.SelectedItems[0]
    $location = $selectedItem.Tag
    
    # Check if it is a country-based location
    $odataType = $location.AdditionalProperties.'@odata.type'
    if ($odataType -ne '#microsoft.graph.countryNamedLocation') {
        [System.Windows.Forms.MessageBox]::Show("Only country-based Named Locations can be edited with this tool.", "Not Supported")
        return
    }
    
    Show-CountryLocationDialog -listView $listView -Mode "Edit" -ExistingLocation $location
}

function Copy-SelectedNamedLocation {
    param($listView)
    
    if (-not $global:isConnected) {
        [System.Windows.Forms.MessageBox]::Show("Please connect to Microsoft Graph first.", "Not Connected")
        return
    }
    
    if ($listView.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a Named Location to copy.", "No Selection")
        return
    }

    $selectedItem = $listView.SelectedItems[0]
    $location = $selectedItem.Tag
    
    # Check if it is a country-based location
    $odataType = $location.AdditionalProperties.'@odata.type'
    if ($odataType -ne '#microsoft.graph.countryNamedLocation') {
        [System.Windows.Forms.MessageBox]::Show("Only country-based Named Locations can be copied with this tool.", "Not Supported")
        return
    }
    
    # Get the country codes and settings from the source
    $sourceCountries = $location.AdditionalProperties['countriesAndRegions']
    $sourceIncludeUnknown = $location.AdditionalProperties['includeUnknownCountriesAndRegions']
    
    # Prompt for new name
    $promptMessage = "Enter a name for the new Named Location:`n`nThis will copy the country codes (" + ($sourceCountries -join ', ') + ") and settings from '" + $location.DisplayName + "'"
    $defaultName = "Copy of " + $location.DisplayName
    
    $newName = Show-InputBox -Prompt $promptMessage -Title "Copy Named Location" -DefaultValue $defaultName
    
    if ([string]::IsNullOrWhiteSpace($newName)) {
        return
    }
    
    try {
        # Try the PowerShell cmdlet approach first (often more reliable)
        try {
            Write-Host "Attempting copy with PowerShell cmdlet..." -ForegroundColor Yellow
            
            $params = @{
                "@odata.type" = "#microsoft.graph.countryNamedLocation"
                DisplayName = $newName
                CountriesAndRegions = $sourceCountries
                IncludeUnknownCountriesAndRegions = $sourceIncludeUnknown
            }
            
            Write-Host ("PowerShell params: " + ($params | ConvertTo-Json)) -ForegroundColor Cyan
            New-MgIdentityConditionalAccessNamedLocation -BodyParameter $params
            
            Write-Host "PowerShell cmdlet copy successful!" -ForegroundColor Green
            [System.Windows.Forms.MessageBox]::Show("Named Location copied successfully as '" + $newName + "'!", "Success")
            
            # Small delay to allow Microsoft Graph to process the new location
            Start-Sleep -Seconds 1
            Refresh-NamedLocationsList $listView
            return
        } catch {
            Write-Host ("PowerShell cmdlet failed: " + $_.Exception.Message) -ForegroundColor Red
            Write-Host "Trying REST API approach..." -ForegroundColor Yellow
        }
        
        # Fallback to REST API with cleaned data
        $createParams = @{
            "@odata.type" = "#microsoft.graph.countryNamedLocation"
            displayName = $newName
            countriesAndRegions = $sourceCountries
            includeUnknownCountriesAndRegions = [bool]$sourceIncludeUnknown
        }
        
        $jsonBody = $createParams | ConvertTo-Json -Depth 10
        Write-Host ("Creating copy with REST API: " + $jsonBody) -ForegroundColor Cyan
        
        $uri = "https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations"
        $response = Invoke-MgGraphRequest -Method POST -Uri $uri -Body $jsonBody -ContentType "application/json"
        
        Write-Host ("REST API copy successful: " + ($response | ConvertTo-Json -Depth 3)) -ForegroundColor Green
        [System.Windows.Forms.MessageBox]::Show("Named Location copied successfully as '" + $newName + "'!", "Success")
        
        # Small delay to allow Microsoft Graph to process the new location
        Start-Sleep -Seconds 1
        Refresh-NamedLocationsList $listView
    } catch {
        $errorMessage = "Error copying Named Location:`n`n"
        $errorMessage += "Error: " + $_.Exception.Message + "`n`n"
        $errorMessage += "Source: " + $location.DisplayName + "`n"
        $errorMessage += "Target Name: " + $newName + "`n"
        $errorMessage += "Countries: " + ($sourceCountries -join ', ') + "`n"
        $errorMessage += "Include Unknown: " + $sourceIncludeUnknown + "`n`n"
        $errorMessage += "Try using a simpler name without special characters."
        
        Write-Host ("Copy error: " + $errorMessage) -ForegroundColor Red
        [System.Windows.Forms.MessageBox]::Show($errorMessage, "Copy Error")
    }
}

function Rename-SelectedNamedLocation {
    param($listView)
    
    if (-not $global:isConnected) {
        [System.Windows.Forms.MessageBox]::Show("Please connect to Microsoft Graph first.", "Not Connected")
        return
    }
    
    if ($listView.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a Named Location to rename.", "No Selection")
        return
    }

    $selectedItem = $listView.SelectedItems[0]
    $currentName = $selectedItem.Text
    $location = $selectedItem.Tag
    $locationId = $location.Id

    $newName = Show-InputBox -Prompt "Enter new display name:" -Title "Rename Named Location" -DefaultValue $currentName
    
    if ([string]::IsNullOrWhiteSpace($newName) -or $newName -eq $currentName) {
        return
    }

    try {
        # Check if the location still exists before trying to rename
        $existingLocation = Get-MgIdentityConditionalAccessNamedLocation -NamedLocationId $locationId -ErrorAction SilentlyContinue
        if (-not $existingLocation) {
            [System.Windows.Forms.MessageBox]::Show("The selected Named Location no longer exists. Refreshing list.", "Not Found")
            Refresh-NamedLocationsList $listView
            return
        }
        
        # Get the odata type from the existing location
        $odataType = $existingLocation.AdditionalProperties.'@odata.type'
        if (-not $odataType) {
            # Fallback: determine type from the location object
            if ($existingLocation.AdditionalProperties.'countriesAndRegions') {
                $odataType = "#microsoft.graph.countryNamedLocation"
            } elseif ($existingLocation.AdditionalProperties.'ipRanges') {
                $odataType = "#microsoft.graph.ipNamedLocation"
            } else {
                $odataType = "#microsoft.graph.namedLocation"
            }
        }
        
        # Use BodyParameter to include the odata type
        $updateParams = @{
            NamedLocationId = $locationId
            BodyParameter = @{
                "@odata.type" = $odataType
                displayName = $newName
            }
        }
        
        Update-MgIdentityConditionalAccessNamedLocation @updateParams
        [System.Windows.Forms.MessageBox]::Show("Named Location renamed successfully!", "Success")
        
        # Small delay before refresh
        Start-Sleep -Seconds 1
        Refresh-NamedLocationsList $listView
    } catch {
        if ($_.Exception.Message -like "*NotFound*" -or $_.Exception.Message -like "*404*") {
            [System.Windows.Forms.MessageBox]::Show("The Named Location no longer exists. Refreshing list.", "Not Found")
            Refresh-NamedLocationsList $listView
        } else {
            [System.Windows.Forms.MessageBox]::Show("Error renaming Named Location: $_", "Error")
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
        [System.Windows.Forms.MessageBox]::Show("Please select one or more Named Locations to delete.", "No Selection")
        return
    }

    $selectedCount = $listView.SelectedItems.Count
    $confirmMessage = "Are you sure you want to delete $selectedCount Named Location(s)?`n`n"
    if ($selectedCount -eq 1) {
        $confirmMessage += "Location: " + $listView.SelectedItems[0].Text
    } else {
        $confirmMessage += "Locations:`n"
        foreach ($item in $listView.SelectedItems) {
            $confirmMessage += "  - " + $item.Text + "`n"
        }
    }
    $confirmMessage += "`nThis action cannot be undone!"
    
    $result = [System.Windows.Forms.MessageBox]::Show($confirmMessage, "Confirm Delete", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        # Load all policies once for reference checking
        Write-Host "Loading policies for reference checking..." -ForegroundColor Yellow
        $allPolicies = @()
        try {
            $allPolicies = Get-MgIdentityConditionalAccessPolicy -All -ErrorAction Stop
        } catch {
            Write-Host "Warning: Could not load policies for reference checking: $_" -ForegroundColor Yellow
        }
        
        $deletedCount = 0
        $skippedCount = 0
        $errorCount = 0
        $referencedLocations = @()
        $errors = @()
        
        # Process each selected location
        foreach ($selectedItem in $listView.SelectedItems) {
            $locationName = $selectedItem.Text
            $location = $selectedItem.Tag
            $locationId = $location.Id
            
            try {
                # Check if the location still exists before trying to delete
                $existingLocation = Get-MgIdentityConditionalAccessNamedLocation -NamedLocationId $locationId -ErrorAction SilentlyContinue
                if (-not $existingLocation) {
                    Write-Host "Skipping '$locationName' - already deleted." -ForegroundColor Yellow
                    $skippedCount++
                    continue
                }
                
                # Check if any policies reference this location
                $referencingPolicies = @()
                foreach ($policy in $allPolicies) {
                    $hasReference = $false
                    if ($policy.Conditions -and $policy.Conditions.Locations) {
                        if ($policy.Conditions.Locations.IncludeLocations) {
                            $includeLocs = $policy.Conditions.Locations.IncludeLocations
                            if ($includeLocs -is [string]) { $includeLocs = @($includeLocs) }
                            if ($includeLocs -contains $locationId) {
                                $hasReference = $true
                            }
                        }
                        if ($policy.Conditions.Locations.ExcludeLocations) {
                            $excludeLocs = $policy.Conditions.Locations.ExcludeLocations
                            if ($excludeLocs -is [string]) { $excludeLocs = @($excludeLocs) }
                            if ($excludeLocs -contains $locationId) {
                                $hasReference = $true
                            }
                        }
                    }
                    if ($hasReference) {
                        $referencingPolicies += $policy
                    }
                }
                
                if ($referencingPolicies.Count -gt 0) {
                    $policyNames = $referencingPolicies | ForEach-Object { $_.DisplayName }
                    $referencedLocations += [PSCustomObject]@{
                        LocationName = $locationName
                        PolicyNames = $policyNames
                    }
                    # Skipping - referenced by policies
                    $skippedCount++
                    continue
                }
                
                # Delete the location
                Remove-MgIdentityConditionalAccessNamedLocation -NamedLocationId $locationId -ErrorAction Stop
                # Deleted successfully
                $deletedCount++
                
                # Small delay between deletions to avoid rate limiting
                if ($deletedCount -lt $selectedCount) {
                    Start-Sleep -Milliseconds 500
                }
                
            } catch {
                if ($_.Exception.Message -like "*NotFound*" -or $_.Exception.Message -like "*404*" -or $_.Exception.Message -like "*does not exist*") {
                    Write-Host "Skipping '$locationName' - already deleted." -ForegroundColor Yellow
                    $skippedCount++
                } elseif ($_.Exception.Message -like "*referenced*" -or $_.Exception.Message -like "*1178*") {
                    # Skipping - referenced by policies
                    $skippedCount++
                } else {
                    # Error deleting location
                    $errors += "$locationName : $($_.Exception.Message)"
                    $errorCount++
                }
            }
        }
        
        # Show summary message
        $summaryMessage = "Deletion Summary:`n`n"
        $summaryMessage += "Successfully deleted: $deletedCount`n"
        if ($skippedCount -gt 0) {
            $summaryMessage += "Skipped: $skippedCount`n"
        }
        if ($errorCount -gt 0) {
            $summaryMessage += "Errors: $errorCount`n"
        }
        
        if ($referencedLocations.Count -gt 0) {
            $summaryMessage += "`nLocations skipped due to policy references:`n"
            foreach ($ref in $referencedLocations) {
                $summaryMessage += "  - $($ref.LocationName): $($ref.PolicyNames -join ', ')`n"
            }
        }
        
        if ($errors.Count -gt 0) {
            $summaryMessage += "`nErrors encountered:`n"
            foreach ($err in $errors) {
                $summaryMessage += "  - $err`n"
            }
        }
        
        if ($deletedCount -gt 0) {
            [System.Windows.Forms.MessageBox]::Show($summaryMessage, "Deletion Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        } else {
            [System.Windows.Forms.MessageBox]::Show($summaryMessage, "No Locations Deleted", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        }
        
        # Refresh the list
        Start-Sleep -Seconds 1
        Refresh-NamedLocationsList $listView
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
        # Cache all locations for reference resolution
        Write-Host "Loading locations for reference checking..." -ForegroundColor Yellow
        $allLocations = @{}
        try {
            $locations = Get-MgIdentityConditionalAccessNamedLocation -All -ErrorAction SilentlyContinue
            foreach ($loc in $locations) {
                $allLocations[$loc.Id] = $loc.DisplayName
            }
            # Also add special locations
            $allLocations["All"] = "All"
            $allLocations["AllTrusted"] = "All Trusted IPs"
        } catch {
            Write-Host "Warning: Could not load locations for reference checking: $_" -ForegroundColor Yellow
        }
        
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
            
            # Get referenced locations
            $referencedLocations = @()
            if ($policy.Conditions -and $policy.Conditions.Locations) {
                if ($policy.Conditions.Locations.IncludeLocations) {
                    $includeLocs = $policy.Conditions.Locations.IncludeLocations
                    if ($includeLocs -is [string]) { $includeLocs = @($includeLocs) }
                    foreach ($locId in $includeLocs) {
                        if ($allLocations.ContainsKey($locId)) {
                            $referencedLocations += $allLocations[$locId] + " (Include)"
                        } else {
                            $referencedLocations += $locId + " (Include)"
                        }
                    }
                }
                if ($policy.Conditions.Locations.ExcludeLocations) {
                    $excludeLocs = $policy.Conditions.Locations.ExcludeLocations
                    if ($excludeLocs -is [string]) { $excludeLocs = @($excludeLocs) }
                    foreach ($locId in $excludeLocs) {
                        if ($allLocations.ContainsKey($locId)) {
                            $referencedLocations += $allLocations[$locId] + " (Exclude)"
                        } else {
                            $referencedLocations += $locId + " (Exclude)"
                        }
                    }
                }
            }
            
            if ($referencedLocations.Count -gt 0) {
                $item.SubItems.Add(($referencedLocations -join '; ')) | Out-Null
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
            
            # Perform the deletion with error handling
            try {
                Remove-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policyId -Confirm:$false -ErrorAction Stop
            } catch {
                # Check if error is "does not exist" - treat as success since policy is already gone
                if ($_.Exception.Message -like "*NotFound*" -or $_.Exception.Message -like "*404*" -or $_.Exception.Message -like "*does not exist*") {
                    Write-Host "Policy was already deleted (does not exist). Treating as successful deletion." -ForegroundColor Yellow
                    # Continue to success message below
                } else {
                    throw  # Re-throw other errors to be caught by outer catch
                }
            }
            
            # Verify deletion succeeded by checking if policy still exists
            Start-Sleep -Milliseconds 500  # Brief delay for API propagation
            $verifyPolicy = Get-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policyId -ErrorAction SilentlyContinue
            if ($verifyPolicy) {
                throw "Policy deletion appeared to succeed but policy still exists. It may have been recreated or deletion is pending."
            }
            
            Write-Host "✅ Policy deleted successfully!" -ForegroundColor Green
            [System.Windows.Forms.MessageBox]::Show("Conditional Access Policy '$policyName' deleted successfully!", "Success")
            
            # Small delay before refresh
            Start-Sleep -Seconds 1
            Refresh-PoliciesList $listView
            
        } catch {
            Write-Host "❌ Error deleting policy: $($_.Exception.Message)" -ForegroundColor Red
            
            if ($_.Exception.Message -like "*NotFound*" -or $_.Exception.Message -like "*404*" -or $_.Exception.Message -like "*does not exist*") {
                [System.Windows.Forms.MessageBox]::Show("The Conditional Access Policy '$policyName' was already deleted or does not exist. Refreshing list.", "Already Deleted")
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
    
    if (-not $global:isConnected) {
        [System.Windows.Forms.MessageBox]::Show("Please connect to Microsoft Graph first.", "Not Connected")
        return
    }
    
    if ($listView.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a Conditional Access Policy to copy.", "No Selection")
        return
    }

    $selectedItem = $listView.SelectedItems[0]
    $policy = $selectedItem.Tag
    $policyId = $policy.Id
    $sourceName = $policy.DisplayName
    
    # Prompt for new name
    $promptMessage = "Enter a name for the new Conditional Access Policy:`n`nThis will copy all settings from '$sourceName'"
    $defaultName = "Copy of $sourceName"
    
    $newName = Show-InputBox -Prompt $promptMessage -Title "Copy Conditional Access Policy" -DefaultValue $defaultName
    
    if ([string]::IsNullOrWhiteSpace($newName)) {
        return
    }
    
    try {
        Write-Host "Copying Conditional Access Policy: $sourceName" -ForegroundColor Yellow
        
        # Get the full policy details
        $fullPolicy = Get-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policyId
        
        if (-not $fullPolicy) {
            [System.Windows.Forms.MessageBox]::Show("The selected policy no longer exists. Refreshing list.", "Not Found")
            Refresh-PoliciesList $listView
            return
        }
        
        Write-Host "Analyzing source policy structure..." -ForegroundColor Cyan
        
        # Start with a minimal working policy structure
        $newPolicyBody = @{
            displayName = $newName
            state = "disabled"
            conditions = @{
                users = @{
                    includeUsers = @("All")
                }
                applications = @{
                    includeApplications = @("All")  
                }
            }
            grantControls = @{
                operator = "OR"
                builtInControls = @("block")
            }
        }
        
        # Try to copy conditions if they exist and are valid
        if ($fullPolicy.Conditions) {
            Write-Host "Copying policy conditions..." -ForegroundColor Cyan
            
            # Copy users
            if ($fullPolicy.Conditions.Users) {
                $userConditions = @{}
                if ($fullPolicy.Conditions.Users.IncludeUsers -and $fullPolicy.Conditions.Users.IncludeUsers.Count -gt 0) {
                    $includeUsers = @($fullPolicy.Conditions.Users.IncludeUsers | Where-Object { $_ -ne $null -and $_ -ne "" })
                    if ($includeUsers.Count -gt 0) {
                        $userConditions.includeUsers = $includeUsers
                    }
                }
                if ($fullPolicy.Conditions.Users.ExcludeUsers -and $fullPolicy.Conditions.Users.ExcludeUsers.Count -gt 0) {
                    $excludeUsers = @($fullPolicy.Conditions.Users.ExcludeUsers | Where-Object { $_ -ne $null -and $_ -ne "" })
                    if ($excludeUsers.Count -gt 0) {
                        $userConditions.excludeUsers = $excludeUsers
                    }
                }
                if ($fullPolicy.Conditions.Users.IncludeGroups -and $fullPolicy.Conditions.Users.IncludeGroups.Count -gt 0) {
                    $includeGroups = @($fullPolicy.Conditions.Users.IncludeGroups | Where-Object { $_ -ne $null -and $_ -ne "" })
                    if ($includeGroups.Count -gt 0) {
                        $userConditions.includeGroups = $includeGroups
                    }
                }
                if ($fullPolicy.Conditions.Users.ExcludeGroups -and $fullPolicy.Conditions.Users.ExcludeGroups.Count -gt 0) {
                    $excludeGroups = @($fullPolicy.Conditions.Users.ExcludeGroups | Where-Object { $_ -ne $null -and $_ -ne "" })
                    if ($excludeGroups.Count -gt 0) {
                        $userConditions.excludeGroups = $excludeGroups
                    }
                }
                if ($userConditions.Count -gt 0) {
                    $newPolicyBody.conditions.users = $userConditions
                }
            }
            
            # Copy applications
            if ($fullPolicy.Conditions.Applications) {
                $appConditions = @{}
                if ($fullPolicy.Conditions.Applications.IncludeApplications -and $fullPolicy.Conditions.Applications.IncludeApplications.Count -gt 0) {
                    $includeApps = @($fullPolicy.Conditions.Applications.IncludeApplications | Where-Object { $_ -ne $null -and $_ -ne "" })
                    if ($includeApps.Count -gt 0) {
                        $appConditions.includeApplications = $includeApps
                    }
                }
                if ($fullPolicy.Conditions.Applications.ExcludeApplications -and $fullPolicy.Conditions.Applications.ExcludeApplications.Count -gt 0) {
                    $excludeApps = @($fullPolicy.Conditions.Applications.ExcludeApplications | Where-Object { $_ -ne $null -and $_ -ne "" })
                    if ($excludeApps.Count -gt 0) {
                        $appConditions.excludeApplications = $excludeApps
                    }
                }
                if ($appConditions.Count -gt 0) {
                    $newPolicyBody.conditions.applications = $appConditions
                }
            }
            
            # Copy locations if they exist
            if ($fullPolicy.Conditions.Locations) {
                $locationConditions = @{}
                if ($fullPolicy.Conditions.Locations.IncludeLocations -and $fullPolicy.Conditions.Locations.IncludeLocations.Count -gt 0) {
                    $includeLocations = @($fullPolicy.Conditions.Locations.IncludeLocations | Where-Object { $_ -ne $null -and $_ -ne "" })
                    if ($includeLocations.Count -gt 0) {
                        $locationConditions.includeLocations = $includeLocations
                    }
                }
                if ($fullPolicy.Conditions.Locations.ExcludeLocations -and $fullPolicy.Conditions.Locations.ExcludeLocations.Count -gt 0) {
                    $excludeLocations = @($fullPolicy.Conditions.Locations.ExcludeLocations | Where-Object { $_ -ne $null -and $_ -ne "" })
                    if ($excludeLocations.Count -gt 0) {
                        $locationConditions.excludeLocations = $excludeLocations
                    }
                }
                if ($locationConditions.Count -gt 0) {
                    $newPolicyBody.conditions.locations = $locationConditions
                }
            }
            
            # Copy platforms if they exist
            if ($fullPolicy.Conditions.Platforms) {
                $platformConditions = @{}
                if ($fullPolicy.Conditions.Platforms.IncludePlatforms -and $fullPolicy.Conditions.Platforms.IncludePlatforms.Count -gt 0) {
                    $includePlatforms = @($fullPolicy.Conditions.Platforms.IncludePlatforms | Where-Object { $_ -ne $null -and $_ -ne "" })
                    if ($includePlatforms.Count -gt 0) {
                        $platformConditions.includePlatforms = $includePlatforms
                    }
                }
                if ($fullPolicy.Conditions.Platforms.ExcludePlatforms -and $fullPolicy.Conditions.Platforms.ExcludePlatforms.Count -gt 0) {
                    $excludePlatforms = @($fullPolicy.Conditions.Platforms.ExcludePlatforms | Where-Object { $_ -ne $null -and $_ -ne "" })
                    if ($excludePlatforms.Count -gt 0) {
                        $platformConditions.excludePlatforms = $excludePlatforms
                    }
                }
                if ($platformConditions.Count -gt 0) {
                    $newPolicyBody.conditions.platforms = $platformConditions
                }
            }
        }
        
        # Try to copy grant controls
        if ($fullPolicy.GrantControls) {
            Write-Host "Copying grant controls..." -ForegroundColor Cyan
            $grantControls = @{}
            
            # Operator is required
            if ($fullPolicy.GrantControls.Operator) {
                $grantControls.operator = $fullPolicy.GrantControls.Operator
            } else {
                $grantControls.operator = "OR"  # Default fallback
            }
            
            if ($fullPolicy.GrantControls.BuiltInControls -and $fullPolicy.GrantControls.BuiltInControls.Count -gt 0) {
                $builtInControls = @($fullPolicy.GrantControls.BuiltInControls | Where-Object { $_ -ne $null -and $_ -ne "" })
                if ($builtInControls.Count -gt 0) {
                    $grantControls.builtInControls = $builtInControls
                }
            }
            
            if ($fullPolicy.GrantControls.CustomAuthenticationFactors -and $fullPolicy.GrantControls.CustomAuthenticationFactors.Count -gt 0) {
                $customFactors = @($fullPolicy.GrantControls.CustomAuthenticationFactors | Where-Object { $_ -ne $null -and $_ -ne "" })
                if ($customFactors.Count -gt 0) {
                    $grantControls.customAuthenticationFactors = $customFactors
                }
            }
            
            if ($fullPolicy.GrantControls.TermsOfUse -and $fullPolicy.GrantControls.TermsOfUse.Count -gt 0) {
                $termsOfUse = @($fullPolicy.GrantControls.TermsOfUse | Where-Object { $_ -ne $null -and $_ -ne "" })
                if ($termsOfUse.Count -gt 0) {
                    $grantControls.termsOfUse = $termsOfUse
                }
            }
            
            $newPolicyBody.grantControls = $grantControls
            Write-Host "Added grant controls with operator: $($grantControls.operator)" -ForegroundColor Gray
        }
        
        # Try to copy session controls (simplified)
        if ($fullPolicy.SessionControls) {
            Write-Host "Copying session controls..." -ForegroundColor Cyan
            $sessionControls = @{}
            
            if ($fullPolicy.SessionControls.ApplicationEnforcedRestrictions -and 
                $fullPolicy.SessionControls.ApplicationEnforcedRestrictions.IsEnabled -ne $null) {
                $sessionControls.applicationEnforcedRestrictions = @{
                    isEnabled = $fullPolicy.SessionControls.ApplicationEnforcedRestrictions.IsEnabled
                }
            }
            
            if ($fullPolicy.SessionControls.CloudAppSecurity -and 
                $fullPolicy.SessionControls.CloudAppSecurity.IsEnabled -ne $null) {
                $cloudAppSecurity = @{
                    isEnabled = $fullPolicy.SessionControls.CloudAppSecurity.IsEnabled
                }
                if ($fullPolicy.SessionControls.CloudAppSecurity.CloudAppSecurityType) {
                    $cloudAppSecurity.cloudAppSecurityType = $fullPolicy.SessionControls.CloudAppSecurity.CloudAppSecurityType
                }
                $sessionControls.cloudAppSecurity = $cloudAppSecurity
            }
            
            if ($fullPolicy.SessionControls.SignInFrequency -and 
                $fullPolicy.SessionControls.SignInFrequency.IsEnabled -ne $null) {
                $signInFreq = @{
                    isEnabled = $fullPolicy.SessionControls.SignInFrequency.IsEnabled
                }
                if ($fullPolicy.SessionControls.SignInFrequency.Type) {
                    $signInFreq.type = $fullPolicy.SessionControls.SignInFrequency.Type
                }
                if ($fullPolicy.SessionControls.SignInFrequency.Value -ne $null) {
                    $signInFreq.value = $fullPolicy.SessionControls.SignInFrequency.Value
                }
                $sessionControls.signInFrequency = $signInFreq
            }
            
            if ($fullPolicy.SessionControls.PersistentBrowser -and 
                $fullPolicy.SessionControls.PersistentBrowser.IsEnabled -ne $null) {
                $persistentBrowser = @{
                    isEnabled = $fullPolicy.SessionControls.PersistentBrowser.IsEnabled
                }
                if ($fullPolicy.SessionControls.PersistentBrowser.Mode) {
                    $persistentBrowser.mode = $fullPolicy.SessionControls.PersistentBrowser.Mode
                }
                $sessionControls.persistentBrowser = $persistentBrowser
            }
            
            # Only add session controls if we have valid controls
            if ($sessionControls.Count -gt 0) {
                $newPolicyBody.sessionControls = $sessionControls
                Write-Host "Added session controls: $($sessionControls.Keys -join ', ')" -ForegroundColor Gray
            } else {
                Write-Host "No valid session controls found, skipping..." -ForegroundColor Gray
            }
        }
        
        Write-Host "Creating new policy: $newName" -ForegroundColor Cyan
        Write-Host "Policy will be created in DISABLED state for safety" -ForegroundColor Yellow
        
        # Function to remove null values recursively and preserve arrays
        function Remove-NullValues {
            param($obj)
            
            if ($obj -eq $null) { return $null }
            
            if ($obj -is [hashtable]) {
                $cleaned = @{}
                foreach ($key in $obj.Keys) {
                    $value = Remove-NullValues $obj[$key]
                    if ($value -ne $null) {
                        $cleaned[$key] = $value
                    }
                }
                if ($cleaned.Count -gt 0) { 
                    return $cleaned 
                } else { 
                    return $null 
                }
            }
            
            if ($obj -is [array] -or $obj -is [System.Collections.ArrayList]) {
                $cleaned = @()
                foreach ($item in $obj) {
                    $value = Remove-NullValues $item
                    if ($value -ne $null) {
                        $cleaned += $value
                    }
                }
                if ($cleaned.Count -gt 0) { 
                    # Force array return even for single items
                    return ,$cleaned
                } else { 
                    return $null 
                }
            }
            
            return $obj
        }
        
        # Clean the policy of null values
        $cleanedPolicy = Remove-NullValues $newPolicyBody
        
        # Debug: Show the array types before JSON conversion
        Write-Host "DEBUG: Checking array preservation..." -ForegroundColor Magenta
        if ($cleanedPolicy.grantControls.builtInControls) {
            Write-Host "builtInControls type: $($cleanedPolicy.grantControls.builtInControls.GetType().Name), Count: $($cleanedPolicy.grantControls.builtInControls.Count)" -ForegroundColor Magenta
        }
        if ($cleanedPolicy.conditions.users.includeUsers) {
            Write-Host "includeUsers type: $($cleanedPolicy.conditions.users.includeUsers.GetType().Name), Count: $($cleanedPolicy.conditions.users.includeUsers.Count)" -ForegroundColor Magenta
        }
        
        # Convert to JSON and show for debugging
        $jsonBody = $cleanedPolicy | ConvertTo-Json -Depth 10
        Write-Host "JSON being sent (first 800 chars):" -ForegroundColor Gray
        Write-Host $jsonBody.Substring(0, [Math]::Min(800, $jsonBody.Length)) -ForegroundColor Gray
        
        if ($jsonBody.Length -gt 800) {
            Write-Host "... (truncated, total length: $($jsonBody.Length) chars)" -ForegroundColor Gray
        }
        
        # Use REST API
        $uri = "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies"
        $response = Invoke-MgGraphRequest -Method POST -Uri $uri -Body $jsonBody -ContentType "application/json"
        
        Write-Host "✅ Policy created successfully!" -ForegroundColor Green
        Write-Host ("New policy ID: " + $response.id) -ForegroundColor Green
        
        $successMessage = "Conditional Access Policy copied successfully as '$newName'!`n`n"
        $successMessage += "⚠️ IMPORTANT: The new policy has been created in DISABLED state for safety.`n"
        $successMessage += "Please review the settings and enable it manually when ready.`n`n"
        $successMessage += "New Policy ID: " + $response.id
        
        [System.Windows.Forms.MessageBox]::Show($successMessage, "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        
        # Refresh the list
        Start-Sleep -Seconds 2
        Refresh-PoliciesList $listView
        
    } catch {
        Write-Host "❌ ERROR: Policy copy failed" -ForegroundColor Red
        Write-Host ("Error details: " + $_.Exception.Message) -ForegroundColor Red
        
        $errorMessage = "Error copying Conditional Access Policy:`n`n"
        $errorMessage += "Error: " + $_.Exception.Message + "`n`n"
        $errorMessage += "Source Policy: $sourceName`n"
        $errorMessage += "Target Name: $newName`n`n"
        $errorMessage += "This might be due to:`n"
        $errorMessage += "• Missing permissions (need Policy.ReadWrite.ConditionalAccess)`n"
        $errorMessage += "• Referenced objects (groups, named locations) that don't exist`n"
        $errorMessage += "• Complex policy conditions that need manual recreation`n"
        $errorMessage += "• Special characters in the policy name`n`n"
        $errorMessage += "Try creating a simple test policy first to verify permissions."
        
        Write-Host $errorMessage -ForegroundColor Red
        [System.Windows.Forms.MessageBox]::Show($errorMessage, "Copy Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    }
}

# New workflow: create a geo-IP exception by duplicating a policy, cloning a country named location with new countries,
# adding specific users to the new policy, and excluding those users from the original policy.
function Create-GeoIpExceptionForPolicy {
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
    $policy = $selectedItem.Tag
    $policyId = $policy.Id

    try {
        $fullPolicy = Get-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policyId
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Unable to load the selected policy: $_", "Error")
        return
    }

    if (-not $fullPolicy.Conditions -or -not $fullPolicy.Conditions.Locations) {
        [System.Windows.Forms.MessageBox]::Show("The selected policy has no location conditions to clone.", "Not Supported")
        return
    }

    # Collect all referenced location IDs (include + exclude) and resolve them
    $referencedLocationIds = @()
    if ($fullPolicy.Conditions.Locations.IncludeLocations) { $referencedLocationIds += $fullPolicy.Conditions.Locations.IncludeLocations }
    if ($fullPolicy.Conditions.Locations.ExcludeLocations) { $referencedLocationIds += $fullPolicy.Conditions.Locations.ExcludeLocations }
    $referencedLocationIds = $referencedLocationIds | Where-Object { $_ -ne $null -and $_ -ne "" } | Select-Object -Unique

    $resolvedLocations = @()
    foreach ($locId in $referencedLocationIds) {
        try {
            $loc = Get-MgIdentityConditionalAccessNamedLocation -NamedLocationId $locId -ErrorAction Stop
            $odataType = $loc.AdditionalProperties.'@odata.type'
            if ($odataType -eq '#microsoft.graph.countryNamedLocation') {
                # Track whether the location is used in include/exclude
                $usage = @()
                if ($fullPolicy.Conditions.Locations.IncludeLocations -contains $locId) { $usage += "Included" }
                if ($fullPolicy.Conditions.Locations.ExcludeLocations -contains $locId) { $usage += "Excluded" }

                $resolvedLocations += [PSCustomObject]@{
                    Id = $loc.Id
                    DisplayName = $loc.DisplayName
                    Countries = $loc.AdditionalProperties['countriesAndRegions']
                    IncludeUnknown = [bool]$loc.AdditionalProperties['includeUnknownCountriesAndRegions']
                    Usage = ($usage -join ', ')
                }
            }
        } catch {
            Write-Host ("Skipping location $locId due to error: " + $_.Exception.Message) -ForegroundColor Yellow
        }
    }

    if ($resolvedLocations.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No country-based named locations were found in this policy.", "Not Supported")
        return
    }

    # Build a wizard-style form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Create Geo-IP Exception"
    $form.Size = New-Object System.Drawing.Size(720, 650)
    $form.StartPosition = "CenterParent"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false

    $infoLabel = New-Object System.Windows.Forms.Label
    $infoLabel.Text = "This will copy the selected policy, clone a named location with new countries, add users to the new policy, and exclude them from the original."
    $infoLabel.Location = New-Object System.Drawing.Point(10, 10)
    $infoLabel.Size = New-Object System.Drawing.Size(690, 30)
    $form.Controls.Add($infoLabel)

    # Policy name input
    $policyNameLabel = New-Object System.Windows.Forms.Label
    $policyNameLabel.Text = "New policy name:"
    $policyNameLabel.Location = New-Object System.Drawing.Point(10, 50)
    $policyNameLabel.Size = New-Object System.Drawing.Size(200, 20)
    $form.Controls.Add($policyNameLabel)

    $policyNameTextBox = New-Object System.Windows.Forms.TextBox
    $policyNameTextBox.Location = New-Object System.Drawing.Point(10, 70)
    $policyNameTextBox.Size = New-Object System.Drawing.Size(680, 20)
    $policyNameTextBox.Text = ($policy.DisplayName + " - Geo Exception")
    $form.Controls.Add($policyNameTextBox)

    # Location selector
    $locationLabel = New-Object System.Windows.Forms.Label
    $locationLabel.Text = "Choose the location to clone and relax:" 
    $locationLabel.Location = New-Object System.Drawing.Point(10, 105)
    $locationLabel.Size = New-Object System.Drawing.Size(300, 20)
    $form.Controls.Add($locationLabel)

    $locationListBox = New-Object System.Windows.Forms.ListBox
    $locationListBox.Location = New-Object System.Drawing.Point(10, 130)
    $locationListBox.Size = New-Object System.Drawing.Size(680, 100)
    $locationListBox.DisplayMember = "Display"
    $locationListBox.SelectionMode = "One"
    foreach ($loc in $resolvedLocations) {
        $display = "$($loc.DisplayName) [$($loc.Id)] - $($loc.Usage)"
        $item = New-Object PSObject -Property @{ Data = $loc; Display = $display }
        $locationListBox.Items.Add($item) | Out-Null
    }
    if ($locationListBox.Items.Count -gt 0) { $locationListBox.SelectedIndex = 0 }
    $form.Controls.Add($locationListBox)

    # New location name
    $locNameLabel = New-Object System.Windows.Forms.Label
    $locNameLabel.Text = "New named location name:"
    $locNameLabel.Location = New-Object System.Drawing.Point(10, 245)
    $locNameLabel.Size = New-Object System.Drawing.Size(200, 20)
    $form.Controls.Add($locNameLabel)

    $locNameTextBox = New-Object System.Windows.Forms.TextBox
    $locNameTextBox.Location = New-Object System.Drawing.Point(10, 265)
    $locNameTextBox.Size = New-Object System.Drawing.Size(680, 20)
    $form.Controls.Add($locNameTextBox)

    # Country selection controls
    $countriesLabel = New-Object System.Windows.Forms.Label
    $countriesLabel.Text = "Countries for the exception (use selector):"
    $countriesLabel.Location = New-Object System.Drawing.Point(10, 295)
    $countriesLabel.Size = New-Object System.Drawing.Size(300, 20)
    $form.Controls.Add($countriesLabel)

    $countriesTextBox = New-Object System.Windows.Forms.TextBox
    $countriesTextBox.Location = New-Object System.Drawing.Point(10, 315)
    $countriesTextBox.Size = New-Object System.Drawing.Size(550, 20)
    $countriesTextBox.ReadOnly = $true
    $form.Controls.Add($countriesTextBox)

    $countriesButton = New-Object System.Windows.Forms.Button
    $countriesButton.Text = "Select Countries"
    $countriesButton.Location = New-Object System.Drawing.Point(570, 313)
    $countriesButton.Size = New-Object System.Drawing.Size(120, 24)
    $countriesButton.BackColor = [System.Drawing.Color]::LightSteelBlue
    $form.Controls.Add($countriesButton)

    $includeUnknownCheckbox = New-Object System.Windows.Forms.CheckBox
    $includeUnknownCheckbox.Text = "Include unknown/future countries"
    $includeUnknownCheckbox.Location = New-Object System.Drawing.Point(10, 345)
    $includeUnknownCheckbox.Size = New-Object System.Drawing.Size(300, 20)
    $form.Controls.Add($includeUnknownCheckbox)

    # Users to add
    $usersLabel = New-Object System.Windows.Forms.Label
    $usersLabel.Text = "Users to include in new policy and exclude from original (one per line: UPN, name, or ID):"
    $usersLabel.Location = New-Object System.Drawing.Point(10, 375)
    $usersLabel.Size = New-Object System.Drawing.Size(690, 20)
    $form.Controls.Add($usersLabel)

    $usersTextBox = New-Object System.Windows.Forms.RichTextBox
    $usersTextBox.Location = New-Object System.Drawing.Point(10, 400)
    $usersTextBox.Size = New-Object System.Drawing.Size(550, 120)
    $form.Controls.Add($usersTextBox)
    
    $usersSearchButton = New-Object System.Windows.Forms.Button
    $usersSearchButton.Text = "Search Users"
    $usersSearchButton.Location = New-Object System.Drawing.Point(570, 400)
    $usersSearchButton.Size = New-Object System.Drawing.Size(120, 30)
    $usersSearchButton.BackColor = [System.Drawing.Color]::LightBlue
    $usersSearchButton.Add_Click({
        $foundUsers = Show-UserSearchDialog -Title "Search Users to Add"
        if ($foundUsers -and $foundUsers.Count -gt 0) {
            $existingText = $usersTextBox.Text.Trim()
            $newLines = @()
            if ($existingText) {
                $newLines += $existingText.Split("`n") | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
            }
            foreach ($user in $foundUsers) {
                # Add userPrincipalName if available, otherwise display name
                $userIdentifier = $user.UserPrincipalName
                if ([string]::IsNullOrWhiteSpace($userIdentifier)) {
                    $userIdentifier = $user.DisplayName
                }
                if ($userIdentifier -notin $newLines) {
                    $newLines += $userIdentifier
                }
            }
            $usersTextBox.Text = ($newLines -join "`n")
        }
    })
    $form.Controls.Add($usersSearchButton)

# Helper to seed UI from selected location
function Invoke-GeoIpLocationUiUpdate {
    if ($locationListBox.SelectedItem -and $locationListBox.SelectedItem.Data) {
        $locData = $locationListBox.SelectedItem.Data
        $countriesTextBox.Text = ($locData.Countries -join ', ')
        $includeUnknownCheckbox.Checked = $locData.IncludeUnknown
        $locNameTextBox.Text = "Exception - " + $locData.DisplayName
    }
}

$locationListBox.Add_SelectedIndexChanged([System.EventHandler]{ param($sender, $eventArgs) Invoke-GeoIpLocationUiUpdate })
Invoke-GeoIpLocationUiUpdate


    # Country picker handler
    $countriesButton.Add_Click({
        $current = @()
        if (-not [string]::IsNullOrWhiteSpace($countriesTextBox.Text)) {
            $current = $countriesTextBox.Text.Split(',') | ForEach-Object { $_.Trim().ToUpper() }
        }
        $selectedCountries = Show-CountrySelectionDialog -PreselectedCountries $current
        if ($selectedCountries) {
            $countriesTextBox.Text = ($selectedCountries -join ', ')
        }
    })

    # Buttons
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "Create Exception"
    $okButton.Location = New-Object System.Drawing.Point(500, 540)
    $okButton.Size = New-Object System.Drawing.Size(90, 30)
    $okButton.BackColor = [System.Drawing.Color]::LightGreen
    $form.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(600, 540)
    $cancelButton.Size = New-Object System.Drawing.Size(90, 30)
    $cancelButton.BackColor = [System.Drawing.Color]::LightGray
    $cancelButton.Add_Click({ $form.Close() })
    $form.Controls.Add($cancelButton)

    $okButton.Add_Click({
        if (-not $locationListBox.SelectedItem) {
            [System.Windows.Forms.MessageBox]::Show("Please select a named location to clone.", "Validation Error")
            return
        }
        if ([string]::IsNullOrWhiteSpace($locNameTextBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please provide a name for the new named location.", "Validation Error")
            return
        }
        if ([string]::IsNullOrWhiteSpace($policyNameTextBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please provide a name for the new policy.", "Validation Error")
            return
        }
        if ([string]::IsNullOrWhiteSpace($countriesTextBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please select at least one country.", "Validation Error")
            return
        }

        $countryCodes = $countriesTextBox.Text.Split(',') | ForEach-Object { $_.Trim().ToUpper() } | Where-Object { $_ -ne "" }
        if ($countryCodes.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select at least one country.", "Validation Error")
            return
        }

        $userInputs = $usersTextBox.Text.Split("`n") | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
        if ($userInputs.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please enter at least one user to add.", "Validation Error")
            return
        }

        $locData = $locationListBox.SelectedItem.Data
        $newLocationName = $locNameTextBox.Text
        $newPolicyName = $policyNameTextBox.Text
        $includeUnknown = $includeUnknownCheckbox.Checked

        $policyJson = $null
        try {
            $resolved = Resolve-UserInput -UserInputs $userInputs
            if ($resolved.NotFoundUsers.Count -gt 0) {
                $msg = "The following users could not be resolved: " + ($resolved.NotFoundUsers -join ', ') + "`nContinue anyway?"
                $continueResult = [System.Windows.Forms.MessageBox]::Show($msg, "Users Not Found", [System.Windows.Forms.MessageBoxButtons]::YesNo)
                if ($continueResult -eq [System.Windows.Forms.DialogResult]::No) { return }
            }
            if ($resolved.ResolvedUserIds.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("No resolvable users were provided.", "Validation Error")
                return
            }

            # 1) Create the new named location
            $createLocationBody = @{
                "@odata.type" = "#microsoft.graph.countryNamedLocation"
                displayName = $newLocationName
                countriesAndRegions = $countryCodes
                includeUnknownCountriesAndRegions = $includeUnknown
            }
            $locJson = $createLocationBody | ConvertTo-Json -Depth 10
            Write-Host ("Creating new named location: " + $locJson) -ForegroundColor Cyan
            $locUri = "https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations"
            $newLocation = Invoke-MgGraphRequest -Method POST -Uri $locUri -Body $locJson -ContentType "application/json"
            $newLocationId = $newLocation.id
            Write-Host ("New location created with ID: " + $newLocationId) -ForegroundColor Green

            # Wait for location to propagate (Microsoft Graph API sometimes needs a moment)
            Write-Host "Waiting for location to propagate..." -ForegroundColor Yellow
            $maxRetries = 5
            $retryCount = 0
            $locationExists = $false
            while ($retryCount -lt $maxRetries -and -not $locationExists) {
                Start-Sleep -Seconds 2
                try {
                    $verifyUri = "https://graph.microsoft.com/v1.0/identity/conditionalAccess/namedLocations/$newLocationId"
                    $verifyLocation = Invoke-MgGraphRequest -Method GET -Uri $verifyUri
                    if ($verifyLocation.id -eq $newLocationId) {
                        $locationExists = $true
                        Write-Host "Location verified successfully." -ForegroundColor Green
                    }
                } catch {
                    $retryCount++
                    if ($retryCount -lt $maxRetries) {
                        Write-Host "Location not yet available, retrying... ($retryCount/$maxRetries)" -ForegroundColor Yellow
                    }
                }
            }
            if (-not $locationExists) {
                throw "Named location $newLocationId was created but could not be verified after $maxRetries attempts. Please try again."
            }

            # 2) Build new policy body based on existing policy, but with adjusted users/locations
            # Start with a minimal working policy structure (like the copy function)
            $newPolicyBody = @{
                displayName = $newPolicyName
                state = "disabled"  # Safety first
                conditions = @{
                    users = @{
                        includeUsers = @("All")
                    }
                    applications = @{
                        includeApplications = @("All")
                    }
                }
                grantControls = @{
                    operator = "OR"
                    builtInControls = @("block")
                }
            }

            # Build user conditions
            $userConditions = @{}
            
            # Include users: remove "All" and add specified users + existing specific users
            $existingInclude = @()
            if ($fullPolicy.Conditions.Users.IncludeUsers) {
                $existingInclude += ($fullPolicy.Conditions.Users.IncludeUsers | Where-Object { $_ -ne $null -and $_ -ne "" -and $_ -ne "All" })
            }
            $includeUsers = @()
            $includeUsers += $existingInclude
            foreach ($uid in $resolved.ResolvedUserIds) {
                if ($uid -notin $includeUsers) { $includeUsers += $uid }
            }
            if ($includeUsers.Count -eq 0) { 
                $includeUsers = @("All") 
            }
            $userConditions.includeUsers = $includeUsers

            # IMPORTANT: According to Microsoft Graph API schema, when includeUsers contains specific users (not "All"),
            # excludeUsers must be empty or omitted. Only preserve excludeUsers if includeUsers is "All"
            if ($includeUsers.Count -eq 1 -and $includeUsers[0] -eq "All") {
                # Only preserve exclusions when including "All" users
                if ($fullPolicy.Conditions.Users.ExcludeUsers) { 
                    $excludeUsers = @($fullPolicy.Conditions.Users.ExcludeUsers | Where-Object { $_ -ne $null -and $_ -ne "" })
                    if ($excludeUsers.Count -gt 0) {
                        $userConditions.excludeUsers = $excludeUsers
                    }
                }
            }
            # Otherwise, excludeUsers should not be set when including specific users
            
            # Preserve groups and roles (these are allowed with specific users)
            if ($fullPolicy.Conditions.Users.IncludeGroups) { $userConditions.includeGroups = $fullPolicy.Conditions.Users.IncludeGroups }
            if ($fullPolicy.Conditions.Users.ExcludeGroups) { $userConditions.excludeGroups = $fullPolicy.Conditions.Users.ExcludeGroups }
            if ($fullPolicy.Conditions.Users.IncludeRoles) { $userConditions.includeRoles = $fullPolicy.Conditions.Users.IncludeRoles }
            if ($fullPolicy.Conditions.Users.ExcludeRoles) { $userConditions.excludeRoles = $fullPolicy.Conditions.Users.ExcludeRoles }

            # Update the conditions.users with the built user conditions
            $newPolicyBody.conditions.users = $userConditions

            # Applications - update existing structure
            if ($fullPolicy.Conditions.Applications) {
                $appConditions = @{}
                if ($fullPolicy.Conditions.Applications.IncludeApplications -and $fullPolicy.Conditions.Applications.IncludeApplications.Count -gt 0) {
                    $includeApps = @($fullPolicy.Conditions.Applications.IncludeApplications | Where-Object { $_ -ne $null -and $_ -ne "" })
                    if ($includeApps.Count -gt 0) {
                        $appConditions.includeApplications = $includeApps
                    }
                }
                if ($fullPolicy.Conditions.Applications.ExcludeApplications -and $fullPolicy.Conditions.Applications.ExcludeApplications.Count -gt 0) {
                    $excludeApps = @($fullPolicy.Conditions.Applications.ExcludeApplications | Where-Object { $_ -ne $null -and $_ -ne "" })
                    if ($excludeApps.Count -gt 0) {
                        $appConditions.excludeApplications = $excludeApps
                    }
                }
                if ($appConditions.Count -gt 0) { 
                    $newPolicyBody.conditions.applications = $appConditions 
                }
            }

            # Locations - update existing structure
            if ($fullPolicy.Conditions.Locations) {
                $locationConditions = @{}
                if ($fullPolicy.Conditions.Locations.IncludeLocations) {
                    # Handle both array and single string values - ensure it's always an array
                    $includeLocationsRaw = $fullPolicy.Conditions.Locations.IncludeLocations
                    if ($includeLocationsRaw -is [string]) {
                        $includeLocationsRaw = @($includeLocationsRaw)
                    }
                    $includeLocations = [System.Collections.ArrayList]::new()
                    foreach ($item in $includeLocationsRaw) {
                        if ($item -ne $null -and $item -ne "") {
                            $loc = if ($item -eq $locData.Id) { $newLocationId } else { $item }
                            [void]$includeLocations.Add($loc)
                        }
                    }
                    if ($includeLocations.Count -gt 0) {
                        $locationConditions.includeLocations = $includeLocations.ToArray()
                    }
                }
                if ($fullPolicy.Conditions.Locations.ExcludeLocations) {
                    # Handle both array and single string values - ensure it's always an array
                    $excludeLocationsRaw = $fullPolicy.Conditions.Locations.ExcludeLocations
                    if ($excludeLocationsRaw -is [string]) {
                        $excludeLocationsRaw = @($excludeLocationsRaw)
                    }
                    $excludeLocations = [System.Collections.ArrayList]::new()
                    foreach ($item in $excludeLocationsRaw) {
                        if ($item -ne $null -and $item -ne "") {
                            $loc = if ($item -eq $locData.Id) { $newLocationId } else { $item }
                            [void]$excludeLocations.Add($loc)
                        }
                    }
                    if ($excludeLocations.Count -gt 0) {
                        $locationConditions.excludeLocations = $excludeLocations.ToArray()
                    }
                }
                if ($locationConditions.Count -gt 0) { 
                    $newPolicyBody.conditions.locations = $locationConditions 
                }
            }

            # Platforms - update existing structure
            if ($fullPolicy.Conditions.Platforms) {
                $platformConditions = @{}
                if ($fullPolicy.Conditions.Platforms.IncludePlatforms -and $fullPolicy.Conditions.Platforms.IncludePlatforms.Count -gt 0) {
                    $includePlatforms = @($fullPolicy.Conditions.Platforms.IncludePlatforms | Where-Object { $_ -ne $null -and $_ -ne "" })
                    if ($includePlatforms.Count -gt 0) {
                        $platformConditions.includePlatforms = $includePlatforms
                    }
                }
                if ($fullPolicy.Conditions.Platforms.ExcludePlatforms -and $fullPolicy.Conditions.Platforms.ExcludePlatforms.Count -gt 0) {
                    $excludePlatforms = @($fullPolicy.Conditions.Platforms.ExcludePlatforms | Where-Object { $_ -ne $null -and $_ -ne "" })
                    if ($excludePlatforms.Count -gt 0) {
                        $platformConditions.excludePlatforms = $excludePlatforms
                    }
                }
                if ($platformConditions.Count -gt 0) { 
                    $newPolicyBody.conditions.platforms = $platformConditions 
                }
            }

            # Grant controls - update existing structure
            if ($fullPolicy.GrantControls) {
                # Operator
                if ($fullPolicy.GrantControls.Operator) {
                    $newPolicyBody.grantControls.operator = $fullPolicy.GrantControls.Operator
                }
                
                # BuiltInControls is required - ensure it has at least one value
                if ($fullPolicy.GrantControls.BuiltInControls -and $fullPolicy.GrantControls.BuiltInControls.Count -gt 0) {
                    $builtInControls = @($fullPolicy.GrantControls.BuiltInControls | Where-Object { $_ -ne $null -and $_ -ne "" })
                    if ($builtInControls.Count -gt 0) {
                        $newPolicyBody.grantControls.builtInControls = $builtInControls
                    }
                }
                
                # CustomAuthenticationFactors
                if ($fullPolicy.GrantControls.CustomAuthenticationFactors -and $fullPolicy.GrantControls.CustomAuthenticationFactors.Count -gt 0) {
                    $customAuth = @($fullPolicy.GrantControls.CustomAuthenticationFactors | Where-Object { $_ -ne $null -and $_ -ne "" })
                    if ($customAuth.Count -gt 0) { 
                        $newPolicyBody.grantControls.customAuthenticationFactors = $customAuth 
                    }
                }
                
                # TermsOfUse
                if ($fullPolicy.GrantControls.TermsOfUse -and $fullPolicy.GrantControls.TermsOfUse.Count -gt 0) {
                    $termsOfUse = @($fullPolicy.GrantControls.TermsOfUse | Where-Object { $_ -ne $null -and $_ -ne "" })
                    if ($termsOfUse.Count -gt 0) { 
                        $newPolicyBody.grantControls.termsOfUse = $termsOfUse 
                    }
                }
            }

            # Session controls (simple copy)
            if ($fullPolicy.SessionControls) {
                $sessionControls = @{}
                if ($fullPolicy.SessionControls.ApplicationEnforcedRestrictions -and $fullPolicy.SessionControls.ApplicationEnforcedRestrictions.IsEnabled -ne $null) {
                    $sessionControls.applicationEnforcedRestrictions = @{ isEnabled = $fullPolicy.SessionControls.ApplicationEnforcedRestrictions.IsEnabled }
                }
                if ($fullPolicy.SessionControls.CloudAppSecurity -and $fullPolicy.SessionControls.CloudAppSecurity.IsEnabled -ne $null) {
                    $cloudApp = @{ isEnabled = $fullPolicy.SessionControls.CloudAppSecurity.IsEnabled }
                    if ($fullPolicy.SessionControls.CloudAppSecurity.CloudAppSecurityType) { $cloudApp.cloudAppSecurityType = $fullPolicy.SessionControls.CloudAppSecurity.CloudAppSecurityType }
                    $sessionControls.cloudAppSecurity = $cloudApp
                }
                if ($fullPolicy.SessionControls.SignInFrequency -and $fullPolicy.SessionControls.SignInFrequency.IsEnabled -ne $null) {
                    $signInFreq = @{ isEnabled = $fullPolicy.SessionControls.SignInFrequency.IsEnabled }
                    if ($fullPolicy.SessionControls.SignInFrequency.Type) { $signInFreq.type = $fullPolicy.SessionControls.SignInFrequency.Type }
                    if ($fullPolicy.SessionControls.SignInFrequency.Value -ne $null) { $signInFreq.value = $fullPolicy.SessionControls.SignInFrequency.Value }
                    $sessionControls.signInFrequency = $signInFreq
                }
                if ($fullPolicy.SessionControls.PersistentBrowser -and $fullPolicy.SessionControls.PersistentBrowser.IsEnabled -ne $null) {
                    $persistent = @{ isEnabled = $fullPolicy.SessionControls.PersistentBrowser.IsEnabled }
                    if ($fullPolicy.SessionControls.PersistentBrowser.Mode) { $persistent.mode = $fullPolicy.SessionControls.PersistentBrowser.Mode }
                    $sessionControls.persistentBrowser = $persistent
                }
                if ($sessionControls.Count -gt 0) { $newPolicyBody.sessionControls = $sessionControls }
            }

            function Remove-NullGraphValues {
                param($obj)
                if ($obj -eq $null) { return $null }
                if ($obj -is [hashtable]) {
                    $clean = @{}
                    foreach ($k in $obj.Keys) {
                        $val = Remove-NullGraphValues $obj[$k]
                        if ($val -ne $null) { $clean[$k] = $val }
                    }
                    if ($clean.Count -gt 0) { return $clean } else { return $null }
                }
                if ($obj -is [array] -or $obj -is [System.Collections.ArrayList]) {
                    $arr = @()
                    foreach ($item in $obj) {
                        $val = Remove-NullGraphValues $item
                        if ($val -ne $null) { $arr += $val }
                    }
                    if ($arr.Count -gt 0) { return ,$arr } else { return $null }
                }
                return $obj
            }

            $cleanPolicy = Remove-NullGraphValues $newPolicyBody
            
            # Validate required fields before sending (shouldn't be needed since we start with valid structure, but safety check)
            if (-not $cleanPolicy.displayName) {
                throw "Policy displayName is required but missing"
            }
            if (-not $cleanPolicy.state) {
                throw "Policy state is required but missing"
            }
            if (-not $cleanPolicy.conditions -or -not $cleanPolicy.conditions.users) {
                throw "Policy conditions.users is required but missing"
            }
            if (-not $cleanPolicy.grantControls -or -not $cleanPolicy.grantControls.builtInControls -or $cleanPolicy.grantControls.builtInControls.Count -eq 0) {
                throw "Policy grantControls.builtInControls is required but missing or empty"
            }
            
            # Ensure includeUsers is not empty (fallback safety)
            if (-not $cleanPolicy.conditions.users.includeUsers -or $cleanPolicy.conditions.users.includeUsers.Count -eq 0) {
                $cleanPolicy.conditions.users.includeUsers = @("All")
            }
            
            # Ensure applications is present (fallback safety)
            if (-not $cleanPolicy.conditions.applications) {
                $cleanPolicy.conditions.applications = @{ includeApplications = @("All") }
            }
            
            $policyJson = $cleanPolicy | ConvertTo-Json -Depth 15
            Write-Host ("Creating new policy with body (truncated to 800 chars): " + $policyJson.Substring(0, [Math]::Min(800, $policyJson.Length))) -ForegroundColor Cyan

            $policyUri = "https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies"
            $createdPolicy = Invoke-MgGraphRequest -Method POST -Uri $policyUri -Body $policyJson -ContentType "application/json"
            $newPolicyId = $createdPolicy.id
            Write-Host ("New policy created with ID: " + $newPolicyId) -ForegroundColor Green

            # 3) Exclude the same users from the original policy
            $origPolicy = Get-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policyId
            $currentExclude = @()
            if ($origPolicy.Conditions.Users.ExcludeUsers) { $currentExclude += $origPolicy.Conditions.Users.ExcludeUsers }
            foreach ($uid in $resolved.ResolvedUserIds) {
                if ($uid -notin $currentExclude) { $currentExclude += $uid }
            }

            $origUserConditions = @{
                IncludeUsers = $origPolicy.Conditions.Users.IncludeUsers
                ExcludeUsers = $currentExclude
            }
            if ($origPolicy.Conditions.Users.IncludeGroups) { $origUserConditions.IncludeGroups = $origPolicy.Conditions.Users.IncludeGroups }
            if ($origPolicy.Conditions.Users.ExcludeGroups) { $origUserConditions.ExcludeGroups = $origPolicy.Conditions.Users.ExcludeGroups }
            if ($origPolicy.Conditions.Users.IncludeRoles) { $origUserConditions.IncludeRoles = $origPolicy.Conditions.Users.IncludeRoles }
            if ($origPolicy.Conditions.Users.ExcludeRoles) { $origUserConditions.ExcludeRoles = $origPolicy.Conditions.Users.ExcludeRoles }

            Update-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policyId -Conditions @{ Users = $origUserConditions }
            Write-Host "Users excluded from original policy." -ForegroundColor Green

            $summary = "Geo-IP exception created successfully!`n`nNew location: $newLocationName`nNew policy: $newPolicyName (disabled)`nUsers were excluded from the original policy."
            [System.Windows.Forms.MessageBox]::Show($summary, "Success", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)

            $form.Close()
        } catch {
            $errorMsg = "Error creating exception flow: $_`n`n"
            if ($policyJson) {
                $errorMsg += "JSON sent (first 1000 chars):`n" + $policyJson.Substring(0, [Math]::Min(1000, $policyJson.Length)) + "`n`n"
            }
            $errorMsg += "Full error details:`n$($_.Exception.Message)"
            
            # Try to get response body from HttpResponseMessage
            if ($_.Exception.Response) {
                try {
                    $response = $_.Exception.Response
                    # Check if it's HttpResponseMessage (from Invoke-MgGraphRequest)
                    if ($response -is [System.Net.Http.HttpResponseMessage]) {
                        $responseBody = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                        if ($responseBody) {
                            $errorMsg += "`n`nResponse: $responseBody"
                        }
                    }
                    # Check if it's HttpWebResponse (legacy)
                    elseif ($response -is [System.Net.HttpWebResponse]) {
                        $reader = New-Object System.IO.StreamReader($response.GetResponseStream())
                        $responseBody = $reader.ReadToEnd()
                        $reader.Close()
                        if ($responseBody) {
                            $errorMsg += "`n`nResponse: $responseBody"
                        }
                    }
                } catch {
                    # If we can't read the response, just include what we have
                    $errorMsg += "`n`n(Unable to read response body: $($_.Exception.Message))"
                }
            }
            
            Write-Host $errorMsg -ForegroundColor Red
            [System.Windows.Forms.MessageBox]::Show($errorMsg, "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })

    $form.ShowDialog() | Out-Null

    # Refresh UI after the dialog closes
    Refresh-NamedLocationsList $script:namedLocationsListView
    Refresh-PoliciesList $listView
}

function Rename-SelectedPolicy {
    param($listView)
    
    if (-not $global:isConnected) {
        [System.Windows.Forms.MessageBox]::Show("Please connect to Microsoft Graph first.", "Not Connected")
        return
    }
    
    if ($listView.SelectedItems.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Please select a Conditional Access Policy to rename.", "No Selection")
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
        $existingPolicy = Get-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policyId -ErrorAction SilentlyContinue
        if (-not $existingPolicy) {
            [System.Windows.Forms.MessageBox]::Show("The selected Conditional Access Policy no longer exists. Refreshing list.", "Not Found")
            Refresh-PoliciesList $listView
            return
        }
        
        Update-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policyId -DisplayName $newName
        [System.Windows.Forms.MessageBox]::Show("Conditional Access Policy renamed successfully!", "Success")
        
        # Small delay before refresh
        Start-Sleep -Seconds 1
        Refresh-PoliciesList $listView
    } catch {
        if ($_.Exception.Message -like "*NotFound*" -or $_.Exception.Message -like "*404*") {
            [System.Windows.Forms.MessageBox]::Show("The Conditional Access Policy no longer exists. Refreshing list.", "Not Found")
            Refresh-PoliciesList $listView
        } else {
            [System.Windows.Forms.MessageBox]::Show("Error renaming Conditional Access Policy: $_", "Error")
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

    $searchButton = New-Object System.Windows.Forms.Button
    $searchButton.Text = "Search Users"
    $searchButton.Location = New-Object System.Drawing.Point(470, 235)
    $searchButton.Size = New-Object System.Drawing.Size(100, 30)
    $searchButton.BackColor = [System.Drawing.Color]::LightBlue
    $searchButton.Add_Click({
        $foundUsers = Show-UserSearchDialog -Title "Search Users to Add"
        if ($foundUsers -and $foundUsers.Count -gt 0) {
            $existingText = $addTextBox.Text.Trim()
            $newLines = @()
            if ($existingText) {
                $newLines += $existingText.Split("`n") | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
            }
            foreach ($user in $foundUsers) {
                # Add userPrincipalName if available, otherwise display name
                $userIdentifier = $user.UserPrincipalName
                if ([string]::IsNullOrWhiteSpace($userIdentifier)) {
                    $userIdentifier = $user.DisplayName
                }
                if ($userIdentifier -notin $newLines) {
                    $newLines += $userIdentifier
                }
            }
            $addTextBox.Text = ($newLines -join "`n")
        }
    })
    $form.Controls.Add($searchButton)
    
    $addButton = New-Object System.Windows.Forms.Button
    $addButton.Text = "Add Users"
    $addButton.Location = New-Object System.Drawing.Point(580, 235)
    $addButton.Size = New-Object System.Drawing.Size(100, 30)
    $addButton.BackColor = [System.Drawing.Color]::LightGreen
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
    $removeButton.BackColor = [System.Drawing.Color]::LightCoral
    $removeButton.ForeColor = [System.Drawing.Color]::DarkRed
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
    $closeButton.BackColor = [System.Drawing.Color]::LightGray
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

function Show-UserSearchDialog {
    param(
        [string]$Title = "Search and Select Users"
    )
    
    if (-not $global:isConnected) {
        [System.Windows.Forms.MessageBox]::Show("Please connect to Microsoft Graph first.", "Not Connected")
        return @()
    }
    
    $selectedUsers = @()
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object System.Drawing.Size(750, 600)
    $form.StartPosition = "CenterParent"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    
    # Search label and box
    $searchLabel = New-Object System.Windows.Forms.Label
    $searchLabel.Text = "Search (name or email):"
    $searchLabel.Location = New-Object System.Drawing.Point(10, 15)
    $searchLabel.Size = New-Object System.Drawing.Size(150, 20)
    $form.Controls.Add($searchLabel)
    
    $searchBox = New-Object System.Windows.Forms.TextBox
    $searchBox.Location = New-Object System.Drawing.Point(170, 13)
    $searchBox.Size = New-Object System.Drawing.Size(400, 20)
    $form.Controls.Add($searchBox)
    
    $searchButton = New-Object System.Windows.Forms.Button
    $searchButton.Text = "Search"
    $searchButton.Location = New-Object System.Drawing.Point(580, 12)
    $searchButton.Size = New-Object System.Drawing.Size(80, 25)
    $searchButton.BackColor = [System.Drawing.Color]::LightBlue
    $form.Controls.Add($searchButton)
    
    # Results list
    $resultsListView = New-Object System.Windows.Forms.ListView
    $resultsListView.Location = New-Object System.Drawing.Point(10, 45)
    $resultsListView.Size = New-Object System.Drawing.Size(720, 450)
    $resultsListView.View = "Details"
    $resultsListView.FullRowSelect = $true
    $resultsListView.MultiSelect = $true
    $resultsListView.GridLines = $true
    $resultsListView.Columns.Add("Display Name", 250) | Out-Null
    $resultsListView.Columns.Add("Email (UPN)", 300) | Out-Null
    $resultsListView.Columns.Add("ID", 150) | Out-Null
    $form.Controls.Add($resultsListView)
    
    # Status label
    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Text = "Enter search terms and click Search or press Enter"
    $statusLabel.Location = New-Object System.Drawing.Point(10, 505)
    $statusLabel.Size = New-Object System.Drawing.Size(600, 20)
    $form.Controls.Add($statusLabel)
    
    # Buttons
    $addButton = New-Object System.Windows.Forms.Button
    $addButton.Text = "Add Selected"
    $addButton.Location = New-Object System.Drawing.Point(550, 535)
    $addButton.Size = New-Object System.Drawing.Size(90, 30)
    $addButton.BackColor = [System.Drawing.Color]::LightGreen
    $addButton.Enabled = $false
    $form.Controls.Add($addButton)
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(650, 535)
    $cancelButton.Size = New-Object System.Drawing.Size(90, 30)
    $cancelButton.BackColor = [System.Drawing.Color]::LightGray
    $cancelButton.Add_Click({ $form.Close() })
    $form.Controls.Add($cancelButton)
    
    # Search function
    function Invoke-UserSearch {
        param([string]$searchTerm)
        
        $resultsListView.Items.Clear()
        $statusLabel.Text = "Searching..."
        $statusLabel.Refresh()
        [System.Windows.Forms.Application]::DoEvents()
        
        try {
            $users = @()
            $searchTerm = $searchTerm.Trim()
            
            if ([string]::IsNullOrWhiteSpace($searchTerm)) {
                $statusLabel.Text = "Please enter a search term (at least 2 characters)"
                return
            }
            
            if ($searchTerm.Length -lt 2) {
                $statusLabel.Text = "Please enter at least 2 characters to search"
                return
            }
            
            # Escape single quotes for OData filter
            $escapedTerm = $searchTerm.Replace("'", "''")
            
            # Search by display name (contains is supported for displayName)
            try {
                $filter = "contains(displayName,'$escapedTerm')"
                $nameResults = Get-MgUser -Filter $filter -Top 50 -ErrorAction SilentlyContinue
                if ($nameResults) {
                    $users += $nameResults
                }
            } catch {
                Write-Host "Error searching by name: $_" -ForegroundColor Yellow
            }
            
            # Search by userPrincipalName (startswith only - contains not supported)
            try {
                $filter = "startswith(userPrincipalName,'$escapedTerm')"
                $emailResults = Get-MgUser -Filter $filter -Top 50 -ErrorAction SilentlyContinue
                if ($emailResults) {
                    foreach ($user in $emailResults) {
                        if ($user.Id -notin ($users | ForEach-Object { $_.Id })) {
                            $users += $user
                        }
                    }
                }
            } catch {
                Write-Host "Error searching by UPN: $_" -ForegroundColor Yellow
            }
            
            # Search by mail (startswith only - contains not supported)
            try {
                $filter = "startswith(mail,'$escapedTerm')"
                $mailResults = Get-MgUser -Filter $filter -Top 50 -ErrorAction SilentlyContinue
                if ($mailResults) {
                    foreach ($user in $mailResults) {
                        if ($user.Id -notin ($users | ForEach-Object { $_.Id })) {
                            $users += $user
                        }
                    }
                }
            } catch {
                Write-Host "Error searching by mail: $_" -ForegroundColor Yellow
            }
            
            # Also try using Search parameter for better partial matching (if available)
            # This uses Microsoft Graph's search capabilities
            try {
                $searchUri = "https://graph.microsoft.com/v1.0/users?`$search=`"$escapedTerm`"&`$top=50"
                $searchResults = Invoke-MgGraphRequest -Method GET -Uri $searchUri -ErrorAction SilentlyContinue
                if ($searchResults -and $searchResults.value) {
                    foreach ($user in $searchResults.value) {
                        if ($user.id -notin ($users | ForEach-Object { $_.Id })) {
                            # Convert to MgUser object format
                            $mgUser = Get-MgUser -UserId $user.id -ErrorAction SilentlyContinue
                            if ($mgUser) {
                                $users += $mgUser
                            }
                        }
                    }
                }
            } catch {
                # Search parameter might not be available, that's okay
                Write-Host "Search parameter not available or error: $_" -ForegroundColor Yellow
            }
            
            # Additional fallback: Get users and filter client-side for partial matches
            # This handles cases where API filters don't support contains for email fields
            if ($users.Count -eq 0) {
                try {
                    Write-Host "Performing client-side search for better partial matching..." -ForegroundColor Yellow
                    $statusLabel.Text = "Searching (this may take a moment)..."
                    $statusLabel.Refresh()
                    [System.Windows.Forms.Application]::DoEvents()
                    
                    # Get users in batches for better performance
                    $allUsers = Get-MgUser -All -Top 1000 -ErrorAction SilentlyContinue
                    $searchLower = $searchTerm.ToLower()
                    $foundCount = 0
                    
                    foreach ($user in $allUsers) {
                        $match = $false
                        if ($user.DisplayName -and $user.DisplayName.ToLower().Contains($searchLower)) {
                            $match = $true
                        }
                        if ($user.UserPrincipalName -and $user.UserPrincipalName.ToLower().Contains($searchLower)) {
                            $match = $true
                        }
                        if ($user.Mail -and $user.Mail.ToLower().Contains($searchLower)) {
                            $match = $true
                        }
                        if ($match) {
                            $users += $user
                            $foundCount++
                            # Limit results for performance
                            if ($foundCount -ge 100) { break }
                        }
                    }
                    Write-Host "Client-side search found $foundCount user(s)" -ForegroundColor Green
                } catch {
                    Write-Host "Error in client-side search: $_" -ForegroundColor Yellow
                }
            }
            
            # Remove duplicates and limit results
            $uniqueUsers = $users | Sort-Object DisplayName -Unique | Select-Object -First 100
            
            foreach ($user in $uniqueUsers) {
                $item = New-Object System.Windows.Forms.ListViewItem($user.DisplayName)
                $item.SubItems.Add($user.UserPrincipalName) | Out-Null
                $item.SubItems.Add($user.Id) | Out-Null
                $item.Tag = $user
                $resultsListView.Items.Add($item) | Out-Null
            }
            
            if ($uniqueUsers.Count -eq 0) {
                $statusLabel.Text = "No users found. Try a different search term."
            } else {
                $statusLabel.Text = "Found $($uniqueUsers.Count) user(s). Select users and click 'Add Selected'."
            }
            
        } catch {
            $statusLabel.Text = "Error searching: $($_.Exception.Message)"
            Write-Host "Error in user search: $_" -ForegroundColor Red
        }
    }
    
    # Search button click
    $searchButton.Add_Click({
        Invoke-UserSearch -searchTerm $searchBox.Text
    })
    
    # Search on Enter key
    $searchBox.Add_KeyDown({
        if ($_.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
            Invoke-UserSearch -searchTerm $searchBox.Text
            $_.SuppressKeyPress = $true
        }
    })
    
    # Enable/disable Add button based on selection
    $resultsListView.Add_SelectedIndexChanged({
        $addButton.Enabled = $resultsListView.SelectedItems.Count -gt 0
    })
    
    # Add selected button
    $addButton.Add_Click({
        $selectedUsers = @()
        foreach ($item in $resultsListView.SelectedItems) {
            $selectedUsers += $item.Tag
        }
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })
    
    $form.ShowDialog() | Out-Null
    return $selectedUsers
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

    $searchButton = New-Object System.Windows.Forms.Button
    $searchButton.Text = "Search Users"
    $searchButton.Location = New-Object System.Drawing.Point(470, 265)
    $searchButton.Size = New-Object System.Drawing.Size(100, 30)
    $searchButton.BackColor = [System.Drawing.Color]::LightBlue
    $searchButton.Add_Click({
        $foundUsers = Show-UserSearchDialog -Title "Search Users to Add"
        if ($foundUsers -and $foundUsers.Count -gt 0) {
            $existingText = $addTextBox.Text.Trim()
            $newLines = @()
            if ($existingText) {
                $newLines += $existingText.Split("`n") | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
            }
            foreach ($user in $foundUsers) {
                # Add userPrincipalName if available, otherwise display name
                $userIdentifier = $user.UserPrincipalName
                if ([string]::IsNullOrWhiteSpace($userIdentifier)) {
                    $userIdentifier = $user.DisplayName
                }
                if ($userIdentifier -notin $newLines) {
                    $newLines += $userIdentifier
                }
            }
            $addTextBox.Text = ($newLines -join "`n")
        }
    })
    $form.Controls.Add($searchButton)
    
    $addButton = New-Object System.Windows.Forms.Button
    $addButton.Text = "Add Users"
    $addButton.Location = New-Object System.Drawing.Point(580, 265)
    $addButton.Size = New-Object System.Drawing.Size(100, 30)
    $addButton.BackColor = [System.Drawing.Color]::LightGreen
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
    $removeButton.BackColor = [System.Drawing.Color]::LightCoral
    $removeButton.ForeColor = [System.Drawing.Color]::DarkRed
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
                
                # Check if policy has groups or roles that can satisfy the requirement
                $hasGroups = $currentPolicy.Conditions.Users.IncludeGroups -and $currentPolicy.Conditions.Users.IncludeGroups.Count -gt 0
                $hasRoles = $currentPolicy.Conditions.Users.IncludeRoles -and $currentPolicy.Conditions.Users.IncludeRoles.Count -gt 0
                
                # If no users left, we need at least something or the policy will be invalid
                if ($newIncludeList.Count -eq 0) {
                    if ($hasGroups -or $hasRoles) {
                        # Policy has groups or roles, so it's OK to have no users
                        Write-Host "Removing all users, but policy has groups/roles so it remains valid." -ForegroundColor Yellow
                    } else {
                        # No groups or roles - must have at least "All" users
                        $confirmEmpty = [System.Windows.Forms.MessageBox]::Show("This will remove all specific users.`n`nThe policy requires at least one of:`n- Included Users`n- Included Groups`n- Included Roles`n`nSince there are no groups or roles, this will default to 'All Users'. Continue?", "Warning", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
                        if ($confirmEmpty -eq [System.Windows.Forms.DialogResult]::No) {
                            return
                        }
                        # Default to "All" when no users, groups, or roles remain
                        $newIncludeList = @("All")
                    }
                }
                
                $userConditions = @{
                    IncludeUsers = $newIncludeList
                    ExcludeUsers = $currentPolicy.Conditions.Users.ExcludeUsers
                }
                
                # Preserve groups and roles if they exist
                if ($hasGroups) {
                    $userConditions.IncludeGroups = $currentPolicy.Conditions.Users.IncludeGroups
                }
                if ($hasRoles) {
                    $userConditions.IncludeRoles = $currentPolicy.Conditions.Users.IncludeRoles
                }
                if ($currentPolicy.Conditions.Users.ExcludeGroups) {
                    $userConditions.ExcludeGroups = $currentPolicy.Conditions.Users.ExcludeGroups
                }
                if ($currentPolicy.Conditions.Users.ExcludeRoles) {
                    $userConditions.ExcludeRoles = $currentPolicy.Conditions.Users.ExcludeRoles
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
    $closeButton.BackColor = [System.Drawing.Color]::LightGray
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

function Show-HelpDialog {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Help - Conditional Access Manager"
    $form.Size = New-Object System.Drawing.Size(750, 650)
    $form.StartPosition = "CenterParent"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    
    # Rich text box for help content
    $helpTextBox = New-Object System.Windows.Forms.RichTextBox
    $helpTextBox.Location = New-Object System.Drawing.Point(10, 10)
    $helpTextBox.Size = New-Object System.Drawing.Size(720, 570)
    $helpTextBox.ReadOnly = $true
    $helpTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $helpTextBox.BackColor = [System.Drawing.Color]::White
    $helpTextBox.Multiline = $true
    $helpTextBox.WordWrap = $true
    $helpTextBox.ScrollBars = "Vertical"
    
    $helpContent = @"
CONDITIONAL ACCESS MANAGER - HELP GUIDE

═══════════════════════════════════════════════════════════════
QUICK START
═══════════════════════════════════════════════════════════════

1. Click "Connect to Microsoft Graph" to authenticate
2. Select a tab: "Named Locations" or "Conditional Access Policies"
3. Use the color-coded buttons to perform actions
4. Refresh lists after making changes

═══════════════════════════════════════════════════════════════
CONNECTION MANAGEMENT
═══════════════════════════════════════════════════════════════

• Connect: Authenticate with Microsoft Graph (browser popup)
• Reconnect/Change Tenant: Switch to a different tenant
• Disconnect: End current session
• Reset Auth: Clear stuck authentication sessions if connection hangs
• Status shows "Connecting..." during authentication, then connection status

═══════════════════════════════════════════════════════════════
NAMED LOCATIONS TAB
═══════════════════════════════════════════════════════════════

CREATE COUNTRY LOCATION (Green button)
  - Click "Create Country Location"
  - Enter a display name
  - Select countries from the picker
  - Optionally include unknown/future countries

EDIT COUNTRIES (Orange button)
  - Select a location from the list
  - Modify countries or settings
  - Save changes

COPY COUNTRIES (Light Blue button)
  - Select a location to duplicate
  - Enter new name
  - Creates exact copy with all settings

RENAME (Orange button)
  - Select a location
  - Enter new display name

DELETE (Red button)
  - Select one or more locations
  - Cannot delete if referenced by policies
  - Shows summary of deleted/skipped items

REFERENCE TRACKING
  - "Referenced Policies" column shows which policies use each location
  - Prevents accidental deletion of in-use locations

═══════════════════════════════════════════════════════════════
CONDITIONAL ACCESS POLICIES TAB
═══════════════════════════════════════════════════════════════

MANAGE INCLUDED USERS (Light Blue button)
  - Add/remove users from policy scope
  - Use search to find users by name or email
  - Supports partial matching (e.g., "john" finds "john.doe@domain.com")
  - Multi-select support for bulk operations

MANAGE USER EXCEPTIONS (Light Blue button)
  - Add users to exclusion list
  - Search and select multiple users
  - Users are excluded from policy enforcement

RENAME (Orange button)
  - Select a policy
  - Update display name

COPY POLICY (Light Blue button)
  - Duplicates policy with all settings
  - IMPORTANT: Copied policies are created as DISABLED for safety
  - Enable manually after reviewing settings

DELETE (Red button)
  - Select one or more policies
  - Extra confirmation for enabled policies
  - Shows summary of deleted items

REFERENCE TRACKING
  - "Referenced Locations" column shows which locations are used
  - Indicates if locations are included or excluded

═══════════════════════════════════════════════════════════════
USER SEARCH FEATURES
═══════════════════════════════════════════════════════════════

• Search by full or partial name (e.g., "john", "smith")
• Search by full or partial email (e.g., "john@", "@domain.com")
• Results show display name and email address
• Select multiple users at once
• Client-side fallback handles partial matches

═══════════════════════════════════════════════════════════════
BUTTON COLOR CODING
═══════════════════════════════════════════════════════════════

🟢 Green: Add/Create actions
🔴 Red: Delete/Remove actions
🔵 Blue: Refresh/Search actions
🟠 Orange: Edit/Rename actions
🔵 Light Blue: Copy/Manage actions
⚪ Gray: Cancel/Close actions

═══════════════════════════════════════════════════════════════
IMPORTANT TIPS
═══════════════════════════════════════════════════════════════

• Always refresh lists after making changes or switching tenants
• Copied policies are DISABLED by default - enable manually if needed
• Cannot delete locations that are referenced by policies
• Use "Reset Auth" if authentication appears stuck or times out
• Multi-select works for bulk deletion operations
• Reference columns help identify dependencies before deletion

═══════════════════════════════════════════════════════════════
TROUBLESHOOTING
═══════════════════════════════════════════════════════════════

Connection Issues:
  - If authentication hangs, click "Reset Auth"
  - Check network connectivity
  - Verify Microsoft.Graph module is installed
  - Ensure you have required permissions

Module Errors:
  - Install: Install-Module Microsoft.Graph -Scope CurrentUser
  - Restart application after installation
  - Check PowerShell execution policy if needed

Search Not Working:
  - Ensure User.Read.All permission is granted
  - Try partial searches instead of exact matches
  - Check user exists in the tenant

═══════════════════════════════════════════════════════════════
REQUIRED PERMISSIONS
═══════════════════════════════════════════════════════════════

• Policy.Read.All
• Policy.ReadWrite.ConditionalAccess
• User.Read.All
• Group.Read.All
• Organization.Read.All

Grant these during the initial authentication flow.
"@
    $helpTextBox.Text = $helpContent
    $form.Controls.Add($helpTextBox)

    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "Close"
    $closeButton.Location = New-Object System.Drawing.Point(655, 590)
    $closeButton.Size = New-Object System.Drawing.Size(75, 23)
    $closeButton.BackColor = [System.Drawing.Color]::LightGray
    $closeButton.Add_Click({ $form.Close() })
    $form.Controls.Add($closeButton)

    $form.ShowDialog() | Out-Null
}

function Show-AboutDialog {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "About - Conditional Access Manager"
    $form.Size = New-Object System.Drawing.Size(520, 420)
    $form.StartPosition = "CenterParent"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false
    
    # Title Label
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Text = "Conditional Access Manager"
    $titleLabel.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(470, 30)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $form.Controls.Add($titleLabel)
    
    # Version Label
    $versionLabel = New-Object System.Windows.Forms.Label
    $versionLabel.Text = "Version 3.5"
    $versionLabel.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $versionLabel.Location = New-Object System.Drawing.Point(20, 55)
    $versionLabel.Size = New-Object System.Drawing.Size(470, 25)
    $versionLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $form.Controls.Add($versionLabel)
    
    # Description Label
    $descLabel = New-Object System.Windows.Forms.Label
    $descLabel.Text = "A powerful GUI-based PowerShell tool for managing Microsoft Entra (Azure AD) Conditional Access policies and Named Locations."
    $descLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $descLabel.Location = New-Object System.Drawing.Point(20, 85)
    $descLabel.Size = New-Object System.Drawing.Size(470, 45)
    $descLabel.TextAlign = [System.Drawing.ContentAlignment]::TopCenter
    $form.Controls.Add($descLabel)
    
    # Info Rich Text Box
    $infoTextBox = New-Object System.Windows.Forms.RichTextBox
    $infoTextBox.Location = New-Object System.Drawing.Point(20, 135)
    $infoTextBox.Size = New-Object System.Drawing.Size(470, 200)
    $infoTextBox.ReadOnly = $true
    $infoTextBox.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $infoTextBox.BackColor = [System.Drawing.Color]::White
    $infoTextBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $infoTextBox.Multiline = $true
    $infoTextBox.WordWrap = $true
    
    $infoContent = @"
Repository: https://github.com/monobrau/ca_manager

License: MIT License

Requirements:
- Windows PowerShell 5.1+
- Microsoft.Graph PowerShell module
- Microsoft Entra admin permissions

Built with:
- PowerShell Windows Forms
- Microsoft Graph PowerShell SDK
- PS2EXE (for executable compilation)
"@
    $infoTextBox.Text = $infoContent
    $form.Controls.Add($infoTextBox)
    
    # Close Button
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Text = "Close"
    $closeButton.Location = New-Object System.Drawing.Point(415, 345)
    $closeButton.Size = New-Object System.Drawing.Size(75, 23)
    $closeButton.BackColor = [System.Drawing.Color]::LightGray
    $closeButton.Add_Click({ $form.Close() })
    $form.Controls.Add($closeButton)
    
    $form.ShowDialog() | Out-Null
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
    $script:disconnectButton.BackColor = [System.Drawing.Color]::LightCoral
    $script:disconnectButton.Add_Click({
        Disconnect-GraphAPI
    })
    $connectionPanel.Controls.Add($script:disconnectButton)

    # Reset Auth Button
    $script:resetAuthButton = New-Object System.Windows.Forms.Button
    $script:resetAuthButton.Text = "Reset Auth"
    $script:resetAuthButton.Location = New-Object System.Drawing.Point(470, 0)
    $script:resetAuthButton.Size = New-Object System.Drawing.Size(100, 30)
    $script:resetAuthButton.BackColor = [System.Drawing.Color]::Orange
    $script:resetAuthButton.Add_Click({
        Reset-Authentication
    })
    $connectionPanel.Controls.Add($script:resetAuthButton)

    # Status Label
    $script:statusLabel = New-Object System.Windows.Forms.Label
    $script:statusLabel.Text = "Not connected"
    $script:statusLabel.Location = New-Object System.Drawing.Point(580, 5)
    $script:statusLabel.Size = New-Object System.Drawing.Size(200, 20)
    $script:statusLabel.ForeColor = [System.Drawing.Color]::Red
    $connectionPanel.Controls.Add($script:statusLabel)

    # About Button
    $aboutButton = New-Object System.Windows.Forms.Button
    $aboutButton.Text = "About"
    $aboutButton.Location = New-Object System.Drawing.Point(790, 0)
    $aboutButton.Size = New-Object System.Drawing.Size(80, 30)
    $aboutButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $aboutButton.BackColor = [System.Drawing.Color]::LightGray
    $aboutButton.Add_Click({
        Show-AboutDialog
    })
    $connectionPanel.Controls.Add($aboutButton)
    $aboutButton.BringToFront()

    # Help Button
    $helpButton = New-Object System.Windows.Forms.Button
    $helpButton.Text = "Help"
    $helpButton.Location = New-Object System.Drawing.Point(880, 0)
    $helpButton.Size = New-Object System.Drawing.Size(80, 30)
    $helpButton.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $helpButton.BackColor = [System.Drawing.Color]::LightGray
    $helpButton.Add_Click({
        Show-HelpDialog
    })
    $connectionPanel.Controls.Add($helpButton)
    $helpButton.BringToFront()

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
    $script:namedLocationsListView.Columns.Add("Type", 100) | Out-Null
    $script:namedLocationsListView.Columns.Add("Details", 300) | Out-Null
    $script:namedLocationsListView.Columns.Add("Referenced Policies", 400) | Out-Null
    $namedLocationsTab.Controls.Add($script:namedLocationsListView)

    # Named Locations Buttons
    $nlRefreshButton = New-Object System.Windows.Forms.Button
    $nlRefreshButton.Text = "Refresh"
    $nlRefreshButton.Location = New-Object System.Drawing.Point(10, 15)
    $nlRefreshButton.Size = New-Object System.Drawing.Size(80, 25)
    $nlRefreshButton.BackColor = [System.Drawing.Color]::LightBlue
    $nlRefreshButton.Add_Click({
        Refresh-NamedLocationsList $script:namedLocationsListView
    })
    $namedLocationsTab.Controls.Add($nlRefreshButton)

    $nlCreateButton = New-Object System.Windows.Forms.Button
    $nlCreateButton.Text = "Create Country Location"
    $nlCreateButton.Location = New-Object System.Drawing.Point(100, 15)
    $nlCreateButton.Size = New-Object System.Drawing.Size(150, 25)
    $nlCreateButton.BackColor = [System.Drawing.Color]::LightGreen
    $nlCreateButton.Add_Click({
        Show-CreateCountryLocationDialog $script:namedLocationsListView
    })
    $namedLocationsTab.Controls.Add($nlCreateButton)

    $nlEditButton = New-Object System.Windows.Forms.Button
    $nlEditButton.Text = "Edit Countries"
    $nlEditButton.Location = New-Object System.Drawing.Point(260, 15)
    $nlEditButton.Size = New-Object System.Drawing.Size(100, 25)
    $nlEditButton.BackColor = [System.Drawing.Color]::Orange
    $nlEditButton.Add_Click({
        Edit-SelectedNamedLocation $script:namedLocationsListView
    })
    $namedLocationsTab.Controls.Add($nlEditButton)

    $nlCopyButton = New-Object System.Windows.Forms.Button
    $nlCopyButton.Text = "Copy Countries"
    $nlCopyButton.Location = New-Object System.Drawing.Point(370, 15)
    $nlCopyButton.Size = New-Object System.Drawing.Size(100, 25)
    $nlCopyButton.BackColor = [System.Drawing.Color]::LightSteelBlue
    $nlCopyButton.Add_Click({
        Copy-SelectedNamedLocation $script:namedLocationsListView
    })
    $namedLocationsTab.Controls.Add($nlCopyButton)

    $nlRenameButton = New-Object System.Windows.Forms.Button
    $nlRenameButton.Text = "Rename"
    $nlRenameButton.Location = New-Object System.Drawing.Point(480, 15)
    $nlRenameButton.Size = New-Object System.Drawing.Size(80, 25)
    $nlRenameButton.BackColor = [System.Drawing.Color]::DarkOrange
    $nlRenameButton.ForeColor = [System.Drawing.Color]::White
    $nlRenameButton.Add_Click({
        Rename-SelectedNamedLocation $script:namedLocationsListView
    })
    $namedLocationsTab.Controls.Add($nlRenameButton)

    $nlDeleteButton = New-Object System.Windows.Forms.Button
    $nlDeleteButton.Text = "Delete"
    $nlDeleteButton.Location = New-Object System.Drawing.Point(570, 15)
    $nlDeleteButton.Size = New-Object System.Drawing.Size(80, 25)
    $nlDeleteButton.BackColor = [System.Drawing.Color]::LightCoral
    $nlDeleteButton.ForeColor = [System.Drawing.Color]::DarkRed
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
    $script:policiesListView.Columns.Add("Included Users", 250) | Out-Null
    $script:policiesListView.Columns.Add("Excluded Users", 250) | Out-Null
    $script:policiesListView.Columns.Add("Referenced Locations", 300) | Out-Null
    $policiesTab.Controls.Add($script:policiesListView)

    # Policies Buttons
    $polRefreshButton = New-Object System.Windows.Forms.Button
    $polRefreshButton.Text = "Refresh"
    $polRefreshButton.Location = New-Object System.Drawing.Point(10, 15)
    $polRefreshButton.Size = New-Object System.Drawing.Size(80, 25)
    $polRefreshButton.BackColor = [System.Drawing.Color]::LightBlue
    $polRefreshButton.Add_Click({
        Refresh-PoliciesList $script:policiesListView
    })
    $policiesTab.Controls.Add($polRefreshButton)

    $polManageIncludedButton = New-Object System.Windows.Forms.Button
    $polManageIncludedButton.Text = "Manage Included Users"
    $polManageIncludedButton.Location = New-Object System.Drawing.Point(100, 15)
    $polManageIncludedButton.Size = New-Object System.Drawing.Size(140, 25)
    $polManageIncludedButton.BackColor = [System.Drawing.Color]::LightSteelBlue
    $polManageIncludedButton.Add_Click({
        Show-ManageIncludedUsersDialog $script:policiesListView
    })
    $policiesTab.Controls.Add($polManageIncludedButton)

    $polManageUsersButton = New-Object System.Windows.Forms.Button
    $polManageUsersButton.Text = "Manage User Exceptions"
    $polManageUsersButton.Location = New-Object System.Drawing.Point(250, 15)
    $polManageUsersButton.Size = New-Object System.Drawing.Size(150, 25)
    $polManageUsersButton.BackColor = [System.Drawing.Color]::LightSteelBlue
    $polManageUsersButton.Add_Click({
        Show-ManageUserExceptionsDialog $script:policiesListView
    })
    $policiesTab.Controls.Add($polManageUsersButton)

    $polRenameButton = New-Object System.Windows.Forms.Button
    $polRenameButton.Text = "Rename Policy"
    $polRenameButton.Location = New-Object System.Drawing.Point(410, 15)
    $polRenameButton.Size = New-Object System.Drawing.Size(100, 25)
    $polRenameButton.BackColor = [System.Drawing.Color]::DarkOrange
    $polRenameButton.ForeColor = [System.Drawing.Color]::White
    $polRenameButton.Add_Click({
        Rename-SelectedPolicy $script:policiesListView
    })
    $policiesTab.Controls.Add($polRenameButton)

    $polCopyButton = New-Object System.Windows.Forms.Button
    $polCopyButton.Text = "Copy Policy"
    $polCopyButton.Location = New-Object System.Drawing.Point(520, 15)
    $polCopyButton.Size = New-Object System.Drawing.Size(100, 25)
    $polCopyButton.BackColor = [System.Drawing.Color]::LightSteelBlue
    $polCopyButton.Add_Click({
        Copy-SelectedPolicy $script:policiesListView
    })
    $policiesTab.Controls.Add($polCopyButton)

    $polExceptionButton = New-Object System.Windows.Forms.Button
    $polExceptionButton.Text = "Geo-IP Exception"
    $polExceptionButton.Location = New-Object System.Drawing.Point(630, 15)
    $polExceptionButton.Size = New-Object System.Drawing.Size(120, 25)
    $polExceptionButton.BackColor = [System.Drawing.Color]::LightGreen
    $polExceptionButton.Add_Click({
        Create-GeoIpExceptionForPolicy $script:policiesListView
    })
    $policiesTab.Controls.Add($polExceptionButton)

    $polDeleteButton = New-Object System.Windows.Forms.Button
    $polDeleteButton.Text = "Delete Policy"
    $polDeleteButton.Location = New-Object System.Drawing.Point(760, 15)
    $polDeleteButton.Size = New-Object System.Drawing.Size(100, 25)
    $polDeleteButton.BackColor = [System.Drawing.Color]::LightCoral
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
    $script:mainForm = $mainForm  # Store reference for authentication
    Write-Host "Showing form..." -ForegroundColor Green
    [void]$mainForm.ShowDialog()
} catch {
    Write-Error "Error: $($_.Exception.Message)"
    Read-Host "Press Enter to exit"
}



