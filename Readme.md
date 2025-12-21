# 🔐 Conditional Access Manager

A powerful GUI-based PowerShell tool for managing Microsoft Entra (Azure AD) Conditional Access policies and Named Locations with ease.

![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue.svg)
![Microsoft Graph](https://img.shields.io/badge/Microsoft%20Graph-API-green.svg)
![Windows](https://img.shields.io/badge/Windows-Forms-lightgrey.svg)
![License](https://img.shields.io/badge/License-MIT-yellow.svg)

## 🚀 Features

### 📍 Named Locations Management
- **Create** country-based named locations with ease
- **Edit** existing country locations and update country codes
- **Copy** named locations to quickly duplicate configurations
- **Rename** locations with simple dialog prompts
- **Delete** named locations (with multi-item support)
- **Bulk operations** - delete multiple locations at once
- **Reference tracking** - see which policies reference each location
- **Visual list view** showing location details, types, and referenced policies
- **Validation** - prevents deletion of locations referenced by policies

### 👥 Conditional Access Policy Management
- **View all policies** with detailed user inclusion/exclusion information
- **Manage included users** - add/remove users from policy scope
- **Manage user exceptions** - easily exclude users from policies
- **Copy policies** - duplicate existing policies with all settings preserved
- **Rename policies** - update policy display names with simple dialogs
- **Delete policies** - remove policies with enhanced safety checks
- **Reference tracking** - see which named locations are referenced by each policy
- **User search** - search for users by name or email with partial matching
- **Bulk user operations** - add multiple users at once
- **Real-time policy updates** with immediate Graph API synchronization

### 🔍 User Search & Management
- **Advanced user search** - search by full or partial name or email address
- **Interactive search dialog** - browse and select users from search results
- **Multi-select support** - select multiple users at once
- **Smart resolution** - automatically resolves users by email, display name, or User ID
- **Client-side fallback** - handles partial matches even when API filters don't support it

### 🌍 Geo-IP Exception Creation
- **One-click exception creation** - create geo-IP exceptions with ease
- **Policy duplication** - automatically copies existing policy structure
- **Location cloning** - clones named locations with new country selections
- **User management** - add users to new policy and exclude from original
- **Country selection** - visual country picker for easy selection
- **Automatic policy configuration** - handles all policy settings automatically

### 🎨 User Interface
- **Colorized buttons** - intuitive color coding for different actions
  - 🟢 Green: Add/Create actions
  - 🔴 Red: Delete/Remove actions
  - 🔵 Blue: Search/Refresh actions
  - 🟠 Orange: Edit/Rename actions
  - 🔵 Light Blue: Copy/Manage actions
  - ⚪ Gray: Cancel/Close actions
- **Clean GUI experience** - no console window, no popup interruptions
- **Reference columns** - see policy/location relationships at a glance
- **Multi-item selection** - select and operate on multiple items

## 📋 Prerequisites

- **Windows 10/11** or **Windows Server 2016+**
- **PowerShell 5.1** or later
- **Microsoft.Graph PowerShell module**
- **Microsoft Entra admin permissions**:
  - `Policy.Read.All`
  - `Policy.ReadWrite.ConditionalAccess`
  - `User.Read.All`
  - `Group.Read.All`
  - `Organization.Read.All`

## 🛠️ Installation

### Option 1: Download Executable (Recommended)

1. **Download the latest release**:
   - Go to [Releases](https://github.com/monobrau/ca_manager/releases)
   - Download `ca_manager.exe` from the latest release

2. **Install Microsoft Graph module** (required):
   ```powershell
   Install-Module Microsoft.Graph -Scope CurrentUser
   ```
   Note: This may take several minutes as it installs multiple sub-modules.

3. **Run the executable**:
   - Double-click `ca_manager.exe`
   - Or run from PowerShell: `.\ca_manager.exe`

### Option 2: Run from Source

1. **Clone the repository**:
   ```powershell
   git clone https://github.com/monobrau/ca_manager.git
   cd ca_manager
   ```

2. **Install Microsoft Graph module** (if not already installed):
   ```powershell
   Install-Module Microsoft.Graph -Scope CurrentUser
   ```

3. **Run the script**:
   ```powershell
   .\ca2.ps1
   ```

## 🎯 Quick Start

1. **Launch the tool** by running `ca_manager.exe` or `ca2.ps1`
2. **Connect to Microsoft Graph** using the "Connect" button
3. **Authenticate** with your Microsoft Entra admin account
4. **Navigate** between the "Named Locations" and "Conditional Access Policies" tabs
5. **Manage your resources** using the intuitive GUI buttons

## 📖 Usage Guide

### Connecting to Microsoft Graph

1. Click **"Connect to Microsoft Graph"**
2. Complete the authentication flow in your browser
3. Grant the required permissions
4. The status will show your connected tenant

### Managing Named Locations

#### Creating a Country Location
1. Go to the **"Named Locations"** tab
2. Click **"Create Country Location"** (green button)
3. Enter a display name (e.g., "Blocked Countries")
4. Click **"Select Countries"** to open the country picker
5. Select countries from the list
6. Optionally check "Include unknown/future countries"
7. Click **"Create"**

#### Editing Existing Locations
1. Select a country-based named location from the list
2. Click **"Edit Countries"** (orange button)
3. Modify the settings as needed
4. Click **"Edit"** to save changes

#### Copying Locations
1. Select a named location from the list
2. Click **"Copy Countries"** (light blue button)
3. Enter a new name for the copy
4. The location will be duplicated with all settings

#### Deleting Locations
1. Select one or more locations from the list
2. Click **"Delete"** (red button)
3. Confirm the deletion
4. Note: Locations referenced by policies cannot be deleted (you'll see which policies reference them)

### Managing Conditional Access Policies

#### Adding User Exceptions
1. Go to the **"Conditional Access Policies"** tab
2. Select a policy from the list
3. Click **"Manage User Exceptions"** (light blue button)
4. Click **"Search Users"** (blue button) to find users, or
5. Add users by email, display name, or User ID (one per line)
6. Click **"Add Users"** (green button)

#### Managing Included Users
1. Select a policy from the list
2. Click **"Manage Included Users"** (light blue button)
3. Choose between "All Users" checkbox or specific user list
4. Click **"Search Users"** (blue button) to find and add users
5. Select users to remove and click **"Remove Selected"** (red button)
6. Click **"Add Users"** to add new users

#### Searching for Users
1. Click **"Search Users"** button in any user management dialog
2. Enter a search term (name or email - partial matches work)
3. Click **"Search"** or press Enter
4. Select one or more users from the results
5. Click **"Add Selected"** to add them to your list

#### Creating Geo-IP Exceptions
1. Select a policy that uses named locations
2. Click **"Geo-IP Exception"** (green button)
3. Enter a name for the new policy
4. Select the location to clone and modify
5. Enter a name for the new named location
6. Click **"Select Countries"** to choose countries for the exception
7. Click **"Search Users"** to add users to the exception
8. Click **"Create Exception"** (green button)
9. The tool will:
   - Create a new policy based on the original
   - Clone the location with new countries
   - Add specified users to the new policy
   - Exclude those users from the original policy

#### Copying Policies
1. Select a policy from the list
2. Click **"Copy Policy"** (light blue button)
3. Enter a name for the new policy
4. The new policy will be created in **DISABLED** state for safety
5. Review settings and enable manually when ready

#### Renaming Policies
1. Select a policy from the list
2. Click **"Rename Policy"** (orange button)
3. Enter the new display name
4. Click OK to save changes

#### Deleting Policies
1. Select one or more policies from the list
2. Click **"Delete Policy"** (red button)
3. Confirm the deletion (extra confirmation for enabled policies)
4. Review the summary of what was deleted/skipped

## ⚙️ Configuration

### Required Microsoft Graph Scopes

The tool automatically requests these permissions:

```powershell
$requiredScopes = @(
    "Policy.Read.All",
    "Policy.ReadWrite.ConditionalAccess", 
    "User.Read.All",
    "Group.Read.All",
    "Organization.Read.All"
)
```

### Multi-Tenant Support

Use the **"Reconnect/Change Tenant"** button to:
- Switch between different tenants
- Specify a particular tenant ID
- Re-authenticate with different credentials

## 🐛 Troubleshooting

### Common Issues

**Microsoft Graph module not found**
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```
After installation, restart the application.

**"Function capacity exceeded" error**
- The EXE version imports only required sub-modules to avoid this
- If running from script, ensure you're not importing the full Microsoft.Graph module

**Authentication failures**
- Verify you have the required admin permissions
- Try the "Reconnect" button to refresh your session
- Check if MFA is properly configured

**Location/Policy deletion blocked**
- Check the "Referenced Policies" or "Referenced Locations" columns
- Remove references before deleting
- The tool will show you which policies/locations are blocking deletion

**User search not finding users**
- Try partial matches (e.g., "beck" for "abecker@sfs.edu")
- The search uses client-side fallback for better partial matching
- Ensure you have User.Read.All permissions

### Error Handling

The tool provides:
- Detailed error messages with suggested fixes
- Reference checking before deletion
- Validation for policy operations
- Clear feedback for all operations

## 🔧 Advanced Features

### Bulk Operations
- **Multi-item deletion** - delete multiple locations or policies at once
- **Bulk user addition** - add multiple users simultaneously
- **Batch operations** - operations show summary of successes/failures

### Policy Safety Features
- **Copied policies created as DISABLED** - Prevents accidental activation
- **Enhanced deletion warnings** - Extra confirmation for enabled policies
- **Reference validation** - Prevents deletion of referenced resources
- **Comprehensive error handling** - Detailed feedback for all operations

### Smart User Resolution
- **Partial matching** - Search works with partial names/emails
- **Multiple search methods** - API search + client-side fallback
- **Automatic resolution** - Handles email, display name, or User ID
- **Clear feedback** - Shows which users were found/not found

### Reference Tracking
- **Policy references** - See which policies use each named location
- **Location references** - See which locations are used by each policy
- **Include/Exclude indicators** - Shows whether locations are included or excluded
- **Visual indicators** - Easy to spot dependencies

## 📊 Version History

- **v3.5** (Current)
  - Added "Reset Auth" button to clear stuck authentication sessions
  - Connection state tracking with visual "Connecting..." status
  - In-app help dialog with quick start guide
  - Tooltips on buttons for better user guidance
  - Enhanced connection error handling with specific timeout/cancel messages
  - Prevention of multiple simultaneous connection attempts
  - Improved UI feedback during authentication process

- **v3.4**
  - User search functionality with partial matching
  - Geo-IP exception creation workflow
  - Reference tracking columns
  - Multi-item deletion support
  - Colorized UI buttons
  - Silent mode (no console, no popups)
  - Enhanced error handling and validation

## 🤝 Contributing

We welcome contributions! Please follow these steps:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

### Development Guidelines

- Follow PowerShell best practices
- Add error handling for new features
- Update documentation for any new functionality
- Test with multiple tenant configurations
- Ensure PowerShell 5.1 compatibility for EXE builds

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- Microsoft Graph PowerShell SDK team
- Microsoft Entra (Azure AD) documentation contributors
- PowerShell community for Windows Forms guidance
- PS2EXE project for executable compilation support

## 🔗 Related Links

- [Microsoft Graph PowerShell SDK](https://docs.microsoft.com/en-us/powershell/microsoftgraph/)
- [Microsoft Entra Conditional Access](https://docs.microsoft.com/en-us/azure/active-directory/conditional-access/)
- [Named Locations Documentation](https://docs.microsoft.com/en-us/azure/active-directory/conditional-access/location-condition)

---

⭐ **Star this repository** if you find it helpful!

🐛 **Report issues** on the [Issues page](https://github.com/monobrau/ca_manager/issues)

💬 **Questions?** Start a [Discussion](https://github.com/monobrau/ca_manager/discussions)
