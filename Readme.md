# 🔐 Microsoft Entra Conditional Access Management Tool (geo-IP blocking focus)

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
- **Delete** unwanted named locations with confirmation
- **Bulk country management** - add multiple countries at once
- **Visual list view** showing location details and types

### 👥 Conditional Access Policy Management
- **View all policies** with detailed user inclusion/exclusion information
- **Manage included users** - add/remove users from policy scope
- **Manage user exceptions** - easily exclude users from policies
- **Copy policies** - duplicate existing policies with all settings preserved
- **Rename policies** - update policy display names with simple dialogs
- **User resolution** - search by email, display name, or User ID
- **Bulk user operations** - add multiple users at once
- **Real-time policy updates** with immediate Graph API synchronization

### 🔌 Microsoft Graph Integration
- **Secure authentication** with required scopes
- **Multi-tenant support** - connect to different tenants
- **Automatic token management** and refresh
- **Comprehensive error handling** with detailed error messages
- **Real-time status updates** showing connection state

## 📋 Prerequisites

- **Windows PowerShell 5.1** (not PowerShell Core)
- **Microsoft.Graph PowerShell module**
- **Windows Forms** support (.NET Framework)
- **Microsoft Entra admin permissions**:
  - `Policy.Read.All`
  - `Policy.ReadWrite.ConditionalAccess`
  - `User.Read.All`
  - `Group.Read.All`
  - `Organization.Read.All`

## 🛠️ Installation

1. **Clone the repository**:
   ```powershell
   git clone https://github.com/yourusername/conditional-access-manager.git
   cd conditional-access-manager
   ```

2. **Install Microsoft Graph module** (if not already installed):
   ```powershell
   Install-Module Microsoft.Graph -Scope CurrentUser
   ```

3. **Run the script**:
   ```powershell
   .\ConditionalAccessManager.ps1
   ```

## 🎯 Quick Start

1. **Launch the tool** by running the PowerShell script
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
2. Click **"Create Country Location"**
3. Enter a display name (e.g., "Blocked Countries")
4. Add country codes separated by commas (e.g., "US,CA,GB,DE")
5. Optionally check "Include unknown/future countries"
6. Click **"Create"**

#### Editing Existing Locations
1. Select a country-based named location from the list
2. Click **"Edit Countries"**
3. Modify the settings as needed
4. Click **"Edit"** to save changes

### Managing Conditional Access Policies

#### Adding User Exceptions
1. Go to the **"Conditional Access Policies"** tab
2. Select a policy from the list
3. Click **"Manage User Exceptions"**
4. Add users by email, display name, or User ID (one per line)
5. Click **"Add Users"**

#### Managing Included Users
1. Select a policy from the list
2. Click **"Manage Included Users"**
3. Choose between "All Users" or specific user list
4. Add/remove users as needed

#### Copying Policies
1. Select a policy from the list
2. Click **"Copy Policy"**
3. Enter a name for the new policy
4. The new policy will be created in **DISABLED** state for safety
5. Review settings and enable manually when ready

#### Renaming Policies
1. Select a policy from the list
2. Click **"Rename Policy"**
3. Enter the new display name
4. Click OK to save changes

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

**"Unexpected token" error**
- Ensure you're using Windows PowerShell 5.1, not PowerShell Core
- Run: `$PSVersionTable.PSEdition` (should show "Desktop")

**Microsoft Graph module not found**
```powershell
Install-Module Microsoft.Graph -Scope CurrentUser -Force
```

**Authentication failures**
- Verify you have the required admin permissions
- Try the "Reconnect" button to refresh your session
- Check if MFA is properly configured

**GUI not appearing**
- Ensure Windows Forms is available
- Try running PowerShell as Administrator
- Restart PowerShell ISE/terminal and try again

### Error Logging

The tool provides detailed error messages with:
- Specific error descriptions
- Suggested fixes for common issues
- API response details for debugging

## 🔧 Advanced Features

### Bulk Operations
- Add multiple users to policies simultaneously
- Copy named locations with all settings preserved
- Copy complete policies with all conditions and controls
- Batch country code updates

### Policy Safety Features
- **Copied policies created as DISABLED** - Prevents accidental activation
- **Comprehensive setting preservation** - All conditions, controls, and sessions copied
- **Validation and error handling** - Detailed feedback for copy operations
- **Permission verification** - Ensures proper Graph API permissions before operations

### Smart User Resolution
- Automatically resolves users by email or display name
- Handles GUID-based User IDs
- Provides clear feedback for users not found

### Real-time Synchronization
- Immediate updates to Microsoft Graph
- Automatic list refreshing after changes
- Connection status monitoring

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

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 👨‍💻 Author

**Gemini** - *Initial work and development*

## 🙏 Acknowledgments

- Microsoft Graph PowerShell SDK team
- Microsoft Entra (Azure AD) documentation contributors
- PowerShell community for Windows Forms guidance

## 📊 Version History

- **v3.5** - Added policy copying and renaming capabilities with safety features
- **v3.4** - Current stable version with all features working
- **v3.3** - Added user management capabilities
- **v3.2** - Enhanced named location management
- **v3.1** - Improved error handling and UI
- **v3.0** - Initial release with basic functionality

## 🔗 Related Links

- [Microsoft Graph PowerShell SDK](https://docs.microsoft.com/en-us/powershell/microsoftgraph/)
- [Microsoft Entra Conditional Access](https://docs.microsoft.com/en-us/azure/active-directory/conditional-access/)
- [Named Locations Documentation](https://docs.microsoft.com/en-us/azure/active-directory/conditional-access/location-condition)

---

⭐ **Star this repository** if you find it helpful!

🐛 **Report issues** on the [Issues page](https://github.com/yourusername/conditional-access-manager/issues)

💬 **Questions?** Start a [Discussion](https://github.com/yourusername/conditional-access-manager/discussions)
