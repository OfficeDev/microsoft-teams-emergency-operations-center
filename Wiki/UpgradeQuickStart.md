# TEOC Upgrade Quick Start Guide

This guide helps you quickly determine what version you have and which upgrade steps you need to follow to upgrade to the latest version (v3.4).

## Step 1: Identify Your Current Version

Before starting the upgrade, you need to know which version you currently have installed:

### Option A: Check in Teams App
1. Open Microsoft Teams
2. Go to your TEOC app
3. Look at the app details or about section for version information

### Option B: Check SharePoint Site
1. Navigate to your TEOC SharePoint site
2. Check the site contents and lists structure:
   - **v0.5/0.5.1**: Basic incident list structure
   - **v1.0**: Has incident types and roles management
   - **v2.0**: Has active dashboard features and role assignments
   - **v3.x**: Has enhanced permissions and workspace-based App Insights

### Option C: Check Azure App Service
1. Go to Azure Portal
2. Navigate to your TEOC resource group
3. Open the App Service
4. Check the Application Settings for version-specific configurations:
   - `REACT_APP_SHAREPOINT_SITE_NAME` (added in v1.0+)
   - `GENERATE_SOURCEMAP` (added in v3.1+)

### Option D: Feature-based Identification
Compare your current features with the [Version Feature Comparison Table](./VersionComparison.md) to identify your version based on available functionality.

## Step 2: Choose Your Upgrade Path

Based on your current version, follow the appropriate upgrade path:

### Upgrading from v3.3 to v3.4
✅ **Simplest upgrade path** - Only requires steps 6, 10, and 11 from the main upgrade guide
- [Sync latest version](#quick-sync-steps)
- [Update Teams app package](#quick-package-update)
- [Launch and test](#quick-launch-test)

### Upgrading from v3.0-v3.2 to v3.4
- Add `GENERATE_SOURCEMAP` setting (Step 5)
- [Sync latest version](#quick-sync-steps)
- [Update Teams app package](#quick-package-update)
- [Launch and test](#quick-launch-test)

### Upgrading from v2.x to v3.4
⚠️ **Medium complexity** - Requires API permissions and App Insights migration
- Follow steps 6-11 from the main upgrade guide
- **Important**: App Insights migration is required (Step 9)
- Additional Graph API permissions needed (Step 7)
- Exchange Online permissions needed (Step 8)

### Upgrading from v1.0 to v3.4
⚠️ **Complex upgrade** - Requires SharePoint changes and scripts
- Follow steps 2-11 from the main upgrade guide
- **Important**: PowerShell scripts required (Step 2)
- Column updates needed (Step 3)
- All permission updates required

### Upgrading from v0.5/v0.5.1 to v3.4
🚨 **Most complex upgrade** - All steps required
- Follow ALL steps 1-11 from the main upgrade guide
- **Important**: Location column modification required (Step 1)
- App Service settings required (Step 4)

## Quick Action Steps

### Quick Sync Steps
1. Azure Portal → TEOC Resource Group → App Service
2. Click "Deployment Center"
3. Click "Sync"
4. Wait for "Success" status

### Quick Package Update
1. Download latest TEOC app package from [releases](https://github.com/OfficeDev/microsoft-teams-emergency-operations-center/releases)
2. Teams Admin Center → Teams apps → Manage apps
3. Find your TEOC app → Upload file
4. Upload the new package

### Quick Launch Test
1. Open TEOC in Teams
2. Click "Login" if prompted
3. Grant any new permissions required
4. Verify dashboard loads properly

## Pre-Upgrade Checklist

Before starting any upgrade:

- [ ] **Backup**: Take a backup of your SharePoint site and note current configuration
- [ ] **Permissions**: Ensure you have admin access to:
  - [ ] Azure subscription and TEOC resource group
  - [ ] SharePoint admin or site collection admin
  - [ ] Teams admin center
  - [ ] Azure AD app registration permissions
- [ ] **Downtime**: Plan for potential downtime during the upgrade
- [ ] **Testing**: Have a plan to test the app after upgrade
- [ ] **Support**: Identify who will handle any issues that arise

## Getting Help

If you encounter issues during upgrade:

1. **Check the detailed [Upgrade Guide](./Upgrade.md)** for complete step-by-step instructions
2. **Review [Version Feature Comparison](./VersionComparison.md)** to understand what you're upgrading to
3. **Review [Troubleshooting Guide](./Troubleshooting.md)** for common issues
4. **Check [FAQ](./FAQ.md)** for upgrade-related questions (see FAQ #7-12)
5. **Review [Release Notes](./ReleaseNotes.md)** to understand what's new in v3.4

## Emergency Rollback

If you need to rollback after upgrade issues:

1. **Teams App**: Upload your previous app package in Teams Admin Center
2. **Azure App Service**: Use Deployment Center to rollback to previous deployment
3. **SharePoint**: Restore from backup if changes were made to lists/columns

> **Need More Help?** 
> If this quick start guide doesn't cover your specific situation, refer to the complete [Upgrade Guide](./Upgrade.md) for detailed instructions, or check the [FAQ](./FAQ.md) for common upgrade questions.