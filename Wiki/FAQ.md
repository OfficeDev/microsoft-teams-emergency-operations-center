# Known Limitations

## 1. Tasks in Active Bridge component
Tasks in Active Bridge component has some known limitations as below -

- User can only search and assign the tasks to the people who are most relevant to the user in the People picker in Add Task section.( *Relevance is determined by the user’s communication and collaboration patterns and business relationships. People can be local contacts or from an organization’s directory, and people from recent communications.* )

![Limitation](./Images/AddTask.png)

- In GCCH tenant, task will not get added when it's assigned to some one through the People Picker, however user can still create a task without an Assignee.

![Limitation](./Images/GCCHTasks.png)

# FAQs

## 1. Azure deployment failed with conflict error, what do I need to do?
Deployment can fail if the base resource name provided is not available and you could see conflict error. Choose a different "Base Resource Name". You can check if your desired name is available by going to the page to [create a new web app](https://portal.azure.com/#create/Microsoft.WebSite) in the Azure Portal. Enter your desired name in the "App name" field. An error message will appear if the name you have chosen is taken or invalid.

## 2. Provisioning failed with error, what do I need to do?
Provisioning can fail due to some network issue. If it failed and TEOC site collection is created in your tenant then delete the site collection from Active sites as well as Deleted sites. Run the provisioning script again.

![ProvisioningError](./Images/ProvisioningError.png)

## 3. Upgrading to latest version failed with error, what do I need to do?
App service deployment fails with the error message “Input string was not in a correct format” due to fsevents package.

![UpgradeError](./Images/NodeError.png)

The fsevents library is causing the app service deployment to fail for already deployed apps.

-	Go to portal.azure.com. 
-   Navigate to resource group where all TEOC resources are deployed.
-	Click on the app service.
-   Click on Configuration.
-	Click on WEBSITE_NODE_DEFAULT_VERSION.
-	Update the default value to 16.13.0 (previous value ~14) and save.
-	Click on overview and re-start the app service.
-	Once the app service is restarted. Navigate to Deployment Center and click on Sync.
-	Wait till the deployment is completed. You can validate this once the status changes to Success under logs. 


 ## 4. Does the app support multiple locales?

 Yes, TEOC v1.0 supports the translations for below 12 languages, 

- Arabic (SA)
- Chinese
- Chinese (TW)
- English (US)
- French
- German
- Hebrew
- Japanese
- Korean
- Portuguese (BR)
- Russian
- Spanish

## 5. Does the app support mobile devices?

Yes, TEOC v1.0 supports the desktop, mobile and tab devices.

## 6. Does the app works in GCC/GCCH tenant?

Yes, TEOC v2.0 works in Commercial, GCC and GCCH tenants.

## 7. How do I know which version of TEOC I currently have installed?

There are several ways to identify your current TEOC version:

**Method 1: Check Teams App Details**
- Open TEOC app in Microsoft Teams
- Look for version information in the app details or about section

**Method 2: Check Azure App Service Configuration**
- Go to Azure Portal → TEOC Resource Group → App Service → Configuration
- Look for version-specific application settings:
  - `REACT_APP_SHAREPOINT_SITE_NAME` (present in v1.0+)
  - `GENERATE_SOURCEMAP` (present in v3.1+)

**Method 3: Check SharePoint Site Structure**
- Navigate to your TEOC SharePoint site
- Examine the lists and columns structure to determine version capabilities

For detailed version identification steps, see the [Upgrade Quick Start Guide](./UpgradeQuickStart.md#step-1-identify-your-current-version).

## 8. What's the difference between upgrading from different versions?

The upgrade complexity depends on your starting version:

- **From v3.3 to v3.4**: Simple upgrade (3 steps) - just sync, update package, and test
- **From v3.0-v3.2 to v3.4**: Minimal upgrade (4 steps) - add one app setting plus sync and update
- **From v2.x to v3.4**: Moderate upgrade - requires API permission updates and App Insights migration
- **From v1.0 to v3.4**: Complex upgrade - requires PowerShell scripts and SharePoint changes
- **From v0.5/v0.5.1 to v3.4**: Most complex - requires all upgrade steps including column modifications

See the [Upgrade Quick Start Guide](./UpgradeQuickStart.md#step-2-choose-your-upgrade-path) for specific step requirements.

## 9. Can I skip intermediate versions and upgrade directly to v3.4?

Yes, you can upgrade directly from any previous version to v3.4. However, you'll need to complete all applicable upgrade steps for the versions you're skipping. The [main upgrade guide](./Upgrade.md) includes all necessary steps with clear indicators of which steps apply to which version ranges.

## 10. What should I do if my upgrade fails or causes issues?

If you encounter problems during upgrade:

1. **Check the error**: Review [FAQ #3](#3-upgrading-to-latest-version-failed-with-error-what-do-i-need-to-do) for common upgrade errors
2. **Review troubleshooting**: Check the [Troubleshooting Guide](./Troubleshooting.md) for deployment and app-specific issues
3. **Emergency rollback**: 
   - Teams App: Upload your previous app package in Teams Admin Center
   - Azure App Service: Use Deployment Center to rollback to previous deployment
   - SharePoint: Restore from backup if changes were made

## 11. Do I need to backup anything before upgrading?

Yes, it's recommended to backup:
- Your TEOC SharePoint site (especially if upgrading from v1.0 or older)
- Note your current Azure App Service configuration settings
- Document your current app package version
- Take screenshots of your current dashboard and incident data

For a complete pre-upgrade checklist, see the [Upgrade Quick Start Guide](./UpgradeQuickStart.md#pre-upgrade-checklist).

## 12. How long does an upgrade typically take?

Upgrade duration varies by version:
- **From v3.3**: 15-30 minutes
- **From v3.0-v3.2**: 30-45 minutes  
- **From v2.x**: 1-2 hours (includes permission updates and App Insights migration)
- **From v1.0**: 2-3 hours (includes PowerShell scripts and testing)
- **From v0.5/v0.5.1**: 3-4 hours (most comprehensive upgrade)

Additional time may be needed for testing and validation after the upgrade.