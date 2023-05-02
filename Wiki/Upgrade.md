## Upgrade to version 2.0 from 1.0 or 0.5/0.5.1 

### 1. Run the upgrade PowerShell script
  
  1. Clone the [repository](https://github.com/OfficeDev/microsoft-teams-emergency-operations-center.git) locally. Open the `Deployment/provisioning/Upgrade` folder to get the latest provisioning files i.e `Update-EOC-Provision.ps1` and `Update-EOC-SiteTemplate-From-V1.0.xml` ( If you are upgrading the app from 1.0 version ) or `Update-EOC-SiteTemplate-From-V0.5.xml` ( If you are upgrading the app from 0.5 or 0.5.1 version )

  2. Run the PowerShell script (Update-EOC-Provision.ps1) as an Administrator, script will ask for below inputs:
    
    * XML file path – enter fully qualified path of the XML file (Ex: C:\Scripts\Update-EOC-SiteTemplate-From-V1.0.xml) 
    * Tenant Name – Name of the tenant where the SharePoint site was already provisioned (Ex: Contoso)
    * SharePoint site name – Enter your existing TEOC site name (Ex: TEOCSite)
 
  ![Provisioning Scripts](./Images/UpgradeScript.png)

  3. Once the above script is completed, then run the script (`EOC-UpdateLookupColumn.ps1`) as an Administrator, script will ask for below inputs:

    * Enter the TEOC Site URL - enter the absolute URL of the TEOC SharePoint site. (Ex: https://contoso.sharepoint.com/sites/TEOCSite)

  ![Provisioning Scripts](./Images/LookupColumnScript.png)

  This script will copy the value from the existing IncidentStatus column to a new Status column in the TEOC-IncidentTransaction List. Please wait for the script to complete as it will take sometime depending on the number of items on the list.

### 2. Add new application setting in the App Service ( **This step is applicable only if you upgrade the app from 0.5 or 0.5.1 version** )

  1. In Azure portal, navigate to TEOC resource group, open app service and click on Configuration.

  2. Click on "New application setting" and add **REACT_APP_SHAREPOINT_SITE_NAME** as the Name and **TEOCSite** as the value and click on OK.
  >Note: If you have already customized the site name during inital deployment, then use that sitename instead of TEOCSite.

 ![AppService](./Images/AddAppSetting.png)
  

### 3. Sync latest version

  1.  In Azure portal, navigate to TEOC resource group, open app service and click on deployment center.

![AppService](./Images/Update1.png)

  2.  Click on sync.

![Sync](./Images/Update2.png)

  3.  Wait until you see status as success for sync.

![SyncLog](./Images/Update3.png)

### 4. Update version

1.  Delete existing app from teams admin center.

2.  Refer [6.Create the Teams app packages](https://github.com/OfficeDev/microsoft-teams-emergency-operations-center/wiki/Deployment-Guide#6-create-the-teams-app-packages) section of deployment guide.

3.  Refer [7.Install the app in Microsoft Teams](https://github.com/OfficeDev/microsoft-teams-emergency-operations-center/wiki/Deployment-Guide#7-install-the-app-in-microsoft-teams) section from deployment guide to upload updated zip.

