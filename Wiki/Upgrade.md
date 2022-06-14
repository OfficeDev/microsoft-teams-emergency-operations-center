## Upgrade to version 0.5.1

### Sync latest version

1.  In Azure portal, navigate to TEOC resource group, open app service and click on deployment center.

    ![AppService](images/Update1.PNG)

2.  Click on sync.

    ![Sync](images/Update2.PNG)

3.  Wait until you see status as success for sync.

    ![SyncLog](images/Update3.PNG)

### Update version

1.  Delete existing app from teams admin center.

2.  Refer [6.Create the Teams app packages](https://github.com/OfficeDev/microsoft-teams-emergency-operations-center/wiki/Deployment-Guide#6-create-the-teams-app-packages) section of deployment guide.

    >Note: If you already have "AppPackage" folder with the manifest file, then update app version in manifest file from 0.5 to 0.5.1 and rezip the files.
    
    ![Version](images/Update4.PNG)

3.  Refer [7.Install the app in Microsoft Teams](https://github.com/OfficeDev/microsoft-teams-emergency-operations-center/wiki/Deployment-Guide#7-install-the-app-in-microsoft-teams) section from deployment guide to upload updated zip.
