# General Template issues

## Problems deploying to Azure

### 1. App service name already exists

`Website with given name <name> already exists.`

`Creating resources of type "Microsoft.Web/sites" failed with status "Conflict"`

This happens when Azure web apps with the name `[BaseResourceName]` or `[BaseResourceName]-config` already exist.

### Fix

Choose a different "Base Resource Name". You can check if your desired name is available by going to the page to [create a new web app](https://portal.azure.com/#create/Microsoft.WebSite) in the Azure Portal. Enter your desired name in the "App name" field. An error message will appear if the name you have chosen is taken or invalid.

### 2. Resource location doesn't support Application Insights

`The subscription is not registered for the resource type 'components' in the region '<region>'. Please re-register for this provider in order to have access to this location.`

`Creating resources of type "microsoft.insights/components" failed with status "Conflict"`

This happens when Application Insights is not available in the Azure datacenter location where you chose to create the app.

### Fix

Create the resource group in a location where Application Insights is available. For an up-to-date list of these locations, refer [here](https://azure.microsoft.com/en-us/global-infrastructure/services/?products=monitor), under "Application Insights".

## Problems installing the app in Teams

### 1. Missing entrypoint to upload custom app

You don't see the entry points to upload an app, either using the [Store](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/apps/apps-upload#upload-your-package-into-a-team-or-conversation-using-the-store) or using the [Apps tab](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/apps/apps-upload#upload-your-package-into-a-team-using-the-apps-tab).

This is most likely because sideloading of external apps is not enabled in your tenant. See [here](https://docs.microsoft.com/en-us/microsoftteams/admin-settings) for more information.

## Didn't find your problem here?

Please report the issue [here](https://github.com/OfficeDev/microsoft-teams-emergency-operations-center/issues/new)

