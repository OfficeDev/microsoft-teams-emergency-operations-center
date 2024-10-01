import React from "react";
import * as microsoftTeams from "@microsoft/teams-js";

/**
 * The 'Config' component is used to display your group tabs
 * user configuration options.  Here you will allow the user to
 * make their choices and once they are done you will need to validate
 * their choices and communicate that to Teams to enable the save button.
 */
class TabConfig extends React.Component {
  render() {
    // Initialize the Microsoft Teams SDK

    microsoftTeams.app.initialize().then(() => {
   
    /**
     * When the user clicks "Save", save the url for your configured tab.
     * This allows for the addition of query string parameters based on
     * the settings selected by the user.
     */
    microsoftTeams.pages.config.registerOnSaveHandler((saveEvent) => {
      const baseUrl = `https://${window.location.hostname}:${window.location.port}`;
      microsoftTeams.pages.config.setConfig({
        suggestedDisplayName: "EOC",
        entityId: "Test",
        contentUrl: baseUrl + "/index.html#/tab",
        websiteUrl: baseUrl + "/index.html#/tab",
      });
      saveEvent.notifySuccess();
    });

    /**
     * After verifying that the settings for your tab are correctly
     * filled in by the user you need to set the state of the dialog
     * to be valid.  This will enable the save button in the configuration
     * dialog.
     */
    microsoftTeams.pages.config.setValidityState(true);
    
  }).catch((error) => {
    console.error("TEOC_TabConfig_Error_Initializing Microsoft Teams SDK:", error);

  });

    return (
      <div>
        <h1> Microsoft Teams Emergency Operations Center</h1>
        <div>       
        App Template to help facilitate the creation of teams and assets for incident response for designated scenarios. In addition to quick team creation and asset deployment, TEOC also delivers a central dashboard to see and manage incidents from and take further action. Helping you to respond and act quicker powered by the solutions you already have.
        </div>
      </div>
    );
  }
}

export default TabConfig;
