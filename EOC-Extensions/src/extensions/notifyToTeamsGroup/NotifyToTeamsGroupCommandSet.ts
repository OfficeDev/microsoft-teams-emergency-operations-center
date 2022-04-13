import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface INotifyToTeamsGroupCommandSetProperties {
  NotifyToTeamsGroup: string;
}

const LOG_SOURCE: string = 'NotifyToTeamsGroupCommandSet';

export default class NotifyToTeamsGroupCommandSet extends BaseListViewCommandSet<INotifyToTeamsGroupCommandSetProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized NotifyToTeamsGroupCommandSet');
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('NotifyToTeamsGroup');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      let currentSiteTitle = this.context.pageContext.web.title.toLowerCase();
      compareOneCommand.visible = (event.selectedRows.length === 1 && currentSiteTitle.indexOf("teoc") >= 0);
    }
  }

  @override
  public async onExecute(event: IListViewCommandSetExecuteEventParameters): Promise<any> {
    //get associated teams group id
    let currentSiteTeamsGroupID = this.context.pageContext.site.group.id._guid;
    try {
      switch (event.itemId) {
        case 'NotifyToTeamsGroup':
          //get channel context to post message           
          let getChannelRes = await this.getChannel(currentSiteTeamsGroupID);
          if (getChannelRes != "Failed") {
            let currentItemProperties = {
              createdDate: event.selectedRows[0].getValueByName("Created"),
              createdBy: event.selectedRows[0].getValueByName("Editor"),
              linkToPage: event.selectedRows[0].getValueByName("FileRef"),
              fileDisplayName: event.selectedRows[0].getValueByName("FileLeafRef")
            };
            let sendMessageRes = await this.sendMessage(currentSiteTeamsGroupID, getChannelRes, currentItemProperties);
            if (sendMessageRes != "Failed") {
              alert("Notification successfully sent");
            } else {
              alert("Failed to notify teams group, please verify if the associated group exists.");
            }
          }
          break;
        default:
          throw new Error('Unknown command');
      }
    }
    catch (error) {
      console.error('EOC App: NotifyToTeamsGroupCommandSet_onExecute_Failed to post message to teams.', error);
    }
  }

  //method to send message to teams group
  public async sendMessage(teamId, channelId, currentItemProperties: any): Promise<any> {
    return new Promise<any>(async (resolve, reject) => {
      try {
        const siteAbsUrl = this.context.pageContext.site.absoluteUrl;
        const siteRelUrl = this.context.pageContext.site.serverRelativeUrl;
        const filePath = `${siteAbsUrl}` + `${currentItemProperties.linkToPage.substr(siteRelUrl.length)}`;
        //construct adaptive card
        var content = {
          "body": {
            "contentType": "html",
            "content": "<attachment id=\"74d20c7f34aa4a7fb74e2b30004247c5\"></attachment>"
          },
          "attachments": [
            {
              "id": "74d20c7f34aa4a7fb74e2b30004247c5",
              "contentType": "application/vnd.microsoft.card.adaptive",
              "contentUrl": null,
              "content": `{
                  "type": "AdaptiveCard",
                  "body": [
                    {
                      "type": "TextBlock",
                      "size": "Large",
                      "weight": "Bolder",
                      "text": "New Post Published"
                  },
                      {
                          "type": "TextBlock",
                          "size": "Medium",
                          "weight": "Bolder",
                          "text": "${currentItemProperties.fileDisplayName.split(".")[0]}"
                      },
                      {
                          "type": "ColumnSet",
                          "columns": [
                              {
                                  "type": "Column",
                                  "items": [
                                      {
                                          "type": "TextBlock",
                                          "text": "Created By -  ${currentItemProperties.createdBy[0].title}",
                                          "wrap": true
                                      },
                                      {
                                          "type": "TextBlock",
                                          "spacing": "None",
                                          "text": "Created Date - ${currentItemProperties.createdDate}",
                                          "isSubtle": true,
                                          "wrap": true
                                      }
                                  ],
                                  "width": "stretch"
                              }
                          ]
                      }
                  ],
                  "actions": [
                      {
                          "type": "Action.OpenUrl",
                          "title": "View",
                          "url": "${filePath}"
                      }
                  ],
                  "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                  "version": "1.3"
              }`,
              "name": null,
              "thumbnailUrl": null
            }
          ]
        };
        //post the message
        let postMessageResp = await this.context.msGraphClientFactory.getClient();
        postMessageResp.api('/teams/' + teamId + '/channels/' + channelId + "/messages/")
          .post(content)
          .then((res) => { resolve(res); })
          .catch(
            (err) => {
              console.log(err);
              reject("Failed");
            }
          );
      }
      catch (error) {
        console.error('TeamEOC_NotifyToTeamsGroupCommandSet_sendMessage_Failed to send message', error);
        reject("Failed");
      }
    });
  }

  // method to get the channel where message will be sent
  public async getChannel(currentSiteTeamID: string): Promise<any> {
    return new Promise<any>(async (resolve, reject) => {
      try {
        let currentSiteTeamChannel: string;
        //get Announcements channel id related to current site teams group
        let channels = await this.context.msGraphClientFactory.getClient();
        channels.api('/teams/' + currentSiteTeamID + '/channels')
          .filter('startsWith(displayName,\'Announcements\')').select('id')
          .get((err, res) => {
            // get the first channel from the array
            currentSiteTeamChannel = res.value[0].id;
            resolve(currentSiteTeamChannel);
          });
      } catch (error) {
        console.error('EOC App: NotifyToTeamsGroupCommandSet_getChannel_Unable to get channel context', error);
        reject("Failed");
      }
    });
  }
}
