import { Icon } from '@fluentui/react';
import { Button, CloseIcon, Dialog } from "@fluentui/react-northstar";
import { Toggle } from '@fluentui/react/lib/Toggle';
import { TooltipHost } from '@fluentui/react/lib/Tooltip';
import { ApplicationInsights } from '@microsoft/applicationinsights-web';
import { Client } from "@microsoft/microsoft-graph-client";
import * as microsoftTeams from "@microsoft/teams-js";
import React from 'react';
import CommonService, { IListItem } from "../common/CommonService";
import * as constants from '../common/Constants';
import * as graphConfig from '../common/graphConfig';
import siteConfig from '../config/siteConfig.json';

//Global Variables
let teamWebURL: string = "";
let incidentId: string = "";
let graphEndpointList: string = "";

export interface BridgeProps {
    incidentData: IListItem;
    appInsights: ApplicationInsights;
    userPrincipalName: any;
    graph: Client;
    siteId: string;
    localeStrings: any;
    onShowIncidentHistory: Function;
    currentUserId: string;
    updateIncidentData: Function;
    onEditButtonClick: Function;
    isOwner: boolean;
    updateMessagebar: Function;
}

export interface BridgeState {
    bridgeID: any;
    bridgeLink: any;
    newsTabLink: string;
    showConfirmDialog: boolean;
    toggleStatus: boolean;
    confirmMessage: string;
}

export default class Bridge extends React.Component<BridgeProps, BridgeState> {
    constructor(props: BridgeProps) {
        super(props);

        //States
        this.state = {
            bridgeID: this.props.incidentData.bridgeID,
            bridgeLink: this.props.incidentData.bridgeLink,
            newsTabLink: "",
            showConfirmDialog: false,
            toggleStatus: false,
            confirmMessage: ""
        }

        //Bind Methods
        this.activateBridge = this.activateBridge.bind(this);
        this.onToggleChange = this.onToggleChange.bind(this);
        this.createNewsTabLink = this.createNewsTabLink.bind(this);
    }

    //Create object for Common Services class
    private commonService = new CommonService();

    //Component Life cycle method
    //If News tab link is not available in Incident Transaction list generate the tab link 
    //and update the list
    public async componentDidMount() {
        incidentId = this.props.incidentData.incidentId ? this.props.incidentData.incidentId.toString() : "";
        graphEndpointList = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}/lists/${siteConfig.incidentsList}/items/${incidentId}/fields`;

        if (this.props.incidentData) {

            //Get group ID from the TeamsWebURL
            teamWebURL = this.props.incidentData.teamWebURL ? this.props.incidentData.teamWebURL : '';
            const teamGroupId = teamWebURL.split("?")[1].split("&")[0].split("=")[1].trim();


            //If News tab link is not available in Incident Transaction list generate the tab link 
            //and update the list
            if (this.props.incidentData.newsTabLink === undefined || this.props.incidentData.newsTabLink === "")
                await this.createNewsTabLink(teamGroupId);
            else
                this.setState({ newsTabLink: this.props.incidentData.newsTabLink });
        }
    }

    //Create the link for News tab if not existing already and update the link in Incident Transaction list
    private async createNewsTabLink(teamGroupId: string) {
        try {
            //Get the Announcements channel ID
            const announcementChannelID = await this.commonService.getChannelId(this.props.graph,
                teamGroupId, constants.Announcements);

            //Get the News tab URL
            const newsTabURL = await this.commonService.getTabURL(this.props.graph, teamGroupId,
                announcementChannelID, constants.News);
            console.log(constants.infoLogPrefix + "Created news tab link for the incident");

            //Update the News tab link in Incident Transaction list
            if (newsTabURL !== null) {
                this.setState({ newsTabLink: newsTabURL });

                const updateItemObj = {
                    NewsTabLink: newsTabURL
                };
                await this.commonService.updateItemInList(graphEndpointList,
                    this.props.graph, updateItemObj);

                //log trace
                console.log(constants.infoLogPrefix + "Updated news tab link for the incident");
            }
        }
        catch (error: any) {
            console.error(
                constants.errorLogPrefix + "ActiveBridge_Bridge_createNewsTabLink \n",
                JSON.stringify(error)
            );

            //log exception to App Insights
            this.commonService.trackException(this.props.appInsights, error,
                constants.componentNames.BridgeComponent, 'ActiveBridge__Bridge_createNewsTabLink',
                this.props.userPrincipalName);
        }
    }

    //Set states based on toggle and open the confirm popup
    private async onToggleChange(_ev: React.MouseEvent<HTMLElement>, checked?: boolean) {
        let toggleStatus = checked ? true : false;
        let confirmMessage = checked ? this.props.localeStrings.activateBridgeMessage : this.props.localeStrings.deActivateBridgeMessage;
        this.setState({ showConfirmDialog: true, toggleStatus: toggleStatus, confirmMessage: confirmMessage })
    }

    //Based on toggle selection activate/deactivate bridge and update the Incident Transaction list
    private async activateBridge(checked: boolean) {
        try {
            //Hide the confirm pop up
            this.setState({ showConfirmDialog: false });
            this.props.updateMessagebar(-1, "", true, false);

            //If toggle turned on create a meeting and update the the Bridge details in Incident Transaction list
            if (checked) {
                //Create online meeting for the incident
                const meetingObj = {
                    "subject": constants.teamEOCPrefix + ": " + this.props.incidentData.incidentId + " - " + this.props.incidentData.incidentName
                };
                const meetingResult = await this.commonService.sendGraphPostRequest(graphConfig.onlineMeetingGraphEndpoint, this.props.graph, meetingObj);

                //log trace
                console.log(constants.infoLogPrefix + "Created bridge for Incident");

                //Update the Bridge details in Incident Transaction list
                if (meetingResult !== null) {
                    //Update BridgeID and BridgeLink                    
                    await this.updateBridgeDetails(meetingResult.id, meetingResult.joinUrl);
                    const incidentData = this.props.incidentData;
                    incidentData.bridgeID = meetingResult.id;
                    incidentData.bridgeLink = meetingResult.joinUrl;
                    this.props.updateIncidentData(incidentData);
                    this.setState({ bridgeID: meetingResult.id, bridgeLink: meetingResult.joinUrl });
                    this.props.updateMessagebar(4, this.props.localeStrings.bridgeActivationMsg, false, false);
                }
            }
            //If toggle turned off delete the online meeting and remove the Bridge details in Incident Transaction list
            else {
                await this.commonService.sendGraphDeleteRequest(graphConfig.onlineMeetingGraphEndpoint + "/" + this.state.bridgeID, this.props.graph);

                //log trace
                console.log(constants.infoLogPrefix + "Deleted bridge for Incident");

                await this.updateBridgeDetails("", "");
                const incidentData = this.props.incidentData;
                incidentData.bridgeID = "";
                incidentData.bridgeLink = "";
                this.props.updateIncidentData(incidentData);
                this.setState({ bridgeID: "", bridgeLink: "" });
                this.props.updateMessagebar(4, this.props.localeStrings.bridgeDeactivationMsg, false, false);
            }
        }
        catch (error: any) {
            console.error(
                constants.errorLogPrefix + "ActiveBridge_Bridge_onToggleSetting \n",
                JSON.stringify(error)
            );

            this.props.updateMessagebar(1, this.props.localeStrings.genericErrorMessage + " "
                + this.props.localeStrings.errMsgForBridgeActivation, false, false);

            //log exception to AppInsights
            this.commonService.trackException(this.props.appInsights, error,
                constants.componentNames.BridgeComponent, 'activateBridge', this.props.userPrincipalName);

        }
    }

    //Update bridge details in the Incident Transaction list
    private async updateBridgeDetails(bridgeID: string, bridgeLink: string) {
        try {
            const updateItemObj = {
                BridgeID: bridgeID,
                BridgeLink: bridgeLink
            }
            await this.commonService.updateItemInList(graphEndpointList, this.props.graph, updateItemObj);

            //log trace
            console.log(constants.infoLogPrefix + "Updated bridge details for the incident");
        }
        catch (error: any) {
            console.error(
                constants.errorLogPrefix + "ActiveBridge_Bridge_updateBridgeDetails \n",
                JSON.stringify(error)
            );
            //log exception to AppInsights
            this.commonService.trackException(this.props.appInsights, error,
                constants.componentNames.BridgeComponent, 'updateBridgeDetails', this.props.userPrincipalName);
        }
    }

    render() {
        return (
            <div className="bridge-wrapper">
                <div className='bridge-container'>
                    <div className='bridge-links'>
                        <div className="preview-img-wrapper">
                            <img src={require("../assets/Images/PreviewIcon.svg").default}
                                alt={this.props.localeStrings.previewLabel}
                                title={this.props.localeStrings.previewLabel}
                            />
                        </div>
                        <div className="links-wrapper">
                            <div className="team-group-name">
                                {this.props.incidentData.incidentName}
                            </div>
                            <div className="links">
                                {this.props.isOwner &&
                                    <Button
                                        icon={<img
                                            src={require("../assets/Images/EditIncidentBoldIcon.svg").default}
                                            alt={this.props.localeStrings.edit}
                                            className="bridge-edit-icon"
                                            title={this.props.localeStrings.edit}
                                        />}
                                        text
                                        content={this.props.localeStrings.edit}
                                        title={this.props.localeStrings.edit}
                                        onClick={() => this.props.onEditButtonClick(this.props.incidentData)}
                                        className="bridge-edit-link"
                                    />
                                }
                                <Button
                                    icon={<img
                                        src={require("../assets/Images/IncidentHistoryBoldIcon.svg").default}
                                        alt={this.props.localeStrings.viewIncidentHistory}
                                        title={this.props.localeStrings.viewIncidentHistory}
                                        className="bridge-history-icon"
                                    />}
                                    text
                                    content={this.props.localeStrings.incidentHistory}
                                    title={this.props.localeStrings.viewIncidentHistory}
                                    onClick={() => this.props.onShowIncidentHistory(this.props.incidentData.incidentId)}
                                    className="bridge-history-link"
                                />
                                <Button
                                    icon={<img
                                        src={require("../assets/Images/TeamChatIcon.svg").default}
                                        alt={this.props.localeStrings.teamChatLabel}
                                        title={this.props.localeStrings.teamChatLabel}
                                        className="bridge-chat-icon"
                                    />}
                                    text
                                    title={this.props.localeStrings.teamChatLabel}
                                    content={this.props.localeStrings.teamChatLabel}
                                    onClick={() => microsoftTeams.executeDeepLink(teamWebURL)}
                                    className="bridge-chat-link"
                                />
                                {this.state.newsTabLink !== "" &&
                                    <Button
                                        icon={<img
                                            src={require("../assets/Images/NewsIcon.svg").default}
                                            alt={this.props.localeStrings.newsLabel}
                                            title={this.props.localeStrings.newsLabel}
                                            className="bridge-news-icon"
                                        />}
                                        text
                                        content={this.props.localeStrings.newsLabel}
                                        title={this.props.localeStrings.newsLabel}
                                        onClick={() => microsoftTeams.executeDeepLink(this.state.newsTabLink)}
                                        className="bridge-news-link"
                                    />
                                }
                            </div>
                        </div>
                    </div>
                    <div className='bridge-buttons'>
                        {this.props.isOwner &&
                            <Toggle
                                checked={this.state.bridgeID === undefined || this.state.bridgeID === "" ? false : true}
                                label={
                                    <div className="toggle-btn-label">
                                        {this.props.localeStrings.activateBridgeLabel}
                                        <span className="bridge-toggle-Info-Icon">
                                            <TooltipHost
                                                content={this.props.localeStrings.bridgeToggleBtnInfoText}
                                                calloutProps={{ gapSpace: 0 }}
                                            >
                                                <Icon iconName="Info" />
                                            </TooltipHost>
                                        </span>
                                    </div>
                                }
                                inlineLabel
                                onChange={this.onToggleChange}
                                className={`bridge-toggle-btn${this.state.bridgeID === undefined || this.state.bridgeID === "" ?
                                    " bridge-toggle-disabled-btn" : ""}`}
                            />
                        }
                        <Button
                            title={this.props.localeStrings.joinBridgeButtonLabel}
                            onClick={() => microsoftTeams.executeDeepLink(this.state.bridgeLink)}
                            disabled={this.state.bridgeID === undefined || this.state.bridgeID === "" ? true : false}
                            content={this.props.localeStrings.joinBridgeButtonLabel}
                            className="join-bridge-btn"
                            primary
                        />
                    </div>
                </div>
                <Dialog
                    cancelButton={{
                        content: this.props.localeStrings.noButton,
                        title: this.props.localeStrings.noButton
                    }}
                    confirmButton={{
                        content: this.props.localeStrings.yesButton,
                        title: this.props.localeStrings.yesButton
                    }}
                    onCancel={() => this.setState({ showConfirmDialog: false })}
                    onConfirm={() => this.activateBridge(this.state.toggleStatus)}
                    open={this.state.showConfirmDialog}
                    header={{ content: this.props.localeStrings.confirmPopupTitle }}
                    headerAction={{
                        icon: <CloseIcon />, title: this.props.localeStrings.btnClose,
                        onClick: () => this.setState({ showConfirmDialog: false })
                    }}
                    content={{ content: this.state.confirmMessage }}
                    closeOnOutsideClick={false}
                    className="bridge-confirm-popup"
                />
            </div>
        );
    }
}
