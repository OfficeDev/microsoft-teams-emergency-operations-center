import { Icon, SelectionMode, TooltipHost } from '@fluentui/react';
import { Button, CloseIcon, Dialog, InfoIcon, TeamCreateIcon } from '@fluentui/react-northstar';
import { GroupedList, IGroup, IGroupHeaderProps } from '@fluentui/react/lib/GroupedList';
import { ApplicationInsights } from '@microsoft/applicationinsights-web';
import { Person } from '@microsoft/mgt-react';
import { Client } from "@microsoft/microsoft-graph-client";
import React from 'react';
import CommonService, { IListItem } from '../common/CommonService';
import * as constants from "../common/Constants";
import * as graphConfig from '../common/graphConfig';
import Communications from './Communications';

export interface MembersProps {
    incidentData: IListItem;
    graph: Client;
    appInsights: ApplicationInsights;
    userPrincipalName: any;
    localeStrings: any;
    isOwner: boolean;
    onEditButtonClick: Function;
}

export interface MembersState {
    roles: Array<IGroup>;
    members: Array<any>;
    showCommunicationsPopup: boolean;
    showLoader: boolean;
    cardMessage: string;
    highImportance: boolean;
    includeLink: boolean;
    validationMessage: string;
    messageType: number;
    statusMessage: string;
}

export default class Members extends React.Component<MembersProps, MembersState> {

    private commonService = new CommonService();

    constructor(props: MembersProps) {
        super(props);

        //States
        this.state = {
            roles: [],
            members: [],
            showCommunicationsPopup: false,
            showLoader: false,
            cardMessage: "",
            highImportance: false,
            includeLink: false,
            validationMessage: "",
            messageType: -1,
            statusMessage: ""
        }

        //Bind Methods
        this.onRenderCell = this.onRenderCell.bind(this);
        this.onRenderHeader = this.onRenderHeader.bind(this);
        this.setState = this.setState.bind(this);
        this.postAnnouncement = this.postAnnouncement.bind(this);
        this.sendMessage = this.sendMessage.bind(this);
        this.resetStates = this.resetStates.bind(this);
    }

    // Component Life Cycle Method
    //Get Assigned roles of the Incident
    public componentDidMount() {
        this.getAssignedRoles();
    }

    // Method to Get Assigned roles of an Incident
    private getAssignedRoles = () => {
        try {
            if (this.props.incidentData && this.props.incidentData.incidentId) {
                const roles: any = [];
                const members: any = [];
                const roleLeads: any[] = [];

                //Separate Role leads into an Array
                const leads = this.props.incidentData.roleLeads ? this.props.incidentData.roleLeads : '';
                if (leads.length > 0 && leads.split(";").length > 1) {
                    leads.split(";").forEach((role) => {
                        if (role.length > 0) {
                            const leadDetailsObj: any[] = [];
                            role.split(":")[1].trim().split(",").forEach(user => {
                                leadDetailsObj.push({
                                    userName: user.split("|")[0].trim(),
                                    userEmail: user.split("|")[2].trim(),
                                    userId: user.split("|")[1].trim(),
                                    role: role.split(":")[0].trim(),
                                    lead: true
                                });
                            });
                            roleLeads.push(...leadDetailsObj);
                        }
                    });
                }

                //Push Incident Commander
                const incidentCommanderObj = this.props.incidentData.incidentCommanderObj ? this.props.incidentData.incidentCommanderObj.split("|") : [];
                members.push({
                    userName: incidentCommanderObj[0].trim(),
                    userEmail: incidentCommanderObj[2].trim().slice(0, incidentCommanderObj[2].trim().length - 1),
                    userId: incidentCommanderObj[1].trim(),
                    lead: false
                });
                roles.push({
                    key: "Incident Commander0",
                    name: "Incident Commander",
                    startIndex: 0,
                    count: 1
                });

                //Push Roles and Assigned Members
                const roleAssignments = this.props.incidentData.roleAssignments ? this.props.incidentData.roleAssignments : '';
                if (roleAssignments.length > 0 && roleAssignments.split(";").length > 1) {
                    roleAssignments.split(";").forEach((role, idx) => {
                        if (role.length > 0) {
                            let roleMembers: any[] = [];
                            const roleName = role.split(":")[0].trim();

                            role.split(":")[1].trim().split(",").forEach(user => {
                                roleMembers.push({
                                    userName: user.split("|")[0].trim(),
                                    userEmail: user.split("|")[2].trim(),
                                    userId: user.split("|")[1].trim(),
                                    role: roleName,
                                    lead: false
                                });
                            });

                            const leadData = roleLeads.find((roleData) => roleName === roleData.role);

                            if (leadData) {
                                roleMembers = roleMembers.filter((user: any) => user.userId !== leadData.userId);
                                roleMembers.unshift(leadData);
                            }

                            roles.push({
                                key: roleName + (idx + 1),
                                name: roleName,
                                startIndex: members.length,
                                count: roleMembers.length
                            });

                            members.push(...roleMembers);
                        }
                    });
                }

                this.setState({
                    roles: roles,
                    members: members
                });
            }
        }
        catch (error) {
            console.error(
                constants.errorLogPrefix + "_Members_getAssignedRoles \n",
                JSON.stringify(error)
            );

            //log exception to AppInsights
            this.commonService.trackException(this.props.appInsights, error, constants.componentNames.MembersComponent, 'getAssignedRoles', this.props.userPrincipalName);
        }
    }

    //Method to reset states
    private resetStates = () => {
        this.setState({
            cardMessage: "",
            includeLink: false,
            highImportance: false,
            validationMessage: "",
            showCommunicationsPopup: false,
            messageType: -1,
            statusMessage: ""
        });
    }

    //Method to be called onclick of post announcement
    private postAnnouncement = async () => {
        try {
            if (this.state.cardMessage.trim() !== "") {
                this.setState({ showLoader: true, messageType: -1, statusMessage: "" });

                const teamGroupId: string | undefined = this.props.incidentData.teamWebURL?.split("?")[1]
                    .split("&")[0].split("=")[1].trim();

                const response = await this.commonService.getGraphData(
                    graphConfig.teamsGraphEndpoint + "/" + teamGroupId, this.props.graph);
                const teamDisplayName = response.displayName;

                const announcementsChannelId = await this.commonService.getChannelId(
                    this.props.graph,
                    teamGroupId,
                    constants.Announcements);

                await this.sendMessage(teamDisplayName, teamGroupId, announcementsChannelId);

                this.setState({
                    cardMessage: "",
                    highImportance: false,
                    includeLink: false,
                    showLoader: false,
                    messageType: 4,
                    statusMessage: this.props.localeStrings.successMessageForPostAnnouncement
                });
            }
            else {
                this.setState({ validationMessage: this.props.localeStrings.announcementMessageValidationText });
            }
        }
        catch (error) {
            console.error(
                constants.errorLogPrefix + "Members_Communications_PostAnnouncement \n",
                JSON.stringify(error)
            );

            this.setState({
                messageType: 1,
                statusMessage: this.props.localeStrings.genericErrorMessage + " " + this.props.localeStrings.errMsgForPostAnnouncement,
                showLoader: false
            });

            //log exception in app insights
            this.commonService.trackException(this.props.appInsights, error, constants.componentNames.MembersComponent, 'Members__Communications_PostAnnouncement', this.props.userPrincipalName);
        }
    }

    //Method to create and send adaptive card to Annoucement channel
    private async sendMessage(teamDisplayName: string, teamGroupId: any, channelId: string) {
        try {
            let cardBody = {
                "importance": this.state.highImportance ? "high" : "normal",
                "body": {
                    "contentType": "html",
                    "content": "<at id='0'>" + teamDisplayName + "</at><attachment id='4465B062-EE1C-4E0F-B944-3B7AF61EAF40'></attachment>",
                },
                "attachments": [
                    {
                        "id": "4465B062-EE1C-4E0F-B944-3B7AF61EAF40",
                        "contentType": "application/vnd.microsoft.card.adaptive",
                        "content": JSON.stringify({
                            "type": "AdaptiveCard",
                            "body": [
                                {
                                    "type": "TextBlock",
                                    "size": "Large",
                                    "weight": "Bolder",
                                    "text": "Teams Emergency Operations Center",
                                    "style": "heading",
                                    "color": "Accent",
                                },
                                {
                                    "type": "TextBlock",
                                    "text": this.state.cardMessage,
                                    "wrap": "true"
                                },
                                {
                                    "type": "ActionSet",
                                    "isVisible": this.state.includeLink,
                                    "actions": [
                                        {
                                            "type": "Action.OpenUrl",
                                            "title": "Join Bridge",
                                            "url": this.props.incidentData.bridgeLink
                                        }
                                    ],
                                }
                            ],
                            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                            "version": "1.5"
                        })
                    },
                ],
                "mentions": [
                    {
                        "id": 0,
                        "mentionText": teamDisplayName,
                        "mentioned": {
                            "conversation": {
                                "id": teamGroupId,
                                "displayName": teamDisplayName,
                                "conversationIdentityType": "team"
                            }
                        }
                    }
                ],
            };

            const endpoint = graphConfig.teamsGraphEndpoint + "/" + teamGroupId +
                graphConfig.channelsGraphEndpoint + "/" + channelId + graphConfig.messagesGraphEndpoint;
            await this.commonService.sendGraphPostRequest(endpoint, this.props.graph, cardBody);
        }
        catch (error) {
            console.error(
                constants.errorLogPrefix + "Members_Communications_sendMessage \n",
                JSON.stringify(error)
            );
            throw error;
        }
    }

    //Method to customize each member
    private onRenderCell = (_nestingDepth?: number, item?: any) => {
        return (
            <div className={`member-profile${item.lead ? " profile-border-left" : ""}`} key={item.userId}>
                <Person
                    userId={item.userId}
                    showPresence={true}
                    view={3}
                    personCardInteraction={1}
                    avatarSize='large'
                />
                {item.lead && <span className='lead-label'>{this.props.localeStrings.leadLabel}</span>}
            </div>
        );
    }

    //Method to customize each role
    private onRenderHeader = (props?: IGroupHeaderProps): JSX.Element | null => {
        if (props) {
            const toggleCollapse = (): void => {
                props.onToggleCollapse!(props.group!);
            };
            return (
                <div className='members-grp-header' onClick={toggleCollapse}>
                    <Icon
                        iconName={props.group!.isCollapsed ? "FlickLeft" : "FlickUp"}
                        className={props.group!.isCollapsed ? "flick-left-icon" : "flick-up-icon"}
                    />
                    <div className='grp-name'>{props.group!.name}</div>
                </div>
            );
        }
        return null;
    }

    //Render Method
    render() {
        return (
            <div className='members'>
                <GroupedList
                    items={this.state.members}
                    onRenderCell={this.onRenderCell}
                    selectionMode={SelectionMode.none}
                    groups={this.state.roles}
                    groupProps={{ onRenderHeader: this.onRenderHeader }}
                    className="members-grouped-list"
                    rootListProps={{ "aria-busy": "true" }}
                    listProps={{ "aria-busy": "true" }}
                />
                {this.props.isOwner &&
                    <Button
                        content={this.props.localeStrings.addMembersBtnlabel}
                        title={this.props.localeStrings.addMembersBtnlabel}
                        icon={<TeamCreateIcon outline />}
                        onClick={() => this.props.onEditButtonClick(this.props.incidentData)}
                        className="add-members-btn"
                    />
                }
                <Dialog
                    cancelButton={{
                        content: this.props.localeStrings.btnClose,
                        icon: <CloseIcon bordered circular size="smallest" />,
                        title: this.props.localeStrings.btnClose,
                        disabled: this.state.showLoader
                    }}
                    confirmButton={{
                        content: this.props.localeStrings.sendButtonLabel,
                        title: this.props.localeStrings.sendButtonLabel,
                        disabled: this.state.showLoader
                    }}
                    onCancel={() => this.resetStates()}
                    onConfirm={() => this.postAnnouncement()}
                    onOpen={() => this.setState({ showCommunicationsPopup: true })}
                    open={this.state.showCommunicationsPopup}
                    header={{
                        content: <>
                            {this.props.localeStrings.announcementLabel} &nbsp;
                            <TooltipHost content={this.props.localeStrings.announcementInfoIconContent}>
                                <InfoIcon outline />
                            </TooltipHost>
                        </>
                    }}
                    headerAction={{
                        icon: <CloseIcon />, title: this.props.localeStrings.btnClose,
                        onClick: () => this.resetStates(),
                        disabled: this.state.showLoader
                    }}
                    content={{
                        content: (<Communications
                            graph={this.props.graph}
                            incidentData={this.props.incidentData}
                            setState={this.setState}
                            cardMessage={this.state.cardMessage}
                            highImportance={this.state.highImportance}
                            includeLink={this.state.includeLink}
                            validationMessage={this.state.validationMessage}
                            showLoader={this.state.showLoader}
                            messageType={this.state.messageType}
                            statusMessage={this.state.statusMessage}
                            localeStrings={this.props.localeStrings}
                        />)
                    }}
                    closeOnOutsideClick={false}
                    trigger={<Button
                        content={this.props.localeStrings.postAnnouncementLabel}
                        icon={<Icon iconName='Message' />}
                        title={this.props.localeStrings.postAnnouncementLabel}
                        className="announcement-popup-trigger-btn"
                    />}
                    className="communications-popup"
                />
            </div>
        );
    }
}
