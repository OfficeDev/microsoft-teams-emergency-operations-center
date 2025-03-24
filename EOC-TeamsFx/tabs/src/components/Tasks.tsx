import { ApplicationInsights } from '@microsoft/applicationinsights-web';
import { Tasks } from '@microsoft/mgt-react';
import { Client } from "@microsoft/microsoft-graph-client";
import React from 'react';
import CommonService, { IListItem } from "../common/CommonService";
import * as constants from '../common/Constants';
import * as graphConfig from '../common/graphConfig';
import siteConfig from '../config/siteConfig.json';

export interface PlannerTasksProps {
    incidentData: IListItem;
    graph: Client;
    siteId: string;
    appInsights: ApplicationInsights;
    userPrincipalName: any;
    updateMessagebar: Function;
    showTasksLoader: boolean;
    localeStrings: any;
    graphContextURL: string;
    tenantID: any;
}

export interface PlannerTasksState {
    planID: string;
}


export default class PlannerTasks extends React.Component<PlannerTasksProps, PlannerTasksState> {
    constructor(props: PlannerTasksProps) {
        super(props);
        this.state = {
            planID: ""
        }
        this.componentDidMount = this.componentDidMount.bind(this);
    }
    //Create object for Common Services class
    private commonService = new CommonService();

    //Component Life cycle method
    //Retrieve Plan details for the incident and create if not existing
    public async componentDidMount() {
        await this.createPlannerPlan();
    }

    //Create Planner Plan for old incidents which does not have a plan created already and update the IncidentTransaction list
    private createPlannerPlan = async () => {
        try {
            //Check if there is plan id for the current incident. If not create a plan for the Incident Group ID    
            if (this.props.incidentData.planID === null || this.props.incidentData.planID === "" ||
                this.props.incidentData.planID === undefined) {
                if (this.props.incidentData && this.props.incidentData.incidentId) {

                    this.props.updateMessagebar(-1, "", false, true);

                    //Get group ID from the TeamsWebURL
                    const teamWebURL = this.props.incidentData.teamWebURL ? this.props.incidentData.teamWebURL : '';
                    const teamGroupId = teamWebURL.split("?")[1].split("&")[0].split("=")[1].trim();

                    let incidentId = this.props.incidentData.incidentId ? this.props.incidentData.incidentId.toString() : "";

                    //For old incidents add Incident commander and Created By user as member to the group so the plan can be created
                    try {
                        let userIds: any = [];
                        let userId = this.props.incidentData.incidentCommanderObj ? this.props.incidentData.incidentCommanderObj.split('|')[1] : "";
                        userIds.push(userId);
                        let createdById = this.props.incidentData.createdById;
                        userIds.push(createdById);
                        await this.addUserAsGroupMember(userIds, teamGroupId);
                    }
                    catch (ex) {
                        console.error(
                            constants.errorLogPrefix + "ActiveBridgeTasks_addUserAsGroupMember \n",
                            JSON.stringify(ex)
                        );
                    }
                    let plannerPlanId = "";
                    //Set time out for membership to take effect before creating the plan
                    await this.timeout(5000);

                    let maxPlanCreationAttempt = 5, isPlanCreated = false;
                    while (isPlanCreated === false && maxPlanCreationAttempt > 0) {
                        try {

                            const result = await this.commonService.createPlannerPlan(teamGroupId, incidentId, this.props.graph,
                                this.props.graphContextURL, this.props.tenantID, "", true);
                            
                            plannerPlanId = result?.planId
                                //Set state variable for Plan ID
                            this.setState({
                                planID: plannerPlanId
                            });

                            // update the result object
                            if (plannerPlanId) {
                                console.log(constants.infoLogPrefix + "Planner Plan created on - " + new Date());
                                isPlanCreated = true;
                            }
                            //Update the PlanID in Incident Transaction list
                            if (plannerPlanId !== null && plannerPlanId !== undefined) {
                                const updateItemObj = {
                                    PlanID: plannerPlanId
                                }

                                let graphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}/lists/${siteConfig.incidentsList}/items/${this.props.incidentData.incidentId}/fields`;
                                await this.commonService.updateItemInList(graphEndpoint, this.props.graph, updateItemObj);

                                //log trace
                                console.log(constants.infoLogPrefix + "Updated Plan Id for the incident");

                                this.props.updateMessagebar(-1, "", false, false);
                            }

                        } catch (updationError: any) {
                            console.log(constants.infoLogPrefix + "Plan creation failed on - " + new Date());
                            console.error(
                                constants.errorLogPrefix + "Tasks_createPlannerPlan \n",
                                JSON.stringify(updationError)
                            );
                        }
                        maxPlanCreationAttempt--;
                        await this.timeout(10000);
                    }
                    console.log(constants.infoLogPrefix + "createPlan_No Of Attempt", (5 - maxPlanCreationAttempt));
                }
            }
            //If the Plan ID already exists set the state
            else {
                this.setState({
                    planID: this.props.incidentData.planID
                });
            }
        }
        catch (ex) {
            console.error(
                constants.errorLogPrefix + "ActiveBridgeTasks_createPlannerPlan \n",
                JSON.stringify(ex)
            );

            this.props.updateMessagebar(-1, "", false, false);

            // Log Exception
            this.commonService.trackException(this.props.appInsights, ex, constants.componentNames.TasksComponent,
                'ActiveBridgeTasks_createPlannerPlan', this.props.userPrincipalName);
        }

    }
    // method to delay the operation by adding timeout
    private timeout = (delay: number): Promise<any> => {
        return new Promise(res => setTimeout(res, delay));
    }

    //For old incidents add Incident Commander and Created By as member to the group so the plan can be created
    private async addUserAsGroupMember(userIds: any, groupID: string): Promise<any> {
        return new Promise(async (resolve, reject) => {
            try {
                //graph endpoint to get members of the group
                let graphEndpointToGetMembers = graphConfig.teamGroupsGraphEndpoint + "/" + groupID + graphConfig.membersGraphEndpoint;

                const response = await this.commonService.getGraphData(graphEndpointToGetMembers, this.props.graph);

                const existingMembers: any = [];
                if (response.value.length > 0) {
                    response.value.forEach((user: any) => {
                        existingMembers.push(user.id);
                    });
                }

                const usersToAdd: any = [];
                const usersToAddAsOwnersToTeam: any = [];
                const uniqueUserArray: any = [];
                //create member object to add it to group
                userIds.forEach((userId: any) => {
                    if (uniqueUserArray.indexOf(userId) === -1 && existingMembers.indexOf(userId) === -1) {
                        {
                            uniqueUserArray.push(userId);
                            usersToAdd.push(this.props.graphContextURL + graphConfig.usersGraphEndpoint + "('" + userId + "')");

                            //fix for GCC tenant - adding the owners again to team as it is getting removed when the group membership is updated
                            usersToAddAsOwnersToTeam.push({
                                "@odata.type": "microsoft.graph.aadUserConversationMember",
                                "roles": ["owner"],
                                "user@odata.bind": this.props.graphContextURL + graphConfig.usersGraphEndpoint + "('" + userId + "')"
                            });
                        }
                    }
                });
                //call the patch request to add members
                if (usersToAdd.length > 0) {
                    //adding the incident creator and incident commander to the group as members
                    const membersObj = {
                        "members@odata.bind": usersToAdd
                    }
                    let graphEndpoint = graphConfig.teamGroupsGraphEndpoint + "/" + groupID;
                    await this.commonService.sendGraphPatchRequest(graphEndpoint, this.props.graph, membersObj);

                    //fix for GCC tenant - adding the incident creator and incident commander to the team as owners
                    await this.timeout(10000);

                    const ownersObj = {
                        "values": usersToAddAsOwnersToTeam
                    }
                    let addUserToTeamEndPoint = graphConfig.teamsGraphEndpoint + "/" + groupID + graphConfig.addMembersGraphEndpoint;
                    await this.commonService.sendGraphPostRequest(addUserToTeamEndPoint, this.props.graph, ownersObj);
                }
                resolve(true);
            } catch (ex) {
                console.error(
                    constants.errorLogPrefix + "ActiveBridgeTasks_addUserAsGroupMember \n",
                    JSON.stringify(ex)
                );
                reject(ex);

                // Log Exception
                this.commonService.trackException(this.props.appInsights, ex,
                    constants.componentNames.TasksComponent, 'ActiveBridgeTasks_addUserAsGroupMember',
                    this.props.userPrincipalName);
            }
        });
    }

    render() {
        return (
            <div className='tasks-wrapper'>
                {this.state.planID &&
                    <Tasks
                        initialId={this.state.planID}
                        targetId={this.state.planID}
                        data-source="TasksSource.planner"
                        className='active-dashboard-task'
                    />
                }
                {(!this.state.planID && !this.props.showTasksLoader) &&
                    <div className='no-tasks-created-msg'>{this.props.localeStrings.noTasksCreatedMessage}</div>
                }
            </div>
        );
    }
}
