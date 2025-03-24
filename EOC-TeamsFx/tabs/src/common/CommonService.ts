import { SeverityLevel } from "@microsoft/applicationinsights-common";
import { Client } from "@microsoft/microsoft-graph-client";
import * as graphConfig from "../common/graphConfig";
import { IIncidentStatus, IInputRegexValidationStates, ILocation } from "../components/ICreateIncident";
import * as constants from "./Constants";

export interface IListItem {
    itemId?: string;
    incidentId?: number;
    incidentName?: string;
    incidentCommander?: string;
    incidentStatusObj?: IIncidentStatus;
    status?: string;
    location?: string;
    startDate?: string;
    startDateUTC?: string;
    modifiedDate?: string;
    teamWebURL?: string;
    roleAssignments?: string;
    roleLeads?: string;
    roleAssignmentsObj?: string;
    roleLeadsObj?: string;
    incidentDescription?: string;
    incidentType?: string;
    incidentCommanderObj?: string;
    severity?: string;
    lastModifiedBy?: string;
    version?: string;
    reasonForUpdate?: string;
    planID?: string;
    bridgeID?: string;
    bridgeLink?: string;
    newsTabLink?: string;
    cloudStorageLink?: string;
    createdById?: string;
}

export interface IInputValidationStates {
    incidentNameHasError: boolean;
    incidentStatusHasError: boolean;
    incidentLocationHasError: boolean;
    incidentTypeHasError: boolean;
    incidentDescriptionHasError: boolean;
    incidentStartDateTimeHasError: boolean;
    incidentCommandarHasError: boolean;
}

export interface IRoleItem {
    itemId: string;
    role: string;
    users: any[];
    lead: any[];
}

export interface IIncidentTypeDefaultRoleItem {
    itemId?: string;
    incidentType?: string;
    roleAssignments?: string;
    roleLeads?: string;
    additionalChannels?: string;
    sharedChannel?: string;
    cloudStorageLink?: string;
}

export interface IConfigSettingItem {
    itemId?: string;
    title?: string;
    value?: string;
}

export default class CommonService {

    //#region Dashboard Methods

    // get data to show on the Dashboard
    public async getDashboardData(graphEndpoint: any, graph: Client): Promise<any> {
        try {

            const incidentsData = await graph.api(graphEndpoint).get();

            // Prepare the output array
            let formattedIncidentsData: Array<IListItem> = new Array<IListItem>();

            // Map the JSON response to the output array
            incidentsData.value.forEach((item: any) => {
                formattedIncidentsData.push({
                    itemId: item.fields.id,
                    incidentId: parseInt(item.fields.IncidentId),
                    incidentName: item.fields.IncidentName,
                    incidentCommander: this.formatIncidentCommander(item.fields.IncidentCommander),
                    incidentCommanderObj: item.fields.IncidentCommander,
                    incidentStatusObj: { status: item.fields.Status, id: item.fields.StatusLookupId },
                    location: item.fields.Location !== "null" ? JSON.stringify(this.formatGeoLocation(item)) : "",
                    startDate: this.formatDate(item.fields.StartDateTime),
                    startDateUTC: new Date(item.fields.StartDateTime).toISOString().slice(0, new Date(item.fields.StartDateTime).toISOString().length - 1),
                    modifiedDate: item.fields.Modified,
                    teamWebURL: item.fields.TeamWebURL,
                    incidentDescription: item.fields.Description,
                    incidentType: item.fields.IncidentType,
                    roleAssignments: item.fields.RoleAssignment,
                    roleLeads: item.fields.RoleLeads,
                    severity: item.fields.Severity ? item.fields.Severity : "",
                    planID: item.fields.PlanID,
                    bridgeID: item.fields.BridgeID,
                    bridgeLink: item.fields.BridgeLink,
                    newsTabLink: item.fields.NewsTabLink,
                    cloudStorageLink: item.fields.CloudStorageLink,
                    createdById: item.createdBy.user.id
                });
                formattedIncidentsData = formattedIncidentsData.filter((e: any) => e.location !== "");
            });
            return formattedIncidentsData;

        } catch (error) {
            console.error(
                constants.errorLogPrefix + "_CommonService_GetDashboardData \n",
                JSON.stringify(error)
            );
        }
    }

    // format the date to show in required format
    private formatDate(inputDate: string): string {
        const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
        const dateStr = inputDate.split('T')[0];
        const yearStr = dateStr.split("-")[0];
        const monthStr = dateStr.split("-")[1];
        const dayStr = dateStr.split("-")[2];
        const timeStr = inputDate.split('T')[1];
        const hourStr = timeStr.split(":")[0];
        const minuteStr = timeStr.split(":")[1];

        // create final date string
        const formattedDate = dayStr + " " + monthNames[(parseInt(monthStr) - 1)] + ", " + yearStr + " " + hourStr + ":" + minuteStr;

        return formattedDate;
    }

    // format incident commander to show in the grid
    private formatIncidentCommander(incidentCommanderStr: string): string {
        let incidentCommanders = "";

        incidentCommanderStr.split(";").forEach(incCom => {
            if (incCom.length > 0) {
                incidentCommanders += incCom.split("|")[0] + ", ";
            }
        });
        incidentCommanders = incidentCommanders.trim();
        incidentCommanders = incidentCommanders.slice(0, -1);

        return incidentCommanders;
    }

    // format roles and users
    private formatRoleAssignments(roleAssignmentStr: string): string {
        let rolesUsersStr = "";
        if (roleAssignmentStr.length > 0 && roleAssignmentStr.split(";").length > 1) {
            roleAssignmentStr.split(";").forEach(role => {
                if (role.length > 0) {
                    let userNamesStr = "";
                    let roleStr = role.split(":")[0].trim();
                    role.split(":")[1].trim().split(",").forEach(user => {
                        userNamesStr += user.split("|")[0].trim() + ", ";
                    });
                    userNamesStr = userNamesStr.trim();
                    userNamesStr = userNamesStr.slice(0, -1);
                    rolesUsersStr += roleStr + " : " + userNamesStr + "\n\n";
                }
            });
        }
        return rolesUsersStr;
    }

    //Format roles to show in Incident history's table view.
    private formatRoleAssignmentsForGrid(roleAssignmentStr: string): any {
        let rolesUsersArray: any[] = [];
        if (roleAssignmentStr.length > 0 && roleAssignmentStr.split(";").length > 1) {
            roleAssignmentStr.split(";").forEach(role => {
                if (role.length > 0) {
                    let userNamesStr = "";
                    let roleStr = role.split(":")[0].trim();
                    role.split(":")[1].trim().split(",").forEach(user => {
                        userNamesStr += user.split("|")[0].trim() + ", ";
                    });

                    userNamesStr = userNamesStr.trim();
                    userNamesStr = userNamesStr.slice(0, -1);
                    rolesUsersArray.push({
                        Role: roleStr,
                        Users: userNamesStr
                    })
                }
            });
        }
        return rolesUsersArray;
    }


    //#endregion

    //#region Create Incident Methods

    // get dropdown options for Incident Type, Status and Role Assignments
    public async getDropdownOptions(graphEndpoint: any, graph: Client, isIncidentStatus?: boolean): Promise<any> {
        try {
            const listData = await graph.api(graphEndpoint).get();

            // Prepare the output array
            const drpdwnOptions: Array<any> = new Array<any>();

            // Map the JSON response to the output array
            listData.value.forEach((item: any) => {
                if (isIncidentStatus) {
                    drpdwnOptions.push({
                        status: item.fields.Title,
                        id: item.fields.id
                    });
                }
                else {
                    drpdwnOptions.push(
                        item.fields.Title
                    );
                }
            });
            return drpdwnOptions;
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "_CommonService_GetDropdownOptions \n",
                JSON.stringify(error)
            );
        }
    }

    // Generic method to update item to SharePoint list
    public async updateItemInList(graphEndpoint: any, graph: Client, listItemObj: any): Promise<any> {
        return await graph.api(graphEndpoint).update(listItemObj);
    }

    // Generic method to add new item to SharePoint list
    public async addItemInList<T>(graphEndpoint: any, graph: Client, listItemObj: any): Promise<T> {
        return await graph.api(graphEndpoint).post(listItemObj) as T;
    }

    // create channel
    public async createChannel(graphEndpoint: any, graph: Client, channelObj: any): Promise<any> {
        return await graph.api(graphEndpoint).post(JSON.stringify(channelObj));
    }

    // generic method for a POST graph query
    /**
     * Sends a POST request to the specified Microsoft Graph endpoint.
     *
     * @param graphEndpoint - The endpoint of the Microsoft Graph API to which the request is sent.
     * @param graph - The Microsoft Graph client used to send the request.
     * @param requestObj - The request object containing the data to be sent in the POST request.
     * @param maxRetry - The maximum number of retry attempts in case of a failure (default is 1).
     * @param timeOut - The time in milliseconds to wait before retrying the request (default is 1000 ms).
     * @returns A promise that resolves to the response of the POST request.
     * @throws Will throw an error if the request fails after the maximum number of retry attempts.
     */
    public async sendGraphPostRequest(graphEndpoint: any, graph: Client, requestObj: any, maxRetry: number = 1, timeOut: number = 1000): Promise<any> {
        let currentAttempt = 0;
        while (currentAttempt < maxRetry) {
            try {
                return await graph.api(graphEndpoint).post(requestObj);
            } catch (error: any) {
                console.error(constants.errorLogPrefix + "CommonServices_sendGraphPostRequest \n", JSON.stringify(error));
                currentAttempt++;
                if (currentAttempt === maxRetry) {
                    throw error;
                }
                await this.timeOut(timeOut);
            }
        }
    }

    /**
     * Pauses the execution for a specified amount of time.
     * 
     * @param ms - The number of milliseconds to pause.
     * @returns A promise that resolves after the specified time has elapsed.
     */
    private async timeOut(ms: number) {
        return new Promise(resolve => setTimeout(resolve, ms));
    }

    // generic method for a PUT graph query
    public async sendGraphPutRequest(graphEndpoint: any, graph: Client, requestObj: any): Promise<any> {
        return await graph.api(graphEndpoint).put(requestObj);
    }

    // generic method for a Delete graph query
    public async sendGraphDeleteRequest(graphEndpoint: any, graph: Client): Promise<any> {
        return await graph.api(graphEndpoint).delete();
    }

    // generic method for a Patch Graph Query
    public async sendGraphPatchRequest(graphEndpoint: any, graph: Client, requestObj: any): Promise<any> {
        return await graph.api(graphEndpoint).patch(requestObj);
    }

    //get Role Default Values
    public async getRoleDefaultData(graphEndpoint: any, graph: Client): Promise<any> {
        try {
            const roleDefaultData = await graph.api(graphEndpoint).get();
            let formattedRolesData: Array<IRoleItem> = new Array<IRoleItem>();

            // Map the JSON response to the output array
            roleDefaultData.value.forEach((item: any) => {
                formattedRolesData.push({
                    itemId: item.fields.id,
                    role: item.fields.Title,
                    users: this.formatPeoplePickerData(item.fields.Users.split(',')),
                    lead: item.fields.RoleLead ? this.formatLeadPeoplePicker(item.fields.RoleLead) : []

                })
            });
            return formattedRolesData;
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "_CommonService_GetRoleDefaultData \n",
                JSON.stringify(error)
            );
        }
    }

    //get Role Default Values
    public async getIncidentTypeDefaultData(graphEndpoint: any, graph: Client): Promise<any> {
        try {
            const incidentTypeRoledata = await graph.api(graphEndpoint).get();
            let formattedRolesData: Array<IIncidentTypeDefaultRoleItem> = new Array<IIncidentTypeDefaultRoleItem>();

            // Map the JSON response to the output array
            incidentTypeRoledata.value.forEach((item: any) => {
                formattedRolesData.push({
                    itemId: item.fields.id,
                    incidentType: item.fields.Title,
                    roleAssignments: item.fields.RoleAssignment,
                    roleLeads: item.fields.RoleLeads ? item.fields.RoleLeads : "",
                    additionalChannels: item.fields.AdditionalChannels ? item.fields.AdditionalChannels : "",
                    sharedChannel: item.fields.SharedChannel ? item.fields.SharedChannel : "",
                    cloudStorageLink: item.fields.CloudStorageLink ? item.fields.CloudStorageLink : ""
                })
            });
            return formattedRolesData;
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "_CommonService_GetIncidentTypeDefaultRolesData \n",
                JSON.stringify(error)
            );
        }
    }

    //format people picker field data
    private formatPeoplePickerData = (usersData: any) => {
        try {
            const selectedUsersArr: any = [];
            usersData.map((user: any) => (
                selectedUsersArr.push({
                    displayName: user ? user.split('|')[0] : "",
                    userPrincipalName: user ? user.split('|')[2] : "",
                    id: user ? user.split('|')[1] : ""
                })
            ));
            return selectedUsersArr;
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "_CommonService_FormatPeoplePickerData \n",
                JSON.stringify(error)
            );
        }
    }

    //format people picker field data for Role Lead
    private formatLeadPeoplePicker = (usersData: any) => {
        try {
            const selectedUsersArr: any = [];
            selectedUsersArr.push({
                displayName: usersData ? usersData.split('|')[0] : "",
                userPrincipalName: usersData ? usersData.split('|')[2] : "",
                id: usersData ? usersData.split('|')[1] : ""
            })
            return selectedUsersArr;
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "_CommonService_formatLeadPeoplePicker \n",
                JSON.stringify(error)
            );
        }
    }

    //#endregion

    //#region Common Methods

    // Get tenant name
    public async getTenantDetails(graphEndpoint: any, graph: Client): Promise<any> {
        try {
            const rootSite = await graph.api(graphEndpoint).get();
            return rootSite;
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "_CommonService_GetTenantDetails \n",
                JSON.stringify(error)
            );
        }
    }

    // this is generic method to return the graph data based on input graph endpoint
    public async getGraphData(graphEndpoint: any, graph: Client): Promise<any> {
        return await graph.api(graphEndpoint).get();
    }

    // sets the initial values for regex validation object
    public getInputRegexValidationInitialState(): IInputRegexValidationStates {
        return {
            incidentNameHasError: false,
            incidentCloudStorageLinkHasError: false
        };
    };

    // perform regex validation on Incident Name and Location
    public regexValidation(incidentInfo: any, isEditMode: boolean): any {
        const regexObject: any = {};
        let inputRegexValidationObj = this.getInputRegexValidationInitialState();
        if (incidentInfo.incidentName.indexOf("#") > -1 || incidentInfo.incidentName.indexOf("&") > -1) {
            inputRegexValidationObj.incidentNameHasError = true;
        }
        //for Guest Users
        const emailRegexString = /^[^~!#$%^&*()+=[\]{}\\/|;:"<>?,.-]+[^~!#$%^&*()+=[\]{}\\/|;:"<>?,]*@[^~!#$%^&*()+=[\]{}\\/|;:"<>?,]+\.[^~!#$%^&*()+=[\]{}\\/|;:"<>?,]*[^~!#$%^&*()+=[\]{}\\/|;:"<>?,.-]+$/;
        const guestUsers = incidentInfo.guestUsers;

        incidentInfo.guestUsers?.forEach((user: any, idx: number) => {
            const trimmedEmail = user?.email?.trim();
            if (trimmedEmail !== "" && trimmedEmail?.indexOf("\\") === -1 && emailRegexString.test(trimmedEmail) &&
                /^\S+$/.test(trimmedEmail)) {
                guestUsers[idx].hasEmailRegexError = false;
            }
            else {
                guestUsers[idx].hasEmailRegexError = trimmedEmail !== "";
            }
        });
        regexObject.guestUsers = guestUsers;

        if (!this.isValidHttpUrl(incidentInfo.cloudStorageLink) && incidentInfo.cloudStorageLink !== undefined &&
            incidentInfo.cloudStorageLink !== "") {
            inputRegexValidationObj.incidentCloudStorageLinkHasError = true;
        }
        regexObject.inputRegexValidationObj = inputRegexValidationObj;
        return { ...regexObject };
    }

    //method to validate a URL
    public isValidHttpUrl(cloudLink: string) {
        let url;
        try {
            url = new URL(cloudLink);
        } catch (_) {
            return false;
        }
        return url.protocol === "http:" || url.protocol === "https:";
    }

    // get existing  members of the team
    public async getExistingTeamMembers(graphEndpoint: string, graph: Client, top: number = 500): Promise<any> {
        try {
            const data: { value: any[] } = { value: [] };
            const members = await this.getGraphData(`${graphEndpoint}?$top=${top}`, graph);
            data.value = members.value;
            let nextLink = members["@odata.nextLink"];
            while (nextLink) {
                const nextData = await this.getGraphData(nextLink, graph);
                data.value = [...data.value, ...nextData.value];
                nextLink = nextData["@odata.nextLink"];
            }
            return data;
        } catch (ex) {
            console.error(
                constants.errorLogPrefix + "CommonService_GetExistingTeamMembers \n",
                JSON.stringify(ex)
            );
            throw ex;
        }
    }

    //Log exception to App Insights
    public trackException(appInsights: any, error: any, componentName: string, methodName: string, userPrincipalName: any) {
        let exception = {
            exception: error,
            severityLevel: SeverityLevel.Error
        };

        appInsights.trackException(exception, { Component: componentName, Method: methodName, User: userPrincipalName })
    }

    //track events in App Insight
    public trackTrace(appInsights: any, message: string, incidentid: string, userPrincipalName: any) {
        let trace = {
            message: message,
            severityLevel: SeverityLevel.Information
        };
        appInsights.trackTrace(trace, { User: userPrincipalName, IncidentID: incidentid })
    }
    //#endregion

    //#region Team Name Configuration

    // get team name configuration list data
    public async getConfigData(graphEndpoint: any, graph: Client, recordsRequired: any): Promise<any> {
        try {
            const configData = await graph.api(graphEndpoint).get();

            // Prepare the output array
            let formattedData: IConfigSettingItem;

            //filter data based on key
            const filteredData = configData.value.filter((e: any) => recordsRequired.includes(e.fields.Title));
            formattedData = filteredData.map((item: any) => {
                return {
                    itemId: item.fields.id,
                    title: item.fields.Title,
                    value: item.fields.Value
                };
            });

            return formattedData;
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "_CommonService_getConfigData \n",
                JSON.stringify(error)
            );
        }
    }

    //update team name configurations
    public async updateTeamNameConfigData(graphEndpoint: any, graph: Client, listItemObj: any): Promise<any> {
        try {
            return await graph.api(graphEndpoint).update(listItemObj);
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "_CommonService_UpdateTeamNameConfigData \n",
                JSON.stringify(error)
            );
        }
    }

    public sortConfigData(configValues: Object) {
        try {
            const filteredObject = Object.fromEntries(Object.entries(configValues).filter(function ([key, value]) {
                return value > 0;
            }));
            //sort object by order
            let obj = Object.entries(filteredObject)
                .sort(([, a], [, b]) => a - b)
                .reduce((r, [k, v]) => ({ ...r, [k]: v }), {});
            return obj;
        } catch (error) {

        }
    }
    //#endregion

    //Get version data of an incident
    public async getVersionsData(graphEndpoint: any, graph: Client): Promise<any> {
        try {
            const incidentsData = await graph.api(graphEndpoint).get();

            // Prepare the output array
            let formattedIncidentsData: Array<IListItem> = this.getFormattedIncidentsData(incidentsData);
            return formattedIncidentsData;
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "_CommonService_GetVersionsData \n",
                JSON.stringify(error)
            );
        }
    }

    private getFormattedIncidentsData(incidentsData: any) {
        let formattedIncidentsData: Array<IListItem> = new Array<IListItem>();

        // Map the JSON response to the output array
        incidentsData.value.forEach((item: any) => {
            //Skipping version 2.0 since those changes are updated by system.
            if (item.id !== "2.0") {
                formattedIncidentsData.push({
                    incidentName: item.fields.IncidentName,
                    incidentCommander: this.formatIncidentCommander(item.fields.IncidentCommander),
                    status: item.fields.Status ? item.fields.Status : item.fields.IncidentStatus,
                    location: item.fields.Location !== "null" ? this.formatGeoLocation(item).DisplayName : "",
                    startDate: this.formatDate(item.fields.StartDateTime),
                    modifiedDate: new Date(item.fields.Modified).toDateString().slice(4) + " " + new Date(item.fields.Modified).toLocaleTimeString(),
                    incidentDescription: item.fields.Description,
                    incidentType: item.fields.IncidentType,
                    roleAssignments: item.fields.RoleAssignment ? this.formatRoleAssignments(item.fields.RoleAssignment) : item.fields.RoleAssignment,
                    roleAssignmentsObj: item.fields.RoleAssignment ? this.formatRoleAssignmentsForGrid(item.fields.RoleAssignment) : item.fields.RoleAssignment,
                    roleLeads: item.fields.RoleLeads ? this.formatRoleAssignments(item.fields.RoleLeads) : item.fields.RoleLeads,
                    roleLeadsObj: item.fields.RoleLeads ? this.formatRoleAssignmentsForGrid(item.fields.RoleLeads) : item.fields.RoleLeads,
                    severity: item.fields.Severity ? item.fields.Severity : "",
                    lastModifiedBy: item.lastModifiedBy.user.displayName,
                    reasonForUpdate: item.fields.ReasonForUpdate,
                    bridgeID: item.fields.BridgeID,
                    cloudStorageLink: item.fields.CloudStorageLink
                });
            }
        });
        return formattedIncidentsData;
    }

    //Get page height of version list in list view.
    public getPageHeight(idx: number | undefined, itemHeight: number, numberOfItemsOnPage: number): number {
        const value = idx ?? 0;
        let height = 0;
        for (let i = value; i < value + numberOfItemsOnPage; ++i) {
            height += itemHeight;
        }
        return height;
    };

//Create default tasks fpr an incident based on the incident type and the defulat tasks list
    public async createdefaultPlannerTasks(planId: string, bucketId: string, graph: Client, siteId: string, defaultTasksList: string, incTypeID: string) {
        try {

            //Get the tasks from the "TEOC-Tasks" lists and filter it based on Incident Type
            const defaultTasksListEndpoint = `${graphConfig.spSiteGraphEndpoint}${siteId}${graphConfig.listsGraphEndpoint}/${defaultTasksList}/items?expand=fields(select=id,Title,IncidentTypeLookupId,IncidentType)&$Top=5000`;
            const defaultTasksListItems = await graph.api(defaultTasksListEndpoint).get();
            const defaultTasksListFiltered = defaultTasksListItems.value.filter((e: any) => e.fields.IncidentType === incTypeID);
            
            //For each default task in the list create a planner task under "To-Do" bucket
            for (const task of defaultTasksListFiltered) {                
                const taskObj = {
                    planId: planId,
                    bucketId: bucketId,
                    title: task.fields.Title
                };
                await this.sendGraphPostRequest(graphConfig.plannerTasksGraphEndpoint, graph, taskObj);
            }
        }
        catch (error) {
            console.error(constants.errorLogPrefix + "CommonServices_createdefaultPlannerTasks \n", JSON.stringify(error));
            return null;
        }
    }

    //Create planner for incident tasks based on group id
    public async createPlannerPlan(group_id: string, incident_id: string, graph: Client,siteId: string, roleAssignmentList: string,
        graphContextURL: string, tenantID: any, generalChannelId?: string, fromTaskModule?: boolean) {
        try {
            const planTitle = "Incident - " + incident_id + ": Tasks";
            const plannerObj = { "owner": group_id, "title": planTitle };

            let plannerResponse = await this.sendGraphPostRequest(graphConfig.plannerGraphEndpoint, graph, plannerObj);
            let planId = plannerResponse.id;
            console.log(constants.infoLogPrefix + "Planner plan created");

            //Get the list of roles
            const roleGraphEndpoint = `${graphConfig.spSiteGraphEndpoint}${siteId}${graphConfig.listsGraphEndpoint}/${roleAssignmentList}/items?$expand=fields&$Top=5000`;          
            let rolesList = await this.getDropdownOptions(roleGraphEndpoint, graph);
            //Remove the 'new role' from the list if it exists
            const newRoleIndex = rolesList.indexOf(constants.newRole);
            if (newRoleIndex > -1) {
                rolesList.splice(newRoleIndex, 1);
            }
            //Add "To-Do" bucket to the list
            rolesList.unshift(constants.plannerBucketTitle);
            let toDoBucketId="";

            //Create a bucket for each role with "To-Do" bucket as first bucket
            for (const role of rolesList) {
                //Order "To Do" as a first bucket
                const order = (role == constants.plannerBucketTitle) ? "5637 !" : "adhg !";
                const bucketObj = { "name": role, "planId": planId, "orderHint": order };
                let plannerResponse = await this.sendGraphPostRequest(graphConfig.plannerGraphEndpoint, graph, plannerObj);
                let bucketResponse = await this.sendGraphPostRequest(graphConfig.bucketsGraphEndpoint, graph, bucketObj);
                //Get the bucket ID for "To do" bucket to create the default tasks
                if (role == constants.plannerBucketTitle)
                    toDoBucketId = bucketResponse.id;
            }

           let general_channel_id = generalChannelId;
            if (fromTaskModule) {
                //Get General channel id
                general_channel_id = await this.getChannelId(graph, group_id, constants.General);
            }

            const graphTabEndPoint = graphConfig.teamsGraphEndpoint + "/" + group_id + graphConfig.channelsGraphEndpoint + "/" + general_channel_id + graphConfig.tabsGraphEndpoint;
            const tasksAppEndPoint = graphContextURL + graphConfig.tasksbyPlannerAppGraphEndPoint;

            let tasksTabObj = {};
            if (graphContextURL === constants.commercialGraphContextURL) {
                //for commercial tenant
                tasksTabObj = {
                    "displayName": planTitle,
                    "teamsApp@odata.bind": tasksAppEndPoint,
                    "configuration": {
                        "entityId": "tt.c_" + generalChannelId + "_p_" + planId,
                        "contentUrl": "https://tasks.office.com/{tid}/Home/PlannerFrame?page=7&auth_pvr=OrgId&auth_upn={userPrincipalName}&groupId={groupId}&planId=" + planId + "&channelId={channelId}&entityId={entityId}&tid={tid}&userObjectId={userObjectId}&subEntityId={subEntityId}&sessionId={sessionId}&theme={theme}&mkt={locale}&ringId={ringId}&PlannerRouteHint={tid}",
                        "removeUrl": "https://tasks.office.com/{tid}/Home/PlannerFrame?page=13&auth_pvr=OrgId&auth_upn={userPrincipalName}&groupId={groupId}&planId=" + planId + "&channelId={channelId}&entityId={entityId}&tid={tid}&userObjectId={userObjectId}&subEntityId={subEntityId}&sessionId={sessionId}&theme={theme}&mkt={locale}&ringId={ringId}&PlannerRouteHint={tid}",
                        "websiteUrl": "https://tasks.office.com/" + tenantID + "/Home/PlanViews/" + planId + "?Type=PlanLink&Channel=TeamsTab"
                    }
                };
            }
            else {
                //for GCCH tenant
                tasksTabObj = {
                    "displayName": planTitle,
                    "teamsApp@odata.bind": tasksAppEndPoint,
                    "configuration": {
                        "entityId": planId,
                        "contentUrl": "https://tasks.office365.us/{tid}/Home/PlannerFrame?page=7&planId=" + planId + "&auth_pvr=Orgid&auth_upn={userPrincipalName}&groupId={groupId}&entityId={entityId}&tid={tid}&userObjectId={userObjectId}&channelId={channelId}&sessionId={sessionId}&theme={theme}&mkt={locale}&ringId={ringId}&PlannerRouteHint={tid}",
                        "removeUrl": "https://tasks.office365.us/{tid}/Home/PlannerFrame?page=13&planId=" + planId + "&auth_pvr=Orgid&auth_upn={userPrincipalName}&groupId={groupId}&entityId={entityId}&tid={tid}&userObjectId={userObjectId}&channelId={channelId}&sessionId={sessionId}&theme={theme}&mkt={locale}&ringId={ringId}&PlannerRouteHint={tid}",
                        "websiteUrl": "https://tasks.office365.us/" + tenantID + "/Home/PlanViews/" + planId,
                    }
                };
            }
            //adding tasks by planner and to do app to general channel
            // Add max retry logic for tab creation to handle Teams Graph API limitations
            await this.sendGraphPostRequest(graphTabEndPoint, graph, tasksTabObj, 3);
            console.log(constants.infoLogPrefix + "Tasks app is added to General Channel");
            return{
                planId:planId,
                toDoBucketId:toDoBucketId
            }
          
        }
        catch (error) {
            console.error(constants.errorLogPrefix + "CommonServices_CreatePlannerPlan \n", JSON.stringify(error));

            return null;
        }
    }

    //Create Active Dashboard Tab in Incident Team General Channel
    public async createActiveDashboardTab(graph: Client, group_id: string, generalChannelId: string,
        graphContextURL: string, appSettings: any) {
        try {
            //Get TEOC app id from Teams app catalog
            const teocAppId = await this.getGraphData(graphContextURL + graphConfig.appCatalogsTEOCAppEndpoint, graph);
            const teocAppDataBindUrl = `${graphContextURL}${graphConfig.teamsAppsEndpoint}/${teocAppId.value[0].id}`;

            //Install TEOC App to Incident Team
            const installAppEndpoint = `${graphContextURL}${graphConfig.teamsGraphEndpoint}/${group_id}${graphConfig.installedAppsEndpoint}`;
            const installAppBody = { "teamsApp@odata.bind": teocAppDataBindUrl };
            await this.sendGraphPostRequest(installAppEndpoint, graph, installAppBody);

            //Add TEOC App to General Channel - Active Dashboard Tab
            const graphTabEndPoint = graphConfig.teamsGraphEndpoint + "/" + group_id
                + graphConfig.channelsGraphEndpoint + "/" + generalChannelId + graphConfig.tabsGraphEndpoint;
            const teocObj = {
                "displayName": constants.activeDashboardTabTitle,
                "teamsApp@odata.bind": teocAppDataBindUrl,
                "configuration": {
                    "entityId": appSettings.entityId,
                    "contentUrl": appSettings.contentUrl,
                    "removeUrl": appSettings.removeUrl,
                    "websiteUrl": appSettings.websiteUrl
                }
            };
            // Add max retry logic for tab creation to handle Teams Graph API limitations
            await this.sendGraphPostRequest(graphTabEndPoint, graph, teocObj, 3);

            console.log(constants.infoLogPrefix + `TEOC App added to Active Dashboard Tab in Incident Team General channel`);
        }
        catch (error) {
            console.error(
                constants.errorLogPrefix + "CommonServices_createAndUpdateActiveDashboardTab \n",
                JSON.stringify(error)
            );
        }

    }

    //Get Channel ID of a particular team channel
    public async getChannelId(graph: any, teamGroupId: any, channelName: string) {
        try {
            //get channel id related to current site teams group
            const endPoint = graphConfig.teamsGraphEndpoint + "/" + teamGroupId +
                graphConfig.channelsGraphEndpoint +
                "?$filter=startsWith(displayName, '" + channelName + "')&$select=id";
            const response = await this.getGraphData(endPoint, graph);
            const channelId = response.value[0].id;
            return channelId;
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "CommonService_getChannelId \n",
                JSON.stringify(error)
            );
            return "Failed " + error;
        }
    }

    //Get URL of a tab in channel of a Team
    public async getTabURL(graph: any, teamGroupId: any, channelID: string, tabName: string) {
        try {
            //get URL for the tab
            const tabsGraphEndPoint = graphConfig.teamsGraphEndpoint + "/" + teamGroupId + graphConfig.channelsGraphEndpoint + "/" + channelID + graphConfig.tabsGraphEndpoint +
                "?$filter=startsWith(displayName, '" + tabName + "')&$select=webUrl";

            const tabResults = await this.getGraphData(tabsGraphEndPoint, graph)
            return tabResults.value[0].webUrl;
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "CommonService_getTabURL \n",
                JSON.stringify(error)
            );
            throw error;
        }
    }

    //method to format location column
    private formatGeoLocation = (item: any) => {
        let formattedData: ILocation = {
            EntityType: "",
            Address: {
                City: "",
                CountryOrRegion: "",
                PostalCode: "",
                State: "",
                Street: ""
            },
            Coordinates: {
                Latitude: "",
                Longitude: ""
            },
            DisplayName: "",
            LocationUri: "",
            UniqueId: ""
        };

        if (this.isJson(item.fields.Location)) {
            let i = JSON.parse(item.fields.Location);
            formattedData.Address = i.Address;
            formattedData.Coordinates = i.Coordinates;
            formattedData.DisplayName = i.DisplayName;
            formattedData.LocationUri = i.LocationUri
            formattedData.EntityType = i.EntityType;
            return formattedData;
        }
        else {
            formattedData = {
                EntityType: "Custom",
                Address: {
                    City: "",
                    CountryOrRegion: "",
                    PostalCode: "",
                    State: "",
                    Street: ""
                },
                Coordinates: {
                    Latitude: '0',
                    Longitude: '0'
                },
                DisplayName: item.fields.Location,
                LocationUri: "",
                UniqueId: ""
            };
            return formattedData;
        }
    }

    //method to verify if string is JSON or not
    private isJson = (str: any) => {
        try {
            JSON.parse(str);
        } catch (e) {
            return false;
        }
        return true;
    }

}



