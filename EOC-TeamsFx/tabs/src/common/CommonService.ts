import { Client } from "@microsoft/microsoft-graph-client";
import * as constants from './Constants';
import { SeverityLevel } from "@microsoft/applicationinsights-common";

export interface IListItem {
    itemId?: string;
    incidentId?: number;
    incidentName?: string;
    incidentCommander?: string;
    status?: string;
    location?: string;
    startDate?: string;
    startDateUTC?: string;
    modifiedDate?: string;
    teamWebURL?: string;
    roleAssignments?: string;
    roleAssignmentsObj?: string;
    incidentDescription?: string;
    incidentType?: string;
    incidentCommanderObj?: string;
    severity?: string;
    lastModifiedBy?: string;
    version?: string;
    reasonForUpdate?: string;
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
}

export interface IIncidentTypeDefaultRoleItem {
    itemId?: string;
    incidentType?: string;
    roleAssignments?: string;
}

export interface IInputRegexValidationStates {
    incidentNameHasError: boolean;
    incidentLocationHasError: boolean;
}

export interface ITeamNameConfigItem {
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
            var formattedIncidentsData: Array<IListItem> = new Array<IListItem>();

            // Map the JSON response to the output array
            incidentsData.value.forEach((item: any) => {
                formattedIncidentsData.push({
                    itemId: item.fields.id,
                    incidentId: parseInt(item.fields.IncidentId),
                    incidentName: item.fields.IncidentName,
                    incidentCommander: this.formatIncidentCommander(item.fields.IncidentCommander),
                    incidentCommanderObj: item.fields.IncidentCommander,
                    status: item.fields.IncidentStatus,
                    location: item.fields.Location,
                    startDate: this.formatDate(item.fields.StartDateTime),
                    startDateUTC: new Date(item.fields.StartDateTime).toISOString().slice(0, new Date(item.fields.StartDateTime).toISOString().length - 1),
                    modifiedDate: item.fields.Modified,
                    teamWebURL: item.fields.TeamWebURL,
                    incidentDescription: item.fields.Description,
                    incidentType: item.fields.IncidentType,
                    roleAssignments: item.fields.RoleAssignment,
                    severity: item.fields.Severity ? item.fields.Severity : ""
                });
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
    public async getDropdownOptions(graphEndpoint: any, graph: Client): Promise<any> {
        try {
            const listData = await graph.api(graphEndpoint).get();

            // Prepare the output array
            const drpdwnOptions: Array<any> = new Array<any>();

            // Map the JSON response to the output array
            listData.value.forEach((item: any) => {
                drpdwnOptions.push(
                    item.fields.Title
                );
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

    // create channel
    public async createChannel(graphEndpoint: any, graph: Client, channelObj: any): Promise<any> {
        return await graph.api(graphEndpoint).post(JSON.stringify(channelObj));
    }

    // generic method for a POST graph query
    public async sendGraphPostRequest(graphEndpoint: any, graph: Client, requestObj: any): Promise<any> {
        return await graph.api(graphEndpoint).post(requestObj);
    }

    // generic method for a PUT graph query
    public async sendGraphPutRequest(graphEndpoint: any, graph: Client, requestObj: any): Promise<any> {
        return await graph.api(graphEndpoint).put(requestObj);
    }

    // generic method for a Delete graph query
    public async sendGraphDeleteRequest(graphEndpoint: any, graph: Client): Promise<any> {
        return await graph.api(graphEndpoint).delete();
    }

    //update teams display name
    public async updateTeamsDisplayName(graphEndpoint: any, graph: Client, requestObj: any): Promise<any> {
        return await graph.api(graphEndpoint).patch(requestObj);
    }

    //get Role Default Values
    public async getRoleDefaultData(graphEndpoint: any, graph: Client): Promise<any> {
        try {
            const roleDefaultData = await graph.api(graphEndpoint).get();
            var formattedRolesData: Array<IRoleItem> = new Array<IRoleItem>();

            // Map the JSON response to the output array
            roleDefaultData.value.forEach((item: any) => {
                formattedRolesData.push({
                    itemId: item.fields.id,
                    role: item.fields.Title,
                    users: this.formatPeoplePickerData(item.fields.Users.split(','))
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
    public async getIncidentTypeDefaultRolesData(graphEndpoint: any, graph: Client): Promise<any> {
        try {
            const incidentTypeRoledata = await graph.api(graphEndpoint).get();
            var formattedRolesData: Array<IIncidentTypeDefaultRoleItem> = new Array<IIncidentTypeDefaultRoleItem>();

            // Map the JSON response to the output array
            incidentTypeRoledata.value.forEach((item: any) => {
                formattedRolesData.push({
                    itemId: item.fields.id,
                    incidentType: item.fields.Title,
                    roleAssignments: item.fields.RoleAssignment
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
            usersData.map((user: any) => {
                selectedUsersArr.push({
                    displayName: user ? user.split('|')[0] : "",
                    userPrincipalName: user ? user.split('|')[2] : "",
                    id: user ? user.split('|')[1] : ""
                });
            });
            return selectedUsersArr;
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "_CommonService_FormatPeoplePickerData \n",
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
            return rootSite.siteCollection.hostname;
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
            incidentLocationHasError: false
        };
    };

    // perform regex validation on Incident Name and Location
    public regexValidation(incidentInfo: any): any {
        let inputRegexValidationObj = this.getInputRegexValidationInitialState();
        let regexvalidationSuccess = true;
        if (incidentInfo.incidentName.indexOf("#") > -1 || incidentInfo.incidentName.indexOf("&") > -1) {
            inputRegexValidationObj.incidentNameHasError = true;
        }
        if (incidentInfo.location.indexOf("#") > -1 || incidentInfo.location.indexOf("&") > -1) {
            inputRegexValidationObj.incidentLocationHasError = true;
        }
        if (inputRegexValidationObj.incidentLocationHasError || inputRegexValidationObj.incidentNameHasError) {
            regexvalidationSuccess = false;
        }
        return inputRegexValidationObj;
    }

    // get existing  members of the team
    public async getExistingTeamMembers(graphEndpoint: string, graph: Client): Promise<any> {
        return new Promise(async (resolve, reject) => {

            // const graphEndpoint = graphConfig.teamsGraphEndpoint + "/" + teamId + graphConfig.membersGraphEndpoint;
            try {
                const members = await this.getGraphData(graphEndpoint, graph);
                resolve(members);
            } catch (ex) {
                console.error(
                    constants.errorLogPrefix + "UpdateIncident_GetExistingTeamMembers \n",
                    JSON.stringify(ex)
                );
                reject(ex);
            }
        });
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
    public async getConfigData(graphEndpoint: any, graph: Client, key: any): Promise<any> {
        try {
            const configData = await graph.api(graphEndpoint).get();

            // Prepare the output array
            var formattedData: Array<ITeamNameConfigItem> = new Array<ITeamNameConfigItem>();

            //filter data based on key
            const filteredObject = configData.value.filter((e: any) => e.fields.Title === key);
            formattedData.push({
                itemId: filteredObject[0].fields.id,
                title: filteredObject[0].fields.Title,
                value: JSON.parse(filteredObject[0].fields.Value)
            });
            return formattedData[0];
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "_CommonService_GetTeamNameConfigData \n",
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

    public async getVersionsData(graphEndpoint: any, graph: Client): Promise<any> {
        try {
            const incidentsData = await graph.api(graphEndpoint).get();

            // Prepare the output array
            var formattedIncidentsData: Array<IListItem> = new Array<IListItem>();

            // Map the JSON response to the output array
            incidentsData.value.forEach((item: any) => {
                //Skipping version 2.0 and 3.0 since those changes are updated by system.
                if (item.id !== "2.0" && item.id !== "3.0") {
                    formattedIncidentsData.push({
                        incidentName: item.fields.IncidentName,
                        incidentCommander: this.formatIncidentCommander(item.fields.IncidentCommander),
                        status: item.fields.IncidentStatus,
                        location: item.fields.Location,
                        startDate: this.formatDate(item.fields.StartDateTime),
                        modifiedDate: new Date(item.fields.Modified).toDateString().slice(4) + " " + new Date(item.fields.Modified).toLocaleTimeString(),
                        incidentDescription: item.fields.Description,
                        incidentType: item.fields.IncidentType,
                        roleAssignments: item.fields.RoleAssignment ? this.formatRoleAssignments(item.fields.RoleAssignment) : item.fields.RoleAssignment,
                        roleAssignmentsObj: item.fields.RoleAssignment ? this.formatRoleAssignmentsForGrid(item.fields.RoleAssignment) : item.fields.RoleAssignment,
                        severity: item.fields.Severity ? item.fields.Severity : "",
                        lastModifiedBy: item.lastModifiedBy.user.displayName,
                        reasonForUpdate: item.fields.ReasonForUpdate
                    });
                }
            });
            return formattedIncidentsData;
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "_CommonService_GetVersionsData \n",
                JSON.stringify(error)
            );
        }
    }

    //Get page height of version list in list view.
    public getPageHeight(idx: number | undefined, itemHeight: number, numberOfItemsOnPage: number): number {
        const value = idx !== undefined ? idx : 0;
        let height = 0;
        for (let i = value; i < value + numberOfItemsOnPage; ++i) {
            height += itemHeight;
        }
        return height;
    };


}

