import { Client } from "@microsoft/microsoft-graph-client";
import * as constants from './Constants';
import moment from "moment";
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
    incidentDescription?: string;
    incidentType?: string;
    incidentCommanderObj?: string;
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

export interface IInputRegexValidationStates {
    incidentNameHasError: boolean;
    incidentLocationHasError: boolean;
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
                    roleAssignments: item.fields.RoleAssignment
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
    formatDate = (inputDate: string): string => {
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
    formatIncidentCommander = (incidentCommanderStr: string): string => {
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

    // generic method for a PATCH graph query
    public async sendGraphPatchRequest(graphEndpoint: any, graph: Client, requestObj: any): Promise<any> {
        return await graph.api(graphEndpoint).patch(requestObj);
    }

    // generic method for a Delete graph query
    public async sendGraphDeleteRequest(graphEndpoint: any, graph: Client): Promise<any> {
        return await graph.api(graphEndpoint).delete();
    }

    //#endregion

    //#region Common Methods

    // Get tenant name
    public async getTenantDetails(graphEndpoint: any, graph: Client): Promise<any> {
        try {
            const domains = await graph.api(graphEndpoint).get();

            let domainName = "";
            domains.value.forEach((element: any) => {
                element.verifiedDomains.forEach((vDomains: any) => {
                    if (vDomains.isInitial) {
                        if (vDomains.name.indexOf('.onmicrosoft.com') > -1) {
                            domainName = vDomains.name.split('.onmicrosoft.com')[0];
                        }
                    }
                });
            });

            return domainName;
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
    public getInputRegexValildationInitialState = (): IInputRegexValidationStates => {
        return {
            incidentNameHasError: false,
            incidentLocationHasError: false
        };
    };

    // perform regex validation on Incident Name and Location
    public regexValidation = (incidentInfo: any): any => {
        let inputRegexValidationObj = this.getInputRegexValildationInitialState();
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
    getExistingTeamMembers = async (graphEndpoint: string, graph: Client): Promise<any> => {
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
    public trackException(appInsights: any, error: any, componentName: string, methodName: string, userPrincipalName:any) {
        let exception = {
            exception: error,
            severityLevel: SeverityLevel.Error
        };

        appInsights.trackException(exception, { Component: componentName, Method: methodName ,User : userPrincipalName})
    }

    //track events in App Insight
    public trackTrace(appInsights: any, message: string, incidentid: string, userPrincipalName:any) {
        let trace = {
            message: message,
            severityLevel: SeverityLevel.Information
        };
        appInsights.trackTrace(trace, { User: userPrincipalName, IncidentID: incidentid})
    }
    //#endregion
}

