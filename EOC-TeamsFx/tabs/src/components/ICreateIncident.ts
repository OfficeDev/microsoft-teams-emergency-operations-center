// interface for incident details form entity
export interface IIncidentEntity {
    incidentId: string;
    incidentName: string;
    incidentStatus: string;
    location: string;
    incidentType: string;
    incidentDesc: string;
    startDateTime: string;
    incidentCommander: UserDetails;
    selectedRole: string;
    assignedUser: UserDetails[];
}

// class for incident details form entity
export class IncidentEntity implements IIncidentEntity {
    public incidentId!: string;
    public incidentName!: string;
    public incidentStatus!: string;
    public location!: string;
    public incidentType!: string;
    public incidentDesc!: string;
    public startDateTime!: string;
    public incidentCommander!: UserDetails;
    public selectedRole!: string;
    public assignedUser!: UserDetails[];
}

export interface UserDetails {
    userName: string;
    userEmail: string;
    userId: string;
}

export interface RoleAssignments {
    role: string;
    userNamesString: string;
    userObjString: string;
    userDetailsObj: UserDetails[];
}

export interface ITeamGroupInfo {
    "displayName": string;
    "mailNickname": string;
    "description": string;
    "visibility": string;
    "groupTypes": Array<string>;
    "mailEnabled": boolean;
    "securityEnabled": boolean;
    "members@odata.bind": Array<string>;
    "owners@odata.bind": Array<string>;
}

export interface ITeamChannel {
    displayName: string;
}

export interface ChannelCreationStatus {
    channelName: string;
    isCreated: boolean;
    rawData: any;
    noOfCreationAttempt: number;
}

export interface ChannelCreationResult {
    isFullyCreated: boolean;
    isPartiallyCreated: boolean;
    failedEntries: any;
    successEntries: any;
}
export interface Tab {
    id: string;
    displayName: string;
    webUrl: string;
    configuration: any;
}

// export interface ITeamCreatedResponse {
//     fully_done: boolean;
//     partially_done: boolean;
//     all_failed: boolean;
//     team_info: any;
//     error: {
//         channel_creations: any;
//         app_installation: any;
//         member_creations: any;
//         all_fail: any;
//     };
// }

export interface ShiftGroupInfo {
    IncidentCommander: Array<any>;
    RoleAssignee: Array<any>
}

export interface ISharedOpenShift {
    notes: string;
    openSlotCount: number;
    displayName: string;
    startDateTime: string;
    endDateTime: string;
    theme: string;
}

export interface IOpenShift {
    schedulingGroupId: string;
    sharedOpenShift: ISharedOpenShift;
}

export interface ITeamCreatedResponse {
    fullyDone: boolean;
    partiallyDone: boolean;
    allFailed: boolean;
    teamInfo: any;
    error: {
        channelCreations: any;
        appInstallation: any;
        memberCreations: any;
        allFail: any;
    };
}

export interface IOpenShift {
    schedulingGroupId: string;
    sharedOpenShift: ISharedOpenShift;
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