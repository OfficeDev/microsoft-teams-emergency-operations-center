// interface for incident details form entity
export interface IIncidentEntity {
    incidentId: string;
    incidentName: string;
    incidentStatus: IIncidentStatus;
    location: string;
    incidentType: string;
    incidentDesc: string;
    startDateTime: string;
    startDate: Date;
    startTime: Date;
    incidentCommander: UserDetails;
    selectedRole: string;
    assignedUser: UserDetails[];
    assignedLead: UserDetails[];
    severity: string;
    reasonForUpdate: string;
    additionalTeamChannels: Array<IAdditionalTeamChannels>;
    cloudStorageLink: string;
    guestUsers: Array<IGuestUsers>;
}

// class for incident details form entity
export class IncidentEntity implements IIncidentEntity {
    public incidentId!: string;
    public incidentName!: string;
    public incidentStatus!: IIncidentStatus;
    public location!: string;
    public incidentType!: string;
    public incidentDesc!: string;
    public startDateTime!: string;
    public startDate!: Date;
    public startTime!: Date;
    public incidentCommander!: UserDetails;
    public selectedRole!: string;
    public assignedUser!: UserDetails[];
    public assignedLead!: UserDetails[];
    public severity!: string;
    public reasonForUpdate!: string;
    public additionalTeamChannels!: IAdditionalTeamChannels[];
    public cloudStorageLink!: string;
    public guestUsers!: IGuestUsers[];
}

export interface UserDetails {
    userName: string;
    userEmail: string;
    userId: string;
}

export interface IIncidentStatus {
    status: string | undefined;
    id: number | undefined;
}

export interface IAdditionalTeamChannels {
    channelName: string;
    channelType?: string;
    hasRegexError: boolean;
    regexErrorMessage: string;
    selectedRoleUsers?: string;
    selectedRoleUserIds?: string;
    expandedGroups?: any;
}

export interface IGuestUsers {
    email: string;
    displayName: string;
    hasEmailRegexError: boolean;
    hasDisplayNameRegexError: boolean;
    hasDisplayNameValidationError: boolean;
    hasEmailValidationError: boolean;
    [propName: string]: any;
}

export interface RoleAssignments {
    role: string;
    userNamesString: string;
    userObjString: string;
    userDetailsObj: UserDetails[];
    leadNameString: string;
    leadObjString: string;
    leadDetailsObj: UserDetails[];
    saveDefault: any;
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
    membershipType?: string;
    members?: Array<any>;
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
    failedChannels: any;
}

export interface Tab {
    id: string;
    displayName: string;
    webUrl: string;
    configuration: any;
}

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
    incidentReasonForUpdateHasError: boolean;
    cloudStorageLinkHasError: boolean;
    guestUsersHasError: boolean;
}

export interface IInputRegexValidationStates {
    incidentNameHasError: boolean;
    incidentLocationHasError?: boolean;
    incidentCloudStorageLinkHasError: boolean;
}

export interface ILocation {
    EntityType: string,
    Address: {
        City: string,
        CountryOrRegion: string,
        PostalCode: string,
        State: string,
        Street: string
    },
    Coordinates: {
        Latitude: string,
        Longitude: string
    },
    DisplayName: string,
    LocationUri: string,
    UniqueId: string
}