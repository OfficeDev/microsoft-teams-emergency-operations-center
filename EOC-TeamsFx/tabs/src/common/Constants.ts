export const defaultGraphBaseURL = "https://graph.microsoft.com/"
export const commercialGraphContextURL = "https://graph.microsoft.com/v1.0/"
export const AppVersion = "3.2.0 ";
export const EOCAppName = "MIFTEOC";
export const all = "All";
export const planning = "Planning";
export const active = "Active";
export const closed = "Closed";
export const dashboardPageSize = 10;
export const mobileWidth = 568;
export const maxCharacterLength = 50;
export const maxCharLengthForSingleLine = 50;
export const maxCharLengthForMultiLine = 500;
export const teamEOCPrefix = "TEOC";
export const messageBarType = {
    "error": "error",
    "success": "success"
}
export const regexValidator = "/[^A-Za-z 0-9 @!]/g";
export const Assessment = "Assessment";
export const GroundAssessments = "Ground Assessments";
export const ViewMap = "View Map";
export const News = "News";
export const Announcements = "Announcements";
export const General = "General";
export const adminTeam = "eoc admin";
export const errorLogPrefix = "EOC-Error: ";
export const infoLogPrefix = "EOC-Log: ";
export const defaultLocale = "en-us";
export const incidentCommanderRoleName = "Incident Commander";
export const secondaryIncidentCommanderRole = "Secondary Incident Commander";
export const newRole = "New Role (Specify a new role to be created and assign users to)";
export const yes = "Yes";
export const no = "No";

//EOC Header 

//Urls
export const helpUrl = "https://aka.ms/m365teocappsupport";
export const feedbackUrl = "https://aka.ms/teocappfeedback";
export const msPublicSectorUrl = "https://aka.ms/MTCPublicSector ";
export const drivingAdoptionUrl = "https://aka.ms/drivingadoption";
export const m365EocUrl = "https://aka.ms/m365teoc";
export const m365EocAppUrl = "https://aka.ms/m365teocapp";
export const githubEocUrl = "https://aka.ms/m365teocapp";

export const teamsWebUrl = "https://teams.microsoft.com";
export const teamsWebUrlGCCH = "https://gov.teams.microsoft.us";

export const componentNames = {
    EOCHomeComponent: 'EOCHome',
    DashboardComponent: 'Dashboard',
    IncidentDetailsComponent: 'IncidentDetails',
    UpdateIncidentCompoent: 'UpdateIncident',
    TeamNameConfiguration: 'TeamNameConfiguration',
    IncidentHistoryComponent: 'IncidentHistory',
    TasksComponent: "Tasks",
    MembersComponent: "Members",
    BridgeComponent: "Bridge",
    AdminSettingsComponent: "AdminSettings",
    ConfigSettingsComponent: "ConfigSettings",
    MapViewer:"MapViewer"
}

export const severity = ["Low", "Medium", "High", "Critical"]
export const teamNameConfigOrderDropdown = ["1", "2", "3", "4", "Don't Include"];
export const maxCharLengthForPrefix = 10;

export const teamNameConfigConstants = {
    IncidentName: "IncidentName",
    IncidentType: "IncidentType",
    Prefix: "Prefix",
    PrefixValue: "PrefixValue",
    StartDate: "StartDate",
    DontInclude: "Don't Include"
}

export const defaultChannelConstants = {
    Logistics: "Logistics",
    Planning: "Planning",
    Recovery: "Recovery"
};

//Incident history constants
export const itemHeight = 30;
export const numberOfItemsOnPage = 10;
export const listViewFieldMinWidth = 190;
export const listViewFieldMaxWidth = 204;
export const listViewNewMinWidth = 210;
export const listViewNewMaxWidth = 308;
export const listViewOldMinWidth = 210;
export const listViewOldMaxWidth = 308;
export const listViewItemInitialCount = 8;
export const modifiedDate = "modifiedDate";
export const lastModifiedBy = "lastModifiedBy";
export const roleAssignmentsObj = "roleAssignmentsObj";
export const roleLeadsObj = "roleLeadsObj";

//Active Dashboard
export const plannerBucketTitle = "To do";
export const activeDashboardTabTitle = "Active Dashboard";

//Admin Settings
export const adminRole = "Admin";

//Teams theme mode
export const defaultMode = "default";
export const darkMode = "dark";
export const contrastMode = "contrast";

//Teams channel types
export const standardChannel = "standard";
export const privateChannel = "private";

//accessibility
export const enterKey = "Enter";
export const tabKey = "Tab";
export const escapeKey = "Escape";
export const addRoleMessage = "Selected users are added to the role successfully";
export const sizePerPageLabel = "Size per page ";
export const sortAscAriaSort = "ascending";
export const sortDescAriaSort = "descending";
export const requiredAriaLabel = " required";
export const sortableAriaLabel = " sortable";
export const startDateAriaLabel = "Start Date";
export const startTimeAriaLabel = "Start Time";
export const selectedAriaLabel = " Selected";

//Config Settings
export const appTitleKey = "AppTitle";
export const enableRoles = "EnableRoles";
export const bingMapsKey = "BingMapsKey";
export const teamNameConfig = "TeamNameConfig";
export const editIncidentAccessRoleKey = "EditIncidentAccessRole";
export const noneOption = "None";


//location picker
export const defaultOutlookBaseURL = "https://outlook.office365.com/.default";
export const outlookAPIFindLocations = "https://outlook.office365.com/SchedulingB2/api/v1.0/me/findmeetinglocations";

//location picker for GCCH
export const defaultOutlookBaseURLGCCH = "https://outlook.office365.us/.default";
export const outlookAPIFindLocationsGCCH = "https://outlook.office365.us/SchedulingB2/api/v1.0/me/findmeetinglocations";

//Exception Codes
export const authorizationRequestDenied = "Authorization_RequestDenied";

//Exception Messages
export const groupCreationAccessDeniedErrorMessage = "Tenant admin has not enabled Unified Group creation.";
