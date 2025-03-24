export const rootSiteGraphEndpoint = "/sites/root";
export const spSiteGraphEndpoint = "/sites/";
export const meGraphEndpoint = "/me";
export const teamGroupsGraphEndpoint = "/groups";
export const teamGraphEndpoint = "/team";
export const teamsGraphEndpoint = "/teams";
export const channelsGraphEndpoint = "/channels";
export const tabsGraphEndpoint = "/tabs";
export const listsGraphEndpoint = "/lists";
export const tagsGraphEndpoint = "/tags";
export const membersGraphEndpoint = "/members";
export const addMembersGraphEndpoint = "/members/add";
export const plannerGraphEndpoint = "/planner/plans";
export const bucketsGraphEndpoint = "/planner/buckets";
export const onlineMeetingGraphEndpoint = "/me/onlineMeetings";
export const messagesGraphEndpoint = "/messages";
export const usersGraphEndpoint = "users";
export const sharepointPageAndListTabGraphEndpoint = "appCatalogs/teamsApps/2a527703-1f6f-4559-a332-d8a7d288cd88";
export const tasksbyPlannerAppGraphEndPoint = "appCatalogs/teamsApps/com.microsoft.teamspace.tab.planner";
export const invitationsGraphEndpoint = "/invitations";
export const emailInvitationsGraphEndpoint = "/me/sendMail";
export const installedAppsEndpoint = "/installedApps";
export const teamsAppsEndpoint = "/appCatalogs/teamsApps";
export const plannerTasksGraphEndpoint = "/planner/tasks";


// 'bef61400-db9b-41d4-a617-403deb7bbe77' is the TEOC app externalId. This is the id from Manifest file thats going to be constant.
export const appCatalogsTEOCAppEndpoint = "/appCatalogs/teamsApps?$filter=externalId eq 'bef61400-db9b-41d4-a617-403deb7bbe77'";

export const scope = [
    "User.Read",
    "Sites.Manage.All",
    "People.Read",
    "Group.ReadWrite.All",
    "TeamMember.ReadWrite.All",
    "TeamsTab.Create",
    "TeamworkTag.ReadWrite",
    "Directory.AccessAsUser.All",
    "User.ReadBasic.All",
    "Tasks.Read",
    "Tasks.ReadWrite",
    "Group.Read.All",
    "OnlineMeetings.ReadWrite",
    "TeamsAppInstallation.ReadWriteSelfForTeam",
    "Mail.Send",
    "AppCatalog.Read.All"
]
