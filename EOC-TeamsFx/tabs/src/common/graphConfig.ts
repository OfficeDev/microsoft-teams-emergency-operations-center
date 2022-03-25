export const organizationGraphEndpoint = "https://graph.microsoft.com/v1.0/organization?$select=verifiedDomains";
export const spSiteGraphEndpoint = "https://graph.microsoft.com/v1.0/sites/";
export const meGraphEndpoint = "/me";
export const teamGroupsGraphEndpoint = "/groups";
export const teamGraphEndpoint = "/team";
export const teamsGraphEndpoint = "/teams";
export const channelsGraphEndpoint = "/channels";
export const appsGraphEndpoint = "/installedApps";
export const tabsGraphEndpoint = "/tabs";
export const scheduleGraphEndpoint = "/schedule";
export const schedulingGroupsGraphEndpoint = "/schedulingGroups";
export const openShiftsGraphEndpoint = "/openShifts";
export const sitesGraphEndpoint = "/sites";
export const listsGraphEndpoint = "/lists";
export const columnsGraphEndpoint = "/columns";
export const usersGraphEndpoint = "https://graph.microsoft.com/v1.0/users/";
export const teamsAppsGraphEndpoint = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/";
export const newsTabTeamsAppIdGraphEndpoint = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/0ae35b36-0fd7-422e-805b-d53af1579093";
export const assessmentTabTeamsAppIdGraphEndpoint = "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/2a527703-1f6f-4559-a332-d8a7d288cd88";
export const allAppsGraphEndpoint = "/appCatalogs/teamsApps?$filter=distributionMethod eq 'organization'";
export const betaGraphEndpoint = "https://graph.microsoft.com/beta/teams/";
export const tagsGraphEndpoint = "/tags";
export const membersGraphEndpoint = "/members";
export const addMembersGraphEndpoint = "/members/add";
export const addUsersGraphEndpoint = "https://graph.microsoft.com/v1.0/users";
export const scope = [
    "User.Read",
    "Sites.Manage.All",
    "People.Read",
    "Group.ReadWrite.All",
    "TeamMember.ReadWrite.All",
    "TeamsTab.Create",
    "TeamworkTag.ReadWrite",
    "Directory.AccessAsUser.All"
]
