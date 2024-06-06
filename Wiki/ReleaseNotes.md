This page contains the release information for Microsoft Teams Emergency Operations Center,

## Version history
| Version | Release Date |
|----|----|
| 3.1 | Jun 6, 2024 |
| 3.0 | Feb 23, 2024 |
| 2.0 | May 02, 2023 |
| 1.0 | Oct 11, 2022 |
| 0.5.1 | Apr 28, 2022 |
| 0.5 | Mar 31, 2022 |

## Release notes

### v3.1 (Jun 6, 2024)

Below are the changes released with this version:

- Upgraded Teams JS to version 2.21.0
- Accessibility improvements
- Bug fix: Update incident fails when same user is assigned as a team member and lead for 'Secondary Incident Commander' role. 

### v3.0 (Feb 19, 2024)

Below are the changes released with this version:

- New Feature: Accessibility compliant UI
- New Feature: Map Viewer on dashboard for Incidents
- New Feature: Ability for the admins to enable/disable map viewer and provide Bing maps key
- New Feature: Location picker while creating incidents
- New Feature: Incident Manager names in dash board replaced by person cards
- New Feature: Sorting implemented for all columns in dashboard table
- New Feature: Export to PDF functionality for Incident History
- New Feature: Dashboard tab created automatically in 'General' channel in the teams created for an incident
- New Feature: Ability to create 'Private' channels while creating incidents
- New Feature: Adaptive cards sent to Incident teams when a bridge is enabled/disabled through 'Active Dashboard'
- New Feature: Admins can modify the App Title
- Migration of classic Application Insights to workspace-based Application Insights.

### v2.0 (May 02, 2023)

Below are the changes released with this version:

- New Feature: Supported in GCC High environment.
- New Feature: Active Dashboard - View Role Assignments and Leads, Post Channel Announcements, Create and Join Bridge, Create and Manage Planner Tasks.
- New Feature: Ability to assign Role Lead for each role.
- New Feature: Ability to enable Role based access to control access to "Manage Settings" and "Create New Incident" features.
- New Feature: Ability to add Guest Users from the incident and assign them to any role except Secondary Incident Commander.
- New Feature: Ability to add Cloud Storage Location on the incident. 
- New Feature: Ability to Create or Modify Team Channels during incident creation.
- Enhancement: Add Secondary Incident Commander as owners to the team.
- Bug Fix: Create incident fails in the tenant where the site creation path was set to /teams/ in SharePoint Admin Portal.

### v1.0 (Oct 11, 2022)

This app is released in General Availability (GA) status. Below are the changes released with this version - 

- New Feature: Ability to configure the Team Name format from dashboard.
- New Feature: Ability to view the version history for each incident from dashboard.
- New Feature: Ability to "Save default users for Roles" and "Save default roles for an Incident Type".
- New Feature: Local language support (translations) available for 12 languages. 
- Upgraded the Teams toolkit version from v3.7.0 to v4.0.5


### v0.5.1 (Apr 28, 2022)

This app is released in _Public Preview_ mode. Below are the changes released with this version - 

- Allow users to manage roles and incident types from dashboard.
- Fix for incident creation failure for GCC tenant.
- Wiki updates

### v0.5 (Mar 31, 2022)

This app is released in _Public Preview_ mode. Below are the features released with this version - 

- Allow users to create incidents based on the type of incidents and location.
- View and manage the incidents from the dashboard.
- Allow users to work in dedicated Teams channels for each incident.
- Allow users to manage the data in their SharePoint tenant

