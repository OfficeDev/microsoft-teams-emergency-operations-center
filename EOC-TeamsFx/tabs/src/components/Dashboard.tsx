import { Icon, IPivotItemProps, Pivot, PivotItem } from '@fluentui/react';
import { Popover, PopoverSurface, PopoverTrigger, SelectTabData, SelectTabEvent, Tab, TabList } from "@fluentui/react-components";
import { Button, Flex, FormInput, Loader, SearchIcon } from "@fluentui/react-northstar";
import { ApplicationInsights } from '@microsoft/applicationinsights-web';
import { Person } from '@microsoft/mgt-react';
import { Client } from "@microsoft/microsoft-graph-client";
import * as microsoftTeams from "@microsoft/teams-js";
import 'bootstrap/dist/css/bootstrap.min.css';
import * as React from "react";
import BootstrapTable from "react-bootstrap-table-next";
import 'react-bootstrap-table-next/dist/react-bootstrap-table2.min.css';
import paginationFactory from 'react-bootstrap-table2-paginator';
import CommonService from "../common/CommonService";
import * as constants from '../common/Constants';
import * as graphConfig from '../common/graphConfig';
import siteConfig from '../config/siteConfig.json';
import '../scss/Dashboard.module.scss';
import { MapViewer } from './MapViewer';

export interface IDashboardProps {
    graph: Client;
    tenantName: string;
    siteId: string;
    onCreateTeamClick: Function;
    onEditButtonClick(incidentData: any): void;
    localeStrings: any;
    onBackClick(showMessageBar: string): void;
    showMessageBar(message: string, type: string): void;
    hideMessageBar(): void;
    appInsights: ApplicationInsights;
    userPrincipalName: any;
    siteName: any;
    onShowAdminSettings: Function;
    onShowIncidentHistory: Function;
    onShowActiveBridge: Function;
    isRolesEnabled: boolean;
    isUserAdmin: boolean;
    settingsLoader: boolean;
    currentThemeName: string;
    activeDashboardIncidentId: string;
    fromActiveDashboardTab: boolean;
    isMapViewerEnabled: boolean;
    azureMapsKeyConfigData: any;
    graphBaseUrl: any;
    currentUserId: string;
}

export interface IDashboardState {
    allIncidents: any;
    planningIncidents: any;
    activeIncidents: any;
    completedIncidents: any;
    filteredAllIncidents: any;
    filteredPlanningIncidents: any;
    filteredActiveIncidents: any;
    filteredCompletedIncidents: any;
    searchText: string | undefined;
    isDesktop: boolean;
    showLoader: boolean;
    loaderMessage: string;
    selectedIncident: any;
    isManageCalloutVisible: boolean;
    currentTab: any;
    incidentIdAriaSort: any;
    incidentNameAriaSort: any;
    locationAriaSort: any;
    severityAriaSort: any;
    incidentCommanderObjAriaSort: any;
    startDateAriaSort: any;
    showMapViewer: boolean;
    selectedTab: any;
}

// interface for Dashboard fields
export interface IListItem {
    itemId: string;
    incidentId: string;
    incidentName: string;
    incidentCommander: string;
    status: string;
    location: string;
    startDate: string;
}

class Dashboard extends React.PureComponent<IDashboardProps, IDashboardState> {
    private dashboardRef: React.RefObject<HTMLDivElement>;
    constructor(props: IDashboardProps) {
        super(props);
        this.dashboardRef = React.createRef();
        this.state = {
            allIncidents: [],
            planningIncidents: [],
            activeIncidents: [],
            completedIncidents: [],
            filteredAllIncidents: [],
            filteredPlanningIncidents: [],
            filteredActiveIncidents: [],
            filteredCompletedIncidents: [],
            searchText: "",
            isDesktop: true,
            showLoader: true,
            loaderMessage: this.props.localeStrings.genericLoaderMessage,
            selectedIncident: [],
            isManageCalloutVisible: false,
            currentTab: "",
            incidentIdAriaSort: "",
            incidentNameAriaSort: "",
            locationAriaSort: "",
            severityAriaSort: "",
            incidentCommanderObjAriaSort: "",
            startDateAriaSort: "",
            showMapViewer: false,
            selectedTab: this.props.localeStrings.incidentDetails
        };

        this.actionFormatter = this.actionFormatter.bind(this);
        this.renderIncidentSettings = this.renderIncidentSettings.bind(this);
        this.getDashboardData = this.getDashboardData.bind(this);
    }

    private dataService = new CommonService();

    // set the state object for screen size
    resize = () => this.setState({ isDesktop: window.innerWidth > constants.mobileWidth });

    // before unmounting, remove event listener
    componentWillUnmount() {
        window.removeEventListener("resize", this.resize.bind(this));
    }

    // component life cycle method
    public async componentDidMount() {
        // Add resize event listener when Dashboard component mounts
        if (!this.props.fromActiveDashboardTab) {
            window.addEventListener("resize", this.resize.bind(this));
            this.resize();
        }
        // Get dashboard data
        this.getDashboardData();
    }


    // connect with service layer to get all incidents
    getDashboardData = async () => {

        this.setState({
            showLoader: true
        })
        try {
             // create graph endpoint for querying Incident Transaction list
            let graphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}` + graphConfig.listsGraphEndpoint + `/${siteConfig.incidentsList}/items?$expand=fields
                ($select=StatusLookupId,Status,id,IncidentId,IncidentName,IncidentCommander,Location,StartDateTime,
                Modified,TeamWebURL,Description,IncidentType,RoleAssignment,RoleLeads,Severity,PlanID,
                BridgeID,BridgeLink,NewsTabLink,CloudStorageLink)&$Top=5000`;

           let allIncidents = this.sortDashboardData(await this.dataService.getDashboardData(graphEndpoint, this.props.graph));
            
           const currentUserId = this.props.currentUserId;
           
           //Filter incidents that current user is part of, either as a member, IncidentCommander or createdBy
           allIncidents = allIncidents.filter((item: any) => 
            (item.roleAssignments?.includes(currentUserId) || 
            item.incidentCommanderObj?.includes(currentUserId) || 
            item.createdById?.includes(currentUserId))
            );
           console.log(constants.infoLogPrefix + "All Incidents retrieved");                        

            // Redirect to current Incident Active Dashboard component
           const activeIncident = allIncidents.find((e: any) => e.incidentId === parseInt(this.props.activeDashboardIncidentId));
            if (this.props.fromActiveDashboardTab && activeIncident !== undefined) {
                this.props.onShowActiveBridge(activeIncident);
            }
            else {
                // filter for Planning tab
                const planningIncidents = allIncidents.filter((e: any) => e.incidentStatusObj.status === constants.planning);
                
                // filter for Active tab
                const activeIncidents = allIncidents.filter((e: any) => e.incidentStatusObj.status === constants.active);

                // filter for Completed tab
                const completedIncidents = allIncidents.filter((e: any) => e.incidentStatusObj.status === constants.closed);
                this.setState({
                    allIncidents: allIncidents,
                    planningIncidents: planningIncidents,
                    activeIncidents: activeIncidents,
                    completedIncidents: completedIncidents,
                    filteredAllIncidents: [...allIncidents],
                    filteredPlanningIncidents: [...planningIncidents],
                    filteredCompletedIncidents: [...completedIncidents],
                    filteredActiveIncidents: [...activeIncidents],
                    showLoader: false
                });
            }
        } catch (error: any) {
            this.setState({
                showLoader: false
            })
            console.error(
                constants.errorLogPrefix + "_Dashboard_GetDashboardData \n",
                JSON.stringify(error)
            );
            this.props.showMessageBar(this.props.localeStrings.genericErrorMessage + ", " + this.props.localeStrings.getIncidentsFailedErrMsg, constants.messageBarType.error);
            // Log Exception
            this.dataService.trackException(this.props.appInsights, error, constants.componentNames.DashboardComponent, 'GetDashboardData', this.props.userPrincipalName);
        }
    }

    // sort data based on order
    sortDashboardData = (allIncidents: any): any => {
        return allIncidents.sort((currIncident: any, prevIncident: any) => (parseInt(currIncident.itemId) < parseInt(prevIncident.itemId)) ? 1 : ((parseInt(prevIncident.itemId) < parseInt(currIncident.itemId)) ? -1 : 0));
    }

    // update icon on pivots(tabs)
    _customRenderer(
        link?: IPivotItemProps,
        defaultRenderer?: (link?: IPivotItemProps) => JSX.Element | null,
    ): JSX.Element | null {
        if (!link || !defaultRenderer) {
            return null;
        }
        return (
            <span>
                <img src={require(`../assets/Images/${link.itemKey}ItemsSelected.svg`)} alt={`${link.headerText}`} id="pivot-icon-selected" />
                <img src={require(`../assets/Images/${link.itemKey}Items.svg`)} alt={`${link.headerText}`} id="pivot-icon" />
                <span id="state">&nbsp;&nbsp;{link.headerText}&nbsp;&nbsp;</span>
                <span id="count">|&nbsp;{link.itemCount}</span>
                {defaultRenderer({ ...link, headerText: undefined, itemCount: undefined })}
            </span>
        );
    }


    //pagination properties for bootstrap table
    private pagination = paginationFactory({
        page: 1,
        sizePerPage: constants.dashboardPageSize,
        showTotal: true,
        alwaysShowAllBtns: false,
        //customized the render options for pagesize button in the pagination for accessbility
        sizePerPageRenderer: ({
            options,
            currSizePerPage,
            onSizePerPageChange
        }) => (
            <div className="btn-group" role="group">
                {
                    options.map((option) => {
                        const isSelect = currSizePerPage === `${option.page}`;
                        return (
                            <button
                                key={option.text}
                                type="button"
                                onClick={() => onSizePerPageChange(option.page)}
                                className={`btn${isSelect ? ' sizeperpage-selected' : ' sizeperpage'}${this.props.currentThemeName === constants.defaultMode ? "" : " selected-darkcontrast"}`}
                                aria-label={isSelect ? constants.sizePerPageLabel + option.text + constants.selectedAriaLabel : constants.sizePerPageLabel + option.text}
                            >
                                {option.text}
                            </button>
                        );
                    })
                }
            </div>
        ),
        //customized the render options for page list in the pagination for accessbility
        pageButtonRenderer: (options: any) => {
            const handleClick = (e: any) => {
                e.preventDefault();
                if (options.disabled) return;
                options.onPageChange(options.page);
            };
            const className = `${options.active ? 'active ' : ''}${options.disabled ? 'disabled ' : ''}`;
            let ariaLabel = "";
            let pageText = "";
            switch (options.title) {
                case "first page":
                    ariaLabel = `Go to ${options.title}`;
                    pageText = '<<';
                    break;
                case "previous page":
                    ariaLabel = `Go to ${options.title}`;
                    pageText = '<';
                    break;
                case "next page":
                    ariaLabel = `Go to ${options.title}`;
                    pageText = '>';
                    break;
                case "last page":
                    ariaLabel = `Go to ${options.title}`;
                    pageText = '>>';
                    break;
                default:
                    ariaLabel = `Go to page ${options.title}`;
                    pageText = options.title;
                    break;
            }
            return (
                <li key={options.title} className={`${className}page-item${this.props.currentThemeName === constants.defaultMode ? "" : " selected-darkcontrast"}`} role="presentation" title={ariaLabel}>
                    <a className="page-link" href="#" onClick={handleClick} role="button" aria-label={options.active ? ariaLabel + constants.selectedAriaLabel : ariaLabel}>
                        <span aria-hidden="true">{pageText}</span>
                    </a>
                </li>
            );
        },
        //customized the page total renderer in the pagination for accessbility
        paginationTotalRenderer: (from, to, size) => {
            const resultsFound = size !== 0 ? `Showing ${from} to ${to} of ${size} Results` : ""
            return (
                <span className="react-bootstrap-table-pagination-total" aria-live="polite" role="status">
                    {resultsFound}
                </span>
            )
        }
    });

    // filter incident based on entered keyword
    searchDashboard = (searchText: any) => {
        let searchKeyword = "";
        // convert to lower case
        if (searchText.target.value) {
            searchKeyword = searchText.target.value.toLowerCase();
        }
        const allIncidents = this.state.allIncidents;
        let filteredAllIncidents = allIncidents.filter((incident: any) => {
            return ((incident["incidentName"] && incident["incidentName"].toLowerCase().indexOf(searchKeyword) > -1) ||
                (incident["incidentId"] && (incident["incidentId"]).toString().toLowerCase().indexOf(searchKeyword) > -1) ||
                (incident["incidentCommander"] && incident["incidentCommander"].toLowerCase().indexOf(searchKeyword) > -1) ||
                (incident["location"] && incident["location"].toLowerCase().indexOf(searchKeyword) > -1) ||
                (incident["severity"] && incident["severity"].toLowerCase().indexOf(searchKeyword) > -1))
        });

        //On Click of Cancel icon
        if (searchText.cancelable) {
            filteredAllIncidents = this.state.allIncidents;
        }
        this.setState({
            searchText: searchText.target.value,
            filteredAllIncidents: filteredAllIncidents,
            filteredPlanningIncidents: filteredAllIncidents.filter((e: any) => e.incidentStatusObj.status === constants.planning),
            filteredActiveIncidents: filteredAllIncidents.filter((e: any) => e.incidentStatusObj.status === constants.active),
            filteredCompletedIncidents: filteredAllIncidents.filter((e: any) => e.incidentStatusObj.status === constants.closed),
        });
    }

    // format the cell for Incident ID column to fix accessibility issues
    incidentIdFormatter = (cell: any, gridRow: any, rowIndex: any, formatExtraData: any) => {
        const ariaLabel = `Row ${rowIndex + 2} ${this.props.localeStrings.incidentId} ${cell}`
        return (
            <span
                aria-label={ariaLabel}
                tabIndex={0}
                role="link"
                className="incIdDeepLink"
                onClick={() => this.onDeepLinkClick(gridRow)}
                onKeyDown={(event) => {
                    if (event.key === constants.enterKey)
                        this.onDeepLinkClick(gridRow)
                }}
            >
                <span title={cell} aria-hidden="true">{cell}</span>
            </span>
        );
    }

    // format the cell for Incident Name column to fix accessibility issues
    incidentNameFormatter = (cell: any, gridRow: any, rowIndex: any, formatExtraData: any) => {
        const ariaLabel = `${this.props.localeStrings.incidentName} ${cell}`
        return (
            <span
                aria-label={ariaLabel}
                tabIndex={0}
                role="link"
                className="incIdDeepLink"
                onClick={() => this.onDeepLinkClick(gridRow)}
                onKeyDown={(event) => {
                    if (event.key === constants.enterKey)
                        this.onDeepLinkClick(gridRow)
                }}
            >
                <span title={cell} aria-hidden="true">{cell}</span>
            </span>
        );
    }

    // format the cell for Severity column to fix accessibility issues
    severityFormatter = (cell: any, gridRow: any, rowIndex: any, formatExtraData: any) => {
        const ariaLabel = `${this.props.localeStrings.fieldSeverity} ${cell}`
        return (
            <span aria-label={ariaLabel}><span title={cell} aria-hidden="true">{cell}</span></span>
        );
    }

    // format the cell for Incident Commander column to fix accessibility issues
    incidentCommanderFormatter = (cell: any, gridRow: any, rowIndex: any, formatExtraData: any) => {
        const incidentCommander = cell ? cell.split("|") : [];
        return (
            <Person
                userId={incidentCommander[1]?.trim()}
                view={3}
                personCardInteraction={1}
                className='incident-commander-person-card'
            />
        );
    }

    // format the cell for Status column to fix accessibility issues
    statusFormatter = (cell: any, row: any, rowIndex: any, formatExtraData: any) => {
        if (row.incidentStatusObj.status === constants.closed) {
            return (
                <span aria-label={`${this.props.localeStrings.status} ${this.props.localeStrings.closed}`}>
                    <img src={require("../assets/Images/ClosedIcon.svg").default} className="status-icon"
                        aria-hidden="true" title={this.props.localeStrings.closed} />
                </span>
            );
        }
        if (row.incidentStatusObj.status === constants.active) {
            return (
                <span aria-label={`${this.props.localeStrings.status} ${this.props.localeStrings.active}`}>
                    <img src={require("../assets/Images/ActiveIcon.svg").default} className="status-icon"
                        aria-hidden="true" title={this.props.localeStrings.active} />
                </span>
            );
        }
        if (row.incidentStatusObj.status === constants.planning) {
            return (
                <span aria-label={`${this.props.localeStrings.status} ${this.props.localeStrings.planning}`}>
                    <img src={require("../assets/Images/PlanningIcon.svg").default} className="status-icon"
                        aria-hidden="true" title={this.props.localeStrings.planning} />
                </span>
            );
        }
    };

    // format the cell for Location column to fix accessibility issues
    locationFormatter = (cell: any, gridRow: any, rowIndex: any, formatExtraData: any) => {
        const ariaLabel = `${this.props.localeStrings.location} ${JSON.parse(cell).DisplayName}`
        if (cell !== "null" || cell !== "") {
            return (
                <span aria-label={ariaLabel}><span aria-hidden="true" title={JSON.parse(cell).DisplayName}>{JSON.parse(cell).DisplayName}</span></span>
            );
        }
        else {
            return (
                <span aria-label={ariaLabel}><span aria-hidden="true" title={"null"}>null</span></span>
            );
        }
    }

    // format the cell for Start Date Time column to fix accessibility issues
    startDateTimeFormatter = (cell: any, gridRow: any, rowIndex: any, formatExtraData: any) => {
        const ariaLabel = `${this.props.localeStrings.startDate} ${cell}`
        return (
            <span aria-label={ariaLabel}><span aria-hidden="true" title={cell}>{cell}</span></span>
        );
    }

    // format the cell for Action column to fix accessibility issues
    public actionFormatter(_cell: any, gridRow: any, _rowIndex: any, _formatExtraData: any) {
        return (
            <span>
                {/* active dashboard icon in dashboard, on click will navigate to edit incident form */}
                <span
                    aria-label={`${this.props.localeStrings.action} ${this.props.localeStrings.activeDashboard}`}
                    onClick={() => this.props.onShowActiveBridge(gridRow)}
                    onKeyDown={(event) => {
                        if (event.key === constants.enterKey)
                            this.props.onShowActiveBridge(gridRow)
                    }}
                    tabIndex={0}
                    role="button"
                >
                    <img
                        src={require("../assets/Images/ActiveBridgeIcon.svg").default}
                        className="grid-active-bridge-icon"
                        aria-hidden="true"
                        title={this.props.localeStrings.activeDashboard}
                    />
                </span>
                {/* edit icon in dashboard, on click will navigate to edit incident form */}
                <span
                    aria-label={`${this.props.localeStrings.action} ${this.props.localeStrings.edit}`}
                    onClick={() => this.props.onEditButtonClick(gridRow)}
                    onKeyDown={(event) => {
                        if (event.key === constants.enterKey)
                            this.props.onEditButtonClick(gridRow)
                    }}
                    tabIndex={0}
                    role="button"
                >
                    <img
                        src={require("../assets/Images/GridEditIcon.svg").default}
                        className="grid-edit-icon"
                        aria-hidden="true"
                        title={this.props.localeStrings.edit}
                    />
                </span>

                {/* view version history icon in dashboard, on click will navigate to incident history form */}
                <span
                    aria-label={`${this.props.localeStrings.action} ${this.props.localeStrings.viewIncidentHistory}`}
                    onClick={() => this.props.onShowIncidentHistory(gridRow.incidentId)}
                    onKeyDown={(event) => {
                        if (event.key === constants.enterKey)
                            this.props.onShowIncidentHistory(gridRow.incidentId)
                    }}
                    tabIndex={0}
                    role="button"
                >
                    <img
                        src={require("../assets/Images/IncidentHistoryIcon.svg").default}
                        className="grid-version-history-icon"
                        aria-hidden="true"
                        title={this.props.localeStrings.viewIncidentHistory}
                    />
                </span>
            </span>
        );
    }


    // create deep link to open the associated Team
    onDeepLinkClick = (rowData: any) => {
        microsoftTeams.app.openLink(rowData.teamWebURL);
    }

    //Incident Settings Area
    public renderIncidentSettings = () => {
        return (
            <Flex space="between" wrap={true}>
                <Popover
                    open={this.state.isManageCalloutVisible}
                    inline={true}
                    onOpenChange={() => this.setState({ isManageCalloutVisible: !this.state.isManageCalloutVisible })}
                    positioning="below"
                    size='medium'
                    closeOnScroll={true}
                >

                    <PopoverTrigger disableButtonEnhancement={true}>

                        <div
                            className={`manage-links${this.state.isManageCalloutVisible ? " callout-visible" : ""}`}
                            onClick={() => this.setState({ isManageCalloutVisible: !this.state.isManageCalloutVisible })}

                            tabIndex={0}
                            onKeyDown={(event) => {
                                if (event.key === constants.enterKey)
                                    this.setState({ isManageCalloutVisible: !this.state.isManageCalloutVisible })
                            }}
                            role="button"
                            title="Manage"
                        >
                            <img
                                src={require("../assets/Images/ManageIcon.svg").default}
                                className={`manage-icon${this.props.currentThemeName === constants.defaultMode ? "" : " manage-icon-darkcontrast"}`}
                                alt=""
                            />
                            <img
                                src={require("../assets/Images/ManageIconActive.svg").default}
                                className='manage-icon-active'
                                alt=""
                            />
                            <div className='manage-label'>{this.props.localeStrings.manageLabel}</div>
                            {this.state.isManageCalloutVisible ?
                                <Icon iconName="ChevronUp" />
                                :
                                <Icon iconName="ChevronDown" />
                            }
                        </div>

                    </PopoverTrigger>
                    <PopoverSurface as="div" className="manage-links-callout" >

                        <div>
                            <div title={this.props.localeStrings.manageIncidentTypesTooltip} className="dashboard-link" onKeyDown={(event) => {
                                if (event.shiftKey)
                                    this.setState({ isManageCalloutVisible: false })
                            }}>
                                <a title={this.props.localeStrings.manageIncidentTypesTooltip} href={`https://${this.props.tenantName}/sites/${this.props.siteName}/lists/${siteConfig.incTypeList}`} target='_blank' rel="noreferrer">
                                    <img src={require("../assets/Images/Manage Incident Types.svg").default} alt="" className={`manage-item-icon${this.props.currentThemeName === constants.defaultMode ? "" : " manage-item-icon-darkcontrast"}`}
                                    />
                                    <span role="button" className="manage-callout-text">{this.props.localeStrings.incidentTypesLabel}</span>
                                </a>
                            </div>
                            <div title={this.props.localeStrings.manageRolesTooltip} className="dashboard-link">
                                <a title={this.props.localeStrings.manageRolesTooltip} href={`https://${this.props.tenantName}/sites/${this.props.siteName}/lists/${siteConfig.roleAssignmentList}`} target='_blank' rel="noreferrer">
                                    <img src={require("../assets/Images/Manage Roles.svg").default} alt="" className={`manage-item-icon${this.props.currentThemeName === constants.defaultMode ? "" : " manage-item-icon-darkcontrast"}`}
                                    />
                                    <span role="button" className="manage-callout-text">{this.props.localeStrings.roles}</span>
                                </a>
                            </div>
                            <div title={this.props.localeStrings.tasksAdminMenuTooltip} className="dashboard-link">
                                <a title={this.props.localeStrings.tasksAdminMenuTooltip} href={`https://${this.props.tenantName}/sites/${this.props.siteName}/lists/${siteConfig.defaulTasksList}`} target='_blank' rel="noreferrer">
                                    <img src={require("../assets/Images/Tasks.svg").default} alt="" className={`manage-item-icon${this.props.currentThemeName === constants.defaultMode ? "" : " manage-item-icon-darkcontrast"}`}
                                    />
                                    <span role="button" className="manage-callout-text">{this.props.localeStrings.tasksAdminMenuLabel}</span>
                                </a>
                            </div>
                            <div
                                className="dashboard-link"
                                title={this.props.localeStrings.adminSettingsLabel}
                                onClick={() => this.props.onShowAdminSettings()}
                                onKeyDown={(event) => {
                                    if (event.key === constants.enterKey)
                                        this.props.onShowAdminSettings()
                                    else if (!event.shiftKey)
                                        this.setState({ isManageCalloutVisible: false })
                                }}
                                role="button"
                                tabIndex={0}
                            >
                                <span className="team-name-link" tabIndex={0}>
                                    <img
                                        src={require("../assets/Images/AdminSettings.svg").default}
                                        alt=""
                                        className={`manage-item-icon${this.props.currentThemeName === constants.defaultMode ? "" : " manage-item-icon-darkcontrast"}`}
                                    />
                                    <span className="manage-callout-text">
                                        {this.props.localeStrings.adminSettingsLabel}
                                    </span>
                                </span>
                            </div>
                        </div>
                    </PopoverSurface>
                </Popover>
                <Button
                    primary className={`create-incident-btn${this.props.currentThemeName === constants.contrastMode ? " create-icon-contrast" : ""}`}
                    fluid={true}
                    onClick={() => this.props.onCreateTeamClick()}
                    title={this.props.localeStrings.btnCreateIncident}
                    onFocus={() => this.setState({ isManageCalloutVisible: false })}
                >
                    <img src={require("../assets/Images/ButtonEditIcon.svg").default} alt={this.props.localeStrings.btnCreateIncident} />
                    {this.props.localeStrings.btnCreateIncident}
                </Button>
            </Flex>
        );
    }

    //render the sort caret on the header column for accessbility
    customSortCaret = (order: any, column: any) => {
        const ariaLabel = navigator.userAgent.match(/iPhone/i) ? "sortable" : "";
        const id = column.dataField;
        if (!order) {
            return (
                <div className="sort-order" id={id} aria-label={ariaLabel}>
                    <span className="dropdown-caret">
                    </span>
                    <span className="dropup-caret">
                    </span>
                </div>);
        }
        else if (order === 'asc') {
            switch (column.dataField) {
                case "incidentId":
                    this.setState({
                        incidentIdAriaSort: constants.sortAscAriaSort, incidentNameAriaSort: "", locationAriaSort: "",
                        severityAriaSort: "", incidentCommanderObjAriaSort: "", startDateAriaSort: ""
                    });
                    break;
                case "incidentName":
                    this.setState({
                        incidentNameAriaSort: constants.sortAscAriaSort, incidentIdAriaSort: "",
                        locationAriaSort: "", severityAriaSort: "", incidentCommanderObjAriaSort: "", startDateAriaSort: ""
                    });
                    break;
                case "location":
                    this.setState({
                        locationAriaSort: constants.sortAscAriaSort, incidentNameAriaSort: "", incidentIdAriaSort: "",
                        severityAriaSort: "", incidentCommanderObjAriaSort: "", startDateAriaSort: ""
                    });
                    break;
                case "severity":
                    this.setState({
                        severityAriaSort: constants.sortAscAriaSort, incidentNameAriaSort: "", incidentIdAriaSort: "",
                        locationAriaSort: "", incidentCommanderObjAriaSort: "", startDateAriaSort: ""
                    })
                    break;
                case "incidentCommanderObj":
                    this.setState({
                        incidentCommanderObjAriaSort: constants.sortAscAriaSort, incidentNameAriaSort: "",
                        incidentIdAriaSort: "", locationAriaSort: "", severityAriaSort: "", startDateAriaSort: ""
                    })
                    break;
                default:
                    this.setState({
                        startDateAriaSort: constants.sortAscAriaSort, incidentNameAriaSort: "", incidentIdAriaSort: "",
                        locationAriaSort: "", severityAriaSort: "", incidentCommanderObjAriaSort: ""
                    });
            }
            return (
                <div className="sort-order">
                    <span className="dropup-caret">
                    </span>
                </div>);
        }
        else if (order === 'desc') {
            switch (column.dataField) {
                case "incidentId":
                    this.setState({
                        incidentIdAriaSort: constants.sortDescAriaSort, incidentNameAriaSort: "", locationAriaSort: "",
                        severityAriaSort: "", incidentCommanderObjAriaSort: "", startDateAriaSort: ""
                    });
                    break;
                case "incidentName":
                    this.setState({
                        incidentNameAriaSort: constants.sortDescAriaSort, incidentIdAriaSort: "", locationAriaSort: "",
                        severityAriaSort: "", incidentCommanderObjAriaSort: "", startDateAriaSort: ""
                    })
                    break;
                case "location":
                    this.setState({
                        locationAriaSort: constants.sortDescAriaSort, incidentNameAriaSort: "", incidentIdAriaSort: "",
                        severityAriaSort: "", incidentCommanderObjAriaSort: "", startDateAriaSort: ""
                    });
                    break;
                case "severity":
                    this.setState({
                        severityAriaSort: constants.sortDescAriaSort, incidentNameAriaSort: "", incidentIdAriaSort: "",
                        locationAriaSort: "", incidentCommanderObjAriaSort: "", startDateAriaSort: ""
                    });
                    break;
                case "incidentCommanderObj":
                    this.setState({
                        incidentCommanderObjAriaSort: constants.sortDescAriaSort, incidentNameAriaSort: "",
                        incidentIdAriaSort: "", locationAriaSort: "", severityAriaSort: "", startDateAriaSort: ""
                    });
                    break;
                default:
                    this.setState({
                        startDateAriaSort: constants.sortDescAriaSort, incidentNameAriaSort: "", incidentIdAriaSort: "",
                        locationAriaSort: "", severityAriaSort: "", incidentCommanderObjAriaSort: ""
                    });
            }
            return (
                <div className="sort-order">
                    <span className="dropdown-caret">
                    </span>
                </div>);
        }
        return null;
    }

    //custom header format for sortable column for accessbility
    headerFormatter(column: any, colIndex: any, { sortElement, filterElement }: any) {
        //adding sortable information to aria-label to fix the accessibility issue in iOS Voiceover
        if (navigator.userAgent.match(/iPhone/i)) {
            const id = column.dataField;
            return (
                <button tabIndex={-1} aria-describedby={id} aria-label={column.text} className='sort-header'>
                    {column.text}
                    {sortElement}
                </button>
            );
        }
        else {
            return (
                <div aria-hidden="true" title={column.text} className='header-div-wrapper'>
                    <span className='header-span-text'>{column.text}</span>
                    {sortElement}
                </div>
            );
        }
    }

    public onTabSelect = (event: SelectTabEvent, data: SelectTabData) => {
        this.setState({
            selectedTab: data.value,
            showMapViewer: data.value === this.props.localeStrings.mapViewer ? true : false
        });
    };

    public render() {
        // Header object for dashboard
        const dashboardHeader: any = [
            {
                dataField: 'incidentId',
                text: this.props.localeStrings.incidentId,
                sort: true,
                sortCaret: this.customSortCaret,
                formatter: this.incidentIdFormatter,
                headerFormatter: this.headerFormatter,
                headerAttrs: { 'aria-sort': this.state.incidentIdAriaSort, 'role': 'columnheader', 'scope': 'col' }
            }, {
                dataField: 'incidentName',
                text: this.props.localeStrings.incidentName,
                sort: true,
                sortCaret: this.customSortCaret,
                formatter: this.incidentNameFormatter,
                headerFormatter: this.headerFormatter,
                headerAttrs: { 'aria-sort': this.state.incidentNameAriaSort, 'role': 'columnheader', 'scope': 'col' }
            }, {
                dataField: 'severity',
                text: this.props.localeStrings.fieldSeverity,
                formatter: this.severityFormatter,
                headerAttrs: { 'aria-sort': this.state.severityAriaSort, 'role': 'columnheader', 'scope': 'col' },
                sort: true,
                sortValue: (cell: any) => constants.severity.indexOf(cell),
                sortCaret: this.customSortCaret,
                headerFormatter: this.headerFormatter
            }, {
                dataField: 'incidentCommanderObj',
                text: this.props.localeStrings.incidentCommander,
                formatter: this.incidentCommanderFormatter,
                headerAttrs: { 'aria-sort': this.state.incidentCommanderObjAriaSort, 'role': 'columnheader', 'scope': 'col' },
                sort: true,
                sortCaret: this.customSortCaret,
                headerFormatter: this.headerFormatter
            }, {
                dataField: 'status',
                text: this.props.localeStrings.status,
                formatter: this.statusFormatter,
                headerAttrs: { 'role': 'columnheader', 'scope': 'col', "aria-label": this.props.localeStrings.status },
                headerFormatter: this.headerFormatter
            }, {
                dataField: 'location',
                text: this.props.localeStrings.location,
                sort: true,
                sortFunc: (a: any, b: any, order: any) => { a = JSON.parse(a).DisplayName; b = JSON.parse(b).DisplayName; return order === 'asc' ? a.localeCompare(b) : b.localeCompare(a) },
                sortCaret: this.customSortCaret,
                headerFormatter: this.headerFormatter,
                formatter: this.locationFormatter,
                headerAttrs: { 'aria-sort': this.state.locationAriaSort, 'role': 'columnheader', 'scope': 'col' }
            }, {
                dataField: 'startDate',
                text: this.props.localeStrings.startDate,
                formatter: this.startDateTimeFormatter,
                headerAttrs: { 'aria-sort': this.state.startDateAriaSort, 'role': 'columnheader', 'scope': 'col' },
                sort: true,
                sortValue: (cell: any) => new Date(cell),
                sortCaret: this.customSortCaret,
                headerFormatter: this.headerFormatter
            }, {
                dataField: 'action',
                text: this.props.localeStrings.action,
                formatter: this.actionFormatter,
                headerAttrs: { 'role': 'columnheader', 'scope': 'col', "aria-label": this.props.localeStrings.action },
                classes: `edit-icon-${this.props.currentThemeName}`,
                headerFormatter: this.headerFormatter
            }
        ]
        const isDarkOrContrastTheme = this.props.currentThemeName === constants.darkMode || this.props.currentThemeName === constants.contrastMode;

        return (
            <>
                {this.state.showLoader ?
                    <Loader label={this.state.loaderMessage} size="largest" />
                    :
                    <div>
                        <div className={`dashboard-search-btn-area${isDarkOrContrastTheme ? " eoc-searcharea-darkcontrast" : ""}`}>
                            <div className="container">
                                <Flex space="between" wrap={true}>
                                    <div className="search-area">
                                        <FormInput
                                            type="text"
                                            icon={<SearchIcon />}
                                            clearable
                                            placeholder={this.props.localeStrings.searchPlaceholder}
                                            fluid={true}
                                            maxLength={constants.maxCharLengthForSingleLine}
                                            required
                                            title={this.props.localeStrings.searchPlaceholder}
                                            onChange={(evt) => this.searchDashboard(evt)}
                                            value={this.state.searchText}
                                            successIndicator={false}
                                            aria-describedby='noincident-all-tab noincident-active-tab noincident-planning-tab noincident-completed-tab'
                                        />
                                    </div>
                                    {this.props.isRolesEnabled ?
                                        this.props.isUserAdmin ? this.renderIncidentSettings() : <></>
                                        : this.props.settingsLoader ? <Loader size="smallest" className="settings-loader" /> : this.renderIncidentSettings()
                                    }
                                </Flex>
                            </div>
                        </div>
                        <div className={`dashboard-pivot-container${isDarkOrContrastTheme ? " eoc-background-darkcontrast" : ""}`}>
                            <div className="container">
                                <TabList defaultSelectedValue={this.state.selectedTab} className="main-tab-list" onTabSelect={this.onTabSelect}>
                                    <Tab value={this.props.localeStrings.incidentDetails} className="grid-heading"><div onClick={() => this.setState({ showMapViewer: false })}>{this.props.localeStrings.incidentDetails}</div>
                                    </Tab>
                                    {this.props.isMapViewerEnabled ?
                                        <Tab value={this.props.localeStrings.mapViewer} className="grid-heading"><div onClick={() => this.setState({ showMapViewer: true })}>{this.props.localeStrings.mapViewer}</div>
                                        </Tab>
                                        : <></>}
                                </TabList>
                                <Flex gap={this.state.isDesktop ? "gap.medium" : "gap.small"} space="evenly" id="status-indicators" wrap={true}>
                                    <Flex gap={this.state.isDesktop ? "gap.small" : "gap.smaller"}>
                                        <img src={require("../assets/Images/AllItems.svg").default} alt="All Items" />
                                        <label>{this.props.localeStrings.all}</label>
                                    </Flex>
                                    <Flex gap={this.state.isDesktop ? "gap.small" : "gap.smaller"}>
                                        <img src={require("../assets/Images/PlanningItems.svg").default} alt="Planning Items" />
                                        <label>{this.props.localeStrings.planning}</label>
                                    </Flex>
                                    <Flex gap={this.state.isDesktop ? "gap.small" : "gap.smaller"}>
                                        <img src={require("../assets/Images/ActiveItems.svg").default} alt="Active Items" />
                                        <label>{this.props.localeStrings.active}</label>
                                    </Flex>
                                    <Flex gap={this.state.isDesktop ? "gap.small" : "gap.smaller"}>
                                        <img src={require("../assets/Images/ClosedItems.svg").default} alt="Completed Items" />
                                        <label>{this.props.localeStrings.closed}</label>
                                    </Flex>
                                </Flex>
                                <Pivot
                                    aria-label="Incidents Details"
                                    linkFormat="tabs"
                                    overflowBehavior='none'
                                    className={`pivot-tabs${isDarkOrContrastTheme ? " pivot-button-darkcontrast" : ""}`}
                                    onLinkClick={(item, ev) => (this.setState({ currentTab: item?.props.headerText }))}
                                    ref={this.dashboardRef}
                                >
                                    <PivotItem
                                        headerText={this.props.localeStrings.all}
                                        itemCount={this.state.filteredAllIncidents.length}
                                        itemKey="All"
                                        onRenderItemLink={this._customRenderer}
                                    >
                                        {!this.state.showMapViewer ?
                                            <BootstrapTable
                                                bootstrap4
                                                bordered={false}
                                                keyField="incidentId"
                                                columns={dashboardHeader}
                                                data={this.state.filteredAllIncidents}
                                                wrapperClasses={isDarkOrContrastTheme ? "table-darkcontrast" : ""}
                                                headerClasses={isDarkOrContrastTheme ? "table-header-darkcontrast" : ""}
                                                pagination={this.pagination}
                                                noDataIndication={() => (<div id="noincident-all-tab" aria-live="polite" role="status">{this.props.localeStrings.noIncidentsFound}</div>)}
                                            />
                                            :
                                            <MapViewer graphBaseUrl={this.props.graphBaseUrl} incidentData={this.state.filteredAllIncidents} azureMapKey={this.props.azureMapsKeyConfigData} showMessageBar={this.props.showMessageBar} userPrincipalName={this.props.userPrincipalName} localeStrings={this.props.localeStrings} appInsights={this.props.appInsights}></MapViewer>
                                        }
                                    </PivotItem>
                                    <PivotItem
                                        headerText={this.props.localeStrings.planning}
                                        itemCount={this.state.filteredPlanningIncidents.length}
                                        itemKey="Planning"
                                        onRenderItemLink={this._customRenderer}
                                    >
                                        {!this.state.showMapViewer ?
                                            <BootstrapTable
                                                bootstrap4
                                                bordered={false}
                                                keyField="incidentId"
                                                columns={dashboardHeader}
                                                data={this.state.filteredPlanningIncidents}
                                                wrapperClasses={isDarkOrContrastTheme ? "table-darkcontrast" : ""}
                                                headerClasses={isDarkOrContrastTheme ? "table-header-darkcontrast" : ""}
                                                pagination={this.pagination}
                                                noDataIndication={() => (<div id="noincident-planning-tab" aria-live="polite" role="status">{this.props.localeStrings.noIncidentsFound}</div>)}
                                            />
                                            :
                                            <MapViewer graphBaseUrl={this.props.graphBaseUrl} incidentData={this.state.filteredPlanningIncidents} azureMapKey={this.props.azureMapsKeyConfigData} showMessageBar={this.props.showMessageBar} userPrincipalName={this.props.userPrincipalName} localeStrings={this.props.localeStrings} appInsights={this.props.appInsights}></MapViewer>
                                        }
                                    </PivotItem>
                                    <PivotItem
                                        headerText={this.props.localeStrings.active}
                                        itemCount={this.state.filteredActiveIncidents.length}
                                        itemKey="Active"
                                        onRenderItemLink={this._customRenderer}
                                    >
                                        {!this.state.showMapViewer ?
                                            <BootstrapTable
                                                bootstrap4
                                                bordered={false}
                                                keyField="incidentId"
                                                columns={dashboardHeader}
                                                data={this.state.filteredActiveIncidents}
                                                wrapperClasses={isDarkOrContrastTheme ? "table-darkcontrast" : ""}
                                                headerClasses={isDarkOrContrastTheme ? "table-header-darkcontrast" : ""}
                                                pagination={this.pagination}
                                                noDataIndication={() => (<div id="noincident-active-tab" aria-live="polite" role="status">{this.props.localeStrings.noIncidentsFound}</div>)}
                                            />
                                            :
                                            <MapViewer graphBaseUrl={this.props.graphBaseUrl} incidentData={this.state.filteredActiveIncidents} azureMapKey={this.props.azureMapsKeyConfigData} showMessageBar={this.props.showMessageBar} userPrincipalName={this.props.userPrincipalName} localeStrings={this.props.localeStrings} appInsights={this.props.appInsights}></MapViewer>
                                        }
                                    </PivotItem>
                                    <PivotItem
                                        headerText={this.props.localeStrings.closed}
                                        itemCount={this.state.filteredCompletedIncidents.length}
                                        itemKey="Closed"
                                        onRenderItemLink={this._customRenderer}
                                    >
                                        {!this.state.showMapViewer ?
                                            <BootstrapTable
                                                bootstrap4
                                                bordered={false}
                                                keyField="incidentId"
                                                columns={dashboardHeader}
                                                data={this.state.filteredCompletedIncidents}
                                                wrapperClasses={isDarkOrContrastTheme ? "table-darkcontrast" : ""}
                                                headerClasses={isDarkOrContrastTheme ? "table-header-darkcontrast" : ""}

                                                pagination={this.pagination}
                                                noDataIndication={() => (<div id="noincident-completed-tab" aria-live="polite" role="status">{this.props.localeStrings.noIncidentsFound}</div>)}
                                            />
                                            :
                                            <MapViewer graphBaseUrl={this.props.graphBaseUrl} incidentData={this.state.filteredCompletedIncidents} azureMapKey={this.props.azureMapsKeyConfigData} showMessageBar={this.props.showMessageBar} userPrincipalName={this.props.userPrincipalName} localeStrings={this.props.localeStrings} appInsights={this.props.appInsights}></MapViewer>
                                        }
                                    </PivotItem>
                                </Pivot>
                            </div>
                        </div>
                    </div>
                }
            </>
        );
    }
}

export default Dashboard;
