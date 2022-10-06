import { Callout, DirectionalHint, Icon, IPivotItemProps, Pivot, PivotItem } from '@fluentui/react';
import { Button, Flex, FormInput, Loader, SearchIcon } from "@fluentui/react-northstar";
import { ApplicationInsights } from '@microsoft/applicationinsights-web';
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

export interface IDashboardProps {
    graph: Client;
    tenantName: string;
    siteId: string;
    onCreateTeamClick: Function;
    onEditButtonClick(incidentData: any): void;
    localeStrings: any;
    onBackClick(showMessageBar: boolean): void;
    showMessageBar(message: string, type: string): void;
    hideMessageBar(): void;
    appInsights: ApplicationInsights;
    userPrincipalName: any;
    siteName: any;
    onShowTeamNameConfigForm: Function;
    onShowIncidentHistory: Function;
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
    constructor(props: IDashboardProps) {
        super(props);

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
            isManageCalloutVisible: false
        };

        this.onActionClick = this.onActionClick.bind(this);
    }

    private dataService = new CommonService();

    // set the state object for screen size
    resize = () => this.setState({ isDesktop: window.innerWidth > constants.mobileWidth });

    // before unmounting, remove event listener
    componentWillUnmount() {
        window.removeEventListener("resize", this.resize.bind(this));
    }

    // On Click of edit or view incident history icons this method gets called.
    public onActionClick(_cell: any, gridRow: any, _rowIndex: any, _formatExtraData: any) {
        return (
            <span>
                {/* bind edit icon to dashboard if status is not 'Completed' */}
                <img
                    src={require("../assets/Images/GridEditIcon.svg").default}
                    alt={this.props.localeStrings.edit}
                    className="grid-edit-icon"
                    title={this.props.localeStrings.edit}
                    onClick={() => this.props.onEditButtonClick(gridRow)}
                />

                {/* view version history of each incident in dashboard when onclick of history icon */}
                <img
                    src={require("../assets/Images/IncidentHistoryIcon.svg").default}
                    alt={this.props.localeStrings.viewIncidentHistory}
                    title={this.props.localeStrings.viewIncidentHistory}
                    className="grid-version-history-icon"
                    onClick={() => this.props.onShowIncidentHistory(gridRow.incidentId)}
                />
            </span>
        );
    }

    // component life cycle method
    public async componentDidMount() {
        window.addEventListener("resize", this.resize.bind(this));
        this.resize();
        // Get dashboard data
        this.getDashboardData();
    }

    // connect with servie layer to get all incidents
    getDashboardData = async () => {

        this.setState({
            showLoader: true
        })

        try {
            // create graph endpoint for querying Incident Transaction list
            let graphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}/lists/${siteConfig.incidentsList}/items?$expand=fields&$Top=5000`;

            const allIncidents = this.sortDashboardData(await this.dataService.getDashboardData(graphEndpoint, this.props.graph));
            console.log(constants.infoLogPrefix + "All Incidents retrieved");

            // filter for Planning tab
            const planningIncidents = allIncidents.filter((e: any) => e.status === constants.planning);

            // filter for Active tab
            const activeIncidents = allIncidents.filter((e: any) => e.status === constants.active);

            // filter for Completed tab
            const completedIncidents = allIncidents.filter((e: any) => e.status === constants.closed);

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
            })
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

    // bind status icon to dashboard
    statusIcon = (cell: any, row: any, rowIndex: any, formatExtraData: any) => {
        if (row.status === constants.closed) {
            return (
                <img src={require("../assets/Images/ClosedIcon.svg").default} alt={this.props.localeStrings.closed} className="status-icon" />
            );
        }
        if (row.status === constants.active) {
            return (
                <img src={require("../assets/Images/ActiveIcon.svg").default} alt={this.props.localeStrings.active} className="status-icon" />
            );
        }
        if (row.status === constants.planning) {
            return (
                <img src={require("../assets/Images/PlanningIcon.svg").default} alt={this.props.localeStrings.planning} className="status-icon" />
            );
        }
    };

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

    //Pagination
    pagination = paginationFactory({
        page: 1,
        sizePerPage: constants.dashboardPageSize,
        lastPageText: '>>',
        firstPageText: '<<',
        nextPageText: '>',
        prePageText: '<',
        showTotal: true,
        alwaysShowAllBtns: false
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
                (incident["location"] && incident["location"].toLowerCase().indexOf(searchKeyword) > -1))
        });

        //On Click of Cancel icon
        if (searchText.cancelable) {
            filteredAllIncidents = this.state.allIncidents;
        }
        this.setState({
            searchText: searchText.target.value,
            filteredAllIncidents: filteredAllIncidents,
            filteredPlanningIncidents: filteredAllIncidents.filter((e: any) => e.status === constants.planning),
            filteredActiveIncidents: filteredAllIncidents.filter((e: any) => e.status === constants.active),
            filteredCompletedIncidents: filteredAllIncidents.filter((e: any) => e.status === constants.closed),
        });
    }

    // bind on click event to incident id
    teamsDeepLink = (cell: any, gridRow: any, rowIndex: any, formatExtraData: any) => {
        return (
            <span className="incIdDeepLink" onClick={() => this.onDeepLinkClick(gridRow)}>{gridRow.incidentId}</span>
        );
    }

    // create deep link to open the associated Team
    onDeepLinkClick = (rowData: any) => {
        microsoftTeams.executeDeepLink(rowData.teamWebURL);
    }

    public render() {
        // Header object for dashboard
        const dashboardHeader = [
            {
                dataField: 'incidentId',
                text: this.props.localeStrings.incidentId,
                sort: true,
                formatter: this.teamsDeepLink,
                headerTitle: true,
                title: true,
            }, {
                dataField: 'incidentName',
                text: this.props.localeStrings.incidentName,
                sort: true,
                headerTitle: true,
                title: true
            }, {
                dataField: 'severity',
                text: this.props.localeStrings.fieldSeverity,
                headerTitle: true,
                title: true
            }, {
                dataField: 'incidentCommander',
                text: this.props.localeStrings.incidentCommander,
                headerTitle: true,
                title: true
            }, {
                dataField: 'status',
                text: this.props.localeStrings.status,
                formatter: this.statusIcon,
                headerTitle: true,
                title: true
            }, {
                dataField: 'location',
                text: this.props.localeStrings.location,
                sort: true,
                headerTitle: true,
                title: true
            }, {
                dataField: 'startDate',
                text: this.props.localeStrings.startDate,
                headerTitle: true,
                title: true
            }, {
                dataField: 'action',
                text: this.props.localeStrings.action,
                formatter: this.onActionClick,
                headerTitle: true,
                title: true
            }
        ]
        // added constants for manage dropdown menu
        const manageButtonId = 'manage-callout-button';
        const manageLabelId = 'manage-callout-label';
        const manageDescriptionId = 'manage-callout-description';
        return (
            <>
                {this.state.showLoader ?
                    <>
                        <Loader label={this.state.loaderMessage} size="largest" />
                    </>
                    :
                    <div>
                        <div id="dashboard-search-btn-area">
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
                                        />
                                    </div>
                                    <Flex space="between" wrap={true}>
                                        <div
                                            className={`manage-links${this.state.isManageCalloutVisible ? " callout-visible" : ""}`}
                                            onClick={() => this.setState({ isManageCalloutVisible: !this.state.isManageCalloutVisible })}
                                            id={manageButtonId}
                                        >
                                            <img
                                                src={require("../assets/Images/ManageIcon.svg").default}
                                                className="manage-icon"
                                                alt="Manage Icon"
                                            />
                                            <img
                                                src={require("../assets/Images/ManageIconActive.svg").default}
                                                className='manage-icon-active'
                                                alt="Manage Active Icon"
                                            />
                                            <div className='manage-label'>{this.props.localeStrings.manageLabel}</div>
                                            {this.state.isManageCalloutVisible ?
                                                <Icon iconName="ChevronUp" />
                                                :
                                                <Icon iconName="ChevronDown" />
                                            }
                                        </div>

                                        <Button
                                            primary id="create-incident-btn"
                                            fluid={true}
                                            onClick={() => this.props.onCreateTeamClick()}
                                            title={this.props.localeStrings.btnCreateIncident}
                                        >
                                            <img src={require("../assets/Images/ButtonEditIcon.svg").default} alt="edit icon" />
                                            {this.props.localeStrings.btnCreateIncident}
                                        </Button>
                                    </Flex>
                                </Flex>
                                {this.state.isManageCalloutVisible ? (
                                    <Callout
                                        ariaLabelledBy={manageLabelId}
                                        ariaDescribedBy={manageDescriptionId}
                                        role="menu"
                                        className="manage-links-callout"
                                        gapSpace={10}
                                        target={`#${manageButtonId}`}
                                        isBeakVisible={false}
                                        onDismiss={() => this.setState({ isManageCalloutVisible: false })}
                                        directionalHint={DirectionalHint.bottomLeftEdge}
                                    >
                                        <div>
                                            <div className="dashboard-link">
                                                <a title={this.props.localeStrings.manageIncidentTypesTooltip} href={`https://${this.props.tenantName}/sites/${this.props.siteName}/lists/${siteConfig.incTypeList}`} target='_blank' rel="noreferrer">
                                                    <img src={require("../assets/Images/Manage Incident Types.svg").default} alt={this.props.localeStrings.manageIncidentTypesTooltip} />
                                                    <span className="manage-callout-text">{this.props.localeStrings.incidentTypesLabel}</span>
                                                </a>
                                            </div>
                                            <div className="dashboard-link">
                                                <a title={this.props.localeStrings.manageRolesTooltip} href={`https://${this.props.tenantName}/sites/${this.props.siteName}/lists/${siteConfig.roleAssignmentList}`} target='_blank' rel="noreferrer">
                                                    <img src={require("../assets/Images/Manage Roles.svg").default} alt={this.props.localeStrings.manageRolesTooltip} />
                                                    <span className="manage-callout-text">{this.props.localeStrings.roles}</span>
                                                </a>
                                            </div>
                                            <div
                                                className="dashboard-link"
                                                title={this.props.localeStrings.manageTeamNameTooltip}
                                                onClick={() => this.props.onShowTeamNameConfigForm()}
                                            >
                                                <span className="team-name-link">
                                                    <img src={require("../assets/Images/TeamNameIcon.svg").default} alt={this.props.localeStrings.manageTeamNameTooltip} />
                                                    <span className="manage-callout-text">{this.props.localeStrings.teamNameLabel}</span>
                                                </span>
                                            </div>
                                        </div>
                                    </Callout>
                                ) : null}
                            </div>
                        </div>
                        <div id="dashboard-pivot-container">
                            <div className="container">
                                <div className="grid-heading">{this.props.localeStrings.incidentDetails}</div>
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
                                    id="piv-tabs"
                                >
                                    <PivotItem
                                        headerText={this.props.localeStrings.all}
                                        itemCount={this.state.filteredAllIncidents.length}
                                        itemKey="All"
                                        onRenderItemLink={this._customRenderer}
                                    >
                                        <BootstrapTable
                                            striped
                                            bootstrap4
                                            bordered={false}
                                            keyField="incidentId"
                                            columns={dashboardHeader}
                                            data={this.state.filteredAllIncidents}
                                            pagination={this.pagination}
                                            noDataIndication={() => (<div>{this.props.localeStrings.noIncidentsFound}</div>)}
                                        />
                                    </PivotItem>
                                    <PivotItem
                                        headerText={this.props.localeStrings.planning}
                                        itemCount={this.state.filteredPlanningIncidents.length}
                                        itemKey="Planning"
                                        onRenderItemLink={this._customRenderer}
                                    >
                                        <BootstrapTable
                                            striped
                                            bootstrap4
                                            bordered={false}
                                            keyField="incidentId"
                                            columns={dashboardHeader}
                                            data={this.state.filteredPlanningIncidents}
                                            pagination={this.pagination}
                                            noDataIndication={() => (<div>{this.props.localeStrings.noIncidentsFound}</div>)}
                                        />
                                    </PivotItem>
                                    <PivotItem
                                        headerText={this.props.localeStrings.active}
                                        itemCount={this.state.filteredActiveIncidents.length}
                                        itemKey="Active"
                                        onRenderItemLink={this._customRenderer}
                                    >
                                        <BootstrapTable
                                            striped
                                            bootstrap4
                                            bordered={false}
                                            keyField="incidentId"
                                            columns={dashboardHeader}
                                            data={this.state.filteredActiveIncidents}
                                            pagination={this.pagination}
                                            noDataIndication={() => (<div>{this.props.localeStrings.noIncidentsFound}</div>)}
                                        />
                                    </PivotItem>
                                    <PivotItem
                                        headerText={this.props.localeStrings.closed}
                                        itemCount={this.state.filteredCompletedIncidents.length}
                                        itemKey="Closed"
                                        onRenderItemLink={this._customRenderer}
                                    >
                                        <BootstrapTable
                                            striped
                                            bootstrap4
                                            bordered={false}
                                            keyField="incidentId"
                                            columns={dashboardHeader}
                                            data={this.state.filteredCompletedIncidents}
                                            pagination={this.pagination}
                                            noDataIndication={() => (<div>{this.props.localeStrings.noIncidentsFound}</div>)}
                                        />
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
