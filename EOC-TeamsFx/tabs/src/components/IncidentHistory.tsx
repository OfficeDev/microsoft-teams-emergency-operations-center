import { Button, ChevronStartIcon, CloseIcon, Dialog } from "@fluentui/react-northstar";
import { CheckboxVisibility, DetailsList, DetailsListLayoutMode, IColumn } from '@fluentui/react/lib/DetailsList';
import { IList, List, ScrollToMode } from '@fluentui/react/lib/List';
import { ApplicationInsights } from '@microsoft/applicationinsights-web';
import { Client } from "@microsoft/microsoft-graph-client";
import React from "react";
import Col from "react-bootstrap/esm/Col";
import Row from "react-bootstrap/esm/Row";
import ReactTable from "react-table";
import withFixedColumns, { ColumnFixed } from "react-table-hoc-fixed-columns";
import "react-table/react-table.css";
import "react-table-hoc-fixed-columns/lib/styles.css";
import CommonService from "../common/CommonService";
import * as constants from "../common/Constants";
import * as graphConfig from '../common/graphConfig';
import siteConfig from '../config/siteConfig.json';
import "../scss/IncidentHistory.module.scss";
import loadable from "@loadable/component";

const IncidentHistoryPDF = loadable(() => import('./IncidentHistoryPDF'));

//Creates table control with fixed coloumns feature using react table control.
const ReactTableFixedColumns = withFixedColumns(ReactTable);

export interface IVersionItem {
    field: string;
    newValue: string;
    oldValue: string;
}
export interface IIncidentHistoryState {
    incidentVersionData: any[];
    showRoles: boolean;
    roleDetails: any;
    versionDetails: any;
    selectedItem: number | undefined;
    seeAllVersions: boolean;
    isListView: boolean;
    gridData: any[];
    isDesktop: boolean;
    showRoleLeads: boolean;
    roleLeadDetails: any;
    versionHistoryPDFData: any
}
export interface IIncidentHistoryProps {
    localeStrings: any;
    onBackClick(showMessageBar: string): void;
    siteId: string;
    graph: Client;
    appInsights: ApplicationInsights;
    userPrincipalName: any;
    showMessageBar(message: string, type: string): void;
    hideMessageBar(): void;
    incidentId: string;
    currentThemeName: string;
}

export default class IncidentHistory extends React.PureComponent<IIncidentHistoryProps, IIncidentHistoryState> {
    private listRef: React.RefObject<IList>;
    private detailsListRef: React.RefObject<HTMLDivElement>;
    private closeIconRef: React.RefObject<any>;
    private itemHeight = constants.itemHeight;
    private numberOfItemsOnPage = constants.numberOfItemsOnPage;

    constructor(props: any) {
        super(props);
        this.listRef = React.createRef();
        this.detailsListRef = React.createRef();
        this.closeIconRef = React.createRef();
        this.state = {
            incidentVersionData: [],
            showRoles: false,
            roleDetails: [],
            versionDetails: {},
            selectedItem: 0,
            seeAllVersions: false,
            isListView: true,
            gridData: [],
            isDesktop: true,
            showRoleLeads: false,
            roleLeadDetails: [],
            versionHistoryPDFData: []
        }

        this.getVersions = this.getVersions.bind(this);
        this.formatVersionData = this.formatVersionData.bind(this);
        this.loadDetails = this.loadDetails.bind(this);
        this.loadRoles = this.loadRoles.bind(this);
        this.hideRoles = this.hideRoles.bind(this);
        this.onRenderCell = this.onRenderCell.bind(this);
    }

    //common service object
    private dataService = new CommonService();

    // set the state object for screen size
    resize = () => this.setState({ isDesktop: window.innerWidth > constants.mobileWidth });

    // before unmounting, remove event listener
    componentWillUnmount() {
        window.removeEventListener("resize", this.resize.bind(this));
    }

    //Component life cycle componentDidMount method.
    public componentDidMount() {
        window.addEventListener("resize", this.resize.bind(this));
        this.resize();
        this.getVersions();
        //set title attribute to pdf download link
        const pdfLink = document.getElementById("incident-history-download-link")?.getElementsByClassName("download-pdf")[0];
        pdfLink?.setAttribute("title", this.props.localeStrings.downloadPDFLinkTooltipContent);
    }

    //Component life cycle componentDidUpdate method.
    public componentDidUpdate(_prevProps: IIncidentHistoryProps, prevState: IIncidentHistoryState) {
        if (prevState.selectedItem !== this.state.selectedItem) {
            this.listRef.current?.scrollToIndex(
                this.state.selectedItem ? this.state.selectedItem : 0,
                (_idx) => this.itemHeight,
                ScrollToMode.top
            );
        }
        //Format version data for pdf
        if (prevState.incidentVersionData !== this.state.incidentVersionData && this.state.incidentVersionData.length > 0) {
            const versionHistoryArray: any = [];
            this.state.incidentVersionData.forEach((versionData: any, index: any) => {
                versionHistoryArray.push(this.formatVersionData(versionData, index, true));
            });
            this.setState({ versionHistoryPDFData: versionHistoryArray });
        }
    }

    //Get Incident Versions
    private async getVersions() {
        try {
            //graph endpoint to get all versions
            let graphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}/lists/${siteConfig.incidentsList}/items/${this.props.incidentId}/versions?$expand=fields
            ($select=StatusLookupId,Status,id,IncidentId,IncidentName,IncidentCommander,Location,StartDateTime,
            Modified,Description,IncidentType,RoleAssignment,Severity,BridgeID,ReasonForUpdate,IncidentStatus,RoleLeads,CloudStorageLink)`;
            const versionsData = await this.dataService.getVersionsData(graphEndpoint, this.props.graph);

            this.setState({
                incidentVersionData: versionsData
            });

            //Format the first version data and display on component load
            this.formatVersionData(versionsData[0], 0);
        }
        catch (error) {
            console.error(
                constants.errorLogPrefix + "IncidentHistory_Get_Versions\n",
                JSON.stringify(error)
            );
            // Log Exception
            this.dataService.trackException(this.props.appInsights, error, constants.componentNames.IncidentHistoryComponent, 'IncidentHistory_Get_Versions', this.props.userPrincipalName);
        }
    }

    //Format version data when on click of each version in the list view.
    private formatVersionData(versionData: any, index: any, forPDf = false) {
        try {
            let diff = require('deep-diff');
            let currentVersionData = versionData;
            let prevVersionData: any;
            if (index !== this.state.incidentVersionData.length - 1) {
                prevVersionData = this.state.incidentVersionData[index + 1];
            } else {
                prevVersionData = {};
            }
            let changes = diff(currentVersionData, prevVersionData);

            let formattedIncidentsData: Array<IVersionItem> = new Array<IVersionItem>();

            changes?.forEach((item: any) => {
                if ((item.path[0] !== constants.modifiedDate &&
                    item.path[0] !== constants.lastModifiedBy) &&
                    item.path[0] !== constants.roleAssignmentsObj &&
                    item.path[0] !== constants.roleLeadsObj &&
                    (item.lhs !== undefined || item.rhs !== undefined)
                ) {

                    //converting camel case string into Pascal case string with space between each word.
                    const pascalCaseText = item.path[0].replace(/(A-Z)/g, " $1").replace(/([A-Z][a-z])/g, " $1");
                    const fieldName = pascalCaseText.charAt(0).toUpperCase() + pascalCaseText.slice(1);
                    formattedIncidentsData.push({
                        field: fieldName,
                        newValue: item.lhs,
                        oldValue: item.rhs
                    });
                }
            });
            if (forPDf) {
                return {
                    modifiedOn: versionData.modifiedDate,
                    modifiedBy: versionData.lastModifiedBy,
                    versionData: formattedIncidentsData
                };
            }
            else {
                this.setState({
                    selectedItem: index,
                    versionDetails: formattedIncidentsData
                });
                return null;
            }
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "IncidentHistory_Format_Version_Data\n",
                JSON.stringify(error)
            );
            // Log Exception
            this.dataService.trackException(this.props.appInsights, error, constants.componentNames.IncidentHistoryComponent, 'IncidentHistory_Format_Version_Data', this.props.userPrincipalName);
        }
    }

    //Format version data for table view. 
    private formatGridData() {
        try {
            let diff = require('deep-diff');
            let dataDifference: any = [];

            for (let i = 0; i < this.state.incidentVersionData.length; i++) {
                let currentVersionData: any;
                let prevVersionData: any;
                if (i !== this.state.incidentVersionData.length - 1) {
                    prevVersionData = this.state.incidentVersionData[i + 1];
                } else {
                    prevVersionData = {};
                }
                currentVersionData = this.state.incidentVersionData[i];
                let changes = diff(currentVersionData, prevVersionData);

                let obj: { [x: string]: any; } = {};
                changes?.forEach((item: { path: (string | number)[]; lhs: any; }) => {
                    if (item.path[0] === constants.roleAssignmentsObj) {
                        obj[constants.roleAssignmentsObj] = currentVersionData[constants.roleAssignmentsObj];
                    }
                    else if (item.path[0] === constants.roleLeadsObj) {
                        obj[constants.roleLeadsObj] = currentVersionData[constants.roleLeadsObj];
                    }
                    else {
                        obj[item.path[0]] = item.lhs
                    }
                });
                if (changes) {
                    //Explicitly adding ModifiedBy to the array since it might have same value in previous version which will not be captured in the difference
                    obj[constants.lastModifiedBy] = currentVersionData[constants.lastModifiedBy];
                    dataDifference.push(obj);
                }
            }
            this.setState({
                gridData: dataDifference
            });
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "IncidentHistory_Format_Grid_Data\n",
                JSON.stringify(error)
            );
            // Log Exception
            this.dataService.trackException(this.props.appInsights, error, constants.componentNames.IncidentHistoryComponent, 'IncidentHistory_Format_Grid_Data', this.props.userPrincipalName);
        }

    }

    //Load version details when on click of version.
    public loadDetails(data: any, index: any) {
        try {
            this.formatVersionData(data, index);
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "IncidentHistory_Load_Details\n",
                JSON.stringify(error)
            );
            // Log Exception
            this.dataService.trackException(this.props.appInsights, error, constants.componentNames.IncidentHistoryComponent, 'IncidentHistory_Load_Details', this.props.userPrincipalName);

        }
    }

    //Load assigned roles of the incident when on click of view roles button in table view.
    private loadRoles(value: any) {
        try {
            this.setState({
                showRoles: true,
                roleDetails: value,
                showRoleLeads: false,
                roleLeadDetails: []
            });

        } catch (error) {
            console.error(
                constants.errorLogPrefix + "IncidentHistory_Load_Roles\n",
                JSON.stringify(error)
            );
            // Log Exception
            this.dataService.trackException(this.props.appInsights, error, constants.componentNames.IncidentHistoryComponent, 'IncidentHistory_Load_Roles', this.props.userPrincipalName);
        }
    }

    //Load assigned role leads of the incident when on click of view role lead button in table view.
    private loadRoleLeads(value: any) {
        try {
            this.setState({
                showRoleLeads: true,
                roleLeadDetails: value,
                showRoles: false,
                roleDetails: [],
            });

        } catch (error) {
            console.error(
                constants.errorLogPrefix + "IncidentHistory_Load_Role_Leads\n",
                JSON.stringify(error)
            );
            // Log Exception
            this.dataService.trackException(this.props.appInsights, error, constants.componentNames.IncidentHistoryComponent, 'IncidentHistory_Load_Role_Leads', this.props.userPrincipalName);
        }
    }

    //Hide roles popup when onclick of cancel/back button.
    private hideRoles() {
        this.setState({
            showRoles: false,
            showRoleLeads: false,
            roleDetails: [],
            roleLeadDetails: []
        })
    }

    //Render each list item in version list.
    private onRenderCell(item: any, index: number | undefined): JSX.Element {
        return (
            <div
                data-is-focusable={true}
                onClick={() => {
                    this.loadDetails(item, index);
                }}
                className={`activity-log-list-item${this.state.selectedItem === index ? " activity-log-list-item-selected" : ""}`}
                title={`${item.modifiedDate}\n${item.lastModifiedBy}`}
            >
                <div className="list-version-date">{item.modifiedDate}</div>
                <div className="list-version-modifiedby">{item.lastModifiedBy}</div>
            </div>
        );
    };

    //Render method
    public render() {

        //Columns for list view
        const listViewColumns: IColumn[] = this.state.isListView ? [
            {
                key: 'field',
                name: this.props.localeStrings.field,
                fieldName: 'field',
                minWidth: constants.listViewFieldMinWidth,
                maxWidth: constants.listViewFieldMaxWidth,
                isResizable: true,
                onRenderHeader: () => <span title={this.props.localeStrings.field}>{this.props.localeStrings.field}</span>,
                onRender: (item: any) => <span title={item.field}>{item.field}</span>
            },
            {
                key: 'newVal',
                name: this.props.localeStrings.new,
                fieldName: 'newValue',
                minWidth: constants.listViewNewMinWidth,
                maxWidth: constants.listViewNewMaxWidth,
                isResizable: true,
                onRenderHeader: () => <span title={this.props.localeStrings.new}>{this.props.localeStrings.new}</span>,
                className: "new-value-cell",
                onRender: (item: any) => <span title={item.newValue}>{item.newValue}</span>
            },
            {
                key: 'oldVal',
                name: this.props.localeStrings.old,
                fieldName: 'oldValue',
                minWidth: constants.listViewOldMinWidth,
                maxWidth: constants.listViewOldMaxWidth,
                isResizable: true,
                onRenderHeader: () => <span title={this.props.localeStrings.old}>{this.props.localeStrings.old}</span>,
                className: "old-value-cell",
                onRender: (item: any) => <span title={item.oldValue}>{item.oldValue}</span>
            }
        ] : [];

        //Columns for table view
        let gridViewColumns: ColumnFixed<any>[] = !this.state.isListView ? [
            {
                fixed: this.state.isDesktop ? "left" : undefined,
                columns: [
                    {
                        Header: () => <div title={this.props.localeStrings.date}>{this.props.localeStrings.date}</div>,
                        accessor: "modifiedDate",
                        Cell: ({ value }: any) => {
                            if (navigator.userAgent.match(/iPhone|Android/i))
                                return (value ? <span tabIndex={0} role="textbox" aria-readonly aria-label={this.props.localeStrings.date}>{value}</span> : "")
                            else
                                return (value ? <span tabIndex={0} title={value}>{value}</span> : "")
                        },
                        width: 210,
                    },
                    {
                        Header: () => <div title={this.props.localeStrings.modifiedBy}>{this.props.localeStrings.modifiedBy}</div>,
                        accessor: "lastModifiedBy",
                        Cell: ({ value }: any) => {
                            if (navigator.userAgent.match(/iPhone|Android/i))
                                return (value ? <span tabIndex={0} role="textbox" aria-readonly aria-label={this.props.localeStrings.modifiedBy}>{value}</span> : "")
                            else
                                return (value ? <span tabIndex={0} title={value}>{value}</span> : "")
                        },
                        width: 200
                    },
                ]
            },
            {
                columns: [
                    {
                        Header: () => <div title={this.props.localeStrings.fieldIncidentName}>{this.props.localeStrings.fieldIncidentName}</div>,
                        accessor: "incidentName",
                        Cell: ({ value }: any) => {
                            if (navigator.userAgent.match(/iPhone|Android/i))
                                return (value ? <span tabIndex={0} role="textbox" aria-readonly aria-label={this.props.localeStrings.fieldIncidentName}>{value}</span> : "")
                            else
                                return (value ? <span tabIndex={0} title={value}>{value}</span> : "")
                        },
                        width: 200,
                    },
                    {
                        Header: () => <div title={this.props.localeStrings.fieldIncidentStatus}>{this.props.localeStrings.fieldIncidentStatus}</div>,
                        accessor: "status",
                        Cell: ({ value }: any) => {
                            if (navigator.userAgent.match(/iPhone|Android/i))
                                return (value ? <span tabIndex={0} role="textbox" aria-readonly aria-label={this.props.localeStrings.fieldIncidentStatus}>{value}</span> : "")
                            else
                                return (value ? <span tabIndex={0} title={value}>{value}</span> : "")
                        },
                        width: 200,
                    },
                    {
                        Header: () => <div title={this.props.localeStrings.fieldSeverity}>{this.props.localeStrings.fieldSeverity}</div>,
                        accessor: "severity",
                        Cell: ({ value }: any) => {
                            if (navigator.userAgent.match(/iPhone|Android/i))
                                return (value ? <span tabIndex={0} role="textbox" aria-readonly aria-label={this.props.localeStrings.fieldSeverity}>{value}</span> : "")
                            else
                                return (value ? <span tabIndex={0} title={value}>{value}</span> : "")
                        },
                        width: 150,
                    },
                    {
                        Header: () => <div title={this.props.localeStrings.fieldLocation}>{this.props.localeStrings.fieldLocation}</div>,
                        accessor: "location",
                        Cell: ({ value }: any) => {
                            if (navigator.userAgent.match(/iPhone|Android/i))
                                return (value ? <span tabIndex={0} role="textbox" aria-readonly aria-label={this.props.localeStrings.fieldLocation}>{value}</span> : "")
                            else
                                return (value ? <span tabIndex={0} title={value}>{value}</span> : "")
                        },
                        width: 200,
                    },
                    {
                        Header: () => <div title={this.props.localeStrings.fieldIncidentCommander}>{this.props.localeStrings.fieldIncidentCommander}</div>,
                        accessor: "incidentCommander",
                        Cell: ({ value }: any) => {
                            if (navigator.userAgent.match(/iPhone|Android/i))
                                return (value ? <span tabIndex={0} role="textbox" aria-readonly aria-label={this.props.localeStrings.fieldIncidentCommander}>{value}</span> : "")
                            else
                                return (value ? <span tabIndex={0} title={value}>{value}</span> : "")
                        },
                        width: 200,
                    },
                    {
                        Header: () => <div title={this.props.localeStrings.roles}>{this.props.localeStrings.roles}</div>,
                        accessor: "roleAssignmentsObj",
                        Cell: ({ value }: any) => value ?
                            <Button
                                tabIndex={0}
                                onClick={() => this.loadRoles(value)}
                                title={this.props.localeStrings.viewLabel}
                                aria-label={`${this.props.localeStrings.roles} ${this.props.localeStrings.viewLabel}`}
                                text
                                className="grid-view-assigned-roles"
                            >
                                {this.props.localeStrings.viewLabel}
                            </Button> : "",
                        width: 100,
                    },
                    {
                        Header: () => <div title={this.props.localeStrings.roleLeadsLabel}>{this.props.localeStrings.roleLeadsLabel}</div>,
                        accessor: "roleLeadsObj",
                        Cell: ({ value }: any) => value ?
                            <Button
                                tabIndex={0}
                                onClick={() => this.loadRoleLeads(value)}
                                title={this.props.localeStrings.viewLabel}
                                aria-label={`${this.props.localeStrings.roleLeadsLabel} ${this.props.localeStrings.viewLabel}`}
                                text
                                className="grid-view-assigned-roles"
                            >
                                {this.props.localeStrings.viewLabel}
                            </Button> : "",
                        width: 120,
                    },
                    {
                        Header: () => <div title={this.props.localeStrings.fieldDescription}>{this.props.localeStrings.fieldDescription}</div>,
                        accessor: "incidentDescription",
                        Cell: ({ value }: any) => {
                            if (navigator.userAgent.match(/iPhone|Android/i))
                                return (value ? <span tabIndex={0} role="textbox" aria-readonly aria-label={this.props.localeStrings.fieldDescription}>{value}</span> : "")
                            else
                                return (value ? <span tabIndex={0} title={value}>{value}</span> : "")
                        },
                        width: 200
                    },
                    {
                        Header: () => <div title={this.props.localeStrings.fieldReasonForUpdate}>{this.props.localeStrings.fieldReasonForUpdate}</div>,
                        accessor: "reasonForUpdate",
                        Cell: ({ value }: any) => {
                            if (navigator.userAgent.match(/iPhone|Android/i))
                                return (value ? <span tabIndex={0} role="textbox" aria-readonly aria-label={this.props.localeStrings.fieldReasonForUpdate}>{value}</span> : "")
                            else
                                return (value ? <span tabIndex={0} title={value}>{value}</span> : "")
                        },
                        width: 200
                    }, {
                        Header: () => <div title={this.props.localeStrings.cloudStorageFieldLabel}>{this.props.localeStrings.cloudStorageFieldLabel}</div>,
                        accessor: "cloudStorageLink",
                        Cell: ({ value }: any) => {
                            if (navigator.userAgent.match(/iPhone|Android/i))
                                return (value ? <span tabIndex={0} role="textbox" aria-readonly aria-label={this.props.localeStrings.cloudStorageFieldLabel}>{value}</span> : "")
                            else
                                return (value ? <span tabIndex={0} title={value}>{value}</span> : "")
                        },
                        width: 200
                    },
                    {
                        Header: () => <div title={this.props.localeStrings.fieldBridgeID}>{this.props.localeStrings.fieldBridgeID}</div>,
                        accessor: "bridgeID",
                        Cell: ({ value }: any) => {
                            if (navigator.userAgent.match(/iPhone|Android/i))
                                return (value ? <span tabIndex={0} role="textbox" aria-readonly aria-label={this.props.localeStrings.fieldBridgeID}>{value}</span> : "")
                            else
                                return (value ? <span tabIndex={0} title={value}>{value}</span> : "")
                        },
                        width: 200
                    }
                ]
            }
        ] : [];

        const isDarkOrContrastTheme = this.props.currentThemeName === constants.darkMode || this.props.currentThemeName === constants.contrastMode;

        return (
            <div className="incident-history">
                <div className=".col-xs-12 .col-sm-8 .col-md-4 container" id="incident-history-path">
                    <label>
                        <span
                            onClick={() => this.props.onBackClick("")}
                            onKeyDown={(event) => {
                                if (event.key === constants.enterKey)
                                    this.props.onBackClick("")
                            }}
                            className="go-back">
                            <ChevronStartIcon id="path-back-icon" />
                            <span className="back-label" role="button" tabIndex={0} title={this.props.localeStrings.back}>{this.props.localeStrings.back}</span>
                        </span> &nbsp;&nbsp;
                        <span className="right-border">|</span>
                        <span title={this.props.localeStrings.incidentHistory}>&nbsp;&nbsp;{this.props.localeStrings.incidentHistory}</span>
                    </label>
                </div>
                <div className={`incident-history-area${isDarkOrContrastTheme ? " incident-history-area-darkcontrast" : ""}`}>
                    <div className="container">
                        <div className="heading-and-view-selection-area">
                            <h1 aria-live="polite" role="alert"> <div className="incident-history-label">{this.props.localeStrings.incidentHistory} - {this.props.incidentId}</div></h1>
                            <div className="view-selection-area">
                                <label htmlFor="list-view-select" className="flip-view" title={this.props.localeStrings.listView}>
                                    <input
                                        type="radio"
                                        name="select view"
                                        id="list-view-select"
                                        onChange={() => this.setState({ isListView: !this.state.isListView })}
                                        checked={this.state.isListView}
                                        aria-label={this.props.localeStrings.listView}
                                    />
                                    <img
                                        src={require("../assets/Images/ListViewIcon.svg").default}
                                        alt={this.props.localeStrings.listView}
                                        className={`view-icons${isDarkOrContrastTheme ? " view-icons-darkcontrast" : ""}`}
                                    />

                                    <span>{this.props.localeStrings.listView}</span>
                                </label>
                                <label htmlFor="table-view-select" className="flip-view" title={this.props.localeStrings.tableView}>
                                    <input
                                        type="radio"
                                        name="select view"
                                        id="table-view-select"
                                        onChange={() => {
                                            this.setState({ isListView: !this.state.isListView });
                                            this.state.gridData.length === 0 && this.formatGridData()
                                        }}
                                        checked={!this.state.isListView}
                                        aria-label={this.props.localeStrings.tableView}
                                    />
                                    <img
                                        src={require("../assets/Images/TableViewIcon.svg").default}
                                        alt={this.props.localeStrings.tableView}
                                        className={`view-icons${isDarkOrContrastTheme ? " view-icons-darkcontrast" : ""}`}
                                    />

                                    <span>{this.props.localeStrings.tableView}</span>
                                </label>
                                <div id="incident-history-download-link">
                                    <IncidentHistoryPDF localeStrings={this.props.localeStrings} incidentId={this.props.incidentId} versionHistoryPDFData={this.state.versionHistoryPDFData} currentThemeName={this.props.currentThemeName} incidentVersionData={this.state.incidentVersionData} />
                                </div>
                            </div>
                        </div>
                        {this.state.isListView ?
                            <div className='activity-version-details-area'>
                                <div className='activity-log-area'>
                                    <div className="activity-log-heading" title={this.props.localeStrings.activityLog}>{this.props.localeStrings.activityLog}</div>
                                    <div className='activity-log-list-main-area' data-is-scrollable>
                                        {this.state.incidentVersionData !== undefined &&
                                            <List
                                                componentRef={this.listRef}
                                                items={this.state.incidentVersionData.slice(0, this.state.seeAllVersions ? this.state.incidentVersionData.length : constants.listViewItemInitialCount)}
                                                onRenderCell={this.onRenderCell}
                                                className="activity-log-list-main"
                                                getPageHeight={(idx) => this.dataService.getPageHeight(idx, this.itemHeight, this.numberOfItemsOnPage)}
                                                version={this.state.selectedItem}
                                            />
                                        }
                                        {!this.state.seeAllVersions && this.state.incidentVersionData.length > constants.listViewItemInitialCount &&
                                            <div
                                                onClick={() => this.setState({ seeAllVersions: true })}
                                                className="see-all-versions"
                                                title={this.props.localeStrings.seeAll}
                                            >
                                                {this.props.localeStrings.seeAll}
                                            </div>
                                        }
                                    </div>
                                </div>
                                <div className='version-details-area' ref={this.detailsListRef}>
                                    {this.state.versionDetails.length > 0 ?
                                        <DetailsList
                                            items={this.state.versionDetails}
                                            columns={listViewColumns}
                                            layoutMode={DetailsListLayoutMode.justified}
                                            checkboxVisibility={CheckboxVisibility.hidden}
                                            onDidUpdate={() => this.detailsListRef.current?.getElementsByClassName("ms-DetailsList-headerWrapper")[0]
                                                ?.getElementsByClassName("ms-DetailsHeader")[0]?.setAttribute("aria-busy", "true")}
                                        />
                                        :
                                        <>
                                            {this.state.versionDetails.length !== undefined ?
                                                <div className="noDataFound" title={this.props.localeStrings.noVersionChangesLabel}>
                                                    {this.props.localeStrings.noVersionChangesLabel}
                                                </div>
                                                :
                                                <div className="noDataFound" title={this.props.localeStrings.loadingLabel}>
                                                    {this.props.localeStrings.loadingLabel}
                                                </div>
                                            }
                                        </>
                                    }
                                </div>
                            </div>
                            :
                            <div>
                                <ReactTableFixedColumns
                                    data={this.state.gridData}
                                    columns={gridViewColumns}
                                    defaultPageSize={this.state.gridData.length}
                                    className="grid-view-table"
                                    showPagination={false}
                                    sortable={false}
                                />
                                {(this.state.showRoles || this.state.showRoleLeads) ?
                                    <Dialog
                                        header={this.state.showRoles ? this.props.localeStrings.roles : this.props.localeStrings.roleLeadsLabel}
                                        headerAction={{
                                            ref: this.closeIconRef,
                                            icon: <CloseIcon onClick={() => this.hideRoles()} />,
                                            title: this.props.localeStrings.btnClose,
                                            onKeyDown: (event: any) => {
                                                if (event.shiftKey || event.key === constants.tabKey) {
                                                    this.closeIconRef.current?.focus();
                                                }
                                                if (event.key === constants.enterKey || event.keyCode === 32) {
                                                    this.hideRoles();
                                                }
                                            }
                                        }}
                                        content={
                                            <div className="role-assignment-table">
                                                <Row id="role-grid-thead" xs={2} sm={2} md={2}>
                                                    <Col md={6} sm={6} xs={6} >{this.props.localeStrings.headerRole}</Col>
                                                    <Col md={6} sm={6} xs={6} className="thead-border-left">
                                                        {this.state.showRoles ? this.props.localeStrings.headerUsers : this.props.localeStrings.leadLabel}
                                                    </Col>
                                                </Row>
                                                <div className="role-grid-tbody-area">
                                                    {(this.state.showRoles ? this.state.roleDetails : this.state.roleLeadDetails).map((item: any, index: any) => (
                                                        <Row xs={2} sm={2} md={2} key={index} id="role-grid-tbody">
                                                            <Col md={6} sm={6} xs={6}>{item.Role}</Col>
                                                            <Col md={6} sm={6} xs={6}>{item.Users}</Col>
                                                        </Row>
                                                    )
                                                    )}
                                                </div>
                                            </div>
                                        }
                                        cancelButton={{
                                            icon: <CloseIcon bordered circular size="smallest" className="roles-popup-btn-close-icon" />,
                                            title: this.props.localeStrings.btnClose,
                                            iconPosition: 'before',
                                            content: this.props.localeStrings.btnClose,
                                            className: "roles-popup-btn-close",
                                            onKeyDown: (event: any) => {
                                                if (event.shiftKey || event.key === constants.tabKey) {
                                                    this.closeIconRef.current?.focus();
                                                }
                                            }
                                        }}
                                        onCancel={(e) => this.hideRoles()}
                                        open={this.state.showRoles || this.state.showRoleLeads}
                                        className={`view-roles-popup${this.props.currentThemeName === constants.darkMode ? " view-roles-popup-dark" : this.props.currentThemeName === constants.contrastMode ? " view-roles-popup-contrast" : ""}`}
                                    />
                                    : null}
                            </div>
                        }
                    </div>
                </div>
            </div>
        );
    }
}