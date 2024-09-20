import { ComboBox, DatePicker, IComboBox, TimePicker } from "@fluentui/react";
import { Checkbox as Fluent9CheckBox, Combobox as Fluent9Combobox, Option } from "@fluentui/react-components";
import { ChevronDown16Regular, ChevronRight16Regular, CloudLink24Regular, Delete24Regular, Dismiss24Regular, PeopleCheckmark24Regular, PeopleEdit24Regular, Save24Regular } from "@fluentui/react-icons";
import { AddIcon } from '@fluentui/react-icons-northstar';
import {
    Button, ChevronStartIcon, Flex,
    FormDropdown, FormInput, FormTextArea, Loader
} from "@fluentui/react-northstar";
import { Checkbox } from '@fluentui/react/lib/Checkbox';
import { Icon } from "@fluentui/react/lib/Icon";
import { Toggle } from '@fluentui/react/lib/Toggle';
import { TooltipHost } from "@fluentui/react/lib/Tooltip";
import { ApplicationInsights } from '@microsoft/applicationinsights-web';
import { LocalizationHelper, PeoplePicker, PersonType, UserType } from '@microsoft/mgt-react';
import { Client } from "@microsoft/microsoft-graph-client";
import 'bootstrap/dist/css/bootstrap.min.css';
import moment from "moment";
import * as React from "react";
import Col from "react-bootstrap/Col";
import Container from 'react-bootstrap/Container';
import Row from "react-bootstrap/Row";
import { renderToStaticMarkup } from 'react-dom/server';
import ReactSlider from 'react-slider';
import { v4 as uuidv4 } from "uuid";
import CommonService, { IListItem } from "../common/CommonService";
import * as constants from '../common/Constants';
import * as graphConfig from '../common/graphConfig';
import siteConfig from '../config/siteConfig.json';
import '../scss/IncidentDetails.module.scss';
import { TeamsFxContext } from "./Context";
import {
    ChannelCreationResult, ChannelCreationStatus,
    IAdditionalTeamChannels, IGuestUsers, IIncidentStatus,
    IInputRegexValidationStates,
    IInputValidationStates, ITeamChannel,
    IncidentEntity,
    RoleAssignments,
    UserDetails
} from "./ICreateIncident";
import { ILocationPickerItem } from "./ILocationPicker";
import { LocationPicker } from "./LocationPicker";

const calloutProps = { gapSpace: 0 };

export interface IIncidentDetailsProps {
    graph: Client;
    graphBaseUrl: any;
    graphContextURL: string;
    siteId: string;
    onBackClick(showMessageBar: string): void;
    showMessageBar(message: string, type: string): void;
    hideMessageBar(): void;
    localeStrings: any;
    currentUserId: string;
    incidentData?: IListItem;
    existingTeamMembers?: any
    isEditMode?: boolean;
    appInsights: ApplicationInsights;
    userPrincipalName: any;
    tenantID: any;
    currentThemeName: string;
    appSettings: any;
    editIncidentAccessRole: string;
}

export interface IIncidentDetailsState {
    dropdownOptions: any;
    incDetailsItem: IncidentEntity;
    newRoleString: string;
    roleAssignments: RoleAssignments[];
    showLoader: boolean;
    loaderMessage: string;
    inputValidation: IInputValidationStates;
    inputRegexValidation: IInputRegexValidationStates;
    isCreateNewRoleBtnDisabled: boolean;
    isAddRoleAssignmentBtnDisabled: boolean;
    isDesktop: boolean;
    formOpacity: number;
    eocAppId: string;
    selectedUsers: any;
    incidentTypeSearchQuery: string;
    selectedIncidentCommander: any;
    isRoleInEditMode: boolean[];
    selectedUsersInEditMode: any;
    teamGroupId: string;
    existingRolesMembers: any;
    existingIncCommander: any;
    isOwner: boolean;
    showNoAccessMessage: boolean;
    teamNameConfigArray: any[];
    prefixValue: string;
    selectedSeverity: number;
    roleDefaultData: any[];
    incidentTypeRoleDefaultData: any[];
    saveDefaultRoleCheck: any;
    saveIncidentTypeDefaultRoleCheck: any;
    isEditMode: boolean;
    graphContextURL: string;
    selectedLead: any;
    selectedLeadInEditMode: any;
    toggleCloudStorageLocation: boolean;
    toggleAdditionalChannels: boolean;
    saveDefaultAdditionalChannels: boolean;
    saveDefaultCloudStorageLink: boolean;
    toggleGuestUsers: boolean;
    incCommanderHasRegexError: boolean;
    secIncCommanderUserHasRegexError: boolean;
    secIncCommanderLeadHasRegexError: boolean;
    secIncCommanderUserInEditModeHasRegexError: boolean;
    secIncCommanderLeadInEditModeHasRegexError: boolean;
    roleAddSuccessMessage: string;
    selectedLocation: any;
}

// sets the initial values for required fields validation object
const getInputValidationInitialState = (): IInputValidationStates => {
    return {
        incidentNameHasError: false,
        incidentStatusHasError: false,
        incidentLocationHasError: false,
        incidentTypeHasError: false,
        incidentDescriptionHasError: false,
        incidentStartDateTimeHasError: false,
        incidentCommandarHasError: false,
        incidentReasonForUpdateHasError: false,
        cloudStorageLinkHasError: false,
        guestUsersHasError: false
    };
};

class IncidentDetails extends React.PureComponent<IIncidentDetailsProps, IIncidentDetailsState> {

    //ref variables declaration to create unique reference for DOM element and after validation focus set. 
    private incidentName: any;
    private incidentLocation: any;
    private incidentType: any;
    private incidentDescription: any;
    private incidentStartDateTime: any;
    private incidentCommandar: any;
    private searchUser: any;
    private searchUserEditMode: any;

    private incCommanderRef: React.RefObject<any>;
    private normalSearchUserRef: React.RefObject<any>;
    private normalSearchLeadRef: React.RefObject<any>;
    private incTypeRef: any;

    constructor(props: IIncidentDetailsProps) {
        super(props);
        this.incCommanderRef = React.createRef();
        this.normalSearchUserRef = React.createRef();
        this.normalSearchLeadRef = React.createRef();
        this.incTypeRef = React.createRef();
        this.state = {
            dropdownOptions: '',
            incDetailsItem: new IncidentEntity(),
            newRoleString: '',
            roleAssignments: [],
            showLoader: true,
            loaderMessage: this.props.localeStrings.genericLoaderMessage,
            inputValidation: getInputValidationInitialState(),
            inputRegexValidation: this.dataService.getInputRegexValidationInitialState(),
            isCreateNewRoleBtnDisabled: true,
            isAddRoleAssignmentBtnDisabled: true,
            isDesktop: true,
            formOpacity: 0.5,
            eocAppId: "",
            selectedUsers: [],
            incidentTypeSearchQuery: "",
            selectedIncidentCommander: [],
            isRoleInEditMode: [],
            selectedUsersInEditMode: [],
            teamGroupId: "",
            existingRolesMembers: [],
            existingIncCommander: [],
            isOwner: false,
            showNoAccessMessage: false,
            teamNameConfigArray: [],
            prefixValue: "",
            selectedSeverity: 0,
            roleDefaultData: [],
            incidentTypeRoleDefaultData: [],
            saveDefaultRoleCheck: false,
            saveIncidentTypeDefaultRoleCheck: false,
            isEditMode: false,
            graphContextURL: this.props.graphContextURL,
            selectedLead: [],
            selectedLeadInEditMode: [],
            toggleCloudStorageLocation: false,
            toggleAdditionalChannels: false,
            saveDefaultAdditionalChannels: false,
            saveDefaultCloudStorageLink: false,
            toggleGuestUsers: false,
            incCommanderHasRegexError: false,
            secIncCommanderUserHasRegexError: false,
            secIncCommanderLeadHasRegexError: false,
            secIncCommanderUserInEditModeHasRegexError: false,
            secIncCommanderLeadInEditModeHasRegexError: false,
            roleAddSuccessMessage: "",
            selectedLocation: ""
        };
        this.onRoleChange = this.onRoleChange.bind(this);
        this.onTextInputChange = this.onTextInputChange.bind(this);
        this.handleIncCommanderChange = this.handleIncCommanderChange.bind(this);
        this.onAddNewRoleChange = this.onAddNewRoleChange.bind(this);
        this.onIncidentTypeChange = this.onIncidentTypeChange.bind(this);
        this.onIncidentStatusChange = this.onIncidentStatusChange.bind(this);

        // localized messages for people pickers
        LocalizationHelper.strings = {
            _components: {
                'people-picker': {
                    noResultsFound: this.props.localeStrings.peoplePickerNoResult,
                    loadingMessage: this.props.localeStrings.peoplePickerLoader
                }
            }
        }
        // initialize ref object to assign unique reference for DOM element.
        // to set focus after validation
        this.incidentName = React.createRef();
        this.incidentCommandar = React.createRef();
        this.incidentDescription = React.createRef();
        this.incidentType = React.createRef();
        this.incidentStartDateTime = React.createRef();
        this.incidentLocation = React.createRef();
        this.searchUser = React.createRef();
        this.searchUserEditMode = React.createRef();
    }

    private dataService = new CommonService();
    private graphEndpoint = "";

    //get all master data and check for edit mode or new record
    public async componentDidMount() {  
        await this.getDropdownOptions();
        //Event listener for screen resizing
        window.addEventListener("resize", this.resize.bind(this));
        this.resize();

        // check if form is in edit mode
        await this.checkIfEditMode();

        //update people picker control with required accessibility attributes
        this.updatePeoplePickerAttributes(this.incCommanderRef.current, this.state.selectedIncidentCommander.length > 0);
        this.updatePeoplePickerAttributes(this.normalSearchUserRef.current, this.state.selectedUsers.length > 0);
        this.updatePeoplePickerAttributes(this.normalSearchLeadRef.current, this.state.selectedLead.length > 0);

        this.getTeamNameConfigData();
        this.getRoleDefaultData();
        if (!this.state.isEditMode) {
            this.getIncidentTypeDefaultData();
            this.onToggleAdditionChannels(true);
            //set the current date and time for the date and time picker variables
            this.setDefaultDateTime();
        }
    }

    //get team name configuration data
    public getTeamNameConfigData = async () => {
        try {
            //graph endpoint to get data from team name configuration list
            let graphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}/lists/${siteConfig.configurationList}/items?$expand=fields&$Top=5000`;
            let configData = await this.dataService.getConfigData(graphEndpoint, this.props.graph, [constants.teamNameConfig]);
            configData = { ...configData[0], value: JSON.parse(configData[0]?.value) };
            let filteredArr: any = Object.keys(configData.value)
                .filter((key) => key.includes(constants.teamNameConfigConstants.IncidentName) || !key.includes(constants.teamNameConfigConstants.PrefixValue) || key.includes(constants.teamNameConfigConstants.IncidentType) || key.includes(constants.teamNameConfigConstants.StartDate))
                .reduce((obj, key) => {
                    return Object.assign(obj, {
                        [key]: configData.value[key]
                    });
                }, {});
            const sortedData: any = this.dataService.sortConfigData(filteredArr);
            this.setState({
                teamNameConfigArray: sortedData,
                prefixValue: configData.value.PrefixValue
            });
        }
        catch (error: any) {
            console.error(
                constants.errorLogPrefix + "IncidentDetails_GetConfiguration \n",
                JSON.stringify(error)
            );
            // Log Exception
            this.dataService.trackException(this.props.appInsights, error, constants.componentNames.IncidentDetailsComponent, 'TeamNameConfiguration_GetConfiguration', this.props.userPrincipalName);
        }
    }

    //get role default values
    private getRoleDefaultData = async () => {
        try {
            //graph endpoint to get data from Role Default list
            let graphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}/lists/${siteConfig.roleDefaultList}/items?$expand=fields&$Top=5000`;
            const rolesDefaultListData = await this.dataService.getRoleDefaultData(graphEndpoint, this.props.graph);
            this.setState({
                roleDefaultData: rolesDefaultListData
            })
        }
        catch (error: any) {
            console.error(
                constants.errorLogPrefix + "IncidentDetails_GetRoleDefaultData \n",
                JSON.stringify(error)
            );
            // Log Exception
            this.dataService.trackException(this.props.appInsights, error, constants.componentNames.IncidentDetailsComponent, 'IncidentDetails_GetRoleDefaultData', this.props.userPrincipalName);
        }
    }

    //get incident type default roles
    private getIncidentTypeDefaultData = async () => {
        try {
            //graph endpoint to get data from Role Default list
            let graphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}/lists/${siteConfig.incidentTypeDefaultRolesList}/items?$expand=fields&$Top=5000`;
            const incidentTypeRolesDefaultListData = await this.dataService.getIncidentTypeDefaultData(graphEndpoint, this.props.graph);
            this.setState({
                incidentTypeRoleDefaultData: incidentTypeRolesDefaultListData
            });
        }
        catch (error: any) {
            console.error(
                constants.errorLogPrefix + "IncidentDetails_getIncidentTypeDefaultData \n",
                JSON.stringify(error)
            );
            // Log Exception
            this.dataService.trackException(this.props.appInsights, error, constants.componentNames.IncidentDetailsComponent, 'IncidentDetails_getIncidentTypeDefaultData', this.props.userPrincipalName);
        }
    }

    //Function for screen Resizing
    resize = () => this.setState({ isDesktop: window.innerWidth > constants.mobileWidth })

    //updating people picker control with accessibility attributes
    private updatePeoplePickerAttributes = (peoplePickerRef: any, selected: boolean) => {
        customElements.whenDefined('mgt-people-picker').then(() => {
            if (selected) {
                const ariaLabel = peoplePickerRef?.getElementsByTagName("mgt-people-picker")[0]?.getAttribute("title");
                peoplePickerRef?.getElementsByTagName("mgt-people-picker")[0]?.shadowRoot
                    ?.querySelector('#selected-list')?.setAttribute('aria-label', ariaLabel);
            }
            else {
                peoplePickerRef?.getElementsByTagName("mgt-people-picker")[0]?.shadowRoot
                    ?.querySelector('mgt-flyout')?.querySelector('#people-picker-input')?.setAttribute('role', 'searchbox');
                peoplePickerRef?.getElementsByTagName("mgt-people-picker")[0]?.shadowRoot
                    ?.querySelector('mgt-flyout')?.querySelector('#people-picker-input')?.removeAttribute('aria-expanded');
            }
        })
    }

    //updating people picker control with accessibility attributes whenever component is updated
    public componentDidUpdate(_prevProps: IIncidentDetailsProps, prevState: IIncidentDetailsState) {
        if (prevState.selectedIncidentCommander !== this.state.selectedIncidentCommander &&
            this.state.selectedIncidentCommander.length === 0) {
            this.updatePeoplePickerAttributes(this.incCommanderRef.current, this.state.selectedIncidentCommander.length > 0);
        }
        if (prevState.selectedUsers !== this.state.selectedUsers) {
            this.updatePeoplePickerAttributes(this.normalSearchUserRef.current, this.state.selectedUsers.length > 0);
        }
        if (prevState.selectedLead !== this.state.selectedLead && this.state.selectedLead.length === 0) {
            this.updatePeoplePickerAttributes(this.normalSearchLeadRef.current, this.state.selectedLead.length > 0);
        }
        //set the state variable with start date and time when the date or time changes in the picker controls        
        if (prevState.incDetailsItem.startDate !== this.state.incDetailsItem.startDate ||
            prevState.incDetailsItem.startTime !== this.state.incDetailsItem.startTime) {

            this.formatStartDateTime(this.state.incDetailsItem.startDate, this.state.incDetailsItem.startTime);
        }
    }

    //set the state variables with default date and time
    private setDefaultDateTime() {
        let incInfo = { ...this.state.incDetailsItem };
        incInfo["startTime"] = new Date();
        incInfo["startDate"] = new Date();
        this.setState({ incDetailsItem: incInfo });
    }

    //set the state variable with date and time in the UTC format
    private formatStartDateTime(startDate: Date, startTime: Date) {
        let incInfo = { ...this.state.incDetailsItem };
        let formattedDate = moment(startDate).format("YYYY-MM-DD");
        let formattedTime;

        if (startTime && startTime.toString() !== "Invalid Date") {
            formattedTime = moment(startTime).format("HH:mm:ss[Z]");
            incInfo["startDateTime"] = formattedDate + "T" + formattedTime;
            this.setState({ incDetailsItem: incInfo });
        }
        else {
            incInfo["startDateTime"] = "";
            this.setState({ incDetailsItem: incInfo });
        }
    }

    // removing event listener for screen resizing on component unmount
    public componentWillUnmount() {
        //Event listener for screen resizing
        window.removeEventListener("resize", this.resize.bind(this));
    }

    // get dropdown options for Incident Status, Incident Type and Roles dropdown
    private getDropdownOptions = async () => {
        try {
            const incStatusGraphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}${graphConfig.listsGraphEndpoint}/${siteConfig.incStatusList}/items?$expand=fields&$Top=5000`;
            const incTypeGraphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}${graphConfig.listsGraphEndpoint}/${siteConfig.incTypeList}/items?$expand=fields&$Top=5000`;
            const roleGraphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}${graphConfig.listsGraphEndpoint}/${siteConfig.roleAssignmentList}/items?$expand=fields&$Top=5000`;

            const statusOptionsPromise = this.dataService.getDropdownOptions(incStatusGraphEndpoint, this.props.graph, true);
            const typeOptionsPromise = this.dataService.getDropdownOptions(incTypeGraphEndpoint, this.props.graph);
            const roleOptionsPromise = this.dataService.getDropdownOptions(roleGraphEndpoint, this.props.graph);

            await Promise.all([statusOptionsPromise, typeOptionsPromise, roleOptionsPromise])
                .then(([statusOptions, typeOptions, roleOptions]) => {
                    const optionsArr: any = [];
                    // remove "Closed" status from options if form is New Form
                    if (!(this.props.incidentData && this.props.incidentData.incidentId)) {
                        optionsArr.statusOptions = statusOptions.filter((statusObj: any) => statusObj.status !== constants.closed);
                    }
                    else {
                        optionsArr.statusOptions = statusOptions
                    }

                    optionsArr.typeOptions = typeOptions.sort();
                    optionsArr.roleOptions = roleOptions.sort();

                    const activeStatus = optionsArr.statusOptions.find((statusObj: IIncidentStatus) =>
                        statusObj.status === constants.active);

                    let incInfo: IncidentEntity = { ...this.state.incDetailsItem };
                    let inputValidationObj = this.state.inputValidation;
                    if (incInfo) {
                        //default the status to Active when the Status dropdown is blank
                        if (incInfo["incidentStatus"] === undefined && activeStatus !== undefined)
                            incInfo["incidentStatus"] = { status: constants.active, id: activeStatus.id };
                        inputValidationObj.incidentStatusHasError = false;
                    }

                    this.setState({
                        dropdownOptions: optionsArr,
                        showLoader: false,
                        incDetailsItem: incInfo,
                        inputValidation: inputValidationObj,
                        formOpacity: 1
                    })
                }, (error: any): void => {
                    this.props.showMessageBar(this.props.localeStrings.genericErrorMessage + ", " + this.props.localeStrings.dropdownRetrievalFailedErrMsg, constants.messageBarType.error);
                    this.setState({
                        showLoader: false,
                        formOpacity: 1
                    })
                });
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "CreateIncident_GetDropdownOptions \n",
                JSON.stringify(error)
            );
            // Log Exception
            this.dataService.trackException(this.props.appInsights, error, constants.componentNames.IncidentDetailsComponent, 'CreateIncident_GetDropdownOptions', this.props.userPrincipalName);
        }
    }

    // check if the form is in edit mode and set the data object if in edit mode
    private checkIfEditMode = async () => {
        // if the object has data, the form is in edit mode
        if (this.props.incidentData && this.props.incidentData.incidentId) {

            const teamWebURL = this.props.incidentData.teamWebURL ? this.props.incidentData.teamWebURL : '';
            const teamGroupId = teamWebURL.split("?")[1].split("&")[0].split("=")[1].trim();

            const incCommanderObj: UserDetails = {
                userName: this.props.incidentData.incidentCommanderObj ? this.props.incidentData.incidentCommanderObj.split('|')[0] : "",
                userEmail: this.props.incidentData.incidentCommanderObj ? this.props.incidentData.incidentCommanderObj.split('|')[2] : "",
                userId: this.props.incidentData.incidentCommanderObj ? this.props.incidentData.incidentCommanderObj.split('|')[1] : ""
            }

            const selectedIncCommander: any = {
                displayName: this.props.incidentData.incidentCommanderObj ? this.props.incidentData.incidentCommanderObj.split('|')[0] : "",
                userPrincipalName: this.props.incidentData.incidentCommanderObj ? this.props.incidentData.incidentCommanderObj.split('|')[2] : "",
                id: this.props.incidentData.incidentCommanderObj ? this.props.incidentData.incidentCommanderObj.split('|')[1] : ""
            }
            let incInfo = { ...this.state.incDetailsItem };
            incInfo.incidentId = this.props.incidentData.incidentId.toString();
            incInfo.incidentName = this.props.incidentData.incidentName ? this.props.incidentData.incidentName : '';
            incInfo.incidentType = this.props.incidentData.incidentType ? this.props.incidentData.incidentType : '';
            incInfo.startDateTime = this.props.incidentData.startDate ? this.props.incidentData.startDate : '';
            incInfo.incidentStatus = this.props.incidentData.incidentStatusObj ? this.props.incidentData.incidentStatusObj : { status: undefined, id: undefined };
            incInfo.incidentCommander = incCommanderObj;
            incInfo.location = this.props.incidentData.location ? this.props.incidentData.location : '';
            incInfo.incidentDesc = this.props.incidentData.incidentDescription ? this.props.incidentData.incidentDescription : '';
            incInfo.severity = this.props.incidentData.severity ? this.props.incidentData.severity.toString() : "";
            incInfo.reasonForUpdate = '';
            incInfo.cloudStorageLink = this.props.incidentData.cloudStorageLink ? this.props.incidentData.cloudStorageLink : '';
            const rolesObj: any[] = [];
            const isRoleInEditMode: boolean[] = [];
            const roleAssignments = this.props.incidentData.roleAssignments ? this.props.incidentData.roleAssignments : '';
            const roleLeads = this.props.incidentData.roleLeads ? this.props.incidentData.roleLeads : '';

            if (roleAssignments.length > 0 && roleAssignments.split(";").length > 1) {
                roleAssignments.split(";").forEach(role => {
                    if (role.length > 0) {
                        let userNamesStr = "";
                        isRoleInEditMode.push(false);
                        const userDetailsObj: any[] = [];
                        role.split(":")[1].trim().split(",").forEach(user => {
                            userNamesStr += user.split("|")[0].trim() + ", ";
                            userDetailsObj.push({
                                userName: user.split("|")[0].trim(),
                                userEmail: user.split("|")[2].trim(),
                                userId: user.split("|")[1].trim(),
                            });
                        });
                        userNamesStr = userNamesStr.trim();
                        userNamesStr = userNamesStr.slice(0, -1);

                        rolesObj.push({
                            role: role.split(":")[0].trim(),
                            userNamesString: userNamesStr,
                            userObjString: role.split(":")[1].trim(),
                            userDetailsObj: userDetailsObj
                        })
                    }
                });
            }

            //adding role lead for each role into the rolesObj array
            if (roleLeads.length > 0 && roleLeads.split(";").length > 1) {
                roleLeads.split(";").forEach(role => {
                    if (role.length > 0) {
                        let leadNameStr = "";
                        isRoleInEditMode.push(false);
                        const leadDetailsObj: any[] = [];
                        role.split(":")[1].trim().split(",").forEach(user => {
                            leadNameStr += user.split("|")[0].trim() + ", ";
                            leadDetailsObj.push({
                                userName: user.split("|")[0].trim(),
                                userEmail: user.split("|")[2].trim(),
                                userId: user.split("|")[1].trim(),
                            });
                        });
                        leadNameStr = leadNameStr.trim();
                        leadNameStr = leadNameStr.slice(0, -1);

                        const roleObj = rolesObj.find(e => e.role === role.split(":")[0].trim());
                        const roleNewObj = {
                            ...roleObj,
                            leadNameString: leadNameStr,
                            leadObjString: role.split(":")[1].trim(),
                            leadDetailsObj: leadDetailsObj
                        };
                        rolesObj[rolesObj.findIndex(e => e.role === role.split(":")[0].trim())] = roleNewObj
                    }
                });
            }

            const selectedRoles = rolesObj.map((roles: any) => roles.role);
            let roleOptions = this.state.dropdownOptions["roleOptions"].filter((role: string) => selectedRoles.indexOf(role) === -1);
            const dropdownOptions = this.state.dropdownOptions;
            dropdownOptions["roleOptions"] = roleOptions;

            this.setState({
                incDetailsItem: incInfo,
                selectedIncidentCommander: [selectedIncCommander],
                roleAssignments: rolesObj,
                existingRolesMembers: rolesObj,
                existingIncCommander: incInfo.incidentCommander,
                isRoleInEditMode: isRoleInEditMode,
                teamGroupId: teamGroupId,
                incidentTypeSearchQuery: this.props.incidentData.incidentType ? this.props.incidentData.incidentType : '',
                dropdownOptions: dropdownOptions,
                selectedSeverity: constants.severity.indexOf(incInfo.severity) === -1 ? 0 : constants.severity.indexOf(incInfo.severity),
                isEditMode: true,
                toggleCloudStorageLocation: incInfo.cloudStorageLink.trim() !== "",
                selectedLocation: JSON.parse(incInfo.location)
            })
        }
    }

    // Update Incident Commander object on change of incident commander
    private handleIncCommanderChange = (selectedValue: any) => {
        let incInfo = { ...this.state.incDetailsItem };
        if (incInfo) {
            let inputValidationObj = this.state.inputValidation;
            let incCommanderHasRegexError = this.state.incCommanderHasRegexError;
            const selctedIncCommander = [];
            //update selected incident commander object 
            if (selectedValue.detail.length > 0) {
                inputValidationObj.incidentCommandarHasError = false;
                // Restrict External users to be added as incident commander
                if (selectedValue.detail[0].userPrincipalName.match("#EXT#") === null) {
                    incCommanderHasRegexError = false;
                    selctedIncCommander.push({
                        displayName: selectedValue.detail[0] ? selectedValue.detail[0].displayName.replace(",", "") : '',
                        userPrincipalName: selectedValue.detail[0] ? selectedValue.detail[0].userPrincipalName : '',
                        id: selectedValue.detail[0] ? selectedValue.detail[0].id.includes("@") ? selectedValue.detail[0].id.split("@")[0] : selectedValue.detail[0].id : ''
                    });
                    // create user object for incident commander
                    incInfo.incidentCommander = {
                        userName: selectedValue.detail[0] ? selectedValue.detail[0].displayName.replace(",", "") : '',
                        userEmail: selectedValue.detail[0] ? selectedValue.detail[0].userPrincipalName : '',
                        userId: selectedValue.detail[0] ? selectedValue.detail[0].id.includes("@") ? selectedValue.detail[0].id.split("@")[0] : selectedValue.detail[0].id : ''

                    }
                }
                else
                    incCommanderHasRegexError = true;
            }
            else {
                inputValidationObj.incidentCommandarHasError = true;
                incCommanderHasRegexError = false;
                // create user object for incident commander
                incInfo.incidentCommander = {
                    userName: '',
                    userEmail: '',
                    userId: ''
                }
            }
            this.setState({
                incDetailsItem: incInfo,
                selectedIncidentCommander: selctedIncCommander,
                inputValidation: inputValidationObj,
                incCommanderHasRegexError: incCommanderHasRegexError
            });
        }
    };

    // on change handler for text input changes
    private onTextInputChange = (event: any, key: string) => {
        let incInfo = { ...this.state.incDetailsItem };
        let inputValidationObj = this.state.inputValidation;
        if (incInfo) {
            switch (key) {
                case "incidentName":
                    incInfo[key] = event.target.value;
                    if (event.target.value.length > 0) {
                        inputValidationObj.incidentNameHasError = false;
                    }
                    else {
                        inputValidationObj.incidentNameHasError = true;
                    }
                    this.setState({
                        incDetailsItem: incInfo,
                        inputValidation: inputValidationObj
                    })
                    break;
                case "assignedUser":
                    incInfo[key] = event.target.value;
                    this.setState({ incDetailsItem: incInfo });
                    break;
                case "incidentDesc":
                    incInfo[key] = event.target.value;
                    if (event.target.value.length > 0) {
                        inputValidationObj.incidentDescriptionHasError = false;
                    }
                    else {
                        inputValidationObj.incidentDescriptionHasError = true;
                    }
                    this.setState({ incDetailsItem: incInfo, inputValidation: inputValidationObj })
                    break;
                case "reasonForUpdate":
                    incInfo[key] = event.target.value;
                    if (event.target.value.length > 0) {
                        inputValidationObj.incidentReasonForUpdateHasError = false;
                    }
                    else {
                        inputValidationObj.incidentReasonForUpdateHasError = true;
                    }
                    this.setState({ incDetailsItem: incInfo, inputValidation: inputValidationObj })
                    break;
                case "cloudStorageLink":
                    incInfo[key] = event.target.value;
                    if (event.target.value.length > 0) {
                        inputValidationObj.cloudStorageLinkHasError = false;
                    }
                    else {
                        inputValidationObj.cloudStorageLinkHasError = true;
                    }
                    this.setState({ incDetailsItem: incInfo, inputValidation: inputValidationObj });
                    break;
                default:
                    break;
            }
        }
    }

    // update state for new role string
    private onAddNewRoleChange = (event: any) => {
        let isButtonDisabled = true;
        if (event.target.value && event.target.value.length > 0) {
            isButtonDisabled = false;
        }
        this.setState({ newRoleString: event.target.value, isCreateNewRoleBtnDisabled: isButtonDisabled });
    }

    // on incident type dropdown value change
    private onIncidentTypeChange = async (event: any, selectedValue: any) => {
        //reset roles dropdown values whenever the Incident type is changed
        await this.getDropdownOptions();

        //check if we have data for selected role
        const filteredincidentTypeDefaultData = this.state.incidentTypeRoleDefaultData.filter((e: any) => e.incidentType === selectedValue.key);
        const rolesObj: any[] = [];
        let defaultAdditionalChannels: IAdditionalTeamChannels[] = [];
        const isRoleInEditMode: boolean[] = [];
        //format roles object
        if (filteredincidentTypeDefaultData.length > 0 && filteredincidentTypeDefaultData[0]?.roleAssignments?.split(";").length > 1) {
            filteredincidentTypeDefaultData[0]?.roleAssignments?.split(";").forEach((role: any) => {
                if (role.length > 0) {
                    let userNamesStr = "";
                    isRoleInEditMode.push(false);
                    const userDetailsObj: any[] = [];
                    role.split(":")[1].trim().split(",").forEach((user: any) => {
                        userNamesStr += user.split("|")[0].trim() + ", ";
                        userDetailsObj.push({
                            userName: user.split("|")[0].trim(),
                            userEmail: user.split("|")[2].trim(),
                            userId: user.split("|")[1].trim(),
                        });
                    });
                    userNamesStr = userNamesStr.trim();
                    userNamesStr = userNamesStr.slice(0, -1);

                    rolesObj.push({
                        role: role.split(":")[0].trim(),
                        userNamesString: userNamesStr,
                        userObjString: role.split(":")[1].trim(),
                        userDetailsObj: userDetailsObj
                    })
                }
            });
        }
        //format roles object with lead details
        if (filteredincidentTypeDefaultData.length > 0 && filteredincidentTypeDefaultData[0]?.roleLeads?.split(";").length > 1) {
            filteredincidentTypeDefaultData[0]?.roleLeads?.split(";").forEach((role: any) => {
                if (role.length > 0) {
                    let leadNameStr = "";
                    isRoleInEditMode.push(false);
                    const leadDetailsObj: any[] = [];
                    role.split(":")[1].trim().split(",").forEach((user: any) => {
                        leadNameStr += user.split("|")[0].trim() + ", ";
                        leadDetailsObj.push({
                            userName: user.split("|")[0].trim(),
                            userEmail: user.split("|")[2].trim(),
                            userId: user.split("|")[1].trim(),
                        });
                    });
                    leadNameStr = leadNameStr.trim();
                    leadNameStr = leadNameStr.slice(0, -1);

                    const roleObj = rolesObj.find(e => e.role === role.split(":")[0].trim());
                    rolesObj.splice(rolesObj.findIndex(e => e.role === role.split(":")[0].trim()), 1);

                    rolesObj.push({
                        ...roleObj,
                        leadNameString: leadNameStr,
                        leadObjString: role.split(":")[1].trim(),
                        leadDetailsObj: leadDetailsObj
                    });
                }
            });
        }
        //Format and Assign default additional channels
        if (filteredincidentTypeDefaultData.length > 0 &&
            filteredincidentTypeDefaultData[0].additionalChannels.split(',').length > 1) {
            filteredincidentTypeDefaultData[0].additionalChannels.split(',').forEach((channel: any) => {
                const channelObj = {
                    channelName: channel.trim(),
                    hasRegexError: false,
                    regexErrorMessage: "",
                    channelType: constants.standardChannel
                }
                defaultAdditionalChannels.push(channelObj);
            });
            if (defaultAdditionalChannels.length === 2) {
                defaultAdditionalChannels.push({ channelName: "", hasRegexError: false, regexErrorMessage: "" });
            }
        }
        else {
            const channel = filteredincidentTypeDefaultData[0]?.additionalChannels?.split(',')[0]?.trim();
            if (channel !== "" && channel !== undefined) {
                defaultAdditionalChannels.push({ channelName: channel, channelType: constants.standardChannel, hasRegexError: false, regexErrorMessage: "" });
                while (defaultAdditionalChannels.length < 3) {
                    defaultAdditionalChannels.push({ channelName: "", channelType: "", hasRegexError: false, regexErrorMessage: "" });
                }
            }
            else {
                defaultAdditionalChannels = [
                    { channelName: constants.defaultChannelConstants.Logistics, channelType: constants.standardChannel, hasRegexError: false, regexErrorMessage: "" },
                    { channelName: constants.defaultChannelConstants.Planning, channelType: constants.standardChannel, hasRegexError: false, regexErrorMessage: "" },
                    { channelName: constants.defaultChannelConstants.Recovery, channelType: constants.standardChannel, hasRegexError: false, regexErrorMessage: "" }
                ];
            }
        }
        //Assign default Cloud Storage link
        if (filteredincidentTypeDefaultData.length > 0) {
            const cloudStorageLink = filteredincidentTypeDefaultData[0]?.cloudStorageLink?.trim();
            this.setState({
                incDetailsItem: {
                    ...this.state.incDetailsItem,
                    cloudStorageLink: cloudStorageLink !== "" ? cloudStorageLink : ""
                }
            });
        }
        else this.setState({ incDetailsItem: { ...this.state.incDetailsItem, cloudStorageLink: "" } })

        const selectedRoles = rolesObj.map((roles: any) => roles.role);
        let roleOptions = this.state.dropdownOptions["roleOptions"].filter((role: string) => selectedRoles.indexOf(role) === -1);
        const dropdownOptions = this.state.dropdownOptions;
        dropdownOptions["roleOptions"] = roleOptions;

        let incInfo = this.state.incDetailsItem;
        if (incInfo) {
            let incInfo = { ...this.state.incDetailsItem };
            if (incInfo) {
                incInfo["incidentType"] = selectedValue.key;
                incInfo["selectedRole"] = "";
                let inputValidationObj = this.state.inputValidation;
                inputValidationObj.incidentTypeHasError = false;
                inputValidationObj.cloudStorageLinkHasError = false;
                this.setState({
                    selectedUsers: [],
                    selectedLead: [],
                    incDetailsItem: {
                        ...incInfo,
                        additionalTeamChannels: [...defaultAdditionalChannels]
                    },
                    inputValidation: inputValidationObj,
                    roleAssignments: filteredincidentTypeDefaultData.length > 0 ? rolesObj : [],
                    existingRolesMembers: filteredincidentTypeDefaultData.length > 0 ? rolesObj : [],
                    isRoleInEditMode: isRoleInEditMode
                });
            }
        }
    }

    // on incident status dropdown value change
    private onIncidentStatusChange = (_event: any, selectedValue: any) => {
        let incInfo = this.state.incDetailsItem;
        if (incInfo) {
            let incInfo = { ...this.state.incDetailsItem };
            if (incInfo) {
                incInfo["incidentStatus"] = { status: selectedValue.text, id: selectedValue.key };
                let inputValidationObj = this.state.inputValidation;
                inputValidationObj.incidentStatusHasError = false;
                this.setState({ incDetailsItem: incInfo, inputValidation: inputValidationObj })
            }
        }
    }

    // on role dropdown value change
    private onRoleChange = (_event: any, selectedRole: any) => {
        this.setState({
            roleAddSuccessMessage: ""
        })
        //check if we have data for selected role
        const filteredRoleData = this.state.roleDefaultData.filter((e: any) => e.role === selectedRole.value);
        let incInfo = this.state.incDetailsItem;

        if (incInfo) {
            let incInfo = { ...this.state.incDetailsItem };
            if (incInfo) {
                incInfo["selectedRole"] = selectedRole.value;
                incInfo["assignedUser"] = filteredRoleData.length > 0 ?
                    filteredRoleData[0].users.map((user: any) => {
                        return {
                            "userName": user ? user.displayName : "",
                            "userEmail": user ? user.userPrincipalName : "",
                            "userId": user ? user.id : "",
                        }
                    })
                    : [];
                incInfo["assignedLead"] = filteredRoleData.length > 0 && filteredRoleData[0].lead.length > 0 ?
                    filteredRoleData[0].lead.map((user: any) => {
                        return {
                            "userName": user ? user.displayName : "",
                            "userEmail": user ? user.userPrincipalName : "",
                            "userId": user ? user.id : "",
                        }
                    })
                    : [];

                this.setState({
                    incDetailsItem: incInfo,
                    selectedUsers: filteredRoleData.length > 0 ? filteredRoleData[0].users : [],
                    selectedLead: filteredRoleData.length > 0 ? filteredRoleData[0].lead : [],
                    secIncCommanderUserHasRegexError: false,
                    secIncCommanderLeadHasRegexError: false
                },
                    //check the state to enable or disable the Add Role button
                    (() => this.checkAddRoleBtnState()))
            }
        }
    }

    // connect with service to create new role in Role Assignments list
    private addNewRole = async () => {
        this.setState({
            showLoader: true,
            formOpacity: 0.5
        })
        // create graph endpoint for role assignment list
        this.graphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}${graphConfig.listsGraphEndpoint}/${siteConfig.roleAssignmentList}/items`;

        // create new item object to add the role
        const listItem = {
            fields: {
                Title: this.state.newRoleString
            }
        };

        try {
            const addedRole = await this.dataService.sendGraphPostRequest(this.graphEndpoint, this.props.graph, listItem);

            if (addedRole) {
                const arr: any = { ...this.state.dropdownOptions };
                arr["roleOptions"].push(this.state.newRoleString);

                let incInfo = this.state.incDetailsItem;
                if (incInfo) {
                    let incInfo = { ...this.state.incDetailsItem };
                    if (incInfo) {
                        incInfo["selectedRole"] = this.state.newRoleString;
                        this.setState({
                            incDetailsItem: incInfo, newRoleString: "", showLoader: false,
                            formOpacity: 1
                        })
                        this.props.showMessageBar(this.props.localeStrings.addRoleSuccessMessage, "success");
                    }
                    else {
                        this.setState({
                            showLoader: false,
                            formOpacity: 1
                        })
                    }
                }
            }
            else {
                this.setState({
                    showLoader: false,
                    formOpacity: 1
                })
            }
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "CreateIncident_AddNewRole \n",
                JSON.stringify(error)
            );
            // Log Exception
            this.dataService.trackException(this.props.appInsights, error, constants.componentNames.IncidentDetailsComponent, 'CreateIncident_AddNewRole', this.props.userPrincipalName);

        }
    }

    // Update Role Assigned User on change of User in role assignment people picker 
    private handleAssignedUserChange = (selectedValue: any) => {
        let incInfo = { ...this.state.incDetailsItem };
        let secIncCommanderUserHasRegexError = this.state.secIncCommanderUserHasRegexError;
        const selectedUsersArr: any = [];
        let assignedUsersArray: any = [];
        if (incInfo) {
            //Restrict External users to be added as secondary incident commander or to "Edit Access Role"
            if (this.state.incDetailsItem.selectedRole === constants.secondaryIncidentCommanderRole || this.state.incDetailsItem.selectedRole === this.props.editIncidentAccessRole) {
                if (selectedValue.detail.length > 0) {
                    selectedValue.detail.forEach((user: any) => {
                        if (user?.userPrincipalName.match("#EXT#") === null) {
                            secIncCommanderUserHasRegexError = false;
                            selectedUsersArr.push({
                                displayName: user.displayName.replace(",", ""),
                                userPrincipalName: user.userPrincipalName,
                                id: user.id.includes("@") ? user.id.split("@")[0] : user.id
                            });
                            assignedUsersArray.push({
                                "userName": user ? user.displayName.replace(",", "") : "",
                                "userEmail": user ? user.userPrincipalName : "",
                                "userId": user ? user.id.includes("@") ? user.id.split("@")[0] : user.id : "",
                            });
                        }
                        else secIncCommanderUserHasRegexError = true
                    });
                }
                else secIncCommanderUserHasRegexError = false;
                incInfo["assignedUser"] = assignedUsersArray;
            }
            else {
                secIncCommanderUserHasRegexError = false;
                incInfo["assignedUser"] = selectedValue.detail.map((user: any) => {
                    selectedUsersArr.push({
                        displayName: user.displayName.replace(",", ""),
                        userPrincipalName: user.userPrincipalName,
                        id: user.id.includes("@") ? user.id.split("@")[0] : user.id
                    });
                    return {
                        "userName": user ? user.displayName.replace(",", "") : "",
                        "userEmail": user ? user.userPrincipalName : "",
                        "userId": user ? user.id.includes("@") ? user.id.split("@")[0] : user.id : "",
                    }
                });
            }

            this.setState({
                incDetailsItem: incInfo, selectedUsers: selectedUsersArr,
                secIncCommanderUserHasRegexError: secIncCommanderUserHasRegexError
            });
            //check the state to enable or disable the Add Role button
            this.checkAddRoleBtnState();
        }
    };

    // Update Role Assigned Lead on change of Lead in role assignment people picker 
    private handleAssignedLeadChange = (selectedValue: any) => {
        let incInfo = { ...this.state.incDetailsItem };
        let secIncCommanderLeadHasRegexError = this.state.secIncCommanderLeadHasRegexError;
        let assignedLeadArray: any = [];
        const selectedRoleLead: any = [];
        if (incInfo) {
            //Restrict External users to be added as secondary incident commander or "Edit Access Role"
            if (this.state.incDetailsItem.selectedRole === constants.secondaryIncidentCommanderRole || this.state.incDetailsItem.selectedRole === this.props.editIncidentAccessRole) {
                if (selectedValue.detail.length > 0) {
                    selectedValue.detail.forEach((user: any) => {
                        if (user?.userPrincipalName?.match("#EXT#") === null) {
                            secIncCommanderLeadHasRegexError = false;
                            selectedRoleLead.push({
                                displayName: user.displayName.replace(",", ""),
                                userPrincipalName: user.userPrincipalName,
                                id: user.id.includes("@") ? user.id.split("@")[0] : user.id
                            });
                            assignedLeadArray.push({
                                "userName": user ? user.displayName.replace(",", "") : "",
                                "userEmail": user ? user.userPrincipalName : "",
                                "userId": user ? user.id.includes("@") ? user.id.split("@")[0] : user.id : "",
                            });
                        }
                        else secIncCommanderLeadHasRegexError = true
                    });
                    incInfo["assignedLead"] = assignedLeadArray;
                }
                else secIncCommanderLeadHasRegexError = false
            }
            else {
                secIncCommanderLeadHasRegexError = false;
                incInfo["assignedLead"] = selectedValue.detail.map((user: any) => {
                    selectedRoleLead.push({
                        displayName: user.displayName.replace(",", ""),
                        userPrincipalName: user.userPrincipalName,
                        id: user.id.includes("@") ? user.id.split("@")[0] : user.id
                    });
                    return {
                        "userName": user ? user.displayName.replace(",", "") : "",
                        "userEmail": user ? user.userPrincipalName : "",
                        "userId": user ? user.id.includes("@") ? user.id.split("@")[0] : user.id : "",
                    }
                });
            }

            this.setState({
                incDetailsItem: incInfo, selectedLead: selectedRoleLead,
                secIncCommanderLeadHasRegexError: secIncCommanderLeadHasRegexError
            });
            //check the state to enable or disable the Add Role button
            this.checkAddRoleBtnState();
        }
    };

    // Update Role Assigned Lead on change of Lead in people picker from role assignment table
    private handleAssignedLeadChangeInEditMode = (selectedValue: any, idx: number) => {
        let incInfo = { ...this.state.incDetailsItem };
        let secIncCommanderLeadInEditModeHasRegexError = this.state.secIncCommanderLeadInEditModeHasRegexError;
        const selectedRoleLead: any = [];
        let assignedLeadArray: any = [];
        if (incInfo) {
            //Restrict External users to be added as secondary incident commander or to "Edit Access Role"
            if (this.state.roleAssignments[idx].role === constants.secondaryIncidentCommanderRole || this.state.roleAssignments[idx].role === this.props.editIncidentAccessRole) {
                if (selectedValue.detail.length > 0) {
                    selectedValue.detail.forEach((user: any) => {
                        if (user?.userPrincipalName.match("#EXT#") === null) {
                            secIncCommanderLeadInEditModeHasRegexError = false;
                            selectedRoleLead.push({
                                displayName: user.displayName.replace(",", ""),
                                userPrincipalName: user.userPrincipalName,
                                id: user.id.includes("@") ? user.id.split("@")[0] : user.id
                            });
                            assignedLeadArray.push({
                                "userName": user ? user.displayName.replace(",", "") : "",
                                "userEmail": user ? user.userPrincipalName : "",
                                "userId": user ? user.id.includes("@") ? user.id.split("@")[0] : user.id : "",
                            });
                        }
                        else secIncCommanderLeadInEditModeHasRegexError = true;
                    });
                }
                else secIncCommanderLeadInEditModeHasRegexError = false
                incInfo["assignedLead"] = assignedLeadArray;
            }
            else {
                secIncCommanderLeadInEditModeHasRegexError = false
                incInfo["assignedLead"] = selectedValue.detail.map((user: any) => {
                    selectedRoleLead.push({
                        displayName: user.displayName.replace(",", ""),
                        userPrincipalName: user.userPrincipalName,
                        id: user.id.includes("@") ? user.id.split("@")[0] : user.id
                    });
                    return {
                        "userName": user ? user.displayName.replace(",", "") : "",
                        "userEmail": user ? user.userPrincipalName : "",
                        "userId": user ? user.id.includes("@") ? user.id.split("@")[0] : user.id : "",
                    }
                });
            }

            this.setState({
                incDetailsItem: incInfo, selectedLeadInEditMode: selectedRoleLead,
                secIncCommanderLeadInEditModeHasRegexError: secIncCommanderLeadInEditModeHasRegexError
            });
        }
    };

    // Update Role Assigned User on change of User in people picker from role assignment table
    private handleAssignedUserChangeInEditMode = (selectedValue: any, idx: number) => {
        let incInfo = { ...this.state.incDetailsItem };
        let secIncCommanderUserInEditModeHasRegexError = this.state.secIncCommanderUserInEditModeHasRegexError;
        const selectedUsersArr: any = [];
        let assignedUsersArray: any = [];
        if (incInfo) {
            //Restrict External users to be added as secondary incident commander or "Edit Access Role"
            if (this.state.roleAssignments[idx].role === constants.secondaryIncidentCommanderRole || this.state.roleAssignments[idx].role === this.props.editIncidentAccessRole) {
                if (selectedValue.detail.length > 0) {
                    selectedValue.detail.forEach((user: any) => {
                        if (user?.userPrincipalName.match("#EXT#") === null) {
                            secIncCommanderUserInEditModeHasRegexError = false;
                            selectedUsersArr.push({
                                displayName: user.displayName.replace(",", ""),
                                userPrincipalName: user.userPrincipalName,
                                id: user.id.includes("@") ? user.id.split("@")[0] : user.id
                            });
                            assignedUsersArray.push({
                                "userName": user ? user.displayName.replace(",", "") : "",
                                "userEmail": user ? user.userPrincipalName : "",
                                "userId": user ? user.id.includes("@") ? user.id.split("@")[0] : user.id : "",
                            });
                        }
                        else secIncCommanderUserInEditModeHasRegexError = true
                    });
                }
                else secIncCommanderUserInEditModeHasRegexError = false
                incInfo["assignedUser"] = assignedUsersArray;
            }
            else {
                secIncCommanderUserInEditModeHasRegexError = false;
                incInfo["assignedUser"] = selectedValue.detail.map((user: any) => {
                    selectedUsersArr.push({
                        displayName: user.displayName.replace(",", ""),
                        userPrincipalName: user.userPrincipalName,
                        id: user.id.includes("@") ? user.id.split("@")[0] : user.id
                    });
                    return {
                        "userName": user ? user.displayName.replace(",", "") : "",
                        "userEmail": user ? user.userPrincipalName : "",
                        "userId": user ? user.id.includes("@") ? user.id.split("@")[0] : user.id : "",
                    }
                });
            }

            this.setState({
                incDetailsItem: incInfo, selectedUsersInEditMode: selectedUsersArr,
                secIncCommanderUserInEditModeHasRegexError: secIncCommanderUserInEditModeHasRegexError
            });
        }
    };

    // update the role assignment array
    private addRoleAssignment = () => {
        let roleAssignment = [...this.state.roleAssignments];
        let userDetailsObj: any = [];
        let userNameString = "";
        let userObjString = "";
        let leadNameString = "";
        let leadObjString = "";
        let leadDetailsObj: any = [];
        // push roles into array to create role object
        this.state.incDetailsItem.assignedUser.forEach(assignedUser => {
            userNameString += assignedUser.userName + ", ";
            userObjString += assignedUser.userName + "|" + assignedUser.userId + "|" + assignedUser.userEmail + ", ";
            userDetailsObj.push({
                userName: assignedUser.userName,
                userEmail: assignedUser.userEmail,
                userId: assignedUser.userId,
            });
        });
        userNameString = userNameString.trim();
        userNameString = userNameString.slice(0, -1);

        userObjString = userObjString.trim();
        userObjString = userObjString.slice(0, -1);

        this.state.incDetailsItem.assignedLead.forEach(assignedLead => {
            leadNameString = assignedLead.userName;
            leadObjString = assignedLead.userName + "|" + assignedLead.userId + "|" + assignedLead.userEmail;
            leadDetailsObj.push({
                userName: assignedLead.userName,
                userEmail: assignedLead.userEmail,
                userId: assignedLead.userId,
            });
        });

        roleAssignment.push({
            role: this.state.incDetailsItem.selectedRole,
            userNamesString: userNameString,
            userObjString: userObjString,
            userDetailsObj: userDetailsObj,
            leadNameString: leadNameString,
            leadObjString: leadObjString,
            leadDetailsObj: leadDetailsObj,
            saveDefault: this.state.saveDefaultRoleCheck
        })

        const isRoleInEditMode = [...this.state.isRoleInEditMode];
        isRoleInEditMode.push(false);

        let roleOptions = this.state.dropdownOptions["roleOptions"].filter((role: string) => role !== this.state.incDetailsItem.selectedRole)

        const dropdownOptions = this.state.dropdownOptions;
        dropdownOptions["roleOptions"] = roleOptions;

        // clear roles control values
        let incInfo = this.state.incDetailsItem;
        if (incInfo) {
            let incInfo = { ...this.state.incDetailsItem };
            if (incInfo) {
                incInfo["selectedRole"] = "";
                this.setState({
                    roleAssignments: roleAssignment,
                    incDetailsItem: incInfo,
                    isRoleInEditMode: isRoleInEditMode,
                    selectedUsers: [],
                    selectedLead: [],
                    isAddRoleAssignmentBtnDisabled: true,
                    dropdownOptions: dropdownOptions,
                    saveDefaultRoleCheck: false,
                    secIncCommanderLeadHasRegexError: false,
                    secIncCommanderUserHasRegexError: false
                })
            }
        }
        this.setState({
            roleAddSuccessMessage: constants.addRoleMessage
        })
    }

    // change add role assignment button disable state
    private checkAddRoleBtnState = () => {
        if (this.state.incDetailsItem.selectedRole !== "" &&
            this.state.incDetailsItem.selectedRole !== undefined &&
            this.state.selectedUsers && this.state.selectedUsers.length > 0) {
            this.setState({
                isAddRoleAssignmentBtnDisabled: false
            })
        }
        else {
            this.setState({
                isAddRoleAssignmentBtnDisabled: true
            })
        }
    }

    // delete added role from RoleAssignment object
    private deleteRoleItem = (itemIndex: number) => {
        const assignments = [...this.state.roleAssignments];

        const dropdownOptions = this.state.dropdownOptions;
        dropdownOptions["roleOptions"].push(assignments[itemIndex].role);
        dropdownOptions["roleOptions"].sort();

        assignments.splice(itemIndex, 1);

        const isRoleInEditMode = [...this.state.isRoleInEditMode];
        isRoleInEditMode.splice(itemIndex, 1);

        this.setState({
            roleAssignments: assignments, isRoleInEditMode: isRoleInEditMode,
            dropdownOptions: dropdownOptions
        });
        this.props.showMessageBar(this.props.localeStrings.removedRoleSuccessMessage, "success");
    }

    // edit added users from RoleAssignment object
    private editRoleItem = (itemIndex: number) => {
        const isRoleInEditMode = [...this.state.isRoleInEditMode];
        isRoleInEditMode.forEach((editMode, index) => {
            if (index === itemIndex) {
                isRoleInEditMode[itemIndex] = true;
            }
            else {
                isRoleInEditMode[index] = false;
            }

        });
        const roles = [...this.state.roleAssignments];
        const selectedUserInRole: any = [];
        const selectedLeadInRole: any = [];
        roles[itemIndex].userDetailsObj.forEach(user => {
            selectedUserInRole.push({
                displayName: user.userName,
                userPrincipalName: user.userEmail,
                id: user.userId
            });
        });

        if (roles[itemIndex].leadDetailsObj !== undefined) {
            roles[itemIndex].leadDetailsObj.forEach(user => {
                selectedLeadInRole.push({
                    displayName: user.userName,
                    userPrincipalName: user.userEmail,
                    id: user.userId
                });
            });
        }
        this.setState({
            isRoleInEditMode: isRoleInEditMode,
            selectedUsersInEditMode: selectedUserInRole,
            selectedLeadInEditMode: selectedLeadInRole,
            secIncCommanderLeadInEditModeHasRegexError: false,
            secIncCommanderUserInEditModeHasRegexError: false
        });
    }

    // exit from edit mode in roles
    private exitEditModeForRoles = (itemIndex: number) => {
        const isRoleInEditMode = [...this.state.isRoleInEditMode];
        isRoleInEditMode[itemIndex] = false;

        this.setState({
            isRoleInEditMode: isRoleInEditMode,
            secIncCommanderLeadInEditModeHasRegexError: false,
            secIncCommanderUserInEditModeHasRegexError: false
        });
    }

    // update the role assignment array
    private updateRoleAssignment = (index: number) => {
        let roleAssignment = [...this.state.roleAssignments];
        let userDetailsObj: any = [];
        let userNameString = "";
        let userObjString = "";
        let leadDetailsObj: any = [];
        let leadNameString = "";
        let leadObjString = "";

        // check if atleast one member is present for the role
        if (this.state.selectedUsersInEditMode.length > 0) {
            // loop throught updated user object        
            this.state.selectedUsersInEditMode.forEach((users: any) => {
                userNameString += users.displayName + ", ";
                userObjString += users.displayName + "|" + users.id + "|" + users.userPrincipalName + ", ";
                userDetailsObj.push({
                    userName: users.displayName,
                    userEmail: users.userPrincipalName,
                    userId: users.id,
                });
            });

            userNameString = userNameString.trim();
            userNameString = userNameString.slice(0, -1);

            userObjString = userObjString.trim();
            userObjString = userObjString.slice(0, -1);

            this.state.selectedLeadInEditMode.forEach((lead: any) => {
                leadNameString = lead.displayName;
                leadObjString = lead.displayName + "|" + lead.id + "|" + lead.userPrincipalName;
                leadDetailsObj.push({
                    userName: lead.displayName,
                    userEmail: lead.userPrincipalName,
                    userId: lead.id,
                });
            });

            roleAssignment[index] = {
                role: roleAssignment[index].role,
                userNamesString: userNameString,
                userObjString: userObjString,
                userDetailsObj: userDetailsObj,
                leadNameString: leadNameString,
                leadObjString: leadObjString,
                leadDetailsObj: leadDetailsObj,
                saveDefault: roleAssignment[index].saveDefault
            }

            const isRoleInEditMode = [...this.state.isRoleInEditMode];
            isRoleInEditMode[index] = false;

            this.setState({
                roleAssignments: roleAssignment,
                isRoleInEditMode: isRoleInEditMode,
                selectedUsersInEditMode: [],
                secIncCommanderLeadInEditModeHasRegexError: false,
                secIncCommanderUserInEditModeHasRegexError: false
            })
        }
        else {
            this.props.showMessageBar(this.props.localeStrings.noMemberForRole, constants.messageBarType.error);
            this.scrollToTop();
        }
    }

    // save or update default role
    private createUpdateDefaultRoles = async () => {
        try {
            //check if role is already present
            this.state.roleAssignments.map((item: any) => {
                let duplicateCount = this.state.roleDefaultData.filter(e => e.role === item.role);
                if (item.saveDefault) {
                    if (duplicateCount.length === 0) {
                        // add default roles
                        const roleObj: any = {
                            fields: {
                                Title: item.role,
                                Users: item.userObjString,
                                RoleLead: item.leadObjString ? item.leadObjString : ""
                            }
                        }
                        this.addDefaultRoles(roleObj);
                    }
                    else {
                        //update default role data
                        const updateRoleObj: any = {
                            Users: item.userObjString,
                            RoleLead: item.leadObjString ? item.leadObjString : ""
                        }
                        this.updateDefaultRoles(updateRoleObj, duplicateCount[0].itemId);
                    }
                }
            })
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "CreateIncident_CreateUpdateDefaultRoles \n",
                JSON.stringify(error)
            );
            // Log Exception
            this.dataService.trackException(this.props.appInsights, error, constants.componentNames.IncidentDetailsComponent, 'CreateIncident_CreateUpdateDefaultRoles', this.props.userPrincipalName);

        }
    }

    //add default user roles
    private addDefaultRoles = async (roleObj: any) => {
        try {
            this.graphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}${graphConfig.listsGraphEndpoint}/${siteConfig.roleDefaultList}/items`;
            await this.dataService.sendGraphPostRequest(this.graphEndpoint, this.props.graph, roleObj);
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "CreateIncident_AddDefaultRoles \n",
                JSON.stringify(error)
            );
            // Log Exception
            this.dataService.trackException(this.props.appInsights, error, constants.componentNames.IncidentDetailsComponent, 'CreateIncident_AddDefaultRoles', this.props.userPrincipalName);

        }
    }

    //update default user roles
    private updateDefaultRoles = async (updateRoleObj: any, itemId: any) => {
        try {
            this.graphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}/lists/${siteConfig.roleDefaultList}/items/${itemId}/fields`;
            await this.dataService.updateItemInList(this.graphEndpoint, this.props.graph, updateRoleObj);
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "CreateIncident_UpdateDefaultRoles \n",
                JSON.stringify(error)
            );
            // Log Exception
            this.dataService.trackException(this.props.appInsights, error, constants.componentNames.IncidentDetailsComponent, 'CreateIncident_UpdateDefaultRoles', this.props.userPrincipalName);

        }
    }

    //set saveDefault for Role Assignment 
    private onChecked = (ev?: React.FormEvent<HTMLElement | HTMLInputElement>, isChecked?: boolean, index?: any) => {
        try {
            let roleAssignment = this.state.roleAssignments;
            roleAssignment[index].saveDefault = isChecked;
            this.setState({
                roleAssignments: roleAssignment
            });
        } catch (ex) {
            console.error(
                constants.errorLogPrefix + "CreateIncident_OnCheckboxChecked \n",
                JSON.stringify(ex)
            );
            // Log Exception
            this.dataService.trackException(this.props.appInsights, ex, constants.componentNames.IncidentDetailsComponent, 'CreateIncident_OnCheckboxChecked', this.props.userPrincipalName);

        }
    };

    //save incidentType default roles
    private saveIncidentTypeDefaultData = async (incidentType: any, roleAssignment: any, roleLead: any,
        cloudStorageLink: string, additionalChannels?: string) => {
        try {
            //check if edit mode & incident type default role data is present 
            if (this.state.isEditMode && this.state.incidentTypeRoleDefaultData.length === 0 &&
                (this.state.saveIncidentTypeDefaultRoleCheck || this.state.saveDefaultCloudStorageLink)) {
                await this.getIncidentTypeDefaultData();
            }
            let duplicateCount = this.state.incidentTypeRoleDefaultData.filter(e => e.incidentType === incidentType);

            if (duplicateCount.length === 0) {
                const newIncidentTypeDefaultDataObj: any = { fields: { Title: incidentType } };
                if (this.state.saveIncidentTypeDefaultRoleCheck) {
                    newIncidentTypeDefaultDataObj.fields.RoleAssignment = roleAssignment.trim();
                    newIncidentTypeDefaultDataObj.fields.RoleLeads = roleLead ? roleLead.trim() : "";
                }
                if (this.state.saveDefaultAdditionalChannels && this.state.toggleAdditionalChannels) {
                    newIncidentTypeDefaultDataObj.fields.AdditionalChannels = additionalChannels;
                }
                if (this.state.saveDefaultCloudStorageLink && this.state.toggleCloudStorageLocation) {
                    newIncidentTypeDefaultDataObj.fields.CloudStorageLink = cloudStorageLink?.trim();
                }
                if (this.state.saveIncidentTypeDefaultRoleCheck ||
                    (this.state.saveDefaultAdditionalChannels && this.state.toggleAdditionalChannels) ||
                    (this.state.saveDefaultCloudStorageLink && this.state.toggleCloudStorageLocation)) {
                    this.addIncidentTypeDefaultRoles(newIncidentTypeDefaultDataObj);
                }
            }
            else {
                const updateIncidentTypeDefaultDataObj: any = {};
                if (this.state.saveIncidentTypeDefaultRoleCheck) {
                    updateIncidentTypeDefaultDataObj.RoleAssignment = roleAssignment.trim();
                    updateIncidentTypeDefaultDataObj.RoleLeads = roleLead ? roleLead.trim() : "";
                }
                if (this.state.saveDefaultAdditionalChannels && this.state.toggleAdditionalChannels) {
                    updateIncidentTypeDefaultDataObj.AdditionalChannels = additionalChannels?.trim();
                }
                if (this.state.saveDefaultCloudStorageLink && this.state.toggleCloudStorageLocation) {
                    updateIncidentTypeDefaultDataObj.CloudStorageLink = cloudStorageLink?.trim();
                }
                if (this.state.saveIncidentTypeDefaultRoleCheck ||
                    (this.state.saveDefaultAdditionalChannels && this.state.toggleAdditionalChannels) ||
                    (this.state.saveDefaultCloudStorageLink && this.state.toggleCloudStorageLocation)) {
                    this.updateIncidentTypeDefaultRoles(updateIncidentTypeDefaultDataObj, duplicateCount[0].itemId);
                }
            }
        } catch (ex) {
            console.error(
                constants.errorLogPrefix + "CreateIncident_saveIncidentTypeDefaultData \n",
                JSON.stringify(ex)
            );
            // Log Exception
            this.dataService.trackException(this.props.appInsights, ex, constants.componentNames.IncidentDetailsComponent, 'CreateIncident_saveIncidentTypeDefaultData', this.props.userPrincipalName);

        }
    }

    //add default roles for an Incident type
    private addIncidentTypeDefaultRoles = async (incidentTypeRoleObj: any) => {
        try {
            this.graphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}${graphConfig.listsGraphEndpoint}/${siteConfig.incidentTypeDefaultRolesList}/items`;
            await this.dataService.sendGraphPostRequest(this.graphEndpoint, this.props.graph, incidentTypeRoleObj);
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "CreateIncident_AddIncidentTypeDefaultRoles \n",
                JSON.stringify(error)
            );
            // Log Exception
            this.dataService.trackException(this.props.appInsights, error, constants.componentNames.IncidentDetailsComponent, 'CreateIncident_AddIncidentTypeDefaultRoles', this.props.userPrincipalName);

        }
    }

    //update default user roles for an Incident Type
    private updateIncidentTypeDefaultRoles = async (incidentTypeRoleObj: any, itemId: any) => {
        try {
            this.graphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}/lists/${siteConfig.incidentTypeDefaultRolesList}/items/${itemId}/fields`;
            await this.dataService.updateItemInList(this.graphEndpoint, this.props.graph, incidentTypeRoleObj);
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "CreateIncident_UpdateIncidentTypeDefaultRoles \n",
                JSON.stringify(error)
            );
            // Log Exception
            this.dataService.trackException(this.props.appInsights, error, constants.componentNames.IncidentDetailsComponent, 'CreateIncident_UpdateIncidentTypeDefaultRoles', this.props.userPrincipalName);

        }
    }

    // create new entry in incident transaction list
    private createNewIncident = async () => {
        this.scrollToTop();
        // incident info object
        let incidentInfo: IncidentEntity = this.state.incDetailsItem;
        this.props.hideMessageBar();
        this.setState({
            showLoader: true,
            formOpacity: 0.5,
            loaderMessage: this.props.localeStrings.genericLoaderMessage,
            inputRegexValidation: this.dataService.getInputRegexValidationInitialState(),
            inputValidation: getInputValidationInitialState()
        });

        // validate for required fields
        if (!this.requiredFieldValidation(incidentInfo)) {
            this.props.showMessageBar(this.props.localeStrings.reqFieldErrorMessage, constants.messageBarType.error);
        }
        else {
            try {
                // validate input strings for incident name and location
                const regexValidation = this.dataService.regexValidation(incidentInfo, this.state.isEditMode);
                if (regexValidation.inputRegexValidationObj.incidentLocationHasError ||
                    regexValidation.inputRegexValidationObj.incidentNameHasError ||
                    this.state.incDetailsItem.additionalTeamChannels.filter((channel: any) => channel.hasRegexError).length > 0 ||
                    this.state.incCommanderHasRegexError ||
                    (this.state.toggleCloudStorageLocation && regexValidation.inputRegexValidationObj.incidentCloudStorageLinkHasError) ||
                    (this.state.toggleGuestUsers && regexValidation.guestUsers.filter((user: any) => user.hasEmailRegexError).length > 0)) {
                    this.props.showMessageBar(this.props.localeStrings.regexErrorMessage, constants.messageBarType.error);
                    this.setState({
                        inputRegexValidation: regexValidation.inputRegexValidationObj,
                        incDetailsItem: {
                            ...this.state.incDetailsItem,
                            guestUsers: regexValidation.guestUsers
                        },
                        showLoader: false,
                        formOpacity: 1
                    });
                }
                else {
                    try {
                        this.setState({
                            loaderMessage: this.props.localeStrings.incidentCreationLoaderMessage
                        });
                        // prepare the role assignment object which will be stored in 
                        // incident transaction list in string format
                        let roleAssignment = "";
                        let roleLead = "";
                        this.state.roleAssignments.forEach(roles => {
                            roleAssignment += roles.role + " : " + roles.userObjString + "; ";
                            roleLead += roles.leadObjString ? roles.role + " : " + roles.leadObjString + "; " : "";
                        });

                        // create object to be passed in graph query
                        const incidentInfoObj: any = {
                            fields: {
                                Title: incidentInfo.incidentName,
                                Description: incidentInfo.incidentDesc,
                                IncidentType: incidentInfo.incidentType,
                                StatusLookupId: incidentInfo.incidentStatus.id,
                                StartDateTime: incidentInfo.startDateTime,
                                Location: JSON.stringify(this.state.selectedLocation),
                                IncidentName: incidentInfo.incidentName,
                                RoleAssignment: roleAssignment.trim(),
                                RoleLeads: roleLead.trim(),
                                IncidentCommander: incidentInfo.incidentCommander.userName + "|" + incidentInfo.incidentCommander.userId + "|" + incidentInfo.incidentCommander.userEmail + ";",
                                Severity: constants.severity[this.state.selectedSeverity],
                                CloudStorageLink: this.state.toggleCloudStorageLocation ? incidentInfo.cloudStorageLink.trim() : ""
                            }
                        }

                        this.graphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}${graphConfig.listsGraphEndpoint}/${siteConfig.incidentsList}/items`;
                        this.setState({
                            loaderMessage: this.props.localeStrings.createIncidentLoaderMessage
                        });
                        const incidentAdded = await this.dataService.sendGraphPostRequest(this.graphEndpoint, this.props.graph, incidentInfoObj);

                        // check if incident is created
                        if (incidentAdded) {
                            console.log(constants.infoLogPrefix + "Incident item created");

                            //call method to add/update default roles
                            this.createUpdateDefaultRoles();

                            //format additional channels into a single string to store 
                            //in IncidentTypeDefaultRoles list
                            let additionalChannels = "";
                            if (this.state.toggleAdditionalChannels && this.state.saveDefaultAdditionalChannels) {
                                additionalChannels = this.state.incDetailsItem.additionalTeamChannels
                                    .map((channel: IAdditionalTeamChannels) => channel.channelName.trim())
                                    .filter((channelName: string) => channelName !== "").join(", ");
                            }

                            //format cloud storage link
                            let cloudStorageLink = "";
                            if (this.state.toggleCloudStorageLocation && this.state.saveDefaultCloudStorageLink) {
                                cloudStorageLink = this.state.incDetailsItem.cloudStorageLink.trim()
                            }
                            //call method to add/update incident type role default values
                            this.saveIncidentTypeDefaultData(incidentInfo.incidentType, roleAssignment,
                                roleLead, cloudStorageLink, additionalChannels);

                            //log trace
                            this.dataService.trackTrace(this.props.appInsights, 'Incident item created ', incidentAdded.id, this.props.userPrincipalName);
                            try {
                                //method to create teams and channels and other related functionalities
                                await this.createTeamAndChannels(incidentAdded.id);

                            } catch (error) {
                                console.error(
                                    constants.errorLogPrefix + "CreateIncident_CreateNewIncident \n",
                                    JSON.stringify(error)
                                );
                                // Log Exception
                                this.dataService.trackException(this.props.appInsights, error, constants.componentNames.IncidentDetailsComponent, 'CreateIncident_CreateNewIncident', this.props.userPrincipalName);
                                // delete the item if error occured
                                await this.deleteIncident(incidentAdded.id);
                                this.setState({
                                    showLoader: false,
                                    formOpacity: 1
                                });
                                this.props.showMessageBar(this.props.localeStrings.genericErrorMessage + ", " + this.props.localeStrings.errMsgForCreateIncident, constants.messageBarType.error);
                            }
                        }
                        else {
                            this.setState({
                                showLoader: false,
                                formOpacity: 1
                            });
                            //log trace
                            this.dataService.trackTrace(this.props.appInsights, 'Incident Creation Failed ', incidentAdded.id, this.props.userPrincipalName);
                            this.props.showMessageBar(this.props.localeStrings.genericErrorMessage + ", " + this.props.localeStrings.errMsgForCreateIncident, constants.messageBarType.error);
                        }
                    } catch (error) {
                        console.error(
                            constants.errorLogPrefix + "CreateIncident_CreateNewIncident \n",
                            JSON.stringify(error)
                        );
                        // Log Exception
                        this.dataService.trackException(this.props.appInsights, error, constants.componentNames.IncidentDetailsComponent, 'CreateIncident_CreateNewIncident', this.props.userPrincipalName);
                        this.setState({
                            showLoader: false,
                            formOpacity: 1
                        });
                        this.props.showMessageBar(this.props.localeStrings.genericErrorMessage + ", " + this.props.localeStrings.errMsgForCreateIncident, constants.messageBarType.error);
                    }
                }
            } catch (error) {
                console.error(
                    constants.errorLogPrefix + "CreateIncident_CreateNewIncident \n",
                    JSON.stringify(error)
                );
                // Log Exception
                this.dataService.trackException(this.props.appInsights, error, constants.componentNames.IncidentDetailsComponent, 'CreateIncident_CreateNewIncident', this.props.userPrincipalName);
            }
        }
    }

    // update the incident in incident transaction list
    private updateIncidentDetails = async () => {
        this.scrollToTop();
        // incident info object
        let incidentInfo: IncidentEntity = this.state.incDetailsItem;
        this.props.hideMessageBar();
        this.setState({
            showLoader: true,
            formOpacity: 0.5,
            loaderMessage: this.props.localeStrings.genericLoaderMessage,
            inputRegexValidation: this.dataService.getInputRegexValidationInitialState(),
            inputValidation: getInputValidationInitialState()
        });

        // validate for required fields
        if (!this.requiredFieldValidation(incidentInfo)) {
            this.props.showMessageBar(this.props.localeStrings.reqFieldErrorMessage, constants.messageBarType.error);
        }
        else {
            try {
                // validate input strings for incident name and location
                const regexValidation = this.dataService.regexValidation(incidentInfo, this.state.isEditMode);
                if (regexValidation.inputRegexValidationObj.incidentLocationHasError ||
                    regexValidation.inputRegexValidationObj.incidentNameHasError ||
                    this.state.incCommanderHasRegexError ||
                    (this.state.toggleCloudStorageLocation && regexValidation.inputRegexValidationObj.incidentCloudStorageLinkHasError) ||
                    (this.state.toggleGuestUsers && regexValidation.guestUsers.filter((user: any) => user.hasEmailRegexError).length > 0)) {
                    this.props.showMessageBar(this.props.localeStrings.regexErrorMessage, constants.messageBarType.error);
                    this.setState({
                        inputRegexValidation: regexValidation.inputRegexValidationObj,
                        incDetailsItem: {
                            ...this.state.incDetailsItem,
                            guestUsers: regexValidation.guestUsers
                        },
                        showLoader: false,
                        formOpacity: 1
                    });
                }
                else {
                    try {
                        this.setState({
                            loaderMessage: this.props.localeStrings.incidentCreationLoaderMessage
                        });
                        // prepare the role assignment object which will be stored in 
                        // incident transaction list in string format
                        let roleAssignment = "";
                        let roleLead = "";
                        this.state.roleAssignments.forEach(roles => {
                            roleAssignment += roles.role + " : " + roles.userObjString + "; ";
                            roleLead += roles.leadObjString ? roles.role + " : " + roles.leadObjString + "; " : "";
                        });


                        // create object to be passed in graph query
                        const incidentInfoObj: any = {
                            Description: incidentInfo.incidentDesc,
                            StatusLookupId: incidentInfo.incidentStatus.id,
                            Location: JSON.stringify(this.state.selectedLocation),
                            IncidentName: incidentInfo.incidentName,
                            IncidentCommander: incidentInfo.incidentCommander.userName + "|" + incidentInfo.incidentCommander.userId + "|" + incidentInfo.incidentCommander.userEmail,
                            RoleAssignment: roleAssignment.trim(),
                            RoleLeads: roleLead.trim(),
                            Severity: constants.severity[this.state.selectedSeverity],
                            ReasonForUpdate: incidentInfo.reasonForUpdate,
                            CloudStorageLink: this.state.toggleCloudStorageLocation ? incidentInfo.cloudStorageLink.trim() : ""
                        }

                        this.graphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}/lists/${siteConfig.incidentsList}/items/${this.state.incDetailsItem.incidentId}/fields`;
                        this.setState({
                            loaderMessage: this.props.localeStrings.createIncidentLoaderMessage
                        });
                        const incidentUpdated = await this.dataService.updateItemInList(this.graphEndpoint, this.props.graph, incidentInfoObj);

                        // check if incident is updated
                        if (incidentUpdated) {
                            let incDetails = this.state.incDetailsItem;

                            //Invite Guest Users
                            let returnInvitationObj: any;
                            if (this.state.toggleGuestUsers) {
                                const teamEndpoint = graphConfig.teamsGraphEndpoint + "/" + this.state.teamGroupId;
                                const teamObj = await this.dataService.getGraphData(teamEndpoint, this.props.graph);
                                returnInvitationObj = this.sendInvitation(this.state.teamGroupId, teamObj.displayName,
                                    this.props.incidentData?.teamWebURL ? this.props.incidentData?.teamWebURL : "");
                            }


                            //call method to add/update default roles
                            this.createUpdateDefaultRoles();

                            //format cloud storage link
                            let cloudStorageLink = "";
                            if (this.state.toggleCloudStorageLocation && this.state.saveDefaultCloudStorageLink) {
                                cloudStorageLink = this.state.incDetailsItem.cloudStorageLink.trim();
                            }

                            //call method to add/update incident type role default values
                            this.saveIncidentTypeDefaultData(incidentInfo.incidentType, roleAssignment, roleLead, cloudStorageLink);

                            // update the date format
                            incDetails.startDateTime = moment(this.state.incDetailsItem.startDateTime).format("DDMMMYYYY");

                            // update team display name
                            let teamDisplayName = this.formatTeamDisplayName(this.state.incDetailsItem.incidentId, this.state.incDetailsItem);
                            this.graphEndpoint = graphConfig.teamsGraphEndpoint + "/" + this.state.teamGroupId;
                            await this.dataService.sendGraphPatchRequest(this.graphEndpoint, this.props.graph, { "displayName": teamDisplayName })

                            const usersObj = this.compareTeamsMembership(this.props.existingTeamMembers);

                            // Get all existing tags
                            let existingTagsList = await this.getAllTags();

                            if (this.state.existingIncCommander.userId !== this.state.incDetailsItem.incidentCommander.userId) {
                                // Add incident commander as owner
                                await this.addUsersToTeam([this.state.incDetailsItem.incidentCommander], true);

                                // add incident commander to tag
                                await this.addUsersToTag([this.state.incDetailsItem.incidentCommander.userId], existingTagsList.value, true);

                            }

                            // check if there are users to remove
                            if (usersObj.removedMembershipIds.length > 0) {
                                // remove users from Team
                                await this.removeUsersFromTeam(usersObj.removedMembershipIds, true);
                            }

                            // check if there are secondary commanders and "Edit Access Role" users to add
                            if (usersObj.newSecondaryIncidentCommanders.length > 0) {
                                // add secondary incident commanders and "Edit Access Role" users as owners
                                await this.addUsersToTeam(usersObj.newSecondaryIncidentCommanders, true);
                            }

                            // check if there are users to add
                            if (usersObj.newAddedUsers.length > 0) {
                                // Add other users as member to Team
                                await this.addUsersToTeam(usersObj.newAddedUsers, false);
                            }

                            // Get all tags after the membership is updated on team
                            let tagsList = await this.getAllTags();

                            // check and get if new tags needs to be created
                            const newRole = this.checkIfNewTagCreationNeeded(tagsList.value);

                            if (newRole.length > 0) {
                                // create the role object from role assignments needed for tag creation
                                const roles = this.createNewRoleObject(newRole);
                                // create the tag for new role
                                await this.createTagObject(this.state.teamGroupId, roles);
                            }

                            const usersForTags: any = [];
                            this.state.roleAssignments.forEach(roles => {
                                roles.userDetailsObj.forEach(users => {
                                    usersForTags.push({ role: roles.role, userId: users.userId });
                                })
                                if (roles.leadDetailsObj !== undefined) {
                                    roles.leadDetailsObj.forEach(lead => {
                                        usersForTags.push({ role: roles.role, userId: lead.userId });
                                    })
                                }
                            });
                            await this.addUsersToTag(usersForTags, tagsList.value, false);

                            console.log(constants.infoLogPrefix + "Incident Updated");
                            //log trace
                            this.dataService.trackTrace(this.props.appInsights, 'Incident Updated', this.state.incDetailsItem.incidentId, this.props.userPrincipalName);
                            Promise.allSettled([returnInvitationObj]).then((promiseObj: any) => {
                                this.setState({
                                    showLoader: false,
                                    formOpacity: 1
                                });

                                // Display success message if incident updated successfully
                                this.props.showMessageBar(this.props.localeStrings.updateSuccessMessage,
                                    constants.messageBarType.success);
                                // Display error message if send invitation fails
                                if ((promiseObj[0]?.value !== undefined && !promiseObj[0].value?.isAllSucceeded))
                                    this.props.showMessageBar(
                                        ((promiseObj[0]?.value !== undefined && !promiseObj[0]?.value?.isAllSucceeded) ? " " + promiseObj[0]?.value?.message : ""),
                                        constants.messageBarType.error);
                                this.props.onBackClick(constants.messageBarType.success);
                            });

                        }
                        else {
                            this.setState({
                                showLoader: false,
                                formOpacity: 1
                            });
                            this.props.showMessageBar(this.props.localeStrings.genericErrorMessage + ", " + this.props.localeStrings.errMsgForUpdateIncident, constants.messageBarType.error);
                        }
                    } catch (error) {
                        console.error(
                            constants.errorLogPrefix + "UpdateIncident_UpdateIncident \n",
                            JSON.stringify(error)
                        );
                        // Log Exception
                        this.dataService.trackException(this.props.appInsights, error, constants.componentNames.IncidentDetailsComponent, 'UpdateIncident_UpdateIncident', this.props.userPrincipalName);
                        this.setState({
                            showLoader: false,
                            formOpacity: 1
                        });
                        this.props.showMessageBar(this.props.localeStrings.genericErrorMessage + ", " + this.props.localeStrings.errMsgForUpdateIncident, constants.messageBarType.error);
                    }
                }
            } catch (error) {
                console.error(
                    constants.errorLogPrefix + "UpdateIncident_UpdateIncident \n",
                    JSON.stringify(error)
                );
                // Log Exception
                this.dataService.trackException(this.props.appInsights, error, constants.componentNames.IncidentDetailsComponent, 'UpdateIncident_UpdateIncident', this.props.userPrincipalName);
            }
        }
    }

    // perform required fields validation
    private requiredFieldValidation = (incidentInfo: IncidentEntity): boolean => {
        let inputValidationObj = getInputValidationInitialState();
        let reqFieldValidationSuccess = true;
        if (incidentInfo.incidentName === "" || incidentInfo.incidentName === undefined ||
            (incidentInfo.incidentName !== undefined && incidentInfo.incidentName.trim() === "")) {
            inputValidationObj.incidentNameHasError = true;
        }
        if (incidentInfo.incidentType === "" || incidentInfo.incidentType === undefined || incidentInfo.incidentType === null) {
            inputValidationObj.incidentTypeHasError = true;
        }
        if (incidentInfo.startDateTime === "" || incidentInfo.startDateTime === undefined) {
            inputValidationObj.incidentStartDateTimeHasError = true;
        }
        if (incidentInfo.incidentStatus.status === "" || incidentInfo.incidentStatus === undefined) {
            inputValidationObj.incidentStatusHasError = true;
        }
        if (incidentInfo.incidentDesc === "" || incidentInfo.incidentDesc === undefined ||
            (incidentInfo.incidentDesc !== undefined && incidentInfo.incidentDesc.trim() === "")) {
            inputValidationObj.incidentDescriptionHasError = true;
        }
        if ((incidentInfo.incidentCommander === undefined || incidentInfo.incidentCommander.userName === '')) {
            inputValidationObj.incidentCommandarHasError = true;
        }
        if (this.props.isEditMode && (incidentInfo.reasonForUpdate === "" || incidentInfo.reasonForUpdate === undefined ||
            (incidentInfo.reasonForUpdate !== undefined && incidentInfo.reasonForUpdate.trim() === ""))) {
            inputValidationObj.incidentReasonForUpdateHasError = true;
        }
        if (this.state.toggleCloudStorageLocation && (incidentInfo.cloudStorageLink?.trim() === "" ||
            incidentInfo.cloudStorageLink === undefined)) {
            inputValidationObj.cloudStorageLinkHasError = true;
        }
        if (this.state.toggleGuestUsers && (incidentInfo.guestUsers?.filter((user: IGuestUsers) =>
            user.displayName.trim() !== "" && user.email.trim() !== "")).length === 0) {
            inputValidationObj.guestUsersHasError = true;
        }
        if (this.state.selectedLocation === "" || this.state.selectedLocation === undefined || this.state.selectedLocation === null) {
            inputValidationObj.incidentLocationHasError = true;
        }
        const guestUsers = incidentInfo.guestUsers;
        if (this.state.toggleGuestUsers) {
            guestUsers?.forEach((user: IGuestUsers, idx: number) => {
                if (user.email.trim() !== "" && user.displayName.trim() === "")
                    guestUsers[idx].hasDisplayNameValidationError = true
                else guestUsers[idx].hasDisplayNameValidationError = false
                if (user.email.trim() === "" && user.displayName.trim() !== "")
                    guestUsers[idx].hasEmailValidationError = true;
                else guestUsers[idx].hasEmailValidationError = false
            });
        }

        if (inputValidationObj.incidentNameHasError || inputValidationObj.incidentTypeHasError ||
            inputValidationObj.incidentStartDateTimeHasError || inputValidationObj.incidentStatusHasError || inputValidationObj.incidentDescriptionHasError ||
            inputValidationObj.incidentCommandarHasError || inputValidationObj.incidentLocationHasError || inputValidationObj.incidentReasonForUpdateHasError ||
            inputValidationObj.cloudStorageLinkHasError || inputValidationObj.guestUsersHasError ||
            (incidentInfo.guestUsers?.filter((user) =>
                (user.hasDisplayNameValidationError || user.hasEmailValidationError)).length > 0)) {
            this.setState({
                inputValidation: inputValidationObj,
                incDetailsItem: {
                    ...this.state.incDetailsItem,
                    guestUsers: incidentInfo.guestUsers !== undefined ? [...guestUsers] : []
                },
                showLoader: false,
                formOpacity: 1
            });
            reqFieldValidationSuccess = false;
        }
        this.setFocusAfterValidation(inputValidationObj);
        return reqFieldValidationSuccess;
    }

    // to set the focus on form after validation
    // to set the focus on non-filled required field after validation 
    private setFocusAfterValidation = (inputValidationObj: any) => {
        if (inputValidationObj.incidentNameHasError) {
            this.incidentName.current?.focus();
        } else if (inputValidationObj.incidentTypeHasError) {
            this.incidentType.current?.querySelector("input").focus();
        } else if (inputValidationObj.incidentStartDateTimeHasError) {
            this.incidentStartDateTime.current?.querySelector("input").focus();
        } else if (inputValidationObj.incidentCommandarHasError) {
            const shadow = this.incidentCommandar.current?.shadowRoot;
            let childInput: HTMLElement = shadow.getElementById("people-picker-input");
            childInput.focus();
        } else if (inputValidationObj.incidentLocationHasError) {
            this.incidentLocation.current?.focus();
        } else if (inputValidationObj.incidentDescriptionHasError) {
            this.incidentDescription.current?.focus();
        }
    }


    //delay the operation by adding timeout
    private timeout = (delay: number): Promise<any> => {
        return new Promise(res => setTimeout(res, delay));
    }

    // wrapper method to perform teams related operations
    private async createTeamAndChannels(incidentId: any) {
        try {
            console.log(constants.infoLogPrefix + "M365 group creation starts");
            // call method to create Teams group
            const groupInfo = await this.createTeamGroup(incidentId);
            try {
                console.log(constants.infoLogPrefix + "M365 group created");

                // create associated team with the group
                const teamInfo = await this.createTeam(groupInfo);
                if (teamInfo.status) {
                    //log trace
                    this.dataService.trackTrace(this.props.appInsights, "Incident Team created ", incidentId, this.props.userPrincipalName);

                    //Send invitations to the guest users
                    let returnInvitationObj: any;
                    if (this.state.toggleGuestUsers)
                        returnInvitationObj = this.sendInvitation(groupInfo.id, teamInfo.data.displayName, teamInfo.data.webUrl)

                    // create channels
                    await this.createChannels(teamInfo.data);

                    this.setState({ loaderMessage: this.props.localeStrings.createPlanloaderMessage });

                    //Get General channel id
                    const generalChannelId = await this.dataService.getChannelId(this.props.graph,
                        groupInfo.id, constants.General);

                    //Create planner with the Group ID                        
                    const planID = await this.dataService.createPlannerPlan(groupInfo.id, incidentId, this.props.graph,
                        this.props.graphContextURL, this.props.tenantID, generalChannelId, false);

                    //Add TEOC app to the Incident Team General channel's Active Dashboard Tab
                    await this.dataService.createActiveDashboardTab(this.props.graph, groupInfo.id,
                        generalChannelId, this.props.graphContextURL, this.props.appSettings);

                    //added for GCCH tenant
                    if (this.props.graphBaseUrl !== constants.defaultGraphBaseURL) {
                        // wait for 5 seconds to ensure the SharePoint site is available via graph API
                        await this.timeout(5000);
                    }

                    // graph endpoint to get team site Id
                    const teamSiteURLGraphEndpoint = graphConfig.teamGroupsGraphEndpoint + "/" +
                        groupInfo.id + graphConfig.rootSiteGraphEndpoint;

                    // retrieve team site details
                    const teamSiteDetails = await this.dataService.getGraphData(teamSiteURLGraphEndpoint, this.props.graph);

                    //get the team site managed path
                    const teamSiteManagedPathURL = teamSiteDetails.webUrl.split(teamSiteDetails.siteCollection.hostname)[1];
                    console.log(constants.infoLogPrefix + "Site ManagedPath", teamSiteManagedPathURL);

                    // create news channel and tab
                    const newsTabLink = await this.createNewsTab(groupInfo, teamSiteDetails.webUrl, teamSiteManagedPathURL);

                    // create assessment channel and tab
                    await this.createAssessmentChannelAndTab(groupInfo.id, teamSiteDetails.webUrl, teamSiteManagedPathURL);

                    // call method to create assessment list
                    await this.createAssessmentList(groupInfo.mailNickname, teamSiteDetails.id);

                    //log trace
                    this.dataService.trackTrace(this.props.appInsights, "Assessment list created ", incidentId, this.props.userPrincipalName);

                    //change the M365 group visibility to Private for GCCH tenant
                    if (this.props.graphBaseUrl !== constants.defaultGraphBaseURL) {
                        this.graphEndpoint = graphConfig.teamGroupsGraphEndpoint + "/" + groupInfo.id;
                        await this.dataService.sendGraphPatchRequest(this.graphEndpoint, this.props.graph, { "visibility": "Private" })
                        console.log(constants.infoLogPrefix + "Group setting changed to Private");
                    }

                    //Update Team details, Plan ID, NewsTabLink in Incident Transation List                                   
                    const updateItemObj = {
                        IncidentId: incidentId,
                        TeamWebURL: teamInfo.data.webUrl,
                        PlanID: planID,
                        NewsTabLink: newsTabLink
                    };

                    await this.updateIncidentItemInList(incidentId, updateItemObj);
                    console.log(constants.infoLogPrefix + "List item updated");

                    let roles: any = this.state.roleAssignments;
                    roles.push({
                        role: constants.incidentCommanderRoleName,
                        userNamesString: this.state.incDetailsItem.incidentCommander.userName,
                        userDetailsObj: [this.state.incDetailsItem.incidentCommander]
                    });

                    //post incident message in General Channel
                    await this.postIncidentMessage(groupInfo.id);

                    // create the tags for incident commander and each selected roles                        
                    await this.createTagObject(teamInfo.data.id, roles);
                    //log trace
                    this.dataService.trackTrace(this.props.appInsights, "Tags are created ", incidentId, this.props.userPrincipalName);
                    Promise.allSettled([returnInvitationObj]).then((promiseObj: any) => {
                        this.setState({
                            showLoader: false,
                            formOpacity: 1
                        });

                        // Display success message if incident updated successfully
                        this.props.showMessageBar(this.props.localeStrings.incidentCreationSuccessMessage, constants.messageBarType.success);

                        // Display error message if guest invitations
                        if ((promiseObj[0]?.value !== undefined && !promiseObj[0].value?.isAllSucceeded))
                            this.props.showMessageBar(
                                ((promiseObj[0]?.value !== undefined && !promiseObj[0]?.value?.isAllSucceeded) ? " " + promiseObj[0]?.value?.message + ". " : ""),
                                constants.messageBarType.error);
                        this.props.onBackClick(constants.messageBarType.success);
                    });

                }
                else {
                    // delete the group if some error occured
                    await this.deleteTeamGroup(groupInfo.id);
                    // delete the item if error occured
                    await this.deleteIncident(incidentId);

                    this.setState({
                        showLoader: false,
                        formOpacity: 1
                    })
                    this.props.showMessageBar(this.props.localeStrings.genericErrorMessage + ", " + this.props.localeStrings.errMsgForCreateIncident, constants.messageBarType.error);
                }
            }
            catch (error) {
                console.error(
                    constants.errorLogPrefix + "CreateIncident_createTeamAndChannels \n",
                    JSON.stringify(error)
                );
                // Log Exception
                this.dataService.trackException(this.props.appInsights, error, constants.componentNames.IncidentDetailsComponent, 'CreateIncident_createTeamAndChannels', this.props.userPrincipalName);
                // delete the group if some error occured
                await this.deleteTeamGroup(groupInfo.id);
                // delete the item if error occured
                await this.deleteIncident(incidentId);

                this.setState({
                    showLoader: false,
                    formOpacity: 1
                })
                this.props.showMessageBar(this.props.localeStrings.genericErrorMessage + ", " + this.props.localeStrings.errMsgForCreateIncident, constants.messageBarType.error);
            }
        }
        catch (error: any) {
            console.error(
                constants.errorLogPrefix + "CreateIncident_createTeamAndChannels \n",
                JSON.stringify(error)
            );
            // Log Exception
            this.dataService.trackException(this.props.appInsights, error, constants.componentNames.IncidentDetailsComponent, 'CreateIncident_createTeamAndChannels', this.props.userPrincipalName);

            // delete the item if error occured
            this.deleteIncident(incidentId);

            this.setState({
                showLoader: false,
                formOpacity: 1
            });

            // Display error message if M365 group creation fails with access denied error
            if (error?.statusCode === 403 && error?.code === constants.authorizationRequestDenied
                && error?.message === constants.groupCreationAccessDeniedErrorMessage) {
                this.props.showMessageBar(this.props.localeStrings.genericErrorMessage + ", " + this.props.localeStrings.m365GroupCreationFailedMessage, constants.messageBarType.error);
            }
            /* Display error message if M365 group creation fails with group already exists error 
            or any other error */
            else {
                this.props.showMessageBar(this.props.localeStrings.genericErrorMessage + ", " + this.props.localeStrings.errMsgForCreateIncident, constants.messageBarType.error);
            }
        }
    }

    //Send invitation to guest users
    private async sendInvitation(teamId: string, teamName: string, teamWebURL: string): Promise<any> {
        const returnInvitationObj = {
            isAllSucceeded: true,
            message: ""
        }
        return new Promise<any>(async (resolve: any) => {
            try {
                this.setState({
                    loaderMessage: this.props.localeStrings.createInvitationLoaderMessage
                });

                let guestIds: any[] = [];

                //Removing duplicates and Filtering valid Guest users email and display name 
                const validGuestUsers = this.state.incDetailsItem.guestUsers.filter((user, index, self) =>
                (user.email.trim() !== "" && user.displayName.trim() !== "" &&
                    index === self.findIndex((userData) => userData.email.trim().toLocaleLowerCase() === user.email.trim().toLocaleLowerCase())));

                //Getting Incident Team Members Data
                const groupMembersEndpoint = graphConfig.teamGroupsGraphEndpoint + "/" + teamId + graphConfig.membersGraphEndpoint;
                const teamGroupMembers = await this.dataService.getGraphData(groupMembersEndpoint, this.props.graph);

                const addMemberEndpoint = graphConfig.teamGroupsGraphEndpoint + "/" + teamId + "/members/$ref";

                //Send guest user invitations
                guestIds = await Promise.allSettled(validGuestUsers.map(async (user: IGuestUsers) => {
                    let inviteRequest = {
                        "invitedUserEmailAddress": user.email.trim(),
                        "invitedUserDisplayName": user.displayName.trim(),
                        "inviteRedirectUrl": this.props.graphBaseUrl !== constants.defaultGraphBaseURL ? constants.teamsWebUrlGCCH : constants.teamsWebUrl
                    }
                    try {
                        //Adding guest users to Azure
                        const res = await this.dataService.sendGraphPostRequest(graphConfig.invitationsGraphEndpoint, this.props.graph, inviteRequest);
                        const isUserNotExistsInTeam = teamGroupMembers?.value?.findIndex((member: any) => member?.id === res.invitedUser.id) === -1;
                        if (isUserNotExistsInTeam) {
                            //Adding guest users to Office 365 group
                            const userToAdd = {
                                "@odata.id": this.state.graphContextURL + graphConfig.usersGraphEndpoint + "/" + res.invitedUser.id
                            }
                            await this.dataService.sendGraphPostRequest(addMemberEndpoint, this.props.graph, userToAdd);
                            console.log(constants.infoLogPrefix + "Guest user ", user.email, " added to group");
                            //Sending invitation mails to guest users
                            const mailObj = this.getEmailTemplate(teamName, this.state.incDetailsItem.incidentDesc.trim(),
                                teamWebURL, user.email);
                            const mailEndpoint = graphConfig.emailInvitationsGraphEndpoint;
                            await this.dataService.sendGraphPostRequest(mailEndpoint, this.props.graph, mailObj);
                        }
                        return { userId: res.invitedUser.id, email: user.email.trim() };
                    }
                    catch (err: any) {
                        console.log(err);
                        return { email: user.email.trim(), statusCode: err.statusCode };
                    }
                }));

                let blockedEntries: any = [];
                let isGuestInviteAccessDenied: boolean = false;

                //get valid user urls to add them to M365 group
                guestIds.forEach((user: any) => {
                    if (isGuestInviteAccessDenied) return
                    if (user?.value?.userId === undefined) {
                        if (user?.value?.statusCode !== 403) {
                            if (user?.value?.statusCode === 400) blockedEntries.push(user?.value?.email);
                        }
                        else isGuestInviteAccessDenied = true
                    }
                });

                if (isGuestInviteAccessDenied) {
                    returnInvitationObj.isAllSucceeded = false;
                    returnInvitationObj.message = this.props.localeStrings.guestInvitesAccessDeniedError;
                }
                else if (blockedEntries.length > 0) {
                    returnInvitationObj.isAllSucceeded = false;
                    returnInvitationObj.message = this.props.localeStrings.guestInvitesBlockedUserError + " : " + blockedEntries.join(", ");
                }
                else returnInvitationObj.isAllSucceeded = true

                resolve(returnInvitationObj);
            }
            catch (error) {
                console.error(
                    constants.errorLogPrefix + "CreateIncident_sendInvitation \n",
                    JSON.stringify(error)
                );
                // Log Exception
                this.dataService.trackException(this.props.appInsights, error, constants.componentNames.IncidentDetailsComponent, 'CreateIncident_sendInvitation', this.props.userPrincipalName);
                returnInvitationObj.isAllSucceeded = false;
                returnInvitationObj.message = this.props.localeStrings.guestInvitesBlockedUserError;
                resolve(returnInvitationObj);
            }
        });
    }

    //Get Email template to send invitation to guest users
    private getEmailTemplate(teamName: string, teamDescription: string, teamWebUrl: string,
        userMailId: string) {
        return ({
            message: {
                subject: 'You have been added to a team in Microsoft Teams',
                body: {
                    contentType: 'HTML',
                    content: renderToStaticMarkup(
                        <React.Fragment>
                            <table style={{ marginLeft: "auto", marginRight: "auto" }}>
                                <td style={{
                                    backgroundColor: "#EEF1F5", borderTop: "4px solid #4F52B2", padding: "15px 15px 20px 15px",
                                    width: 604, minHeight: 450
                                }}>
                                    <dt style={{
                                        font: "normal normal normal 24px/32px Segoe UI", letterSpacing: "0px",
                                        color: "#4F52B2", opacity: 1, textAlign: "center", margin: "10px auto"
                                    }}>
                                        Microsoft Teams
                                    </dt>
                                    <dt style={{
                                        font: "normal normal 600 18px/24px Segoe UI", letterSpacing: "0px",
                                        color: "#242424", opacity: 1, textAlign: "center", margin: "10px auto"
                                    }}>
                                        System added you to the {teamName} of Teams Emergency Operations Center!
                                    </dt>
                                    <br />
                                    <table style={{
                                        display: "flex", justifyContent: "center", marginTop: 10
                                    }} >
                                        <tbody style={{ marginLeft: "auto", marginRight: "auto" }}>
                                            <td style={{
                                                opacity: 1, width: 280, backgroundColor: "#FFFFFF",
                                                minHeight: 190, padding: 10
                                            }}>
                                                <p style={{
                                                    font: "normal normal 600 18px/24px Segoe UI", textAlign: "center",
                                                    letterSpacing: 0, color: "#242424", opacity: 1
                                                }}>
                                                    {teamName} team!</p>
                                                <hr />
                                                <p style={{
                                                    font: "normal normal normal 14px/19px Segoe UI", textAlign: "center",
                                                    letterSpacing: 0, color: "#424242", opacity: 1, marginBottom: 10
                                                }}>
                                                    {teamDescription}
                                                </p>
                                            </td>
                                        </tbody>
                                    </table>
                                    <br />
                                    <table style={{ display: "flex", justifyContent: "center" }} >
                                        <tbody style={{ marginLeft: "auto", marginRight: "auto" }}>
                                            <td style={{
                                                backgroundColor: "#4F52B2", opacity: 1, width: 404, height: 30,
                                                textAlign: "center", marginTop: "10px",
                                                marginBottom: "10px"
                                            }}
                                            >
                                                <a
                                                    href={teamWebUrl} target="_blank" rel="noreferrer"
                                                    style={{
                                                        color: "#FFFFFF", textDecoration: "none", width: 404,
                                                        height: 30, display: "block", paddingTop: 7, paddingBottom: 3
                                                    }}
                                                    title="Open Microsoft Teams"
                                                >
                                                    Open Microsoft Teams
                                                </a>
                                            </td>
                                        </tbody>
                                    </table>
                                </td>
                            </table>
                        </React.Fragment> as any
                    )
                },
                toRecipients: [
                    {
                        emailAddress: {
                            address: userMailId
                        }
                    }
                ]
            },
            saveToSentItems: 'false'
        });
    }

    // Check string is JSON or not
    private isJSON = (str: string) => {
        try {
            JSON.parse(str);
        } catch (e) {
            return false;
        }
        return true;
    }

    // get the location display name
    private getLocationDisplayName = (location: any): string => {
        let displayName: string;
        if (typeof location === "string") {
            if (this.isJSON(location)) {
                displayName = JSON.parse(location).DisplayName;
            } else {
                displayName = location;
            }
        } else {
            displayName = location.DisplayName;
        }
        return displayName;
    }


    //post adaptive card message to General channel
    private async postIncidentMessage(teamGroupId: any) {
        try {
            //get the Team display name to @mention in the adaptive card
            const response = await this.dataService.getGraphData(
                graphConfig.teamsGraphEndpoint + "/" + teamGroupId, this.props.graph);
            const teamDisplayName = response.displayName;

            //get General channel ID        
            let generalChannelId = await this.dataService.getChannelId(this.props.graph, teamGroupId, constants.General);

            await this.sendMessage(teamDisplayName, teamGroupId, generalChannelId);
            console.log(constants.infoLogPrefix + "Adaptive Card message is posted to General channel");
        }
        catch (error) {
            console.error(
                constants.errorLogPrefix + "CreateIncident_postIncidentMessage \n",
                JSON.stringify(error)
            );
        }
    }

    //create and send adaptive card to General channel
    private async sendMessage(teamDisplayName: string, teamGroupId: any, channelId: string) {
        try {
            let cardBody = {
                "body": {
                    "contentType": "html",
                    "content": "<at id='0'>" + teamDisplayName + "</at> - This new team has been created to respond to the following incident: <attachment id='9649cac0-49bb-4406-9527-5d328f255750'></attachment>",
                },
                "attachments": [
                    {
                        "id": "9649cac0-49bb-4406-9527-5d328f255750",
                        "contentType": "application/vnd.microsoft.card.adaptive",
                        "content": JSON.stringify({
                            "type": "AdaptiveCard",
                            "body": [
                                {
                                    "type": "ColumnSet",
                                    "columns": [
                                        {
                                            "type": "Column",
                                            "items": [
                                                {
                                                    "type": "TextBlock",
                                                    "text": "**Incident Name:**  " + this.state.incDetailsItem.incidentName,
                                                    "wrap": true
                                                },
                                                {
                                                    "type": "TextBlock",
                                                    "spacing": "None",
                                                    "text": "**Severity:**  " + constants.severity[this.state.selectedSeverity],
                                                    "wrap": true
                                                },
                                                {
                                                    "type": "TextBlock",
                                                    "spacing": "None",
                                                    "text": "**Location:**  " + this.getLocationDisplayName(this.state.selectedLocation),
                                                    "wrap": true
                                                },
                                                {
                                                    "type": "TextBlock",
                                                    "isVisible": this.state.toggleCloudStorageLocation ? true : false,
                                                    "spacing": "None",
                                                    "text": "**Cloud Storage:**   [Open Link](" + this.state.incDetailsItem.cloudStorageLink + ")",
                                                    "wrap": true
                                                }
                                            ],
                                            "width": "stretch"
                                        }
                                    ]
                                },
                                {
                                    "type": "TextBlock",
                                    "wrap": true,
                                    "text": "Please reach out to **" + this.state.incDetailsItem.incidentCommander.userName + "**, your Incident Commander, for any additional information and engage here in this team for response."
                                }
                            ],
                            "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                            "version": "1.5"
                        })
                    },
                ],
                "mentions": [
                    {
                        "id": 0,
                        "mentionText": teamDisplayName,
                        "mentioned": {
                            "conversation": {
                                "id": teamGroupId,
                                "displayName": teamDisplayName,
                                "conversationIdentityType": "team"
                            }
                        }
                    }
                ],
            };

            const endpoint = graphConfig.teamsGraphEndpoint + "/" + teamGroupId +
                graphConfig.channelsGraphEndpoint + "/" + channelId + graphConfig.messagesGraphEndpoint;
            await this.dataService.sendGraphPostRequest(endpoint, this.props.graph, cardBody);
        }
        catch (error) {
            console.error(
                constants.errorLogPrefix + "CreateIncident_sendMessage \n",
                JSON.stringify(error)
            );
        }
    }


    // update item in Incident Transaction list
    private updateIncidentItemInList = async (itemId: number, updateItemObj: any): Promise<any> => {
        try {
            this.graphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}/lists/${siteConfig.incidentsList}/items/${itemId}/fields`;
            const updatedIncident = await this.dataService.updateItemInList(this.graphEndpoint, this.props.graph, updateItemObj);
            return updatedIncident;
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "CreateIncident_UpdatedTeamIdInList \n",
                JSON.stringify(error)
            );
            // Log Exception
            this.dataService.trackException(this.props.appInsights, error, constants.componentNames.IncidentDetailsComponent, 'CreateIncident_UpdatedTeamIdInList', this.props.userPrincipalName);
        }
    }

    // create a Teams group
    private createTeamGroup = async (incId: string) => {
        try {
            this.setState({
                loaderMessage: this.props.localeStrings.createGroupLoaderMessage
            });
            let incDetails = this.state.incDetailsItem;
            // update the date format
            incDetails.startDateTime = moment(this.state.incDetailsItem.startDateTime).format("DDMMMYYYY");

            // create an array for owners and members
            const ownerArr: any = [];
            const membersArr: any = [];

            //adding Incident Commander as Owner and Member to the group
            ownerArr.push(this.state.graphContextURL + graphConfig.usersGraphEndpoint + "/" + incDetails.incidentCommander.userId);
            membersArr.push(this.state.graphContextURL + graphConfig.usersGraphEndpoint + "/" + incDetails.incidentCommander.userId);

            this.state.roleAssignments.forEach(roles => {
                //adding users of secondary incident commander role and "Edit Access Role" as owners and members
                if (roles.role === constants.secondaryIncidentCommanderRole || roles.role === this.props.editIncidentAccessRole ) {
                    roles.userDetailsObj.forEach(user => {
                        if (ownerArr.indexOf(this.state.graphContextURL + graphConfig.usersGraphEndpoint + "/" + user.userId) === -1) {
                            ownerArr.push(this.state.graphContextURL + graphConfig.usersGraphEndpoint + "/" + user.userId);
                        }
                        if (membersArr.indexOf(this.state.graphContextURL + graphConfig.usersGraphEndpoint + "/" + user.userId) === -1) {
                            membersArr.push(this.state.graphContextURL + graphConfig.usersGraphEndpoint + "/" + user.userId);
                        }
                    });
                    if (roles.leadDetailsObj !== undefined) {
                        roles.leadDetailsObj.forEach(lead => {
                            if (ownerArr.indexOf(this.state.graphContextURL + graphConfig.usersGraphEndpoint + "/" + lead.userId) === -1) {
                                ownerArr.push(this.state.graphContextURL + graphConfig.usersGraphEndpoint + "/" + lead.userId);
                            }
                            if (membersArr.indexOf(this.state.graphContextURL + graphConfig.usersGraphEndpoint + "/" + lead.userId) === -1) {
                                membersArr.push(this.state.graphContextURL + graphConfig.usersGraphEndpoint + "/" + lead.userId);
                            }
                        });
                    }
                } // adding users of other roles in role assignments as members
                else {
                    roles.userDetailsObj.forEach(user => {
                        if (membersArr.indexOf(this.state.graphContextURL + graphConfig.usersGraphEndpoint + "/" + user.userId) === -1) {
                            membersArr.push(this.state.graphContextURL + graphConfig.usersGraphEndpoint + "/" + user.userId);
                        }
                    });
                    if (roles.leadDetailsObj !== undefined) {
                        roles.leadDetailsObj.forEach(lead => {
                            if (membersArr.indexOf(this.state.graphContextURL + graphConfig.usersGraphEndpoint + "/" + lead.userId) === -1) {
                                membersArr.push(this.state.graphContextURL + graphConfig.usersGraphEndpoint + "/" + lead.userId);
                            }
                        });
                    }
                }
            });


            // add current user as a owner if already not present so that we can perform teams creation
            // and sharepoint site related operations on associated team site
            if (ownerArr.indexOf(this.state.graphContextURL + graphConfig.usersGraphEndpoint + "/" + this.props.currentUserId) === -1) {
                ownerArr.push(this.state.graphContextURL + graphConfig.usersGraphEndpoint + "/" + this.props.currentUserId);
            }

            // add current user as a member to be able to create planner plan. 
            if (membersArr.indexOf(this.state.graphContextURL + graphConfig.usersGraphEndpoint + "/" + this.props.currentUserId) === -1) {
                membersArr.push(this.state.graphContextURL + graphConfig.usersGraphEndpoint + "/" + this.props.currentUserId)
            }

            //format team display name
            let teamDisplayName = this.formatTeamDisplayName(incId, incDetails);
            let groupVisibility = this.props.graphBaseUrl === constants.defaultGraphBaseURL ? "Private" : "Public"

            if (membersArr.length > 0) {
                // create object to create teams group
                //update display name based on team name configuration
                let incidentobj = {
                    displayName: teamDisplayName,
                    mailNickname: `${constants.teamEOCPrefix}_${incId}`,
                    description: incDetails.incidentDesc,
                    visibility: groupVisibility,
                    groupTypes: ["Unified"],
                    mailEnabled: true,
                    securityEnabled: true,
                    "members@odata.bind": membersArr,
                    "owners@odata.bind": ownerArr
                }

                // call method to create team group
                let groupResponse = await this.dataService.sendGraphPostRequest(graphConfig.teamGroupsGraphEndpoint, this.props.graph, incidentobj);
                return groupResponse;
            }
            else {
                // create object to create teams group
                let incidentobj = {
                    displayName: teamDisplayName, // `${constants.teamEOCPrefix}-${incId}-${incDetails.incidentType}-${incDetails.startDateTime}`,
                    mailNickname: `${constants.teamEOCPrefix}_${incId}`,
                    description: incDetails.incidentDesc,
                    visibility: groupVisibility,
                    groupTypes: ["Unified"],
                    mailEnabled: true,
                    securityEnabled: true,
                    "owners@odata.bind": ownerArr
                }
                // call method to create team group
                let groupResponse = await this.dataService.sendGraphPostRequest(graphConfig.teamGroupsGraphEndpoint, this.props.graph, incidentobj);
                return groupResponse;
            }
        }
        catch (ex: any) {
            console.error(
                constants.errorLogPrefix + "CreateIncident_CreateTeamGroup \n",
                JSON.stringify(ex)
            );
            console.error("EOC App - CreateTeamGroup_Failed to create M365 group \n" + ex);
            // Log Exception
            this.dataService.trackException(this.props.appInsights, ex, constants.componentNames.IncidentDetailsComponent, 'CreateIncident_CreateTeamGroup', this.props.userPrincipalName);

            throw ex;
        }
    }

    //format team display name
    private formatTeamDisplayName = (incId: any, incDetails: any) => {
        //format team display name
        let teamDisplayName = `${incId}`;
        Object.keys(this.state.teamNameConfigArray).forEach((key: any) => {
            if (this.state.teamNameConfigArray[key] !== constants.teamNameConfigConstants.DontInclude) {
                if (key === constants.teamNameConfigConstants.Prefix) {
                    //prefix should not be more than 10 characters
                    let prefixVal = this.state.prefixValue.substring(0, 11);
                    teamDisplayName = teamDisplayName.concat(`-${prefixVal}`)
                }
                if (key === constants.teamNameConfigConstants.IncidentName) {
                    teamDisplayName = teamDisplayName.concat(`-${incDetails.incidentName}`);
                }
                if (key === constants.teamNameConfigConstants.IncidentType) {
                    //incident type should not be more than 170 character
                    let incTypeVal = incDetails.incidentType.substring(0, 171);
                    teamDisplayName = teamDisplayName.concat(`-${incTypeVal}`);
                }
                if (key === constants.teamNameConfigConstants.StartDate) {
                    teamDisplayName = teamDisplayName.concat(`-${incDetails.startDateTime}`);
                }
            }
        })
        return teamDisplayName;
    }

    // compare the teams membership with old and new roles
    private compareTeamsMembership = (allExistingMembers: any): any => {

        this.setState({
            loaderMessage: this.props.localeStrings.updateTeamMembershipLoaderMessage
        });

        /*creating an array of existing members except created by user to avoid removing it from team owners
        when removing the user from the role assignments when duplicate exists*/
        const allMembersExceptCreator = allExistingMembers.filter((user: any) => user.userId !== this.props.incidentData?.createdById);

        //creating an array for existing Team Members
        const teamsMembers = allExistingMembers.filter((user: any) => {
            return user.roles.length === 0;
        });

        //creating an array for existing Team Owners
        const teamOwners = allExistingMembers.filter((user: any) => {
            return user.roles.length > 0;
        });

        //creating an array of existing secondary commanders, "Edit Access Role" users and role users
        const existingRoleUsers: any = [];
        const existingSecondaryCommanders: any = [];
        this.state.existingRolesMembers.forEach((role: any) => {
            if (role.role === constants.secondaryIncidentCommanderRole || role.role === this.props.editIncidentAccessRole) {
                role.userDetailsObj.forEach((user: any) => {
                    existingSecondaryCommanders.push(user.userId);
                })
                if (role.leadDetailsObj !== undefined) {
                    role.leadDetailsObj.forEach((lead: any) => {
                        existingSecondaryCommanders.push(lead.userId);
                    })
                }
            }
            else {
                role.userDetailsObj.forEach((user: any) => {
                    existingRoleUsers.push(user.userId);
                })
                if (role.leadDetailsObj !== undefined) {
                    role.leadDetailsObj.forEach((lead: any) => {
                        existingRoleUsers.push(lead.userId);
                    })
                }
            }
        });

        //creating an array of new secondary commanders,"Edit Access Role" users and role users
        const newRoleUsers: any = [];
        const newSecondaryCommanders: any = [];
        this.state.roleAssignments.forEach((role: any) => {
            if (role.role === constants.secondaryIncidentCommanderRole || role.role === this.props.editIncidentAccessRole) {
                role.userDetailsObj.forEach((user: any) => {
                    newSecondaryCommanders.push({ role: role.role, userId: user.userId });
                })
                if (role.leadDetailsObj !== undefined) {
                    role.leadDetailsObj.forEach((lead: any) => {
                        newSecondaryCommanders.push({ role: role.role, userId: lead.userId });
                    })
                }
            } else {
                role.userDetailsObj.forEach((user: any) => {
                    newRoleUsers.push({ role: role.role, userId: user.userId });
                })
                if (role.leadDetailsObj !== undefined) {
                    role.leadDetailsObj.forEach((lead: any) => {
                        newRoleUsers.push({ role: role.role, userId: lead.userId });
                    })
                }
            }
        });

        //checking if all new role users are part of Team Members and adding it to array if the user is not a team member
        const users: any = [];
        newRoleUsers.forEach((user: any) => {
            let isExisting = false;
            teamsMembers.forEach((existingUser: any) => {
                if (existingUser.userId === user.userId) {
                    isExisting = true;
                }
            });
            if (!isExisting) {
                users.push(user);
            }
        });

        //checking if all new secondary commanders and "Edit Access Role" users are part of Team Owners and adding it to array if the user is not a team owner
        const secondaryCommanderUsers: any = [];
        newSecondaryCommanders.forEach((user: any) => {
            let isExisting = false;
            teamOwners.forEach((existingUser: any) => {
                if (existingUser.userId === user.userId) {
                    isExisting = true;
                }
            });
            if (!isExisting) {
                secondaryCommanderUsers.push(user);
            }
        });

        let newAddedUsers = users;
        let newSecondaryCommanderUsers = secondaryCommanderUsers;

        //creating an array to add the users to remove from Team Owners and Members
        let removedUsers: any = [];

        //check if Inc commander has changed and remove the old Inc Commander from Owners
        if (this.state.existingIncCommander.userId !== this.state.incDetailsItem.incidentCommander.userId) {

            //check if the old Inc commander is also a Secondary Incident commander or "Edit Access Role" user, if yes dont remove it from Owners
            const isIncCommanderASecCommander = newSecondaryCommanders.filter((user: any) => user.userId === this.state.existingIncCommander.userId);
            //remove from owners if the old Inc commander is not a Secondary Incident commander
            if (isIncCommanderASecCommander.length === 0)
                removedUsers.push(this.state.existingIncCommander.userId);
        }

        //check if any user is removed from secondary inc commander role or "Edit Access Role" and add it to array
        existingSecondaryCommanders.forEach((user: string) => {
            let isFound = false;
            newSecondaryCommanders.forEach((newUser: any) => {
                if (user === newUser.userId)
                    isFound = true;
            });
            if (!isFound) {
                //remove only if the user is not an incident commander
                if (user !== this.state.existingIncCommander.userId)
                    removedUsers.push(user);
            }
        });

        //check if any user is removed from any of the roles and add it to array
        existingRoleUsers.forEach((user: string) => {
            let isFound = false;
            newRoleUsers.forEach((newUser: any) => {
                if (user === newUser.userId)
                    isFound = true;
            });
            if (!isFound) {
                removedUsers.push(user);
            }
        });

        //create an array with id of an user from Teams, to remove them from membership
        const removedMembershipIds: any = [];
        removedUsers.forEach((user: any) => {
            allMembersExceptCreator.filter((member: any) => {
                if (member.userId === user) {
                    removedMembershipIds.push(member.id);
                }
            })
        });

        if (this.state.existingIncCommander.userId === this.state.incDetailsItem.incidentCommander.userId) {
            let usersObj = {
                newAddedUsers: newAddedUsers,
                removedMembershipIds: removedMembershipIds,
                removedUsers: removedUsers,
                removeIncCommander: [],
                newSecondaryIncidentCommanders: newSecondaryCommanderUsers
            }
            return usersObj;
        }
        else {
            const currentIncCommander = allExistingMembers.filter((user: any) => user.userId === this.state.existingIncCommander.userId);

            let usersObj = {
                newAddedUsers: newAddedUsers,
                removedMembershipIds: removedMembershipIds,
                removedUsers: removedUsers,
                removeIncCommander: [currentIncCommander[0].id],
                newSecondaryIncidentCommanders: newSecondaryCommanderUsers
            }

            return usersObj;
        }
    }

    // remove users from Teams members
    private async removeUsersFromTeam(userIds: [], isTeam: boolean, channelId?: string): Promise<any> {
        let result: any = {
            isAllDeleted: false,
            failedEntries: [],
            successEntries: []
        };
        return new Promise(async (resolve, reject) => {
            let allDone = false;
            let counter = 0;

            if (userIds.length > 0) {
                while (!allDone) {
                    let user = userIds[counter];
                    try {
                        this.graphEndpoint = isTeam ? graphConfig.teamsGraphEndpoint + "/" + this.state.teamGroupId + graphConfig.membersGraphEndpoint + "/" + user :
                            //building graph end point to remove user from shared channel
                            graphConfig.teamsGraphEndpoint + "/" + this.state.teamGroupId + graphConfig.channelsGraphEndpoint + "/" + channelId + graphConfig.membersGraphEndpoint + "/" + user

                        await this.dataService.sendGraphDeleteRequest(this.graphEndpoint, this.props.graph);

                        counter++;
                    } catch (ex: any) {
                        console.error(
                            constants.errorLogPrefix + "UpdateIncident_DeleteMemberFromTeam \n",
                            JSON.stringify(ex)
                        );
                        result.failedEntries.push(user);
                        counter++;

                    }
                    allDone = userIds.length === counter;
                }
            }
            result.isAllDeleted = result.failedEntries.length === 0 ? true : false;
            resolve(result);
        });
    }

    // add users to teams members
    private async addUsersToTeam(userIds: any, isOwner: boolean): Promise<any> {
        return new Promise(async (resolve, reject) => {
            this.graphEndpoint = graphConfig.teamsGraphEndpoint + "/" + this.state.teamGroupId + graphConfig.addMembersGraphEndpoint;

            const usersToAdd: any = [];
            const uniqueUserArray: any = [];

            if (isOwner) {
                userIds.forEach((user: any) => {
                    if (uniqueUserArray.indexOf(user.userId) === -1) {
                        uniqueUserArray.push(user.userId);
                        usersToAdd.push({
                            "@odata.type": "microsoft.graph.aadUserConversationMember",
                            "roles": ["owner"],
                            "user@odata.bind": this.state.graphContextURL + graphConfig.usersGraphEndpoint + "('" + user.userId + "')"
                        });
                    }
                });
            }
            else {
                // logic to identify uniqe users 

                userIds.forEach((user: any) => {
                    if (uniqueUserArray.indexOf(user.userId) === -1) {
                        uniqueUserArray.push(user.userId);
                        usersToAdd.push({
                            "@odata.type": "microsoft.graph.aadUserConversationMember",
                            "roles": [],
                            "user@odata.bind": this.state.graphContextURL + graphConfig.usersGraphEndpoint + "('" + user.userId + "')"
                        });
                    }
                });
            }

            const memmbersObj = {
                "values": usersToAdd
            }

            try {
                const usersAdded = await this.dataService.sendGraphPostRequest(this.graphEndpoint, this.props.graph, memmbersObj);
                resolve(usersAdded);
            } catch (ex) {
                console.error(
                    constants.errorLogPrefix + "UpdateIncident_AddUsersToTeam \n",
                    JSON.stringify(ex)
                );
                reject(ex);
                // Log Exception
                this.dataService.trackException(this.props.appInsights, ex, constants.componentNames.IncidentDetailsComponent, 'UpdateIncident_AddUsersToTeam', this.props.userPrincipalName);
            }
        });
    }

    // add users to tag groups
    private async addUsersToTag(userIds: any, tagsList: any, isIncCommander: boolean): Promise<any> {
        let result: any = {
            isFullyCreated: false,
            isPartiallyCreated: false,
            failedEntries: [],
            successEntries: []
        };

        return new Promise(async (resolve, reject) => {
            let allDone = false;
            let counter = 0;

            if (userIds.length > 0) {
                while (!allDone) {
                    let user = userIds[counter];
                    try {
                        let existingTagDetails = [];
                        const members: any = [];
                        if (isIncCommander) {
                            existingTagDetails = tagsList.filter((tags: any) => tags.displayName === constants.incidentCommanderRoleName);
                            members.push({
                                "userId": user
                            })
                        }
                        else {
                            existingTagDetails = tagsList.filter((tags: any) => tags.displayName === user.role);
                            members.push({
                                "userId": user.userId
                            })
                        }

                        if (existingTagDetails.length > 0) {
                            this.graphEndpoint = graphConfig.teamsGraphEndpoint + "/" + this.state.teamGroupId + graphConfig.tagsGraphEndpoint + "/" + existingTagDetails[0].id + graphConfig.membersGraphEndpoint;

                            let addMember = await this.dataService.sendGraphPostRequest(this.graphEndpoint, this.props.graph, members[0]);

                            if (addMember) {
                                result.successEntries.push(user.userId);
                            }
                            else {
                                result.failedEntries.push(user.userId);
                            }
                        }
                        counter++;
                    } catch (ex: any) {
                        console.error(
                            constants.errorLogPrefix + "UpdateIncident_UpdateTagMember \n",
                            JSON.stringify(ex)
                        );
                        result.failedEntries.push(user.userId);
                        counter++;
                    }
                    allDone = userIds.length === counter;
                }
            }
            result.isFullyCreated = result.failedEntries.length === 0 ? true : false;
            resolve(result);
        });
    }

    // create Team associated with Teams group
    private async createTeam(groupInfo: any): Promise<any> {
        return new Promise(async (resolve) => {
            let maxTeamCreationAttempt = 15, isTeamCreated = false;

            let result = {
                status: false,
                data: {}
            };

            this.setState({
                loaderMessage: this.props.localeStrings.createTeamLoaderMessage
            });

            // loop till the team is created
            // attempting multiple times as sometimes teams group doesn't reflect immediately after creation
            while (isTeamCreated === false && maxTeamCreationAttempt > 0) {
                try {
                    // create the team setting object
                    let teamSettings = JSON.stringify(this.getTeamSettings());
                    this.graphEndpoint = graphConfig.teamGroupsGraphEndpoint + "/" + groupInfo.id + graphConfig.teamGraphEndpoint

                    // call method to create team
                    let updatedTeamInfo = await this.dataService.sendGraphPutRequest(this.graphEndpoint, this.props.graph, teamSettings)

                    // update the result object
                    if (updatedTeamInfo) {
                        console.log(constants.infoLogPrefix + "Incident Team created");
                        isTeamCreated = true;
                        result.data = updatedTeamInfo;
                        result.status = true;
                    }
                } catch (updationError: any) {
                    console.log(constants.infoLogPrefix + "Incident Team creation failed");
                    console.error(
                        constants.errorLogPrefix + "CreateIncident_CreateTeam \n",
                        JSON.stringify(updationError)
                    );
                    // Log Exception
                    this.dataService.trackException(this.props.appInsights, updationError, constants.componentNames.IncidentDetailsComponent, 'CreateIncident_CreateTeam', this.props.userPrincipalName);
                    if (updationError.statusCode === 409 && updationError.message === "Team already exists") {
                        isTeamCreated = true;
                        this.graphEndpoint = graphConfig.teamGroupsGraphEndpoint + "/" + groupInfo.id;
                        result.data = await this.dataService.getGraphData(this.graphEndpoint, this.props.graph)
                    }
                }
                maxTeamCreationAttempt--;
                await this.timeout(5000);
            }
            console.log(constants.infoLogPrefix + "createTeam_No Of Attempt", (15 - maxTeamCreationAttempt), result);
            resolve(result);
        });
    }

    // set the Teams properties
    private getTeamSettings = (): any => {
        return {
            "memberSettings": {
                "allowCreateUpdateChannels": true,
                "allowDeleteChannels": true,
                "allowAddRemoveApps": true,
                "allowCreateUpdateRemoveTabs": true,
                "allowCreateUpdateRemoveConnectors": true
            },
            "guestSettings": {
                "allowCreateUpdateChannels": true,
                "allowDeleteChannels": true
            },
            "messagingSettings": {
                "allowUserEditMessages": true,
                "allowUserDeleteMessages": true,
                "allowOwnerDeleteMessages": true,
                "allowTeamMentions": true,
                "allowChannelMentions": true
            },
            "funSettings": {
                "allowGiphy": true,
                "giphyContentRating": "strict",
                "allowStickersAndMemes": true,
                "allowCustomMemes": true
            }
        };
    }

    // get channels to be created
    private getFixedChannel(): Array<ITeamChannel> {
        let res: Array<ITeamChannel> = [];
        try {
            if (this.state.toggleAdditionalChannels) {
                let uniqueRoleUserIds: any = [];
                this.state.incDetailsItem.additionalTeamChannels.forEach((channel: IAdditionalTeamChannels) => {
                    if (channel.channelName.trim() !== "") {
                        if (channel.channelType === constants.privateChannel) {

                            //Get unique user ids from role assignments table
                            if (uniqueRoleUserIds.length === 0 && this.state.roleAssignments.length > 0) {
                                const allRoleUserIds: any = [];
                                //Get all user ids from all roles
                                this.state.roleAssignments?.forEach((role: any) => {
                                    let currentRoleAllIds: any = [];

                                    //Get unique user ids from each role
                                    const roleUsers = role?.userDetailsObj ? role.userDetailsObj.map((user: any) => user?.userId) : [];
                                    const roleLeads = role?.leadDetailsObj ? role.leadDetailsObj.map((lead: any) => lead?.userId) : [];
                                    [...roleUsers, ...roleLeads].forEach((userId: any) => {
                                        if (currentRoleAllIds.indexOf(userId) === -1) {
                                            currentRoleAllIds.push(userId);
                                        }
                                    });
                                    allRoleUserIds.push(...currentRoleAllIds);
                                });

                                //Get unique user ids from all roles
                                allRoleUserIds.forEach((userId: any) => {
                                    if (uniqueRoleUserIds.indexOf(userId) === -1) {
                                        uniqueRoleUserIds.push(userId);
                                    }
                                });
                            }

                            //Get valid user ids from selected user ids state
                            const filteredSelectedUserIds = channel?.selectedRoleUserIds?.split(',').filter((userId: any) =>
                                uniqueRoleUserIds.indexOf(userId?.split("|")[0]) !== -1);

                            let membersArr = this.addMembersToPrivateChannel(filteredSelectedUserIds);

                            res.push({
                                "displayName": channel.channelName.trim(),
                                "membershipType": constants.privateChannel,
                                "members": membersArr
                            });
                        } else {
                            res.push({
                                "displayName": channel.channelName.trim()
                            });
                        }

                    }
                });
            }
            else {
                Object.values(constants.defaultChannelConstants).forEach((channel: string) => {
                    res.push({ "displayName": channel })
                });
            }
            return res;
        }
        catch (error: any) {
            console.error(
                constants.errorLogPrefix + "CreateIncident_getFixedChannel \n",
                JSON.stringify(error)
            );
            return res;
        }
    }


    // Create member object for private channel
    private addMembersToPrivateChannel(filteredSelectedUserIds: any = []) {

        const usersToAdd: any = [];
        const uniqueUserArray: any = [];

        //adding current user as owner to the Private Channel
        uniqueUserArray.push(this.props.currentUserId);
        usersToAdd.push({
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            "user@odata.bind": this.props.graphContextURL + graphConfig.usersGraphEndpoint + "('" + this.props.currentUserId + "')",
            "roles": ["owner"]
        });

        //adding incident commander as owner to the Private Channel
        if (uniqueUserArray.indexOf(this.state.selectedIncidentCommander[0].id) === -1) {
            uniqueUserArray.push(this.state.selectedIncidentCommander[0].id);
            usersToAdd.push({
                "@odata.type": "#microsoft.graph.aadUserConversationMember",
                "user@odata.bind": this.props.graphContextURL + graphConfig.usersGraphEndpoint + "('" + this.state.selectedIncidentCommander[0].id + "')",
                "roles": ["owner"]
            });
        }

        //Adding Secondary Incident Commanders and "Edit Access Role" users as owners to the Private Channel
        const secondaryIncidentCommanderObj = this.state.roleAssignments.find((role: any) => role.role === constants.secondaryIncidentCommanderRole || role.role === this.props.editIncidentAccessRole);
        if (secondaryIncidentCommanderObj) {
            secondaryIncidentCommanderObj?.userDetailsObj?.forEach((user: UserDetails) => {
                if (uniqueUserArray.indexOf(user.userId) === -1) {
                    uniqueUserArray.push(user.userId);
                    usersToAdd.push({
                        "@odata.type": "#microsoft.graph.aadUserConversationMember",
                        "user@odata.bind": this.props.graphContextURL + graphConfig.usersGraphEndpoint + "('" + user.userId + "')",
                        "roles": ["owner"]
                    });
                }
            });
            secondaryIncidentCommanderObj?.leadDetailsObj?.forEach((lead: UserDetails) => {
                if (uniqueUserArray.indexOf(lead.userId) === -1) {
                    uniqueUserArray.push(lead.userId);
                    usersToAdd.push({
                        "@odata.type": "#microsoft.graph.aadUserConversationMember",
                        "user@odata.bind": this.props.graphContextURL + graphConfig.usersGraphEndpoint + "('" + lead.userId + "')",
                        "roles": ["owner"]
                    });
                }
            });
        }

        // adding selected users as member or guest to the Private Channel
        filteredSelectedUserIds.forEach((userId: any) => {
            if (uniqueUserArray.indexOf(userId?.split("|")[0]) === -1) {
                uniqueUserArray.push(userId?.split("|")[0]);
                usersToAdd.push({
                    "@odata.type": "#microsoft.graph.aadUserConversationMember",
                    "user@odata.bind": this.props.graphContextURL + graphConfig.usersGraphEndpoint + "('" + userId?.split("|")[0] + "')",
                    "roles": userId?.split("|")[1] === "guest" ? ["guest"] : []
                });
            }
        });
        return usersToAdd;
    }

    //create channels
    private async createChannels(group_details: any): Promise<any> {
        //some time graph api does't create the channel 
        //thats why we need to re-try 2 time if again it failed then need to take this into failed item. otherwise simply add into 
        //created list, we need to show end-use if something failed then need to pop those error.
        this.setState({
            loaderMessage: this.props.localeStrings.createChannelLoaderMessage
        });
        let channels = this.getFixedChannel();

        let result: ChannelCreationResult = {
            isFullyCreated: false,
            isPartiallyCreated: false,
            failedEntries: [],
            successEntries: [],
            failedChannels: []
        };
        return new Promise(async (resolve) => {
            if (channels.length > 0) {
                const MAX_NUMBER_OF_ATTEMPT = 3;
                let noOfAttempt = 1;
                let allDone = false;
                let counter = 0;

                // loop atleast 3 times or till the channel is created
                while (!allDone) {
                    let channel = channels[counter];
                    try {
                        this.graphEndpoint = graphConfig.teamsGraphEndpoint + "/" + group_details.id + graphConfig.channelsGraphEndpoint;
                        let createdChannel = await this.dataService.createChannel(this.graphEndpoint, this.props.graph, channel);

                        if (createdChannel) {
                            // set channel object
                            let channelObj: ChannelCreationStatus = {
                                channelName: channel.displayName,
                                isCreated: true,
                                noOfCreationAttempt: noOfAttempt,
                                rawData: createdChannel
                            };
                            noOfAttempt = 1;
                            result.successEntries.push(channelObj);
                        }
                        counter++;
                    } catch (ex: any) {
                        console.error(
                            constants.errorLogPrefix + "CreateIncident_CreateChannels \n",
                            JSON.stringify(ex)
                        );
                        // Log Exception
                        this.dataService.trackException(this.props.appInsights, ex, constants.componentNames.IncidentDetailsComponent, 'CreateIncident_CreateChannels', this.props.userPrincipalName);

                        if (noOfAttempt >= MAX_NUMBER_OF_ATTEMPT) {
                            let channelObj: ChannelCreationStatus = {
                                channelName: channel.displayName,
                                isCreated: false,
                                noOfCreationAttempt: noOfAttempt,
                                rawData: ex.message
                            };
                            noOfAttempt = 1;
                            result.isFullyCreated = false;
                            result.failedEntries.push(channelObj);
                            result.failedChannels.push(channel.displayName)
                            counter++;
                        } else {
                            noOfAttempt++;
                        }
                    }
                    allDone = (channels.length) === counter;
                }
                result.isFullyCreated = result.failedEntries.length === 0 ? true : false;
                resolve(result);
            }
            else {
                resolve(result);
            }
        });
    }

    // create assessment channel and tab
    private async createAssessmentChannelAndTab(team_id: string, site_base_url: string, site_name: string): Promise<any> {
        return new Promise(async (resolve, reject) => {
            try {
                this.setState({
                    loaderMessage: this.props.localeStrings.createAssessmentChannelLoaderMessage
                });
                const channelGraphEndpoint = graphConfig.teamsGraphEndpoint + "/" + team_id + graphConfig.channelsGraphEndpoint;
                const channelObj = {
                    "displayName": constants.Assessment
                };

                const channelResult = await this.dataService.createChannel(channelGraphEndpoint, this.props.graph, channelObj);
                console.log(constants.infoLogPrefix + "Assessment channel created");

                const tabGraphEndpoint = graphConfig.teamsGraphEndpoint + "/" + team_id + graphConfig.channelsGraphEndpoint + "/" + channelResult.id + graphConfig.tabsGraphEndpoint;

                //Associate Assessment via sharepoint app
                const assessmentTabObj = {
                    "displayName": constants.GroundAssessments,
                    "teamsApp@odata.bind": this.state.graphContextURL + graphConfig.sharepointPageAndListTabGraphEndpoint,
                    "configuration": {
                        "entityId": uuidv4(),
                        "contentUrl": `${site_base_url}/_layouts/15/teamslogon.aspx?spfx=true&dest=${site_name}/Lists/${siteConfig.lists[0].listURL}/AllItems.aspx`,
                        "removeUrl": null,
                        "websiteUrl": null
                    }
                }

                await this.dataService.sendGraphPostRequest(tabGraphEndpoint, this.props.graph, assessmentTabObj);
                console.log(constants.infoLogPrefix + "Ground Assessments tab is added to the Assessment channel");
                resolve({
                    status: true,
                    message: "channel and tab created also installed app into tab"
                });
            } catch (ex) {
                console.error(
                    constants.errorLogPrefix + "CreateIncident_CreateAssessmentChannelAndTab \n",
                    JSON.stringify(ex)
                );
                reject(ex);
                // Log Exception
                this.dataService.trackException(this.props.appInsights, ex, constants.componentNames.IncidentDetailsComponent, 'CreateIncident_CreateAssessmentChannelAndTab', this.props.userPrincipalName);

            }
        });
    }

    // create News tab
    private createNewsTab(team_info: any, teamSiteURL: string, teamSiteManagedPath: string): Promise<any> {
        return new Promise(async (resolve, reject) => {
            try {
                this.setState({
                    loaderMessage: this.props.localeStrings.createAnnouncementChannelLoaderMessage
                });
                this.graphEndpoint = graphConfig.teamsGraphEndpoint + "/" + team_info.id + graphConfig.channelsGraphEndpoint;

                const tabObj = {
                    "displayName": constants.Announcements,
                    "description": ""
                };
                const channelResult = await this.dataService.sendGraphPostRequest(this.graphEndpoint, this.props.graph, tabObj);
                console.log(constants.infoLogPrefix + "Announcements channel created");

                const addTabObj = {
                    "displayName": constants.News,
                    "teamsApp@odata.bind": this.state.graphContextURL + graphConfig.sharepointPageAndListTabGraphEndpoint,
                    "configuration": {
                        "entityId": uuidv4(),
                        "contentUrl": `${teamSiteURL}/_layouts/15/teamslogon.aspx?spfx=true&dest=${teamSiteManagedPath}/_layouts/15/news.aspx`,
                        "removeUrl": null,
                        "websiteUrl": `${teamSiteURL}/_layouts/15/news.aspx`
                    }
                }
                const addTabGraphEndpoint = graphConfig.teamsGraphEndpoint + "/" + team_info.id + graphConfig.channelsGraphEndpoint + "/" + channelResult.id + graphConfig.tabsGraphEndpoint;

                // calling a generic method which is send a post query to the graph endpoint
                const tabResult = await this.dataService.sendGraphPostRequest(addTabGraphEndpoint, this.props.graph, addTabObj);
                console.log(constants.infoLogPrefix + "News tab is added to the Announcements channel");
                resolve(tabResult.webUrl);
            } catch (ex) {
                console.error(
                    constants.errorLogPrefix + "CreateIncident_CreateNewsTab \n",
                    JSON.stringify(ex)
                );
                reject(ex);
                // Log Exception
                this.dataService.trackException(this.props.appInsights, ex, constants.componentNames.IncidentDetailsComponent, 'CreateIncident_CreateNewsTab', this.props.userPrincipalName);
            }
        });
    }

    // this method creates assessment list in the new team site for incident
    private async createAssessmentList(siteName: string, siteId: string): Promise<any> {
        return new Promise(async (resolve, reject) => {
            /* List,Field,View Creation */
            try {
                let listColumns: any = [];

                siteConfig.lists[0].columns.forEach(column => {
                    listColumns.push(column);
                });

                let listSchema = {
                    displayName: siteConfig.lists[0].listName,
                    columns: listColumns,
                    list: {
                        template: "genericList",
                    },
                };

                this.graphEndpoint = graphConfig.spSiteGraphEndpoint + siteId + graphConfig.listsGraphEndpoint;

                const listCreationRes = await this.dataService.sendGraphPostRequest(this.graphEndpoint, this.props.graph, listSchema);
                console.log(constants.infoLogPrefix + "Ground Assessments list is created in SharePoint site");

                resolve(listCreationRes);
            } catch (ex) {
                console.error(
                    constants.errorLogPrefix + "CreateIncident_CreateAssessmentList \n",
                    JSON.stringify(ex)
                );
                reject(ex);
                // Log Exception
                this.dataService.trackException(this.props.appInsights, ex, constants.componentNames.IncidentDetailsComponent, 'CreateIncident_CreateAssessmentList', this.props.userPrincipalName);
            }
        });
    }

    // loop through selected roles and create tag object
    private async createTagObject(teamId: any, roles: any): Promise<any> {

        let result: any = {
            isFullyCreated: false,
            isPartiallyCreated: false,
            failedEntries: [],
            successEntries: []
        };

        this.setState({
            loaderMessage: this.props.localeStrings.createTagsLoaderMessage
        });

        return new Promise(async (resolve, reject) => {
            let allDone = false;
            let counter = 0;

            if (roles.length > 0) {
                while (!allDone) {
                    let role = roles[counter];
                    try {
                        this.graphEndpoint = graphConfig.teamsGraphEndpoint + "/" + teamId + graphConfig.tagsGraphEndpoint;
                        const members: any = [];
                        role.userDetailsObj.forEach((users: any) => {
                            members.push({
                                "userId": users.userId
                            })
                        });
                        if (role.leadDetailsObj !== undefined) {
                            role.leadDetailsObj.forEach((lead: any) => {
                                if (!members.find((user: any) => user.userId === lead.userId)) {
                                    members.push({
                                        "userId": lead.userId
                                    })
                                }
                            });
                        }
                        const tagObj = {
                            "displayName": role.role,
                            "members": members
                        }
                        let createdTag = await this.createTag(this.graphEndpoint, tagObj)

                        if (createdTag && createdTag.status) {
                            // set tag object
                            let tagCreationObj: any = {
                                tagName: role.role,
                                isCreated: true
                            };
                            result.successEntries.push(tagCreationObj);
                        }
                        else {
                            // set tag object
                            let tagCreationObj: any = {
                                tagName: role.role,
                                isCreated: false
                            };
                            result.failedEntries.push(tagCreationObj);
                        }
                        counter++;
                    } catch (ex: any) {
                        console.error(
                            constants.errorLogPrefix + "CreateIncident_CreateTag \n",
                            JSON.stringify(ex)
                        );

                        let tagCreationObj: any = {
                            tagName: role.role,
                            isCreated: false,
                            rawData: ex.message
                        };
                        result.isFullyCreated = false;
                        result.failedEntries.push(tagCreationObj);
                        counter++;

                    }
                    allDone = roles.length === counter;
                }
            }
            result.isFullyCreated = result.failedEntries.length === 0 ? true : false;
            resolve(result);
        });
    }

    // create tags for selected roles
    private async createTag(graphEndpoint: string, tagObj: any): Promise<any> {
        return new Promise(async (resolve) => {
            let maxTagCreationAttempt = 5, isTagCreated = false;

            let result = {
                status: false,
                data: {}
            };

            // loop till the tag is created
            // attempting multiple times as sometimes teams group doesn't reflect immediately after creation
            while (isTagCreated === false && maxTagCreationAttempt > 0) {
                try {
                    // logging date time stamp for debug
                    console.log(new Date());
                    // call method to create tag
                    let tagCreationInfo = await this.dataService.sendGraphPostRequest(graphEndpoint, this.props.graph, tagObj);

                    // update the result object
                    if (tagCreationInfo) {
                        isTagCreated = true;
                        result.data = tagCreationInfo;
                        result.status = true;
                    }
                } catch (creationError: any) {
                    console.error(
                        constants.errorLogPrefix + "CreateIncident_CreateTag \n" + new Date() + "\n",
                        JSON.stringify(creationError)
                    );
                    // Log Exception
                    this.dataService.trackException(this.props.appInsights, creationError, constants.componentNames.IncidentDetailsComponent, 'CreateIncident_CreateTag', this.props.userPrincipalName);
                    if (creationError.statusCode === 409 && creationError.message === "Tag already exists") {
                        isTagCreated = true;
                    }
                    result.status = false;
                }
                maxTagCreationAttempt--;
            }
            console.log(constants.infoLogPrefix + "createTag_No Of Attempt", (5 - maxTagCreationAttempt), result);
            resolve(result);
        });
    }

    // Get all existing tags for the team
    private async getAllTags(): Promise<any> {
        return new Promise(async (resolve, reject) => {
            this.graphEndpoint = graphConfig.teamsGraphEndpoint + "/" + this.state.teamGroupId + graphConfig.tagsGraphEndpoint;
            try {
                const existingTags = await this.dataService.getGraphData(this.graphEndpoint, this.props.graph);
                resolve(existingTags);
            } catch (ex) {
                console.error(
                    constants.errorLogPrefix + "UpdateIncident_GetAllTags \n",
                    JSON.stringify(ex)
                );
                reject(ex);
                // Log Exception
                this.dataService.trackException(this.props.appInsights, ex, constants.componentNames.IncidentDetailsComponent, 'UpdateIncident_GetAllTags', this.props.userPrincipalName);
            }
        });
    }

    // check if any new role is added which requires new tag creation
    private checkIfNewTagCreationNeeded(tagList: any): any {
        const tags = tagList.map((tag: any) => {
            return tag.displayName;
        });

        const roles = this.state.roleAssignments.map((role: any) => {
            return role.role;
        })

        return roles.filter((tag: any) => tags.indexOf(tag) === -1);
    }

    // create a role object to be used in tags creation
    private createNewRoleObject(newRole: any): any {
        const rolesToAdd: any = [];
        newRole.forEach((roleName: string) => {
            this.state.roleAssignments.forEach(role => {
                if (role.role === roleName) {
                    rolesToAdd.push(role);
                }
            });
        });
        return rolesToAdd;
    }

    //delete team group
    private async deleteTeamGroup(group_id: string): Promise<any> {
        return new Promise(async (resolve, reject) => {
            this.graphEndpoint = graphConfig.teamGroupsGraphEndpoint + "/" + group_id;
            try {
                const deleteResult = await this.dataService.sendGraphDeleteRequest(this.graphEndpoint, this.props.graph);
                resolve(deleteResult);

            } catch (ex) {
                console.error(
                    constants.errorLogPrefix + "CreateIncident_DeleteTeamGroup \n",
                    JSON.stringify(ex)
                );
                //log trace
                this.dataService.trackTrace(this.props.appInsights, "CreateIncident_DeleteTeamGroup ", '', this.props.userPrincipalName);
                reject(ex);
            }
        });
    }

    //delete created incident
    private async deleteIncident(incidentId: number): Promise<any> {
        return new Promise(async (resolve, reject) => {
            this.graphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}/lists/${siteConfig.incidentsList}/items/${incidentId}`;
            try {
                const deleteResult = await this.dataService.sendGraphDeleteRequest(this.graphEndpoint, this.props.graph);
                resolve(deleteResult);
            } catch (ex) {
                console.error(
                    constants.errorLogPrefix + "CreateIncident_DeleteIncident \n",
                    JSON.stringify(ex)
                );
                //log trace
                this.dataService.trackTrace(this.props.appInsights, "CreateIncident_DeleteIncident ", incidentId.toString(), this.props.userPrincipalName);
                reject(ex);
            }
        });
    }

    //toggle Additional Channel fields
    private onToggleAdditionChannels(checked: any) {
        if (this.state.incDetailsItem.additionalTeamChannels === undefined && checked) {
            this.setState({
                incDetailsItem: {
                    ...this.state.incDetailsItem,
                    additionalTeamChannels: [
                        {
                            channelName: constants.defaultChannelConstants.Logistics,
                            hasRegexError: false, regexErrorMessage: ""
                        },
                        {
                            channelName: constants.defaultChannelConstants.Planning,
                            hasRegexError: false, regexErrorMessage: ""
                        },
                        {
                            channelName: constants.defaultChannelConstants.Recovery,
                            hasRegexError: false, regexErrorMessage: ""
                        }
                    ]
                }
            });
        }
        this.setState({ toggleAdditionalChannels: checked });
    }

    //update additional channel name into array
    private updateAdditionalChannel(_evt: any, inputObj: any, idx: number) {
        const additionalChannels = this.state.incDetailsItem.additionalTeamChannels;
        additionalChannels[idx].channelName = inputObj.value;
        //validate the channel names
        if (additionalChannels[idx].channelName.trim() !== "") {
            //check if the channel name has underscore (_) or period (.) at the beginning of the string
            if (additionalChannels[idx].channelName.trim().charAt(0) === "." || additionalChannels[idx].channelName.trim().charAt(0) === "_") {
                additionalChannels[idx].hasRegexError = true;
                additionalChannels[idx].regexErrorMessage = this.props.localeStrings.channelNameStartLetterRegexError;
            }
            //check if the channel name has period (.) at the end of the string
            else if (additionalChannels[idx].channelName.trim().charAt(additionalChannels[idx].channelName.trim().length - 1) === ".") {
                additionalChannels[idx].hasRegexError = true;
                additionalChannels[idx].regexErrorMessage = this.props.localeStrings.channelNameLastLetterRegexError;
            }
            //check if the channel name has any of these restricted characters ~ # % & * { } , + / \\ : < > .. ? | '" in the string  
            else if (!/^(?!.*\.\.)[^._~#%&*{},+/\\:<>?|'"][^~#%&*{},+/\\:<>?|'"]*[^~#%&*{},+/\\:<>?|'".]$/.test(additionalChannels[idx].channelName.trim()) ||
                additionalChannels[idx].channelName.trim()?.indexOf('\\') > -1) {
                if (additionalChannels[idx].channelName.trim()?.length === 1 && additionalChannels[idx].channelName.trim()?.indexOf("\\") === -1 &&
                    /[^._~#%&*{},+/\\:<>?|'"]/.test(additionalChannels[idx].channelName.trim())) {
                    additionalChannels[idx].hasRegexError = false;
                    additionalChannels[idx].regexErrorMessage = "";
                }
                else {
                    additionalChannels[idx].hasRegexError = true;
                    additionalChannels[idx].regexErrorMessage = this.props.localeStrings.ChannelNameRegexError;
                }
            }
            else {
                additionalChannels[idx].hasRegexError = false;
                additionalChannels[idx].regexErrorMessage = ""
            }
        }
        else {
            additionalChannels[idx].hasRegexError = false;
            additionalChannels[idx].regexErrorMessage = ""
        }
        this.setState({
            incDetailsItem: { ...this.state.incDetailsItem, additionalTeamChannels: [...additionalChannels] },
        });
    }

    //Set Channel Type when the private is checked
    private onPrivateChanged(_evt: any, checked: any, idx: number) {
        const additionalChannels = this.state.incDetailsItem.additionalTeamChannels;

        if (checked) {
            additionalChannels[idx].channelType = constants.privateChannel;
        }
        else {
            additionalChannels[idx].channelType = constants.standardChannel;
        }

        console.log("roleassignments", this.state.roleAssignments);

        const allGroupLabels = this.state.roleAssignments.map((option: any) => option.role);
        const expandedGroups = allGroupLabels.reduce(
            (acc: any, label: any) => ({ ...acc, [label]: true }),
            {}
        );
        additionalChannels[idx].expandedGroups = expandedGroups;

        this.setState({
            incDetailsItem: { ...this.state.incDetailsItem, additionalTeamChannels: [...additionalChannels] },
        });
    }
    //function to get the user data from the selected roles for a channel
    private onRolesSelect(_evt: any, data: any, idx: number) {
        try {
            if (data.optionValue?.includes('{}')) {
                const additionalChannels = this.state.incDetailsItem.additionalTeamChannels;
                const selectedRoleUsersData = data.selectedOptions.filter((option: any) => option?.includes('{}'));
                const namesOnly = selectedRoleUsersData.map((option: any) => option?.split('|')[0]);
                const idsOnly = selectedRoleUsersData.map((option: any) => {
                    if (option?.split('|')[2]?.includes("#EXT#")) {
                        return option?.split('|')[1] + "|guest";
                    }
                    else {
                        return option?.split('|')[1] + "|member";
                    }
                });
                additionalChannels[idx].selectedRoleUsers = namesOnly.join(', ');
                additionalChannels[idx].selectedRoleUserIds = idsOnly.join(',');
                this.setState({
                    incDetailsItem: { ...this.state.incDetailsItem, additionalTeamChannels: [...additionalChannels] }
                });
            }
            else {
                this.setState((prevState) => {
                    const prevAdditionalChannels = prevState.incDetailsItem.additionalTeamChannels;
                    prevAdditionalChannels[idx].expandedGroups[data.optionValue] = !prevAdditionalChannels[idx].expandedGroups[data.optionValue];
                    return { incDetailsItem: { ...prevState.incDetailsItem, additionalTeamChannels: [...prevAdditionalChannels] } }
                });
            }
        }
        catch (error: any) {
            console.error(constants.errorLogPrefix + "CreateIncident_onRolesSelect \n", JSON.stringify(error));
        }
    }

    //On click of 'Add Channel' button add a new Channel Input control
    private addChannelInput() {
        const additionalChannels = this.state.incDetailsItem.additionalTeamChannels;
        additionalChannels.push({ channelName: "", hasRegexError: false, regexErrorMessage: "" });
        this.setState({ incDetailsItem: { ...this.state.incDetailsItem, additionalTeamChannels: [...additionalChannels] } });
    }

    //Remove Channel input Control on click of Remove icon
    private removeChannelInput(index: number) {
        const additionalChannels = this.state.incDetailsItem.additionalTeamChannels;
        additionalChannels.splice(index, 1);
        this.setState({
            incDetailsItem: {
                ...this.state.incDetailsItem,
                additionalTeamChannels: [...additionalChannels]
            }
        });
    }

    //On Toggle cloud storage link
    private async onToggleCloudStorageLink(checked: any) {
        this.setState({ toggleCloudStorageLocation: checked });
        if (checked && this.state.isEditMode && this.props.incidentData?.cloudStorageLink === undefined) {
            if (this.state.incidentTypeRoleDefaultData.length === 0) {
                await this.getIncidentTypeDefaultData();
                //check if we have data for selected incident type
                const filteredincidentTypeDefaultData = this.state.incidentTypeRoleDefaultData
                    .filter((e: any) => e.incidentType === this.props.incidentData?.incidentType);
                //Assign default Cloud Storage link
                if (filteredincidentTypeDefaultData.length > 0) {
                    const cloudStorageLink = filteredincidentTypeDefaultData[0]?.cloudStorageLink?.trim();
                    this.setState({
                        incDetailsItem: {
                            ...this.state.incDetailsItem,
                            cloudStorageLink: cloudStorageLink !== "" ? cloudStorageLink : "",
                        }
                    });
                }
                else this.setState({ incDetailsItem: { ...this.state.incDetailsItem, cloudStorageLink: "" } })
            }
        }
    }

    //On Toggle guest users button
    private onToggleGuestUsers(checked: any) {
        if ((this.state.incDetailsItem.guestUsers === undefined || this.state.incDetailsItem.guestUsers.length === 0) && checked) {
            this.setState({
                incDetailsItem: {
                    ...this.state.incDetailsItem,
                    guestUsers: [{
                        email: "", displayName: "", hasDisplayNameRegexError: false, hasEmailRegexError: false,
                        hasDisplayNameValidationError: false, hasEmailValidationError: false
                    }]
                }
            });
        }
        this.setState({ toggleGuestUsers: checked });
    }

    //On click of 'Add More' button add a new Guest User Input control
    private addGuestUserInput() {
        const guestUsers = this.state.incDetailsItem.guestUsers;
        guestUsers.push({
            email: "", displayName: "", hasDisplayNameRegexError: false,
            hasEmailRegexError: false, hasDisplayNameValidationError: false, hasEmailValidationError: false
        });
        this.setState({ incDetailsItem: { ...this.state.incDetailsItem, guestUsers: [...guestUsers] } });
    }

    //update guest user details into array
    private updateGuestUser(inputObj: any, idx: number, fieldName: string) {
        const guestUsers: IGuestUsers[] = this.state.incDetailsItem.guestUsers;
        guestUsers[idx][fieldName] = inputObj.value

        if (guestUsers.filter((user: IGuestUsers) => user.email.trim() !== "" || user.displayName.trim() !== "").length > 0) {
            this.setState({ inputValidation: { ...this.state.inputValidation, guestUsersHasError: false } });
        }
        else {
            this.setState({ inputValidation: { ...this.state.inputValidation, guestUsersHasError: true } });
        }
        if ((guestUsers[idx].email.trim() !== "" && guestUsers[idx].displayName.trim() !== "") ||
            (guestUsers[idx].email.trim() === "" && guestUsers[idx].displayName.trim() === "")) {
            guestUsers[idx].hasDisplayNameValidationError = false;
            guestUsers[idx].hasEmailValidationError = false;
        }
        this.setState({
            incDetailsItem: { ...this.state.incDetailsItem, guestUsers: [...guestUsers] },
        });
    }

    // move focus to top of page to show loader or message bar
    private scrollToTop = () => {
        window.scrollTo({
            top: 0,
            behavior: 'auto'
        });
    };

    //Tooltip for info Icon
    private iconWithTooltip(iconName: string, tooltipContent: string, className: string, id: string, tabindex: number = 0) {
        return (
            <span className={className}>
                <TooltipHost
                    content={tooltipContent}
                    calloutProps={calloutProps}
                    hostClassName="tooltip-host-class"
                    id={id}
                >
                    <Icon iconName={iconName} tabIndex={tabindex} aria-label={tooltipContent}
                        aria-describedby={id} role="img" aria-hidden={false} />
                </TooltipHost>
            </span>
        );
    }

    //format the date to show in the date picker
    private onFormatDate = (date?: Date): string => {
        let formattedDate = moment(date).format("MMM DD YYYY");
        return formattedDate;
    }

    //update the state variable whenever the date is changed in the date picker control
    private onChangeStartDate = (date: Date | null | undefined) => {
        if (date) {
            let incInfo = { ...this.state.incDetailsItem };
            let inputValidationObj = this.state.inputValidation;
            incInfo["startDate"] = date;
            this.setState({ incDetailsItem: incInfo, inputValidation: inputValidationObj });
        }
    }

    //update the state variable whenever the time is changed in the time picker control
    private onChangeStartTime = (_ev: React.FormEvent<IComboBox>, time: Date) => {
        let incInfo = { ...this.state.incDetailsItem };
        let inputValidationObj = this.state.inputValidation;
        incInfo["startTime"] = time;
        if (time?.toString() === "Invalid Date" || null || undefined || "") {
            inputValidationObj.incidentStartDateTimeHasError = true;
        }
        else {
            inputValidationObj.incidentStartDateTimeHasError = false;
        }
        this.setState({ incDetailsItem: incInfo, inputValidation: inputValidationObj });
    }

    //adding key for all incident types
    private options = (optionArray: any) => {
        let myOptions: { key: any; text: any; }[] = [];
        optionArray.forEach((element: any) => {
            myOptions.push({ key: element, text: element });
        });
        return myOptions;
    }

    //adding key for all incident status
    private statusOptions = (optionArray: any) => {
        let myOptions: { key: any; text: any; }[] = [];
        optionArray.forEach((element: any) => {
            myOptions.push({ key: element.id, text: element.status });
        });
        return myOptions;
    }

    //method is called when the menu opens for incident type combo box
    private onMenuOpen = () => {

        //adding option position information to aria attribute to fix the accessibility issue in iOS Voiceover
        if (navigator.userAgent.match(/iPhone/i)) {
            const listBoxElement: any = document.getElementById("incident-type-listbox-list")?.children;
            if (listBoxElement?.length > 0) {
                for (let i = 0; i < listBoxElement?.length; i++) {
                    const buttonId = `incident-type-listbox-list${i}`;
                    const buttonElement: any = document.getElementById(buttonId);
                    const ariaLabel = `${buttonElement.innerText} ${i + 1} of ${listBoxElement.length}`;
                    buttonElement?.setAttribute("aria-label", ariaLabel);
                }
            }
        }

    }

    //on menu open, add the ariaLabel attribute to fix the position issue in iOS for accessbility
    private onStatusMenuOpen = () => {

        //adding option position information to aria attribute to fix the accessibility issue in iOS Voiceover
        if (navigator.userAgent.match(/iPhone/i)) {
            const listBoxElement: any = document.getElementById("incident-status-listbox-list")?.children;
            if (listBoxElement?.length > 0) {
                for (let i = 0; i < listBoxElement?.length; i++) {
                    const buttonId = `incident-status-listbox-list${i}`;
                    const buttonElement: any = document.getElementById(buttonId);
                    const ariaLabel = `${buttonElement.innerText} ${i + 1} of ${listBoxElement.length}`;
                    buttonElement?.setAttribute("aria-label", ariaLabel);
                }
            }
        }

    }

    //onClick or onKeydown event in Incident Commander field, update the ariaLabel attribute to fix the position issue in iOS for accessbility
    private setSuggestionsAttributes = () => {
        //adding option position information to aria attribute to fix the accessibility issue in iOS Voiceover
        setTimeout(() => {
            if (navigator.userAgent.match(/iPhone/i)) {
                const shadow = this.incidentCommandar.current?.shadowRoot;
                let suggestionsListItems: any = shadow?.getElementById("suggestions-list")?.children;

                if (suggestionsListItems?.length > 0 && suggestionsListItems !== "undefined") {
                    for (let i = 0; i < suggestionsListItems?.length; i++) {
                        const ariaLabel = suggestionsListItems[i]?.getAttribute("aria-label");
                        if (!ariaLabel.includes(`${i + 1} of ${suggestionsListItems.length}`)) {
                            let newAriaLabel = `${ariaLabel} ${i + 1} of ${suggestionsListItems.length}`
                            suggestionsListItems[i]?.setAttribute("aria-label", newAriaLabel);
                        }
                    }
                }
                else {
                    let loadingMessage: any = shadow?.querySelector('.loading-text');
                    if (loadingMessage) {
                        this.setSuggestionsAttributes();
                    }
                }
            }
        }, 1000)
    }

    //onClick or onKeydown event in Search User field, update the ariaLabel attribute to fix the position issue in iOS for accessbility
    private setSearchUserAttributes = () => {
        //adding option position information to aria attribute to fix the accessibility issue in iOS Voiceover
        setTimeout(() => {
            if (navigator.userAgent.match(/iPhone/i)) {
                const shadow = this.searchUser.current?.shadowRoot;
                let suggestionsListItems: any = shadow?.getElementById("suggestions-list")?.children;

                if (suggestionsListItems?.length > 0 && suggestionsListItems !== "undefined") {
                    for (let i = 0; i < suggestionsListItems?.length; i++) {
                        const ariaLabel = suggestionsListItems[i]?.getAttribute("aria-label");
                        if (!ariaLabel.includes(`${i + 1} of ${suggestionsListItems.length}`)) {
                            let newAriaLabel = `${ariaLabel} ${i + 1} of ${suggestionsListItems.length}`
                            suggestionsListItems[i]?.setAttribute("aria-label", newAriaLabel);
                        }
                    }
                }
                else {
                    let loadingMessage: any = shadow?.querySelector('.loading-text');
                    if (loadingMessage) {
                        this.setSuggestionsAttributes();
                    }
                }
            }
        }, 1000)

    }

    //onClick or onKeydown event in Search User edit mode field, update the ariaLabel attribute to fix the position issue in iOS for accessbility
    private setSearchUserEditModeAttributes = () => {
        //adding option position information to aria attribute to fix the accessibility issue in iOS Voiceover
        setTimeout(() => {
            if (navigator.userAgent.match(/iPhone/i)) {
                const shadow = this.searchUserEditMode.current?.shadowRoot;
                let suggestionsListItems: any = shadow?.getElementById("suggestions-list")?.children;

                if (suggestionsListItems?.length > 0 && suggestionsListItems !== "undefined") {
                    for (let i = 0; i < suggestionsListItems?.length; i++) {
                        const ariaLabel = suggestionsListItems[i]?.getAttribute("aria-label");
                        if (!ariaLabel.includes(`${i + 1} of ${suggestionsListItems.length}`)) {
                            let newAriaLabel = `${ariaLabel} ${i + 1} of ${suggestionsListItems.length}`
                            suggestionsListItems[i]?.setAttribute("aria-label", newAriaLabel);
                        }
                    }
                }
                else {
                    let loadingMessage: any = shadow?.querySelector('.loading-text');
                    if (loadingMessage) {
                        this.setSuggestionsAttributes();
                    }
                }
            }
        }, 1000)

    }

    //method will be called when location is updated
    private onLocationChange = (selectedLocation: ILocationPickerItem) => {
        try {
            this.validateLocationField(selectedLocation);
            // incident info object
            let incidentInfo: IncidentEntity = this.state.incDetailsItem;
            let location: ILocationPickerItem;
            if (selectedLocation !== null) {
                if (selectedLocation?.EntityType !== "Custom") {
                    location = { 'DisplayName': selectedLocation?.DisplayName, 'Coordinates': selectedLocation?.Coordinates, 'EntityType': selectedLocation?.EntityType };
                }
                else {
                    location = {
                        'DisplayName': selectedLocation?.DisplayName, 'Coordinates':
                            { Latitude: 0, Longitude: 0 }, 'EntityType': selectedLocation?.EntityType
                    };
                }
                incidentInfo.location = JSON.stringify(location);
                this.setState({
                    selectedLocation: location,
                    incDetailsItem: incidentInfo
                });
            }
            incidentInfo.location = "";
            this.setState({
                selectedLocation: selectedLocation,
                incDetailsItem: incidentInfo
            })
        } catch (error: any) {
            console.error(
                constants.errorLogPrefix + "IncidentDetails_LocationChange \n",
                JSON.stringify(error)
            );
            // Log Exception
            this.dataService.trackException(this.props.appInsights, error, constants.componentNames.IncidentDetailsComponent, 'IncidentDetails_LocationChange', this.props.userPrincipalName);
        }
    }

    //method to validate location field on change 
    private validateLocationField = (selectedLocation: ILocationPickerItem) => {
        let inputValidationObj = this.state.inputValidation;
        if (selectedLocation !== null) {
            inputValidationObj.incidentLocationHasError = false;
        }
        else {
            inputValidationObj.incidentLocationHasError = true;
        }
        this.setState({ inputValidation: inputValidationObj });
    }


    //main render method
    public render() {
        const isDarkOrContrastTheme = this.props.currentThemeName === constants.darkMode || this.props.currentThemeName === constants.contrastMode;
        return (
            <>
                <div className="incident-details">
                    <>
                        {this.state.showLoader &&
                            <div className="loader-bg">
                                <div className="loaderStyle">
                                    {this.state.loaderMessage === this.props.localeStrings.genericLoaderMessage ?
                                        <Loader label={this.state.loaderMessage} size="largest" />
                                        :
                                        <Loader aria-live="polite" role="alert" label={this.state.loaderMessage} size="largest" />
                                    }
                                </div>
                            </div>
                        }
                        <div style={{ opacity: this.state.formOpacity }}>
                            <div className=".col-xs-12 .col-sm-8 .col-md-4 container" id="incident-details-path">
                                <label>
                                    <span
                                        onClick={() => this.props.onBackClick("")}
                                        onKeyDown={(event) => {
                                            if (event.key === constants.enterKey)
                                                this.props.onBackClick("")
                                        }}
                                        className="go-back">
                                        <ChevronStartIcon id="path-back-icon" />
                                        <span className="back-label" role="button" tabIndex={0} title="Back">{this.props.localeStrings.back}</span>
                                    </span> &nbsp;&nbsp;
                                    <span className="right-border">|</span>
                                    <span>&nbsp;&nbsp;{this.props.localeStrings.formTitle}</span>
                                </label>
                            </div>
                            <div className={`incident-details-form-area${isDarkOrContrastTheme ? " incident-details-form-area-darkcontrast" : ""}`}>
                                <div className="container">
                                    <h2 aria-live="polite" role="alert"> <div className="incident-form-head-text">
                                        {!this.props.isEditMode ?
                                            <>{this.props.localeStrings.formTitle}</>
                                            :
                                            <>{this.props.localeStrings.formTitleEditMode} - {this.props.incidentData?.incidentId}</>
                                        }
                                    </div></h2>
                                    <Row xs={1} sm={1} md={2} lg={2} xl={2} >
                                        <Col xs={12} sm={12} md={8} lg={8} xl={8}>
                                            <Row xs={1} sm={1} md={2} lg={2} xl={2}>
                                                <Col xs={12} sm={12} md={6} lg={6} xl={6}>
                                                    <div className="incident-grid-item">
                                                        <label className="FormInput-label">{this.props.localeStrings.fieldIncidentName}</label>
                                                        <TooltipHost
                                                            content={this.props.localeStrings.infoIncName}
                                                            calloutProps={calloutProps}
                                                            hostClassName="tooltip-host-class"
                                                        >
                                                            <Icon aria-label={this.props.localeStrings.infoIncName} tabIndex={0} role="img" iconName="Info" className="incNameInfoIcon" />
                                                        </TooltipHost>
                                                        <FormInput
                                                            type="text"
                                                            ref={this.incidentName}
                                                            placeholder={this.props.localeStrings.phIncidentName}
                                                            fluid={true}
                                                            maxLength={constants.maxCharLengthForSingleLine}
                                                            aria-label={this.props.localeStrings.fieldIncidentName + constants.requiredAriaLabel}
                                                            onChange={(evt) => this.onTextInputChange(evt, "incidentName")}
                                                            value={this.state.incDetailsItem ? (this.state.incDetailsItem.incidentName ? this.state.incDetailsItem.incidentName : '') : ''}
                                                            className="incident-details-input-field"
                                                            successIndicator={false}
                                                        />
                                                        {this.state.inputValidation.incidentNameHasError && (
                                                            <label aria-live="polite" role="alert" className="message-label">{this.props.localeStrings.incidentNameRequired}</label>
                                                        )}
                                                        {this.state.inputRegexValidation.incidentNameHasError && (
                                                            <label aria-live="polite" role="alert" className="message-label">{this.props.localeStrings.incidentNameRegex}</label>
                                                        )}
                                                    </div>
                                                </Col>
                                                <Col xs={12} sm={12} md={6} lg={6} xl={6}>
                                                    <div className="incident-grid-item">
                                                        <label className="FormInput-label">{this.props.localeStrings.fieldIncidentStatus}</label>
                                                        <ComboBox
                                                            placeholder={this.props.localeStrings.phIncidentStatus}
                                                            options={this.state.dropdownOptions["statusOptions"] ? this.statusOptions(this.state.dropdownOptions["statusOptions"]) : []}
                                                            selectedKey={this.state.incDetailsItem ? (this.state.incDetailsItem.incidentStatus ? this.state.incDetailsItem.incidentStatus.id : "") : ""}
                                                            onChange={this.onIncidentStatusChange}
                                                            className={"incident-status-dropdown"}
                                                            useComboBoxAsMenuWidth={true}
                                                            persistMenu={true}
                                                            calloutProps={{ directionalHintFixed: true, doNotLayer: true }}
                                                            ref={this.incidentType}
                                                            ariaLabel={this.props.localeStrings.fieldIncidentStatus + constants.requiredAriaLabel}
                                                            id="incident-status-listbox"
                                                            onMenuOpen={this.onStatusMenuOpen}
                                                            styles={{
                                                                optionsContainer: {
                                                                    "button span": {
                                                                        maxHeight: "35px",
                                                                        height: "auto"
                                                                    }
                                                                }
                                                            }}
                                                        />
                                                        {this.state.inputValidation.incidentStatusHasError && (
                                                            <label aria-live="polite" role="alert" className="message-label">{this.props.localeStrings.statusRequired}</label>
                                                        )}
                                                    </div>
                                                </Col>
                                            </Row>
                                            <Row xs={1} sm={1} md={2} lg={2} xl={2}>
                                                <Col xs={12} sm={12} md={6} lg={6} xl={6}>
                                                    <div className="incident-grid-item" ref={this.incTypeRef}>
                                                        <label className="FormInput-label">{this.props.localeStrings.fieldIncidentType}</label>
                                                        {this.props.isEditMode ?
                                                            <FormDropdown
                                                                placeholder={this.props.localeStrings.phIncidentType}
                                                                fluid={true}
                                                                value={this.state.incDetailsItem ? (this.state.incDetailsItem.incidentType ? this.state.incDetailsItem.incidentType : '') : ''}
                                                                className={"incident-type-dropdown-disabled"}
                                                                disabled={true}
                                                                aria-label={this.props.localeStrings.fieldIncidentType + constants.requiredAriaLabel}
                                                            />
                                                            :
                                                            <ComboBox
                                                                placeholder={this.props.localeStrings.phIncidentType}
                                                                options={this.state.dropdownOptions["typeOptions"] ? this.options(this.state.dropdownOptions["typeOptions"]) : []}
                                                                onChange={this.onIncidentTypeChange}
                                                                className={"incident-type-dropdown"}
                                                                useComboBoxAsMenuWidth={true}
                                                                allowFreeInput={true}
                                                                persistMenu={true}
                                                                calloutProps={{ directionalHintFixed: true, doNotLayer: true }}
                                                                ref={this.incidentType}
                                                                ariaLabel={this.props.localeStrings.fieldIncidentType + constants.requiredAriaLabel}
                                                                id="incident-type-listbox"
                                                                onMenuOpen={this.onMenuOpen}
                                                                styles={{
                                                                    optionsContainer: {
                                                                        "button span": {
                                                                            maxHeight: "35px",
                                                                            height: "auto"
                                                                        }
                                                                    }
                                                                }}
                                                            />
                                                        }
                                                        {this.state.inputValidation.incidentTypeHasError && (
                                                            <label aria-live="polite" role="alert" className="message-label">{this.props.localeStrings.incidentTypeRequired}</label>
                                                        )}
                                                    </div>
                                                </Col>
                                                <Col xs={12} sm={12} md={6} lg={6} xl={6}>
                                                    <div className="incident-grid-item" ref={this.incCommanderRef}>
                                                        <label className="FormInput-label">{this.props.localeStrings.fieldIncidentCommander}</label>

                                                        <TooltipHost
                                                            content={this.props.localeStrings.infoIncCommander}
                                                            calloutProps={calloutProps}
                                                            hostClassName="tooltip-host-class"
                                                        >
                                                            <Icon aria-label={this.props.localeStrings.infoIncCommander} tabIndex={0} role="img" iconName="Info" className="incCommanderInfoIcon" />
                                                        </TooltipHost>
                                                        <PeoplePicker
                                                            title={this.props.localeStrings.fieldIncidentCommander}
                                                            ariaLabel={this.props.localeStrings.fieldIncidentCommander + constants.requiredAriaLabel}
                                                            selectionMode="single"
                                                            type={PersonType.person}
                                                            userType={UserType.user}
                                                            selectionChanged={this.handleIncCommanderChange}
                                                            placeholder={this.props.localeStrings.phIncidentCommander}
                                                            className="incident-details-people-picker"
                                                            selectedPeople={this.state.selectedIncidentCommander}
                                                            ref={this.incidentCommandar}
                                                            onKeyDown={this.setSuggestionsAttributes}
                                                            onClick={this.setSuggestionsAttributes}
                                                        />
                                                        {this.state.inputValidation.incidentCommandarHasError && (
                                                            <label aria-live="polite" role="alert" className="message-label">{this.props.localeStrings.incidentCommanderRequired}</label>
                                                        )}
                                                        {this.state.incCommanderHasRegexError && (
                                                            <label aria-live="polite" role="alert" className="message-label">{this.props.localeStrings.guestUsersNotAllowedAsIncCommanderErrorMsg}</label>
                                                        )}
                                                    </div>
                                                </Col>
                                            </Row>
                                            <Row xs={1} sm={1} md={2} lg={2} xl={2}>
                                                <Col xs={12} sm={12} md={6} lg={6} xl={6}>
                                                    <div className="incident-grid-item">
                                                        <label className="FormInput-label">{this.props.localeStrings.fieldStartDate}</label>
                                                        {(this.props.incidentData && this.props.incidentData.incidentId) ?
                                                            <FormInput
                                                                aria-label={this.props.localeStrings.fieldStartDate + constants.requiredAriaLabel}
                                                                type="text"
                                                                placeholder={this.props.localeStrings.phStartDate}
                                                                fluid={true}
                                                                value={this.state.incDetailsItem.startDateTime ? this.state.incDetailsItem.startDateTime : ''}
                                                                disabled
                                                                className="incident-details-input-field-disabled"
                                                            />
                                                            :
                                                            <>
                                                                <div ref={this.incidentStartDateTime} className="incident-startdatetime">
                                                                    <DatePicker
                                                                        value={this.state.incDetailsItem.startDate ? this.state.incDetailsItem.startDate : new Date()}
                                                                        onSelectDate={this.onChangeStartDate}
                                                                        placeholder="Select a date"
                                                                        ariaLabel={constants.startDateAriaLabel + constants.requiredAriaLabel}
                                                                        className="incident-datepicker"
                                                                        formatDate={this.onFormatDate}
                                                                        calloutProps={{ className: `incidentdatepicker-callout${this.props.currentThemeName === constants.darkMode ? " incidentdatepicker-callout-dark" : `${this.props.currentThemeName === constants.contrastMode ? " incidentdatepicker-callout-contrast" : ""}`}` }}
                                                                    />
                                                                    <TimePicker
                                                                        dateAnchor={this.state.incDetailsItem.startDate}
                                                                        value={this.state.incDetailsItem.startTime ? this.state.incDetailsItem.startTime : new Date()}
                                                                        placeholder="Select a time"
                                                                        onChange={this.onChangeStartTime}
                                                                        calloutProps={{ directionalHintFixed: true, doNotLayer: true }}
                                                                        ariaLabel={constants.startTimeAriaLabel + constants.requiredAriaLabel}
                                                                        className="incident-timepicker"
                                                                        styles={{
                                                                            optionsContainer: {
                                                                                "button span": {
                                                                                    height: "auto"
                                                                                }
                                                                            }
                                                                        }}
                                                                    />
                                                                </div>
                                                                {this.state.inputValidation.incidentStartDateTimeHasError && (
                                                                    <label aria-live="polite" role="alert" className="message-label">{this.props.localeStrings.startDateRequired}</label>
                                                                )}
                                                            </>
                                                        }
                                                    </div>
                                                </Col>
                                                <Col xs={12} sm={12} md={6} lg={6} xl={6}>
                                                    <div className="incident-grid-item">
                                                        <label className="FormInput-label">{this.props.localeStrings.fieldLocation}</label>
                                                        <TeamsFxContext.Consumer>
                                                            {(value) =>
                                                                <LocationPicker
                                                                    onChange={this.onLocationChange}
                                                                    defaultValue={this.state.selectedLocation}
                                                                    className="incident-location-picker"
                                                                    placeholder={this.props.localeStrings.phLocation}
                                                                    appInsights={this.props.appInsights}
                                                                    userPrincipalName={this.props.userPrincipalName}
                                                                    graphBaseUrl={this.props.graphBaseUrl}
                                                                    teamsUserCredential={value.teamsUserCredential!}
                                                                />
                                                            }
                                                        </TeamsFxContext.Consumer>
                                                        {this.state.inputValidation.incidentLocationHasError && (
                                                            <label aria-live="polite" role="alert" className="message-label">{this.props.localeStrings.locationRequired}</label>
                                                        )}
                                                        {this.state.inputRegexValidation.incidentLocationHasError && (
                                                            <label aria-live="polite" role="alert" className="message-label">{this.props.localeStrings.locationRegex}</label>
                                                        )}
                                                    </div>
                                                </Col>
                                            </Row>
                                        </Col>
                                        <Col xs={12} sm={12} md={4} lg={4} xl={4}>
                                            <div className="incident-grid-item">
                                                <label className="FormInput-label">{this.props.localeStrings.fieldDescription}</label>
                                                <FormTextArea
                                                    aria-label={this.props.localeStrings.fieldDescription + constants.requiredAriaLabel}
                                                    placeholder={this.props.localeStrings.phDescription}
                                                    fluid={true}
                                                    maxLength={constants.maxCharLengthForMultiLine}
                                                    onChange={(evt) => this.onTextInputChange(evt, "incidentDesc")}
                                                    value={this.state.incDetailsItem ? (this.state.incDetailsItem.incidentDesc ? this.state.incDetailsItem.incidentDesc : '') : ''}
                                                    className="incident-details-description-area"
                                                    ref={this.incidentDescription}
                                                />
                                                {this.state.inputValidation.incidentDescriptionHasError && (
                                                    <label aria-live="polite" role="alert" className="message-label">{this.props.localeStrings.incidentDescRequired}</label>
                                                )}
                                            </div>
                                        </Col>
                                    </Row>
                                    <Row xs={1} sm={2} md={3}>
                                        <Col md={4} sm={8} xs={12}>
                                            <div className="incident-grid-item">
                                                <label className="severity-label">{this.props.localeStrings.fieldSeverity}</label>
                                                <div className="slider_labels">
                                                    {constants.severity.map((item, index) => {
                                                        return (
                                                            <div
                                                                className={index === this.state.selectedSeverity ? "slider_labels-label bold" : "slider_labels-label"}
                                                                key={index}
                                                            >
                                                                {item}
                                                            </div>
                                                        );
                                                    })}
                                                </div>
                                                <ReactSlider
                                                    className="horizontal-slider"
                                                    marks
                                                    markClassName="example-mark"
                                                    min={0}
                                                    max={3}
                                                    value={this.state.selectedSeverity}
                                                    ariaLabel={this.props.localeStrings.fieldSeverity + constants.severity[this.state.selectedSeverity]}
                                                    thumbActiveClassName="focus-indicator"
                                                    trackClassName="example-track"
                                                    onChange={(index) => this.setState({
                                                        selectedSeverity: index
                                                    })}
                                                    renderMark={(props: any) => {
                                                        if (props.key < this.state.selectedSeverity) {
                                                            props.className = "example-mark example-mark-completed";
                                                            props.tabIndex = -1
                                                        } else if (props.key > this.state.selectedSeverity) {
                                                            props.className = "example-mark example-mark-active";
                                                            props.tabIndex = -1
                                                        }
                                                        else if (props.key === this.state.selectedSeverity) {
                                                            props.className = `example-mark ${constants.severity[props.key]}`;
                                                            props.tabIndex = 0
                                                        }
                                                        return <span aria-label={props.key === this.state.selectedSeverity ? this.props.localeStrings.fieldSeverity + constants.severity[props.key] + constants.selectedAriaLabel : this.props.localeStrings.fieldSeverity + constants.severity[props.key]} {...props} />;
                                                    }}
                                                />
                                            </div>
                                        </Col>
                                        {this.props.isEditMode &&
                                            <Col md={8} sm={8} xs={12}>
                                                <div className="incident-grid-item">
                                                    <label className="FormInput-label">{this.props.localeStrings.fieldReasonForUpdate}</label>
                                                    <FormTextArea
                                                        aria-label={this.props.localeStrings.fieldReasonForUpdate + constants.requiredAriaLabel}
                                                        placeholder={this.props.localeStrings.phReasonForUpdate}
                                                        fluid={true}
                                                        maxLength={constants.maxCharLengthForMultiLine}
                                                        onChange={(evt) => this.onTextInputChange(evt, "reasonForUpdate")}
                                                        className="incident-details-reason-update-area"
                                                    />
                                                    {this.state.inputValidation.incidentReasonForUpdateHasError && (
                                                        <label aria-live="polite" role="alert" className="message-label">{this.props.localeStrings.reasonForUpdateRequired}</label>
                                                    )}
                                                </div>
                                            </Col>
                                        }
                                    </Row>
                                    <h2><div className="incident-form-head-text">{this.props.localeStrings.headerRoleAssignment}</div></h2>
                                    <Row xs={1} sm={1} md={2}>
                                        <Col md={4} sm={8} xs={12}>
                                            <div className="incident-grid-item">
                                                <label className="FormInput-label">{this.props.localeStrings.fieldAdditionalRoles}</label>
                                                <FormDropdown
                                                    aria-label={this.props.localeStrings.fieldAdditionalRoles + constants.requiredAriaLabel}
                                                    placeholder={this.props.localeStrings.phRoles}
                                                    items={this.state.dropdownOptions ? this.state.dropdownOptions["roleOptions"] : []}
                                                    fluid={true}
                                                    autoSize
                                                    onChange={this.onRoleChange}
                                                    value={this.state.incDetailsItem ? (this.state.incDetailsItem.selectedRole ? this.state.incDetailsItem.selectedRole : '') : ''}
                                                    className={this.state.incDetailsItem && this.state.incDetailsItem.selectedRole ? "select-role-dropdown" : "select-role-placeholder"}
                                                    id="addRole-dropdown"
                                                    aria-labelledby="addRole-dropdown"
                                                />
                                            </div>
                                            {this.state.incDetailsItem.selectedRole && this.state.incDetailsItem.selectedRole.indexOf("New Role") > -1 ?
                                                <>
                                                    <div className="incident-grid-item">
                                                        <label className="FormInput-label">{this.props.localeStrings.fieldAddRoleName}</label>
                                                        <FormInput
                                                            aria-label={this.props.localeStrings.fieldAddRoleName + constants.requiredAriaLabel}
                                                            placeholder={this.props.localeStrings.phAddRoleName}
                                                            fluid={true}
                                                            maxLength={constants.maxCharLengthForSingleLine}
                                                            onChange={(evt) => this.onAddNewRoleChange(evt)}
                                                            value={this.state.newRoleString}
                                                            className="incident-details-input-field"
                                                            successIndicator={false}
                                                        />
                                                    </div>
                                                    <div className="incident-grid-item">
                                                        <Button
                                                            primary
                                                            onClick={this.addNewRole}
                                                            disabled={this.state.isCreateNewRoleBtnDisabled}
                                                            id={this.state.isCreateNewRoleBtnDisabled ? "manage-role-disabled-btn" : "manage-role-btn"}
                                                            fluid={!this.state.isDesktop}
                                                            title={this.props.localeStrings.btnCreateRole}
                                                        >
                                                            <img src={require("../assets/Images/AddIcon.svg").default}
                                                                alt="add"
                                                                className={`manage-role-btn-icon${this.props.currentThemeName === constants.contrastMode ? " add-icon-contrast" : ""}`}
                                                            />
                                                            &nbsp;&nbsp;&nbsp;
                                                            <label className="manage-role-btn-label">{this.props.localeStrings.btnCreateRole}</label>
                                                        </Button>
                                                    </div>
                                                </>
                                                :
                                                <>
                                                    <div className="incident-grid-item" ref={this.normalSearchUserRef}>
                                                        <label className="FormInput-label">{this.props.localeStrings.fieldSearchUser}</label>
                                                        <>{this.iconWithTooltip("info", this.props.localeStrings.roleUserInfoTooltipContent, "role-user-info-icon", "role-user-info-icon-tooltip")}</>
                                                        <PeoplePicker
                                                            ariaLabel={this.props.localeStrings.fieldSearchUser + constants.requiredAriaLabel}
                                                            selectionMode="multiple"
                                                            type={PersonType.person}
                                                            userType={UserType.user}
                                                            selectionChanged={this.handleAssignedUserChange}
                                                            placeholder={this.props.localeStrings.phSearchUser}
                                                            className="incident-details-people-picker"
                                                            selectedPeople={this.state.selectedUsers}
                                                            title={this.props.localeStrings.fieldSearchUser}
                                                            ref={this.searchUser}
                                                            onKeyDown={this.setSearchUserAttributes}
                                                            onClick={this.setSearchUserAttributes}
                                                        />
                                                        {this.state.secIncCommanderUserHasRegexError && (
                                                            <label className="message-label">{this.props.localeStrings.guestUsersNotAllowedAsSecIncCommanderErrorMsg}</label>
                                                        )}
                                                    </div>
                                                    <div className="incident-grid-item" ref={this.normalSearchLeadRef}>
                                                        <label className="lead-people-picker">{this.props.localeStrings.roleLeadLabel}</label>
                                                        <PeoplePicker
                                                            ariaLabel={this.props.localeStrings.roleLeadLabel}
                                                            selectionMode="single"
                                                            type={PersonType.person}
                                                            userType={UserType.user}
                                                            selectionChanged={this.handleAssignedLeadChange}
                                                            placeholder={this.props.localeStrings.phSearchUser}
                                                            className="incident-details-people-picker"
                                                            selectedPeople={this.state.selectedLead}
                                                            title={this.props.localeStrings.roleLeadLabel}
                                                        />
                                                        {this.state.secIncCommanderLeadHasRegexError && (
                                                            <label className="message-label">{this.props.localeStrings.guestUsersNotAllowedAsSecIncCommanderErrorMsg}</label>
                                                        )}
                                                    </div>
                                                    <div className="incident-grid-item">
                                                        <Fluent9CheckBox
                                                            label={this.props.localeStrings.roleCheckboxTooltip}
                                                            aria-label={this.props.localeStrings.roleCheckboxTooltip}
                                                            onChange={(_, data) => this.setState({ saveDefaultRoleCheck: data.checked })}
                                                            className="role-checkbox"
                                                            checked={this.state.saveDefaultRoleCheck}
                                                        />
                                                    </div>
                                                    <div className="incident-grid-item">
                                                        <Button
                                                            primary
                                                            onClick={this.addRoleAssignment}
                                                            disabled={this.state.isAddRoleAssignmentBtnDisabled}
                                                            id={this.state.isAddRoleAssignmentBtnDisabled ? "manage-role-disabled-btn" : "manage-role-btn"}
                                                            fluid={!this.state.isDesktop}
                                                            title={this.props.localeStrings.btnAddUser}>
                                                            <img src={require("../assets/Images/AddIcon.svg").default}
                                                                alt=""
                                                                className={`manage-role-btn-icon${this.props.currentThemeName === constants.contrastMode ? " add-icon-contrast" : ""}`}
                                                            />
                                                            &nbsp;&nbsp;&nbsp;
                                                            <label className="manage-role-btn-label">{this.props.localeStrings.btnAddUser}</label>
                                                        </Button>
                                                    </div>
                                                    <div role="status" className="add-role-message" aria-live="polite">{this.state.roleAddSuccessMessage}</div>
                                                </>
                                            }
                                        </Col>
                                        <Col md={8} sm={12} xs={12}>
                                            <Container as="table" className="role-assignment-table">
                                                <Row as="tr" xs={3} sm={3} md={3} className={`role-grid-thead${isDarkOrContrastTheme ? " table-header-darkcontrast" : ""}`}>
                                                    <Col as="th" md={3} sm={3} xs={3} key={0}>{this.props.localeStrings.headerRole}</Col>
                                                    <Col as="th" md={3} sm={3} xs={3} key={1} className="thead-border-left">{this.props.localeStrings.headerUsers}</Col>
                                                    <Col as="th" md={3} sm={3} xs={3} key={2} className="thead-border-left">{this.props.localeStrings.leadLabel}</Col>
                                                    <Col as="th" md={1} sm={1} xs={1} key={3} className="thead-border-left col-center">
                                                        <PeopleCheckmark24Regular className="role-header-icon" title={this.props.localeStrings.roleCheckboxTooltip} />
                                                    </Col>
                                                    <Col as="th" md={1} sm={1} xs={1} key={4} className="thead-border-left col-center">
                                                        <PeopleEdit24Regular className="role-header-icon" title={this.props.localeStrings.headerEdit} />
                                                    </Col>
                                                    <Col as="th" md={1} sm={1} xs={1} key={5} className="thead-border-left col-center">
                                                        <Delete24Regular title={this.props.localeStrings.headerDelete} className="role-header-icon" />
                                                    </Col>
                                                </Row>
                                                {this.state.roleAssignments.map((item, index) => (
                                                    <>
                                                        {this.state.isRoleInEditMode[index] ?
                                                            <>
                                                                <Row as="tr" xs={4} sm={4} md={4} key={"edit-" + item.role} className="role-grid-tbody">
                                                                    <Col as="td" md={10} sm={8} xs={8}>
                                                                        <label className="role-grid-tbody-peoplepicker-label"> {this.props.localeStrings.headerUsers}: </label>
                                                                        <PeoplePicker
                                                                            ariaLabel={this.props.localeStrings.fieldSearchUser}
                                                                            title={this.props.localeStrings.fieldSearchUser}
                                                                            selectionMode="multiple"
                                                                            type={PersonType.person}
                                                                            userType={UserType.user}
                                                                            selectionChanged={(selectedValue) => this.handleAssignedUserChangeInEditMode(selectedValue, index)}
                                                                            placeholder={this.props.localeStrings.phSearchUser}
                                                                            className="incident-details-people-picker"
                                                                            selectedPeople={this.state.selectedUsersInEditMode}
                                                                            tabIndex={0}
                                                                            ref={this.searchUserEditMode}
                                                                            onKeyDown={this.setSearchUserEditModeAttributes}
                                                                            onClick={this.setSearchUserEditModeAttributes}
                                                                        />
                                                                        {this.state.secIncCommanderUserInEditModeHasRegexError && (
                                                                            <>
                                                                                <label className="error-message-label">{this.props.localeStrings.guestUsersNotAllowedAsSecIncCommanderErrorMsg}</label>
                                                                                <br />
                                                                            </>
                                                                        )}
                                                                        <label className="role-grid-tbody-peoplepicker-label"> {this.props.localeStrings.leadLabel}: </label>
                                                                        <PeoplePicker
                                                                            ariaLabel={this.props.localeStrings.roleLeadLabel}
                                                                            title={this.props.localeStrings.roleLeadLabel}
                                                                            selectionMode="single"
                                                                            type={PersonType.person}
                                                                            userType={UserType.user}
                                                                            selectionChanged={(selectedValue) => this.handleAssignedLeadChangeInEditMode(selectedValue, index)}
                                                                            placeholder={this.props.localeStrings.phSearchUser}
                                                                            className="incident-details-people-picker"
                                                                            selectedPeople={this.state.selectedLeadInEditMode}
                                                                            tabIndex={0}
                                                                        />
                                                                        {this.state.secIncCommanderLeadInEditModeHasRegexError && (
                                                                            <label className="error-message-label">{this.props.localeStrings.guestUsersNotAllowedAsSecIncCommanderErrorMsg}</label>
                                                                        )}
                                                                    </Col>
                                                                    <Col as="td" md={1} sm={2} xs={2} className="editRoleCol">
                                                                        <Save24Regular aria-label="Save"
                                                                            className="role-icon"
                                                                            onClick={(e: any) => this.updateRoleAssignment(index)}
                                                                            onKeyDown={(event: any) => {
                                                                                if (event.key === constants.enterKey)
                                                                                    this.updateRoleAssignment(index)
                                                                            }}
                                                                            title={this.props.localeStrings.saveIcon}
                                                                            tabIndex={0}
                                                                            role="button"
                                                                        />
                                                                    </Col>
                                                                    <Col as="td" md={1} sm={2} xs={2} className="editRoleCol">
                                                                        <Dismiss24Regular aria-label="Cancel"
                                                                            className="role-icon"
                                                                            onClick={(e: any) => this.exitEditModeForRoles(index)}
                                                                            onKeyDown={(event: any) => {
                                                                                if (event.key === constants.enterKey)
                                                                                    this.exitEditModeForRoles(index)
                                                                            }}
                                                                            title={this.props.localeStrings.cancelIcon}
                                                                            tabIndex={0}
                                                                            role="button"
                                                                        />
                                                                    </Col>
                                                                </Row>
                                                            </>
                                                            :
                                                            <Row as="tr" xs={3} sm={3} md={3} key={"role-table-" + item.role} className="role-grid-tbody">
                                                                <Col as="td" md={3} sm={3} xs={3}>{item.role}</Col>
                                                                <Col as="td" md={3} sm={3} xs={3}>{item.userNamesString}</Col>
                                                                <Col as="td" md={3} sm={3} xs={3}>{item.leadNameString}</Col>
                                                                <Col as="td" md={1} sm={1} xs={1} className="col-center role-body-checkbox">
                                                                    <Fluent9CheckBox
                                                                        title={this.props.localeStrings.incidentTypeDefaultRoleCheckboxLabel}
                                                                        aria-label={this.props.localeStrings.roleCheckboxTooltip}
                                                                        onChange={(ev, isChecked) => this.onChecked(ev, Boolean(isChecked.checked), index)}
                                                                        defaultChecked={item.saveDefault}
                                                                    />
                                                                </Col>
                                                                <Col as="td" md={1} sm={1} xs={1} className="col-center role-body-icons">
                                                                    <PeopleEdit24Regular
                                                                        className="role-icon"
                                                                        onClick={(e: any) => this.editRoleItem(index)}
                                                                        onKeyDown={(event: any) => {
                                                                            if (event.key === constants.enterKey)
                                                                                this.editRoleItem(index)
                                                                        }}
                                                                        title={this.props.localeStrings.headerEdit}
                                                                        tabIndex={0}
                                                                        role="button"
                                                                    />
                                                                </Col>
                                                                <Col as="td" md={1} sm={1} xs={1} className="col-center role-body-icons">
                                                                    <Delete24Regular
                                                                        className="role-icon"
                                                                        onClick={(e: any) => this.deleteRoleItem(index)}
                                                                        onKeyDown={(event: any) => {
                                                                            if (event.key === constants.enterKey)
                                                                                this.deleteRoleItem(index)
                                                                        }}

                                                                        title={this.props.localeStrings.headerDelete}
                                                                        tabIndex={0}
                                                                        role="button"
                                                                    />
                                                                </Col>
                                                            </Row>
                                                        }
                                                    </>
                                                ))}
                                            </Container>
                                            {this.state.roleAssignments.length > 0 ?
                                                <div className="role-assignment-table">
                                                    <Fluent9CheckBox
                                                        label={this.props.localeStrings.incidentTypeDefaultRoleCheckboxLabel}
                                                        aria-label={this.props.localeStrings.incidentTypeDefaultRoleCheckboxLabel}
                                                        onChange={(_ev, isChecked) => this.setState({ saveIncidentTypeDefaultRoleCheck: isChecked.checked })}
                                                        className="assets-save-default-checkbox"
                                                        checked={this.state.saveIncidentTypeDefaultRoleCheck}
                                                    />
                                                </div>
                                                : null}
                                        </Col>
                                    </Row>
                                    <div className="incident-form-head-text">{this.props.localeStrings.assetsLabel}</div>
                                    <Row className="inc-assets-wrapper">
                                        <Col xl={10} lg={12} className={`${this.state.toggleCloudStorageLocation ? "cloud-storage-enabled" : ""}`}>
                                            <div
                                                className={`cloud-link-asset-field-wrapper${this.state.inputRegexValidation.incidentCloudStorageLinkHasError ||
                                                    this.state.inputValidation.cloudStorageLinkHasError ? " cloud-link-asset-field-with-error-wrapper" : ""}`}
                                            >
                                                <Toggle
                                                    checked={this.state.toggleCloudStorageLocation}
                                                    label={this.props.localeStrings.cloudStorageFieldLabel}
                                                    inlineLabel
                                                    onChange={(_, checked: any) => this.onToggleCloudStorageLink(checked)}
                                                    className="inc-assets-tgle-btn"
                                                    disabled={this.state.isEditMode ? true : false}
                                                    id="cloud-storage-toggle-btn"
                                                />
                                                {this.state.toggleCloudStorageLocation &&
                                                    <>
                                                        <div className="cloud-field-with-icon">
                                                            <div className="cloud-storage-field">
                                                                <FormInput
                                                                    type="text"
                                                                    placeholder={this.props.localeStrings.cloudStorageFieldPlaceholder}
                                                                    fluid={true}
                                                                    onChange={(event: any) => this.onTextInputChange(event, "cloudStorageLink")}
                                                                    value={this.state.incDetailsItem.cloudStorageLink}
                                                                    className={this.state.isEditMode ? "incident-details-input-field disabled-input" : "incident-details-input-field"}
                                                                    successIndicator={false}
                                                                    disabled={this.state.isEditMode}
                                                                />
                                                                {this.state.inputRegexValidation.incidentCloudStorageLinkHasError &&
                                                                    <label className="error-message-label">{this.props.localeStrings.cloudStorageFieldRegexMessage}</label>}
                                                                {this.state.inputValidation.cloudStorageLinkHasError &&
                                                                    <label className="error-message-label">{this.props.localeStrings.cloudStorageFieldErrorMessage}</label>
                                                                }
                                                            </div>
                                                            <span className="cloud-icon-area">
                                                                <a
                                                                    href={this.dataService.isValidHttpUrl(this.state.incDetailsItem.cloudStorageLink) ?
                                                                        new URL(this.state.incDetailsItem.cloudStorageLink).href : "/"}
                                                                    target="_blank"
                                                                    rel="noreferrer"
                                                                    className={!this.dataService.isValidHttpUrl(this.state.incDetailsItem.cloudStorageLink) ?
                                                                        "disabled-link" : ""}
                                                                    title={this.props.localeStrings.testCloudStorageLocation}
                                                                >
                                                                    <CloudLink24Regular className="cloud-icon" />
                                                                </a>
                                                            </span>
                                                        </div>
                                                        {!this.state.isEditMode &&
                                                            <Fluent9CheckBox
                                                                label={<div className={this.state.isEditMode ? "save-default-checkbox-label disabled-label" : "save-default-checkbox-label"}>
                                                                    {this.props.localeStrings.saveDefaultLabel}
                                                                    {this.iconWithTooltip(
                                                                        "Info", //Icon library name
                                                                        this.props.localeStrings.cloudStorageFieldSaveDefaultTooltipContent,
                                                                        "save-default-checkbox-info-icon", //Class name
                                                                        "cloud-storage-save-default-tooltip"
                                                                    )}
                                                                </div>}
                                                                aria-label={this.props.localeStrings.saveDefaultLabel}
                                                                disabled={this.state.isEditMode}
                                                                className="assets-save-default-checkbox"
                                                                onChange={(_: any, checked: any) => this.setState({ saveDefaultCloudStorageLink: checked })}
                                                            />
                                                        }
                                                    </>
                                                }
                                            </div>
                                        </Col>
                                        <Col xl={10} lg={12} className='guest-users-toggle'>
                                            <div className="inc-assets-field-flex-column">
                                                <Toggle
                                                    checked={this.state.toggleGuestUsers}
                                                    label={
                                                        <div className="tgle-btn-label">
                                                            {this.props.localeStrings.guestUsersLabel}
                                                            {this.iconWithTooltip(
                                                                "Info", //Icon library name
                                                                this.props.localeStrings.guestUsersInfoIconTooltipContent,
                                                                "tgle-btn-info-icon", //Class name
                                                                "guest-users-info-tooltip"
                                                            )}
                                                        </div>
                                                    }
                                                    inlineLabel
                                                    onChange={(_, checked: any) => this.onToggleGuestUsers(checked)}
                                                    className="inc-assets-tgle-btn"
                                                />
                                                {this.state.toggleGuestUsers &&
                                                    <div className="guest-user-fields-wrapper">
                                                        {this.state.incDetailsItem.guestUsers.map((user: IGuestUsers, idx: number) => {
                                                            return (
                                                                <div className="guest-field-wrapper" key={"guest-user-field-row-" + idx}>
                                                                    <div>
                                                                        <FormInput
                                                                            type="text"
                                                                            label={this.props.localeStrings.emailIdLabel}
                                                                            aria-label={this.props.localeStrings.emailIdLabel}
                                                                            placeholder={this.props.localeStrings.guestEmailIdPlaceholder}
                                                                            fluid={true}
                                                                            maxLength={254}
                                                                            onChange={(_, value) => this.updateGuestUser(value, idx, "email")}
                                                                            value={user.email}
                                                                            className="incident-details-input-field"
                                                                            successIndicator={false}
                                                                            required={true}
                                                                        />
                                                                        {user.hasEmailValidationError && <label className="error-message-label">
                                                                            {this.props.localeStrings.guestemailIdValidationError}
                                                                        </label>}
                                                                        {user.hasEmailRegexError && <label className="error-message-label">
                                                                            {this.props.localeStrings.guestEmailIdRegexError}
                                                                        </label>}
                                                                    </div>
                                                                    <div>
                                                                        <FormInput
                                                                            type="text"
                                                                            label={this.props.localeStrings.displayNameLabel}
                                                                            placeholder={this.props.localeStrings.guestDisplayNamePlaceholder}
                                                                            fluid={true}
                                                                            maxLength={200}
                                                                            onChange={(_, value) => this.updateGuestUser(value, idx, "displayName")}
                                                                            value={user.displayName}
                                                                            className="incident-details-input-field"
                                                                            successIndicator={false}
                                                                            required={true}
                                                                        />
                                                                        {user.hasDisplayNameValidationError && <label className="error-message-label">
                                                                            {this.props.localeStrings.guestDisplayNameValidationError}
                                                                        </label>}
                                                                    </div>
                                                                </div>
                                                            );
                                                        })}
                                                        {this.state.incDetailsItem.guestUsers.length < 10 &&
                                                            <Button
                                                                icon={<AddIcon />}
                                                                content={this.props.localeStrings.addMoreBtnLabel}
                                                                iconPosition="before"
                                                                onClick={() => this.addGuestUserInput()}
                                                                className="add-chnl-btn"
                                                                title={this.props.localeStrings.addMoreBtnLabel}
                                                            />
                                                        }
                                                        {
                                                            this.state.inputValidation.guestUsersHasError &&
                                                            <label className="error-message-label">
                                                                {this.props.localeStrings.guestUsersValidationError}
                                                            </label>
                                                        }
                                                    </div>
                                                }
                                            </div>
                                        </Col>
                                        {!this.state.isEditMode &&
                                            <Col xl={10} lg={12} className="additional-channels-toggle">
                                                <div className="inc-assets-field-flex-column">
                                                    <div className="additional-chnls-label">
                                                        <Toggle
                                                            checked={this.state.toggleAdditionalChannels}
                                                            label={
                                                                <div className="tgle-btn-label">
                                                                    {this.props.localeStrings.additionalChannelsFieldLabel}
                                                                    {this.iconWithTooltip(
                                                                        "Info", //Icon library name
                                                                        this.props.localeStrings.additionalChannelsFieldInfoIconTooltipContent,
                                                                        "tgle-btn-info-icon", //Class name
                                                                        "additional-channels-info-tooltip"
                                                                    )}
                                                                </div>
                                                            }
                                                            inlineLabel
                                                            onChange={(_, checked: any) => this.onToggleAdditionChannels(checked)}
                                                            className="inc-assets-tgle-btn"
                                                            tabIndex={0}
                                                        />
                                                        <Fluent9CheckBox
                                                            label={<div className="save-default-checkbox-label">
                                                                {this.props.localeStrings.saveDefaultLabel}
                                                                {this.iconWithTooltip(
                                                                    "Info", //Icon library name
                                                                    this.props.localeStrings.additionalChannelsFieldSaveDefaultTooltipContent,
                                                                    "save-default-checkbox-info-icon", //Class name
                                                                    "additional-channels-save-default-tooltip"
                                                                )}
                                                            </div>}
                                                            aria-label={this.props.localeStrings.saveDefaultLabel}
                                                            onChange={(_: any, checked: any) => this.setState({ saveDefaultAdditionalChannels: checked })}
                                                            className="assets-save-default-checkbox"
                                                        />
                                                    </div>
                                                    {this.state.toggleAdditionalChannels &&
                                                        <div className="additional-chnl-fields">
                                                            {this.state.incDetailsItem.additionalTeamChannels.map((channel: IAdditionalTeamChannels,
                                                                idx: number) => {
                                                                return (
                                                                    <div className={`chnl-field-combobox-wrapper${channel.hasRegexError ? " field-has-error" : ""}`}>
                                                                        <div className='field-with-error-msg-wrapper'>
                                                                            <div className="field-with-cancel-icon">
                                                                                <FormInput
                                                                                    type="text"
                                                                                    placeholder={this.props.localeStrings.additionalChannelsFieldPlaceholder}
                                                                                    fluid={true}
                                                                                    maxLength={constants.maxCharLengthForSingleLine}
                                                                                    onChange={(eve, value) => this.updateAdditionalChannel(eve, value, idx)}
                                                                                    value={channel.channelName}
                                                                                    className="incident-details-input-field"
                                                                                />
                                                                                <span
                                                                                    className="chnl-remove-icon"
                                                                                    title={this.props.localeStrings.headerDelete}
                                                                                    onClick={() => this.removeChannelInput(idx)}
                                                                                    onKeyDown={(evt: any) => { if (evt.key === constants.enterKey) this.removeChannelInput(idx) }}
                                                                                    tabIndex={0}
                                                                                    role="button"
                                                                                >
                                                                                    <Dismiss24Regular />
                                                                                </span>
                                                                            </div>
                                                                            {channel.hasRegexError &&
                                                                                <label className="error-message-label">
                                                                                    {channel.regexErrorMessage}
                                                                                </label>
                                                                            }
                                                                        </div>
                                                                        <Fluent9CheckBox
                                                                            label={this.props.localeStrings.privateLabel}
                                                                            aria-label={this.props.localeStrings.privateLabel}
                                                                            checked={this.state.incDetailsItem.additionalTeamChannels[idx].channelType === constants.privateChannel}
                                                                            onChange={(eve: any, { checked }: any) => this.onPrivateChanged(eve, checked, idx)}
                                                                            className="assets-private-channel-checkbox"
                                                                        />
                                                                        {this.state.incDetailsItem.additionalTeamChannels[idx].channelType === constants.privateChannel &&
                                                                            <Fluent9Combobox
                                                                                className="private-channel-role-combobox"
                                                                                multiselect={true}
                                                                                positioning="after"
                                                                                inlinePopup={true}
                                                                                placeholder={this.props.localeStrings.privateChannelPlaceholder}
                                                                                value={this.state.incDetailsItem.additionalTeamChannels[idx].selectedRoleUsers ? this.state.incDetailsItem.additionalTeamChannels[idx].selectedRoleUsers : ""}
                                                                                title={this.state.incDetailsItem.additionalTeamChannels[idx].selectedRoleUsers}
                                                                                onOptionSelect={(eve: any, data: any) => this.onRolesSelect(eve, data, idx)}
                                                                                listbox={{ id: "private-channel-role-listbox" }}
                                                                                disabled={this.state.roleAssignments.length === 0}
                                                                            >
                                                                                <div aria-live="polite" role="alert" className="private-channel-callout-note">
                                                                                    <b>{this.props.localeStrings.NoteLabel}: </b>
                                                                                    {this.props.localeStrings.privateChannelCalloutNote}
                                                                                </div>
                                                                                {this.state.roleAssignments?.map((option: any) => {
                                                                                    return (
                                                                                        option.role !== constants.secondaryIncidentCommanderRole ?
                                                                                            <>
                                                                                                <Option
                                                                                                    checkIcon={this.state.incDetailsItem.additionalTeamChannels[idx]?.expandedGroups[option.role] ? <ChevronDown16Regular /> : <ChevronRight16Regular />}
                                                                                                    value={option.role}
                                                                                                    className="role-option-header"
                                                                                                    aria-roledescription='group header'
                                                                                                    role="menuitem"
                                                                                                    aria-expanded={this.state.incDetailsItem.additionalTeamChannels[idx]?.expandedGroups[option.role]}
                                                                                                >
                                                                                                    {option.role}
                                                                                                </Option>
                                                                                                {this.state.incDetailsItem.additionalTeamChannels[idx]?.expandedGroups[option.role] && option.userDetailsObj?.map((user: any) => (
                                                                                                    <Option value={user.userName + "|" + user.userId + "|" + user.userEmail + "|{}"} aria-label={option.role + " " + user.userName}>{user.userName}</Option>
                                                                                                ))}
                                                                                                {this.state.incDetailsItem.additionalTeamChannels[idx]?.expandedGroups[option.role] && option.leadDetailsObj?.map((lead: any) => (
                                                                                                    <Option value={lead.userName + "|" + lead.userId + "|" + lead.userEmail + "|{}"} aria-label={option.role + " " + lead.userName}>{lead.userName}</Option>
                                                                                                ))}
                                                                                            </> : <></>
                                                                                    );
                                                                                })}
                                                                            </Fluent9Combobox>
                                                                        }
                                                                    </div>
                                                                );
                                                            })}
                                                            {this.state.incDetailsItem.additionalTeamChannels.length < 5 &&
                                                                <Button
                                                                    icon={<AddIcon />}
                                                                    content={this.props.localeStrings.addChannelBtnLabel}
                                                                    title={this.props.localeStrings.addChannelBtnLabel}
                                                                    iconPosition="before"
                                                                    onClick={() => this.addChannelInput()}
                                                                    className="add-chnl-btn"
                                                                />
                                                            }
                                                        </div>
                                                    }
                                                </div>
                                            </Col>
                                        }
                                    </Row>
                                    <br />
                                    <Row xs={1} sm={1} md={1}>
                                        <Col md={12} sm={12} xs={12}>
                                            <div className="new-incident-btn-area">
                                                <Flex hAlign="end" gap="gap.large" wrap={true}>
                                                    <Button
                                                        onClick={() => this.props.onBackClick("")}
                                                        className="new-incident-back-btn"
                                                        fluid={true}
                                                        title={this.props.localeStrings.btnBack}
                                                    >
                                                        <ChevronStartIcon /> &nbsp;
                                                        <label>{this.props.localeStrings.btnBack}</label>
                                                    </Button>
                                                    {this.props.isEditMode ?
                                                        <Button
                                                            primary
                                                            onClick={this.updateIncidentDetails}
                                                            fluid={true}
                                                            className={`new-incident-create-btn${this.props.currentThemeName === constants.contrastMode ? " create-icon-contrast" : ""}`}

                                                            title={this.props.localeStrings.btnUpdateIncident}
                                                        >
                                                            <img src={require("../assets/Images/ButtonEditIcon.svg").default} alt="" /> &nbsp;
                                                            <label>{this.props.localeStrings.btnUpdateIncident}</label>
                                                        </Button>
                                                        :
                                                        <Button
                                                            primary
                                                            onClick={this.createNewIncident}
                                                            fluid={true}
                                                            className={`new-incident-create-btn${this.props.currentThemeName === constants.contrastMode ? " create-icon-contrast" : ""}`}

                                                            title={this.props.localeStrings.btnCreateIncident}
                                                        >
                                                            <img src={require("../assets/Images/ButtonEditIcon.svg").default} alt="" /> &nbsp;
                                                            <label>{this.props.localeStrings.btnCreateIncident}</label>
                                                        </Button>
                                                    }
                                                </Flex>
                                            </div>
                                        </Col>
                                    </Row>
                                </div>
                            </div>
                        </div>
                    </>
                </div>
            </>
        );
    }
}

export default IncidentDetails;
