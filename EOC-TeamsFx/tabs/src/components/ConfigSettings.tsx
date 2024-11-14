import { MessageBar } from "@fluentui/react";
import { Input, Spinner } from "@fluentui/react-components";
import { Button, FormDropdown } from "@fluentui/react-northstar";
import { Icon } from "@fluentui/react/lib/Icon";
import { Label } from "@fluentui/react/lib/Label";
import { Toggle } from "@fluentui/react/lib/Toggle";
import { TooltipHost } from "@fluentui/react/lib/Tooltip";
import { ApplicationInsights } from "@microsoft/applicationinsights-web";
import { Client } from "@microsoft/microsoft-graph-client";
import React from "react";
import Col from "react-bootstrap/esm/Col";
import CommonService from "../common/CommonService";
import * as constants from "../common/Constants";
import * as graphConfig from "../common/graphConfig";
import siteConfig from "../config/siteConfig.json";

export interface IConfigSettingsProps {
    localeStrings: any;
    onBackClick(showMessageBar: string): void;
    currentUserDisplayName: string;
    currentUserEmail: string;
    isRolesEnabled: boolean;
    isUserAdmin: boolean;
    siteId: string;
    graph: Client;
    configRoleData: any;
    setState: any;
    tenantName: string;
    siteName: any;
    appInsights: ApplicationInsights;
    userPrincipalName: any;
    isMapViewerEnabled: boolean;
    azureMapsKeyConfigData: any;
    appTitle: string;
    appTitleData: any;
    editIncidentAccessRole: string;
    editIncidentAccessRoleData: any;
};
export interface IConfigSettingsState {
    enableRoles: boolean;
    messages: IMessages;
    showAssignRolesLink: boolean;
    showLoader: boolean;
    enableMapViewer: boolean;
    azureMapsKey: string;
    azureMapsKeyError: boolean;
    appTitle: string;
    appTitleKeyError: boolean;
    roleDropdownOptions: any;
    selectedRole: string;
}
export interface IMessages {
    roles: IMessageData;
    mapViewer: IMessageData;
    genericMessage: IMessageData;
    appTitle: IMessageData;
}
export interface IMessageData {
    messageType: number;
    message: string;
}
export default class ConfigSettings extends React.Component<IConfigSettingsProps, IConfigSettingsState> {
    constructor(props: IConfigSettingsProps) {
        super(props);

        //States
        this.state = {
            enableRoles: this.props.isRolesEnabled,
            messages: {
                roles: { messageType: -1, message: "" },
                mapViewer: { messageType: -1, message: "" },
                genericMessage: { messageType: -1, message: "" },
                appTitle: { messageType: -1, message: "" }
            },
            showAssignRolesLink: this.props.isRolesEnabled,
            showLoader: false,
            enableMapViewer: this.props.isMapViewerEnabled,
            azureMapsKey: this.props.azureMapsKeyConfigData?.value?.trim()?.length > 0 ? this.props.azureMapsKeyConfigData?.value : "",
            azureMapsKeyError: false,
            appTitle: this.props.appTitle,
            appTitleKeyError: false,
            roleDropdownOptions: '',
            selectedRole: this.props.editIncidentAccessRole
        }

        //bind methods
        this.updateSettings = this.updateSettings.bind(this);
    }

    //get roles for dropdown list on load
    public async componentDidMount() {
        await this.getRoleDropdownOptions();
    }

    //Create object for Common Services class
    private commonService = new CommonService();

    //Update Config Settings
    private updateSettings = async () => {
        try {
            let noChanges: boolean = true;
            // create graph endpoint for TEOC Config list
            const configListNewGraphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}${graphConfig.listsGraphEndpoint}/${siteConfig.configurationList}/items`;

            //Reset Generic Message
            this.setState((prevState) => ({
                messages: {
                    ...prevState.messages,
                    genericMessage: { messageType: -1, message: "" }
                }
            }));
            if (this.state.appTitle?.trim() !== this.props.appTitle?.trim() ||
                this.props.appTitleData.itemId === undefined) {
                if (this.state.appTitle === undefined || this.state.appTitle?.trim() === "") {
                    this.setState({ appTitleKeyError: true });
                } else {
                    noChanges = false;
                    this.setState({ showLoader: true });

                    //If AppTitle key is missing in config list add new item to the list
                    if (this.props.appTitleData.itemId === undefined) {
                        //create graph endpoint for TEOC Config list to add app title
                        const newItemAppTitleObj = { fields: { Value: this.state.appTitle.trim(), Title: constants.appTitleKey } };
                        const item = await this.commonService.addItemInList
                            <{ fields: { id: number; Title: string; Value: string } }>(
                                configListNewGraphEndpoint,
                                this.props.graph, newItemAppTitleObj
                            );
                        this.props.setState({
                            appTitleData: {
                                itemId: item?.fields?.id,
                                title: item?.fields?.Title,
                                value: item?.fields?.Value
                            }
                        });
                    } else {
                        //Update App Title in TEOC-config list
                        const graphConfigListEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}/lists/${siteConfig.configurationList}/items/${this.props.appTitleData.itemId}/fields`;
                        const updateAppTitleObj = { Value: this.state.appTitle.trim() };
                        await this.commonService.updateItemInList(graphConfigListEndpoint,
                            this.props.graph, updateAppTitleObj);
                    }
                    //Update App Title in Home Component states
                    this.props.setState({
                        appTitle: this.state.appTitle.trim(),
                        configRoleData: { ...this.props.configRoleData, value: this.state.appTitle.trim() }
                    });

                    //Update Config Settings States
                    this.setState(prevState => ({
                        messages: {
                            ...prevState.messages,
                            genericMessage: { messageType: 4, message: this.props.localeStrings.settingsSavedmessage }
                        }
                    }));
                }
            }

            //Add/Update EditAccessRole value in Config List
            if (this.state.selectedRole?.trim() !== this.props.editIncidentAccessRole?.trim() ||
                this.props.editIncidentAccessRoleData.itemId === undefined) {
                if (this.state.selectedRole != undefined && this.state.selectedRole?.trim() != "") {
                    noChanges = false;
                    this.setState({ showLoader: true });
                    //If EditAccessRole key is missing in config list add new item to the list
                    if (this.props.editIncidentAccessRoleData.itemId === undefined) {
                        const newEditAccessRoleObj = { fields: { Value: this.state.selectedRole.trim(), Title: constants.editIncidentAccessRoleKey } };
                        const objListItem = await this.commonService.addItemInList
                            <{ fields: { id: number; Title: string; Value: string } }>(
                                configListNewGraphEndpoint,
                                this.props.graph, newEditAccessRoleObj
                            );
                        this.props.setState({
                            editIncidentAccessRoleData: {
                                itemId: objListItem?.fields?.id,
                                title: objListItem?.fields?.Title,
                                value: objListItem?.fields?.Value
                            }
                        });
                    }
                    //If EditAccessRole key is already in config list update the item
                    else {
                        //Update Edit Access Role in TEOC-config list
                        const graphConfigListEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}/lists/${siteConfig.configurationList}/items/${this.props.editIncidentAccessRoleData.itemId}/fields`;
                        const updateEditAccessRoleObj = { Value: this.state.selectedRole.trim() };
                        await this.commonService.updateItemInList(graphConfigListEndpoint,
                            this.props.graph, updateEditAccessRoleObj);
                    }
                    //Update Edit Access Role in Home Component states
                    this.props.setState({
                        editIncidentAccessRole: this.state.selectedRole.trim(),
                        configRoleData: { ...this.props.configRoleData, value: this.state.selectedRole.trim() }
                    });

                    //Update Config Settings States
                    this.setState(prevState => ({
                        messages: {
                            ...prevState.messages,
                            genericMessage: { messageType: 4, message: this.props.localeStrings.settingsSavedmessage }
                        }
                    }));
                }
            }

            if (this.props.isRolesEnabled !== this.state.enableRoles) {
                noChanges = false;
                this.setState({ showLoader: true });

                //Update Role Settings in TEOC-config list
                const graphConfigListEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}/lists/${siteConfig.configurationList}/items/${this.props.configRoleData.itemId}/fields`;
                const updateEnableRolesObj = { Value: this.props.isRolesEnabled ? "False" : "True" };
                await this.commonService.updateItemInList(graphConfigListEndpoint,
                    this.props.graph, updateEnableRolesObj);

                //Disabling Roles
                if (this.props.isRolesEnabled) {
                    //Update Role Setting States
                    this.setState((prevState) => ({
                        showAssignRolesLink: false,
                        messages: {
                            ...prevState.messages,
                            roles: { messageType: 4, message: this.props.localeStrings.roleSettingsDisabledMessage }
                        }
                    }));

                    //Update Home Component states
                    this.props.setState({
                        isRolesEnabled: false,
                        configRoleData: { ...this.props.configRoleData, value: "False" }
                    });
                }
                //Enabling Roles
                else {
                    //Update user role as admin if the user doesn't exist in TEOC-UserRoles List
                    if (!this.props.isUserAdmin) {
                        const graphUserRolesListEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}${graphConfig.listsGraphEndpoint}/${siteConfig.userRolesList}/items`;
                        const updateUserRolesObj = {
                            fields: {
                                Title: this.props.currentUserEmail,
                                Name: this.props.currentUserDisplayName,
                                Role: constants.adminRole
                            }
                        };
                        await this.commonService.sendGraphPostRequest(graphUserRolesListEndpoint,
                            this.props.graph, updateUserRolesObj);
                        this.props.setState({
                            isUserAdmin: true
                        });
                    }

                    //Update Role Setting States
                    this.setState(prevState => ({
                        showAssignRolesLink: true,
                        messages: {
                            ...prevState.messages,
                            roles: { messageType: 4, message: this.props.localeStrings.roleSettingsEnabledMessage }
                        }
                    }));

                    //Update Home Component states
                    this.props.setState({
                        isRolesEnabled: true,
                        configRoleData: { ...this.props.configRoleData, value: "True" }
                    });
                }
            }

            //Add azure Map subscription Key record in TEOC-config list
            if (this.state.enableMapViewer && this.props.azureMapsKeyConfigData?.title === undefined &&
                this.state.azureMapsKey?.trim() !== "") {
                noChanges = false;
                this.setState({ showLoader: true });

                const listItem = { fields: { Title: constants.azureMapsKey, Value: this.state.azureMapsKey } };
                const configResponse = await this.commonService.sendGraphPostRequest(configListNewGraphEndpoint, this.props.graph, listItem);
                this.props.setState({
                    isMapViewerEnabled: this.state.enableMapViewer,
                    azureMapsKeyConfigData: {
                        ...this.props.azureMapsKeyConfigData,
                        title: configResponse.fields.Title,
                        value: this.state.azureMapsKey,
                        itemId: configResponse.fields.id
                    }
                });
                //Update Config Settings States
                this.setState(prevState => ({
                    messages: {
                        ...prevState.messages,
                        mapViewer: {
                            messageType: 4,
                            message: this.props.localeStrings.mapViewerKeyEnabledMessage
                        }
                    },
                    azureMapsKey: prevState.azureMapsKey?.trim()
                }));
            }
            //Update azure Map subscription Key in TEOC-config list
            else if ((this.props.azureMapsKeyConfigData?.value &&
                this.props.azureMapsKeyConfigData?.value?.trim() !== this.state.azureMapsKey?.trim()) ||
                this.props.isMapViewerEnabled !== this.state.enableMapViewer) {
                noChanges = false;
                //Validate azure Map subscription Key
                if (this.state.enableMapViewer &&
                    (this.state.azureMapsKey?.trim() === "" || this.state.azureMapsKey?.trim() === undefined)) {
                    this.setState(prevState => ({
                        azureMapsKeyError: true,
                        messages: {
                            ...prevState.messages,
                            mapViewer: { messageType: 1, message: this.props.localeStrings.mapViewerKeyRequiredMessage }
                        }
                    }));
                }
                //Update azure Map subscription Key
                else {
                    this.setState({ showLoader: true });
                    //Endpoint to update azure Map API Key in TEOC-config list
                    const graphConfigListEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}/lists/${siteConfig.configurationList}/items/${this.props.azureMapsKeyConfigData?.itemId}/fields`;
                    const updatedValue = this.state.enableMapViewer ? this.state.azureMapsKey?.trim() : "";
                    let updatedazureAPIKeyObj = { Value: updatedValue }
                    //Update azure Map subscription Key in TEOC-config list API Call
                    await this.commonService.updateItemInList(graphConfigListEndpoint, this.props.graph, updatedazureAPIKeyObj);

                    let messageToDisplay: string;
                    if (this.state.enableMapViewer && this.props.azureMapsKeyConfigData?.value?.trim() !== "" &&
                        this.state.azureMapsKey?.trim() !== "") {
                        messageToDisplay = this.props.localeStrings.mapViewerKeyUpdatedMessage;
                    }
                    else {
                        messageToDisplay = this.state.enableMapViewer ? this.props.localeStrings.mapViewerKeyEnabledMessage : this.props.localeStrings.mapViewerKeyDisabledMessage;
                    }

                    //Update Home Component states
                    this.props.setState({
                        isMapViewerEnabled: this.state.enableMapViewer,
                        azureMapsKeyConfigData: { ...this.props.azureMapsKeyConfigData, value: updatedValue }
                    });

                    //Update Config Settings States
                    this.setState({
                        messages: {
                            ...this.state.messages,
                            mapViewer: {
                                messageType: 4,
                                message: messageToDisplay
                            }
                        },
                        azureMapsKey: updatedValue
                    });
                }
            }

            if (noChanges) {
                this.setState({
                    messages: {
                        ...this.state.messages,
                        genericMessage: { messageType: 4, message: this.props.localeStrings.settingsSavedmessage },
                        mapViewer: { messageType: -1, message: "" },
                        roles: { messageType: -1, message: "" }
                    }
                });
            }
            this.setState({ showLoader: false });
        }
        catch (error: any) {
            this.setState({
                messages: {
                    ...this.state.messages,
                    genericMessage: { messageType: 1, message: error.message }
                },
                showLoader: false
            });

            console.error(
                constants.errorLogPrefix + "_" + constants.componentNames.AdminSettingsComponent + "_" +
                constants.componentNames.ConfigSettingsComponent + "_updateSettings \n",
                JSON.stringify(error)
            );

            //log exception to AppInsights
            this.commonService.trackException(this.props.appInsights, error,
                constants.componentNames.ConfigSettingsComponent, 'updateSettings', this.props.userPrincipalName);
        }
    }

    //Get dropdown options for Incident Status, Incident Type and Roles dropdown
    private getRoleDropdownOptions = async () => {
        try {
            const roleGraphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}${graphConfig.listsGraphEndpoint}/${siteConfig.roleAssignmentList}/items?$expand=fields&$Top=5000`;

            let rolesList = await this.commonService.getDropdownOptions(roleGraphEndpoint, this.props.graph);
            rolesList = rolesList.sort();

            //Remove the Secondarycommander and 'new role' from the dropdown list
            rolesList.splice(rolesList.indexOf(constants.secondaryIncidentCommanderRole), 1,);
            rolesList.splice(rolesList.indexOf(constants.newRole), 1);
            //Add 'None' to the dropdown. This is to allow users to remove any old mapping 
            rolesList.splice(0, 0, constants.noneOption);

            this.setState({
                roleDropdownOptions: rolesList,
                showLoader: false,
            })

        } catch (error) {
            console.error(
                constants.errorLogPrefix + "ConfigSettings_getRoleDropdownOptions \n",
                JSON.stringify(error)
            );
            // Log Exception
            this.commonService.trackException(this.props.appInsights, error, constants.componentNames.ConfigSettingsComponent, 'ConfigSettings_getRoleDropdownOptions', this.props.userPrincipalName);
        }
    }

    // on change of dropdown set the state of selected role
    private onRoleChange = (_event: any, selectedRole: any) => {
        this.setState({
            selectedRole: selectedRole.value
        })
    }

    //Render Method
    render() {
        return (
            <div className='config-settings-wrapper'>
                {Object.entries(this.state.messages).map((messageItem: any) => {
                    if (messageItem[1].messageType === -1) return <></>;
                    return (<Col xl={6} lg={6} md={8} sm={10} xs={12} key={messageItem[0]}>
                        <MessageBar
                            messageBarType={messageItem[1].messageType}
                            title={messageItem[1].message}
                            className="config-settings-message-bar"
                            isMultiline={true}
                            onDismiss={() => {
                                this.setState(prevState => ({
                                    messages: {
                                        ...prevState.messages,
                                        [messageItem[0]]: { messageType: -1, message: "" }
                                    }
                                }))
                            }}
                        >
                            {messageItem[1].message}
                        </MessageBar>
                    </Col>);
                })}
                <div className='settings-sub-wrapper'>
                    <div className={`config-settings-app-title-wrapper ${this.state.appTitleKeyError ? "field-with-error" : ""}`}>
                        <div className="app-title-label">
                            <Label>
                                {this.props.localeStrings.AppTitleLabel}
                                <span className='info-icon'>
                                    <TooltipHost
                                        content={<span dangerouslySetInnerHTML={{ __html: this.props.localeStrings.AppTitleInfoIconText }} />}
                                        calloutProps={{ gapSpace: 0 }}
                                        id="app-title-tooltip"
                                    >
                                        <Icon iconName='info' tabIndex={0} aria-label="Info"
                                            aria-describedby="app-title-tooltip" role="button" />
                                    </TooltipHost>
                                </span>
                            </Label>
                        </div>
                        <div className='app-title-input-wrapper'>
                            <Input                               
                                placeholder={this.props.localeStrings.AppTitlePlaceholderText}
                                className="app-title-input-box"
                                value={this.state.appTitle}
                                maxLength={50}
                                onChange={(_ev, data: any) => (this.setState((prevState) => ({
                                    appTitle: data.value,
                                    appTitleKeyError: data?.value?.trim() === "",
                                    messages: {
                                        ...prevState.messages,
                                        appTitle: { messageType: -1, message: "" }
                                    }
                                })))}
                            />
                            {this.state.appTitleKeyError &&
                                <span className='app-title-error-msg' aria-live="polite" role="alert">
                                    {this.props.localeStrings.AppTitleErrorLabel}</span>
                            }
                        </div>
                    </div>
                    <div className="config-settings-toggle-btn-wrapper">
                        <Toggle
                            checked={this.state.enableRoles}
                            label={
                                <div className="toggle-btn-label">
                                    {this.props.localeStrings.enablesRoles}
                                    <span className='toggle-info-icon'>
                                        <TooltipHost
                                            content={this.props.localeStrings.roleSettingsInfoIconTooltip}
                                            calloutProps={{ gapSpace: 0 }}
                                            id="enable-roles-tooltip"
                                        >
                                            <Icon iconName='info' tabIndex={0} aria-label={this.props.localeStrings.roleSettingsInfoIconTooltip}
                                                aria-describedby="enable-roles-tooltip" role="img" />
                                        </TooltipHost>
                                    </span>
                                </div>
                            }
                            inlineLabel
                            onChange={(_ev: React.MouseEvent<HTMLElement>, checked?: boolean) =>
                                this.setState({
                                    enableRoles: checked as any,
                                    messages: {
                                        ...this.state.messages,
                                        roles: { messageType: -1, message: "" }
                                    }
                                })}
                            className="config-settings-toggle-btn"
                        />
                        {this.state.showAssignRolesLink &&
                            <a
                                href={this.state.showLoader ? "/" : `https://${this.props.tenantName}/sites/${this.props.siteName}/lists/${siteConfig.userRolesList.split('-').join("")}`}
                                target='_blank'
                                rel="noreferrer"
                                className='assign-roles-link'
                            >
                                <Button
                                    text
                                    icon={<Icon iconName='Go' />}
                                    content={this.props.localeStrings.assignRolesBtnLabel}
                                    title={this.props.localeStrings.assignRolesBtnLabel}
                                    iconPosition="after"
                                    className="assign-roles-btn"
                                    tabIndex={-1}
                                />
                            </a>
                        }
                    </div>
                    <div className={`config-settings-toggle-btn-wrapper map-viewer-setting${this.state.azureMapsKeyError ? " field-with-error" : ""}`}>
                        <Toggle
                            checked={this.state.enableMapViewer}
                            label={
                                <div className="toggle-btn-label">
                                    {this.props.localeStrings.enableMapViewerLabel}
                                    <span className='toggle-info-icon'>
                                        <TooltipHost
                                            content={<span dangerouslySetInnerHTML={{ __html: this.props.localeStrings.enableMapViewerTooltipContent }} />}
                                            calloutProps={{ gapSpace: 0 }}
                                            id="map-viewer-tooltip"
                                        >
                                            <Icon iconName='info' tabIndex={0} aria-label={this.props.localeStrings.enableMapViewerTooltipContent}
                                                aria-describedby="map-viewer-tooltip" role="img" />
                                        </TooltipHost>
                                    </span>
                                </div>
                            }
                            inlineLabel
                            onChange={(_ev: React.MouseEvent<HTMLElement>, checked?: boolean) =>
                                this.setState({
                                    enableMapViewer: checked as any,
                                    messages: {
                                        ...this.state.messages,
                                        mapViewer: { messageType: -1, message: "" }
                                    }
                                })}
                            className="config-settings-toggle-btn"
                        />
                        {this.state.enableMapViewer &&
                            <div className='api-key-input-wrapper'>
                                <Input
                                    type="password"
                                    placeholder={this.props.localeStrings.mapViewerPlaceholder}
                                    className="api-key-input-box"
                                    value={this.state.azureMapsKey}
                                    onChange={(_ev, data: any) => {
                                        this.setState((prevState) => ({
                                            azureMapsKey: data.value,
                                            azureMapsKeyError: data?.value?.trim() === "",
                                            messages: {
                                                ...prevState.messages,
                                                mapViewer: { messageType: -1, message: "" }
                                            }
                                        }))
                                    }}
                                />
                                {this.state.azureMapsKeyError &&
                                    <span className='api-key-error-msg' aria-live="polite" role="alert">
                                        {this.props.localeStrings.mapViewerKeyRequiredMessage}</span>
                                }
                            </div>
                        }
                    </div>

                    <div className={`config-settings-app-edit-role-wrapper`}>
                        <div className="role-title-label">
                            <Label>
                                {this.props.localeStrings.editAccessRoleLabel}
                                <span className='info-icon'>
                                    <TooltipHost
                                        content={<span dangerouslySetInnerHTML={{ __html: this.props.localeStrings.editAccessRoleInfoIconText }} />}
                                        calloutProps={{ gapSpace: 0 }}
                                        id="role-title-tooltip"
                                    >
                                        <Icon iconName='info' tabIndex={0} aria-label="Info"
                                            aria-describedby="role-title-tooltip"
                                            role="button"
                                        />
                                    </TooltipHost>
                                </span>
                            </Label>
                        </div>
                        <div className='app-role-input-wrapper'>
                            <FormDropdown
                                aria-label={this.props.localeStrings.fieldAdditionalRoles + constants.requiredAriaLabel}
                                placeholder={this.props.localeStrings.phRoles}
                                items={this.state.roleDropdownOptions ? this.state.roleDropdownOptions : []}
                                fluid={true}
                                autoSize
                                onChange={this.onRoleChange}
                                value={this.state.selectedRole}
                                id="addRole-dropdown"
                                aria-labelledby="addRole-dropdown"
                                className="select-role-with-edit-dropdown"
                            />
                        </div>
                    </div>
                </div>
                {this.state.showLoader &&
                    <Spinner size='large' className='config-settings-spinner'
                        label={this.props.localeStrings.savingConfigSettingsLabel} labelPosition='below' />
                }
                <div className="admin-settings-btn-wrapper">
                    <Button
                        onClick={() => this.props.onBackClick("")}
                        className='admin-settings-back-btn'
                        title={this.props.localeStrings.btnBack}
                        content={this.props.localeStrings.btnBack}
                        disabled={this.state.showLoader}
                    />
                    <Button
                        primary
                        onClick={this.updateSettings}
                        className='admin-settings-save-btn'
                        title={this.props.localeStrings.saveIcon}
                        content={this.props.localeStrings.saveIcon}
                        disabled={this.state.showLoader}
                    />
                </div>
            </div>
        )
    }
}
