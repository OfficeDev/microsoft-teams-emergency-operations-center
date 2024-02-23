import { MessageBar } from "@fluentui/react";
import { Input, Spinner } from "@fluentui/react-components";
import { Button } from "@fluentui/react-northstar";
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
    bingMapsKeyConfigData: any;
    appTitle: string;
    appTitleData: any;
};
export interface IConfigSettingsState {
    enableRoles: boolean;
    messages: IMessages;
    showAssignRolesLink: boolean;
    showLoader: boolean;
    enableMapViewer: boolean;
    bingMapsKey: string;
    bingMapsKeyError: boolean;
    appTitle: string;
    appTitleKeyError: boolean;
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
            bingMapsKey: this.props.bingMapsKeyConfigData?.value?.trim()?.length > 0 ? this.props.bingMapsKeyConfigData?.value : "",
            bingMapsKeyError: false,
            appTitle: this.props.appTitle,
            appTitleKeyError: false
        }

        //bind methods
        this.updateSettings = this.updateSettings.bind(this);
    }

    //Create object for Common Services class
    private commonService = new CommonService();

    //Update Config Settings
    private updateSettings = async () => {
        try {
            let noChanges: boolean = true;
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

                    if (this.props.appTitleData.itemId === undefined) {
                        // create graph endpoint for TEOC Config list to add app title
                        const graphConfigListEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}/lists/${siteConfig.configurationList}/items`;
                        const newItemAppTitleObj = { fields: { Value: this.state.appTitle.trim(), Title: constants.appTitleKey } };
                        const item = await this.commonService.addItemInList
                            <{ fields: { id: number; Title: string; Value: string } }>(
                                graphConfigListEndpoint,
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

            //Add Bing Map API Key record in TEOC-config list
            if (this.state.enableMapViewer && this.props.bingMapsKeyConfigData?.title === undefined &&
                this.state.bingMapsKey?.trim() !== "") {
                noChanges = false;
                this.setState({ showLoader: true });

                // create graph endpoint for TEOC Config list
                const configGraphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}${graphConfig.listsGraphEndpoint}/${siteConfig.configurationList}/items`;
                // create Bing Map API Key list item object
                const listItem = { fields: { Title: constants.bingMapsKey, Value: this.state.bingMapsKey } };
                const configResponse = await this.commonService.sendGraphPostRequest(configGraphEndpoint, this.props.graph, listItem);
                this.props.setState({
                    isMapViewerEnabled: this.state.enableMapViewer,
                    bingMapsKeyConfigData: {
                        ...this.props.bingMapsKeyConfigData,
                        title: configResponse.fields.Title,
                        value: this.state.bingMapsKey,
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
                    bingMapsKey: prevState.bingMapsKey?.trim()
                }));
            }
            //Update Bing Map API Key in TEOC-config list
            else if ((this.props.bingMapsKeyConfigData?.value &&
                this.props.bingMapsKeyConfigData?.value?.trim() !== this.state.bingMapsKey?.trim()) ||
                this.props.isMapViewerEnabled !== this.state.enableMapViewer) {
                noChanges = false;
                //Validate Bing Map API Key
                if (this.state.enableMapViewer &&
                    (this.state.bingMapsKey?.trim() === "" || this.state.bingMapsKey?.trim() === undefined)) {
                    this.setState(prevState => ({
                        bingMapsKeyError: true,
                        messages: {
                            ...prevState.messages,
                            mapViewer: { messageType: 1, message: this.props.localeStrings.mapViewerKeyRequiredMessage }
                        }
                    }));
                }
                //Update Bing Map API Key
                else {
                    this.setState({ showLoader: true });
                    //Endpoint to update Bing Map API Key in TEOC-config list
                    const graphConfigListEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}/lists/${siteConfig.configurationList}/items/${this.props.bingMapsKeyConfigData?.itemId}/fields`;
                    const updatedValue = this.state.enableMapViewer ? this.state.bingMapsKey?.trim() : "";
                    let updatedBingAPIKeyObj = { Value: updatedValue }
                    //Update Bing Map API Key in TEOC-config list API Call
                    await this.commonService.updateItemInList(graphConfigListEndpoint, this.props.graph, updatedBingAPIKeyObj);

                    let messageToDisplay: string;
                    if (this.state.enableMapViewer && this.props.bingMapsKeyConfigData?.value?.trim() !== "" &&
                        this.state.bingMapsKey?.trim() !== "") {
                        messageToDisplay = this.props.localeStrings.mapViewerKeyUpdatedMessage;
                    }
                    else {
                        messageToDisplay = this.state.enableMapViewer ? this.props.localeStrings.mapViewerKeyEnabledMessage : this.props.localeStrings.mapViewerKeyDisabledMessage;
                    }

                    //Update Home Component states
                    this.props.setState({
                        isMapViewerEnabled: this.state.enableMapViewer,
                        bingMapsKeyConfigData: { ...this.props.bingMapsKeyConfigData, value: updatedValue }
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
                        bingMapsKey: updatedValue
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
                    <div className={`config-settings-toggle-btn-wrapper map-viewer-setting${this.state.bingMapsKeyError ? " field-with-error" : ""}`}>
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
                                    placeholder={this.props.localeStrings.mapViewerPlaceholder}
                                    className="api-key-input-box"
                                    value={this.state.bingMapsKey}
                                    onChange={(_ev, data: any) => {
                                        this.setState((prevState) => ({
                                            bingMapsKey: data.value,
                                            bingMapsKeyError: data?.value?.trim() === "",
                                            messages: {
                                                ...prevState.messages,
                                                mapViewer: { messageType: -1, message: "" }
                                            }
                                        }))
                                    }}
                                />
                                {this.state.bingMapsKeyError &&
                                    <span className='api-key-error-msg' aria-live="polite" role="alert">
                                        {this.props.localeStrings.mapViewerKeyRequiredMessage}</span>
                                }
                            </div>
                        }
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
