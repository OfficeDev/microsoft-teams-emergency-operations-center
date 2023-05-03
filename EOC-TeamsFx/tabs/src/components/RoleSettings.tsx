import { MessageBar } from '@fluentui/react';
import { Button } from '@fluentui/react-northstar';
import { Icon } from '@fluentui/react/lib/Icon';
import { Toggle } from '@fluentui/react/lib/Toggle';
import { TooltipHost } from '@fluentui/react/lib/Tooltip';
import { ApplicationInsights } from '@microsoft/applicationinsights-web';
import { Client } from "@microsoft/microsoft-graph-client";
import React from 'react';
import Col from 'react-bootstrap/esm/Col';
import CommonService from '../common/CommonService';
import * as constants from "../common/Constants";
import * as graphConfig from '../common/graphConfig';
import siteConfig from '../config/siteConfig.json';

export interface IRoleSettingsProps {
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
};
export interface IRoleSettingsState {
    enableRoles: boolean;
    messageType: number;
    message: string;
    showAssignRolesLink: boolean;
    showLoader: boolean;
}

export default class RoleSettings extends React.Component<IRoleSettingsProps, IRoleSettingsState> {
    constructor(props: IRoleSettingsProps) {
        super(props);

        //States
        this.state = {
            enableRoles: this.props.isRolesEnabled,
            messageType: -1,
            message: "",
            showAssignRolesLink: this.props.isRolesEnabled,
            showLoader: false
        }

        //bind methods
        this.updateSetting = this.updateSetting.bind(this);
    }

    //Create object for Common Services class
    private commonService = new CommonService();

    //Update Roles Setting
    private updateSetting = async () => {
        try {
            if (this.props.isRolesEnabled !== this.state.enableRoles) {
                this.setState({ showLoader: true });

                //Update Role Settings in TEOC-config list
                const graphConfigListEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}/lists/${siteConfig.configurationList}/items/${this.props.configRoleData.itemId}/fields`;
                const updateEnableRolesObj = { Value: this.props.isRolesEnabled ? "False" : "True" };
                await this.commonService.updateItemInList(graphConfigListEndpoint,
                    this.props.graph, updateEnableRolesObj);

                //Disabling Roles
                if (this.props.isRolesEnabled) {
                    //Update Role Setting States
                    this.setState({
                        showAssignRolesLink: false,
                        message: this.props.localeStrings.roleSettingsDisabledMessage
                    });

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
                    this.setState({
                        showAssignRolesLink: true,
                        message: this.props.localeStrings.roleSettingsEnabledMessage
                    });

                    //Update Home Component states
                    this.props.setState({
                        isRolesEnabled: true,
                        configRoleData: { ...this.props.configRoleData, value: "True" }
                    });
                }

                this.setState({
                    messageType: 4,
                    showLoader: false
                });
            }
            else {
                this.setState({
                    messageType: 4,
                    message: this.props.isRolesEnabled ?
                        this.props.localeStrings.roleSettingsEnabledMessage :
                        this.props.localeStrings.roleSettingsDisabledMessage
                });
            }
        }
        catch (error) {
            this.setState({
                messageType: 1,
                message: this.props.localeStrings.roleSettingsErrorMessage,
                showLoader: false
            });

            console.error(
                constants.errorLogPrefix + "_" + constants.componentNames.AdminSettingsComponent + "_" +
                constants.componentNames.RoleSettingsComponent + "_updateSetting \n",
                JSON.stringify(error)
            );

            //log exception to AppInsights
            this.commonService.trackException(this.props.appInsights, error,
                constants.componentNames.RoleSettingsComponent, 'updateSetting', this.props.userPrincipalName);
        }
    }

    //Render Method
    render() {
        return (
            <div className='role-settings-wrapper'>
                {this.state.messageType !== -1 &&
                    <Col xl={6} lg={6} md={8} sm={10} xs={12}>
                        <MessageBar
                            messageBarType={this.state.messageType}
                            title={this.state.message}
                            className="role-settings-message-bar"
                            isMultiline={true}
                        >
                            {this.state.message}
                        </MessageBar>
                    </Col>
                }
                <Col xl={6} lg={6} md={8} sm={10} xs={12}>
                    <div className="role-settings-toggle-btn-wrapper">
                        <Toggle
                            checked={this.state.enableRoles}
                            label={
                                <div className="toggle-btn-label">
                                    {this.props.localeStrings.enablesRoles}
                                    <span className='toggle-info-icon'>
                                        <TooltipHost
                                            content={this.props.localeStrings.roleSettingsInfoIconTooltip}
                                            calloutProps={{ gapSpace: 0 }}
                                        >
                                            <Icon iconName='info' />
                                        </TooltipHost>
                                    </span>
                                </div>
                            }
                            inlineLabel
                            onChange={(_ev: React.MouseEvent<HTMLElement>, checked?: boolean) =>
                                this.setState({ enableRoles: checked ? true : false, messageType: -1 })}
                            className="role-settings-toggle-btn"
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
                                />
                            </a>
                        }
                    </div>
                </Col>

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
                        onClick={this.updateSetting}
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
