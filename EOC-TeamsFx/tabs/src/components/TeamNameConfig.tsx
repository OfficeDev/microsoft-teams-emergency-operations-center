import React from 'react';
import '../scss/TeamNameConfig.module.scss'
import { Loader, ChevronStartIcon, FormDropdown, FormInput, Button } from "@fluentui/react-northstar";
import Col from "react-bootstrap/Col";
import Row from "react-bootstrap/Row";
import * as constants from '../common/Constants';
import 'bootstrap/dist/css/bootstrap.min.css';
import CommonService from "../common/CommonService";
import siteConfig from '../config/siteConfig.json';
import * as graphConfig from '../common/graphConfig';
import { Client } from "@microsoft/microsoft-graph-client";
import { ApplicationInsights } from '@microsoft/applicationinsights-web';
import { localizedStrings } from '../locale/LocaleStrings';

export interface ITeamNameConfigState {
    showLoader: boolean;
    loaderMessage: string;
    configListData: any;
    configValue: any;
    previewString: any;
    prefixIsMissing: boolean;
}

export interface ITeamNameConfigProps {
    localeStrings: any;
    onBackClick(showMessageBar: boolean): void;
    siteId: string;
    graph: Client;
    appInsights: ApplicationInsights;
    userPrincipalName: any;
    showMessageBar(message: string, type: string): void;
    hideMessageBar(): void;
}

const teamNameConfigString = "TeamNameConfiguration";
export class TeamNameConfig extends React.PureComponent<ITeamNameConfigProps, ITeamNameConfigState>  {
    constructor(props: any) {
        super(props);
        this.state = {
            showLoader: true,
            loaderMessage: this.props.localeStrings.loaderMessage,
            configListData: { itemId: 0, Title: '', value: {} },
            configValue: {},
            previewString: '',
            prefixIsMissing: true
        }
    }

    //common service object
    private dataService = new CommonService();

    // component life cycle method
    public async componentDidMount() {
        // Get Team Name Config List Data
        this.getTeamNameConfigData();
    }

    //get data from team name configurations list
    getTeamNameConfigData = async () => {
        try {
            //graph endpoint to get data from team name configuration list
            let graphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}/lists/${siteConfig.teamNameConfigList}/items?$expand=fields&$Top=5000`;
            const configData = await this.dataService.getConfigData(graphEndpoint, this.props.graph, 'TeamNameConfig');
            const sortedData = this.dataService.sortConfigData(configData.value);
            const previewString = this.formatPreviewString(sortedData);
            this.setState({
                configListData: configData,
                configValue: configData.value,
                previewString: previewString,
                showLoader: false,
                prefixIsMissing: configData.value[constants.teamNameConfigConstants.Prefix] !== constants.teamNameConfigConstants.DontInclude && configData.value[constants.teamNameConfigConstants.Prefix] === '' ? true : false
            })
        }
        catch (error: any) {
            console.error(
                constants.errorLogPrefix + `${teamNameConfigString}_GetConfiguration \n`,
                JSON.stringify(error)
            );
            // Log Exception
            this.dataService.trackException(this.props.appInsights, error, constants.componentNames.IncidentDetailsComponent, `${teamNameConfigString}_GetConfiguration`, this.props.userPrincipalName);
        }
    }

    //method to update team name configuration
    updateConfiguration = async () => {
        // validate for required fields
        if (this.state.prefixIsMissing) {
            this.props.showMessageBar(this.props.localeStrings.reqFieldErrorMessage, constants.messageBarType.error);
        }
        else {
            try {
                this.props.hideMessageBar();
                this.setState({
                    showLoader: true,
                    loaderMessage: this.props.localeStrings.incidentCreationLoaderMessage
                })
                let updatedValues = {
                    Value: JSON.stringify(this.state.configValue)
                }
                //graph endpoint to update team name config data
                let graphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}/lists/${siteConfig.teamNameConfigList}/items/${this.state.configListData.itemId}/fields`;
                const configUpdated = await this.dataService.updateTeamNameConfigData(graphEndpoint, this.props.graph, updatedValues);
                if (configUpdated) {
                    console.log(constants.infoLogPrefix + "Team Name Configurations Updated");
                    this.props.showMessageBar(this.props.localeStrings.updatedConfigSuccessMessage, constants.messageBarType.success);
                    this.setState({
                        showLoader: false
                    });
                    //log trace
                    this.dataService.trackTrace(this.props.appInsights, 'Team Name Configurations Updated', '', this.props.userPrincipalName);
                    this.props.onBackClick(true);
                }
                else {
                    console.log(constants.infoLogPrefix + "Team Name Configurations Update Failed");
                    this.setState({
                        showLoader: false
                    });
                    this.props.showMessageBar(this.props.localeStrings.genericErrorMessage + ", " + this.props.localeStrings.updatedConfigErrorMessage, constants.messageBarType.error);

                    //log trace
                    this.dataService.trackTrace(this.props.appInsights, 'Team Name Configurations Update Failed', '', this.props.userPrincipalName);
                }

            }
            catch (error: any) {
                console.error(
                    constants.errorLogPrefix + `${teamNameConfigString}_UpdateConfiguration`,
                    JSON.stringify(error)
                );
                // Log Exception
                this.dataService.trackException(this.props.appInsights, error, constants.componentNames.IncidentDetailsComponent, `${teamNameConfigString}_UpdateConfiguration`, this.props.userPrincipalName);
            }
        }
    }

    //Format Preview String based on order
    updateOrder = (selectedVal: any, configData: any, key: string) => {
        try {
            //Filter PrefixValue key from object
            let filteredArr: any = Object.keys(configData)
                .filter((key) => key.includes(constants.teamNameConfigConstants.IncidentName) || !key.includes(constants.teamNameConfigConstants.PrefixValue) || key.includes(constants.teamNameConfigConstants.IncidentType) || key.includes(constants.teamNameConfigConstants.StartDate))
                .reduce((obj, key) => {
                    return Object.assign(obj, {
                        [key]: configData[key]
                    });
                }, {});

            let minVal = parseInt(configData[key]) > parseInt(selectedVal.value) ? parseInt(selectedVal.value) : parseInt(configData[key]);
            let maxVal = parseInt(configData[key]) > parseInt(selectedVal.value) ? parseInt(configData[key]) : parseInt(selectedVal.value);
            let isDuplicateValuePresent = Object.values(filteredArr).includes(parseInt(selectedVal.value));
            filteredArr[key] = selectedVal.value === constants.teamNameConfigConstants.DontInclude ? selectedVal.value : parseInt(selectedVal.value);

            let selectedCount = 0;
            Object.values(filteredArr).forEach((val: any) => {
                if (val !== constants.teamNameConfigConstants.DontInclude) {
                    selectedCount++;
                }
            });

            let ifOneIsPresent = Object.values(filteredArr).includes(1);
            let ifThreeIsPresent = Object.values(filteredArr).includes(3);
            let ifFourIsPresent = Object.values(filteredArr).includes(4);

            Object.keys(filteredArr)
                .forEach((k: any) => {
                    if (k !== "PrefixValue" && k !== key) {
                        if (selectedVal.value < parseInt(configData[key])) {
                            if (parseInt(filteredArr[k]) >= minVal && parseInt(filteredArr[k]) < maxVal) {
                                console.log(k)
                                filteredArr[k] = parseInt(filteredArr[k]) + 1;
                            }
                        }
                        else if (selectedVal.value > parseInt(configData[key])) {
                            if (parseInt(filteredArr[k]) > minVal && parseInt(filteredArr[k]) <= maxVal) {
                                console.log(k)
                                filteredArr[k] = parseInt(filteredArr[k]) - 1;
                            }
                        }
                        else if (selectedVal.value === constants.teamNameConfigConstants.DontInclude) {
                            if (filteredArr[k] !== constants.teamNameConfigConstants.DontInclude) {
                                if (parseInt(configData[key]) < parseInt(filteredArr[k])) {
                                    console.log(k)
                                    filteredArr[k] = parseInt(filteredArr[k]) - 1;
                                }
                            }
                        }
                        else if (configData[key] === constants.teamNameConfigConstants.DontInclude) {
                            if (isDuplicateValuePresent) {
                                if (filteredArr[k] !== constants.teamNameConfigConstants.DontInclude) {
                                    if (selectedCount === 2) {
                                        if (parseInt(selectedVal.value) === 1)
                                            filteredArr[k] = parseInt(filteredArr[k]) + 1;
                                        else
                                            filteredArr[k] = parseInt(filteredArr[k]) - 1;
                                    }
                                    else if (selectedCount === 3) {
                                        if (parseInt(selectedVal.value) === 1) {
                                            if (parseInt(filteredArr[k]) >= 1 && parseInt(filteredArr[k]) <= 2)
                                                filteredArr[k] = parseInt(filteredArr[k]) + 1;
                                        }
                                        if (parseInt(selectedVal.value) === 2) {
                                            if (parseInt(filteredArr[k]) >= 2 && parseInt(filteredArr[k]) <= 3)
                                                filteredArr[k] = parseInt(filteredArr[k]) + 1;
                                        }
                                        if (parseInt(selectedVal.value) === 3) {
                                            if (parseInt(filteredArr[k]) >= 2 && parseInt(filteredArr[k]) <= 3)
                                                filteredArr[k] = parseInt(filteredArr[k]) - 1;
                                        }
                                        if (parseInt(selectedVal.value) === 4) {
                                            if (parseInt(filteredArr[k]) >= 3 && parseInt(filteredArr[k]) <= 4)
                                                filteredArr[k] = parseInt(filteredArr[k]) - 1;
                                        }
                                    }
                                    else if (selectedCount === 4) {
                                        if (parseInt(selectedVal.value) === 1) {
                                            if (ifFourIsPresent) {
                                                if (parseInt(filteredArr[k]) >= 1 && parseInt(filteredArr[k]) <= 2)
                                                    filteredArr[k] = parseInt(filteredArr[k]) + 1;
                                            }
                                            else {
                                                if (parseInt(filteredArr[k]) >= 1 && parseInt(filteredArr[k]) <= 3)
                                                    filteredArr[k] = parseInt(filteredArr[k]) + 1;
                                            }
                                        }
                                        if (parseInt(selectedVal.value) === 2) {
                                            if (ifOneIsPresent) {
                                                if (parseInt(filteredArr[k]) >= 2 && parseInt(filteredArr[k]) <= 3)
                                                    filteredArr[k] = parseInt(filteredArr[k]) + 1;
                                            } else {
                                                if (parseInt(filteredArr[k]) === 2)
                                                    filteredArr[k] = parseInt(filteredArr[k]) - 1;
                                            }
                                        }
                                        if (parseInt(selectedVal.value) === 3) {
                                            if (ifFourIsPresent) {
                                                if (parseInt(filteredArr[k]) >= 2 && parseInt(filteredArr[k]) <= 3)
                                                    filteredArr[k] = parseInt(filteredArr[k]) - 1;
                                            } else {
                                                if (parseInt(filteredArr[k]) === 3)
                                                    filteredArr[k] = parseInt(filteredArr[k]) + 1;
                                            }
                                        }
                                        if (parseInt(selectedVal.value) === 4) {
                                            if (ifThreeIsPresent) {
                                                if (parseInt(filteredArr[k]) >= 2 && parseInt(filteredArr[k]) <= 4)
                                                    filteredArr[k] = parseInt(filteredArr[k]) - 1;
                                            } else {
                                                if (parseInt(filteredArr[k]) === 4)
                                                    filteredArr[k] = parseInt(filteredArr[k]) - 1;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                });

            this.setState({
                ...this.state, configValue: {
                    ...this.state.configValue,
                    Prefix: filteredArr[constants.teamNameConfigConstants.Prefix],
                    IncidentName: filteredArr[constants.teamNameConfigConstants.IncidentName],
                    IncidentType: filteredArr[constants.teamNameConfigConstants.IncidentType],
                    StartDate: filteredArr[constants.teamNameConfigConstants.StartDate]
                },
                prefixIsMissing: filteredArr[constants.teamNameConfigConstants.Prefix] !== constants.teamNameConfigConstants.DontInclude && this.state.configValue.PrefixValue === '' ? true : false
            });

            let formattedData = this.dataService.sortConfigData(filteredArr);

            let formattedString = this.formatPreviewString(formattedData);

            this.setState({
                previewString: formattedString
            });
        } catch (error) {
            console.error(
                constants.errorLogPrefix + `${teamNameConfigString}_UpdateOrder`,
                JSON.stringify(error)
            );
            // Log Exception
            this.dataService.trackException(this.props.appInsights, error, constants.componentNames.TeamNameConfiguration, `${teamNameConfigString}_UpdateOrder`, this.props.userPrincipalName);
        }
    }

    // method to form preview string
    formatPreviewString = (configData: any) => {
        try {
            //format preview string
            let string = "Incident ID"
            Object.keys(configData).forEach(key => {
                if (configData[key] !== constants.teamNameConfigConstants.DontInclude) {
                    if (key === constants.teamNameConfigConstants.IncidentName)
                        string = string + " - Incident Name";
                    if (key === constants.teamNameConfigConstants.IncidentType)
                        string = string + " - Incident Type";
                    if (key === constants.teamNameConfigConstants.StartDate)
                        string = string + " - Start Date";
                    if (key === constants.teamNameConfigConstants.Prefix)
                        string = string + " - Prefix";
                }
            })
            return string;
        } catch (error) {
            console.error(
                constants.errorLogPrefix + `${teamNameConfigString}_FormatPreviewString`,
                JSON.stringify(error)
            );
            // Log Exception
            this.dataService.trackException(this.props.appInsights, error, constants.componentNames.TeamNameConfiguration, `${teamNameConfigString}_FormatPreviewString`, this.props.userPrincipalName);

        }
    }

    public render() {
        return (<>
            <div className="team-name-config">
                <>
                    {this.state.showLoader &&
                        <div className="loader-bg">
                            <div className="loaderStyle">
                                <Loader label={this.state.loaderMessage} size="largest" />
                            </div>
                        </div>
                    }
                    <div>
                        <div className=".col-xs-12 .col-sm-8 .col-md-4 container" id="team-name-config-path">
                            <label>
                                <span onClick={() => this.props.onBackClick(false)} className="go-back">
                                    <ChevronStartIcon id="path-back-icon" />
                                    <span className="back-label" title="Back">Back</span>
                                </span> &nbsp;&nbsp;
                                <span className="right-border">|</span>
                                <span>&nbsp;&nbsp;{this.props.localeStrings.labelTeamNameConfig}</span>
                            </label>
                        </div>
                        <div className='team-name-config-form-area'>
                            <div className="container">
                                <div className="team-name-config-page-heading">
                                    {this.props.localeStrings.formTitleTeamNameConfig}
                                </div>
                                <div>
                                    <Row xl={2} lg={2} md={2} sm={2} xs={2}>
                                        <Col xl={5} lg={6} md={7} sm={12} xs={12}>
                                            <Row xl={2} lg={2} md={2} sm={2} xs={2}>
                                                <Col xl={5} lg={5} md={5} sm={5} xs={5}>
                                                    <FormDropdown
                                                        label={this.props.localeStrings.orderLabel}
                                                        items={constants.teamNameConfigOrderDropdown}
                                                        className="team-name-config-order-dropdown"
                                                        value={this.state.configValue.Prefix}
                                                        onChange={(evt, val) => this.updateOrder(val, this.state.configValue, constants.teamNameConfigConstants.Prefix)}

                                                    />
                                                </Col>
                                                <Col xl={7} lg={7} md={5} sm={7} xs={7}>
                                                    <FormInput
                                                        type="text"
                                                        placeholder={this.props.localeStrings.prefixValuePlaceholder}
                                                        fluid={true}
                                                        className="team-name-config-input-field"
                                                        successIndicator={false}
                                                        label={this.props.localeStrings.prefixLabel}
                                                        value={this.state.configValue.PrefixValue}
                                                        maxLength={constants.maxCharLengthForPrefix}
                                                        onChange={(event: any) => this.setState({
                                                            ...this.state, configValue: {
                                                                ...this.state.configValue,
                                                                PrefixValue: event.target.value
                                                            },
                                                            prefixIsMissing: this.state.configValue.Prefix !== constants.teamNameConfigConstants.DontInclude && event.target.value === '' ? true : false
                                                        })}
                                                    />
                                                    {this.state.prefixIsMissing ? <label className="message-label" >{this.props.localeStrings.prefixValueRequired}</label> : <></>}
                                                </Col>
                                            </Row>
                                            <Row xl={2} lg={2} md={2} sm={2} xs={2}>
                                                <Col xl={5} lg={5} md={5} sm={5} xs={5}>
                                                    <FormDropdown
                                                        items={constants.teamNameConfigOrderDropdown}
                                                        className="team-name-config-order-dropdown"
                                                        value={this.state.configValue.IncidentName}
                                                        onChange={(evt, val) => this.updateOrder(val, this.state.configValue, constants.teamNameConfigConstants.IncidentName)}
                                                    />
                                                </Col>
                                                <Col xl={7} lg={7} md={5} sm={7} xs={7}>
                                                    <label className="team-name-config-dropdown-label">{this.props.localeStrings.incidentNameLabel}</label>
                                                </Col>
                                            </Row>
                                            <Row xl={2} lg={2} md={2} sm={2} xs={2}>
                                                <Col xl={5} lg={5} md={5} sm={5} xs={5}>
                                                    <FormDropdown
                                                        items={constants.teamNameConfigOrderDropdown}
                                                        className="team-name-config-order-dropdown"
                                                        value={this.state.configValue.IncidentType}
                                                        onChange={(evt, val) => this.updateOrder(val, this.state.configValue, constants.teamNameConfigConstants.IncidentType)}
                                                    />
                                                </Col>
                                                <Col xl={7} lg={7} md={5} sm={7} xs={7}>
                                                    <label className="team-name-config-dropdown-label">{this.props.localeStrings.incidentTypeLabel}</label>
                                                </Col>
                                            </Row>
                                            <Row xl={2} lg={2} md={2} sm={2} xs={2}>
                                                <Col xl={5} lg={5} md={5} sm={5} xs={5}>
                                                    <FormDropdown
                                                        items={constants.teamNameConfigOrderDropdown}
                                                        className="team-name-config-order-dropdown"
                                                        value={this.state.configValue.StartDate}
                                                        onChange={(evt, val) => this.updateOrder(val, this.state.configValue, constants.teamNameConfigConstants.StartDate)}
                                                    />
                                                </Col>
                                                <Col xl={7} lg={7} md={5} sm={7} xs={7}>
                                                    <label className="team-name-config-dropdown-label">{this.props.localeStrings.startDate}</label>
                                                </Col>
                                            </Row>
                                        </Col>
                                        <Col xl={7} lg={6} md={5} sm={12} xs={12}>
                                            <div className='team-name-config-preview-heading'>{this.props.localeStrings.previewLabel}</div>
                                            <Row xl={2} lg={2} md={2} sm={2} xs={2}>
                                                <Col xl={2} lg={3} md={12} sm={2} xs={12}>
                                                    <img src={require("../assets/Images/PreviewIcon.svg").default}
                                                        alt="Preview"
                                                        className="team-name-config-preview-img"
                                                        title="Preview"
                                                    />
                                                </Col>
                                                <Col xl={10} lg={9} md={12} sm={10} xs={12} className="team-name-config-order-preview-area">
                                                    <div
                                                        className="team-name-config-order-preview"
                                                        title={this.state.previewString}
                                                    >
                                                        {this.state.previewString}
                                                    </div>
                                                </Col>
                                            </Row>
                                        </Col>
                                    </Row>
                                </div>
                                <div className="team-name-config-btn-area">
                                    <Button
                                        onClick={() => this.props.onBackClick(false)}
                                        className='team-name-config-back-btn'
                                        title={this.props.localeStrings.btnBack}
                                    >
                                        <label className="team-name-config-back-btn-label">{this.props.localeStrings.btnBack}</label>
                                    </Button>
                                    <Button
                                        primary
                                        onClick={() => this.updateConfiguration()}
                                        className='team-name-config-save-btn'
                                        title={this.props.localeStrings.btnSaveChanges}
                                    >
                                        <label className="team-name-config-save-btn-label">{this.props.localeStrings.btnSaveChanges}</label>
                                    </Button>
                                </div>
                            </div>
                        </div>
                    </div>
                </>
            </div>
        </>)
    }
}