import { ChevronStartIcon, Loader } from '@fluentui/react-northstar';
import { IconButton } from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';
import { MessageBar } from '@fluentui/react/lib/MessageBar';
import { TooltipHost } from '@fluentui/react/lib/Tooltip';
import { ApplicationInsights } from '@microsoft/applicationinsights-web';
import { Client } from "@microsoft/microsoft-graph-client";
import React from 'react';
import Col from 'react-bootstrap/esm/Col';
import Row from 'react-bootstrap/esm/Row';
import { IListItem } from '../common/CommonService';
import "../scss/ActiveBridge.module.scss";
import Bridge from './Bridge';
import Members from './Members';
import Tasks from './Tasks';

export interface ActiveBridgeProps {
    onBackClick(showMessageBar: string): void;
    localeStrings: any;
    incidentData: IListItem;
    graph: Client;
    siteId: string;
    appInsights: ApplicationInsights;
    userPrincipalName: any;
    onShowIncidentHistory: Function;
    currentUserId: string;
    updateIncidentData: Function;
    onEditButtonClick: Function;
    isOwner: boolean;
    graphContextURL: string;
    tenantID: any;
    fromActiveDashboardTab: boolean;
}

export interface ActiveBridgeState {
    messageType: number;
    messageText: string;
    showBridgeLoader: boolean;
    showTasksLoader: boolean;
}

export default class ActiveBridge extends React.Component<ActiveBridgeProps, ActiveBridgeState> {

    constructor(props: ActiveBridgeProps) {
        super(props);

        //states
        this.state = {
            messageType: -1,
            messageText: "",
            showBridgeLoader: false,
            showTasksLoader: false
        }

        //Bind Methods
        this.updateMessagebar = this.updateMessagebar.bind(this);
    }

    //Update messagebar states
    private updateMessagebar = (messageType: number, message: string,
        showBridgeLoader = false, showTasksLoader = false) => {
        this.setState({
            messageType: messageType,
            messageText: message,
            showBridgeLoader: showBridgeLoader,
            showTasksLoader: showTasksLoader
        });
    }
    render() {
        return (
            <div
                className={`active-bridge-wrapper${(this.state.showBridgeLoader || this.state.showTasksLoader) ? " disable-active-bridge" : ""}`}>
                {!this.props.fromActiveDashboardTab &&
                    <div className=".col-xs-12 .col-sm-8 .col-md-4 container" id="active-bridge-path">
                        <label>
                            <span onClick={() => this.props.onBackClick("")} className="go-back">
                                <ChevronStartIcon id="path-back-icon" />
                                <span className="back-label" title={this.props.localeStrings.back}>{this.props.localeStrings.back}</span>
                            </span> &nbsp;&nbsp;
                            <span className="right-border">|</span>
                            <span title={this.props.localeStrings.activeDashboard}>&nbsp;&nbsp;{this.props.localeStrings.activeDashboard}</span>
                        </label>
                    </div>
                }
                <div className="active-bridge-area">
                    <div className="container">
                        <div className='active-bridge-heading'>
                            {this.props.localeStrings.activeDashboard} - {this.props.incidentData.incidentId}
                        </div>

                        <Row xl={2} lg={2} md={1}>
                            <Col xl={4} lg={4} md={12} className="members-tab-wrapper">
                                <div className='members-tab'>
                                    <div className="members-tab-heading">{this.props.localeStrings.teamLabel}</div>
                                    <Members
                                        incidentData={this.props.incidentData}
                                        graph={this.props.graph}
                                        appInsights={this.props.appInsights}
                                        userPrincipalName={this.props.userPrincipalName}
                                        localeStrings={this.props.localeStrings}
                                        isOwner={this.props.isOwner}
                                        onEditButtonClick={this.props.onEditButtonClick}
                                    />
                                </div>
                            </Col>
                            <Col xl={8} lg={8} md={12} className="bridge-tasks-wrapper">
                                <div>
                                    <div className='bridge-tab'>
                                        <div className="bridge-tab-heading">{this.props.localeStrings.bridgeLabel}</div>
                                        {this.state.showBridgeLoader && (
                                            <Loader
                                                label={this.props.localeStrings.processingLabel}
                                                size="smallest"
                                                labelPosition="start"
                                                className="bridge-spinner"
                                            />
                                        )}
                                        {this.state.messageType !== -1 &&
                                            <MessageBar
                                                messageBarType={this.state.messageType}
                                                title={this.state.messageText}
                                                className="bridge-message-bar"
                                                actions={
                                                    <IconButton
                                                        iconProps={{ iconName: "Cancel" }}
                                                        title={this.props.localeStrings.cancelIcon}
                                                        ariaLabel={this.props.localeStrings.cancelIcon}
                                                        onClick={() =>
                                                            this.setState({ messageText: "", messageType: -1 })}
                                                    />
                                                }
                                                isMultiline={false}
                                                role="status"
                                            >
                                                {this.state.messageText}
                                            </MessageBar>
                                        }
                                        <Bridge
                                            currentUserId={this.props.currentUserId}
                                            onShowIncidentHistory={this.props.onShowIncidentHistory}
                                            incidentData={this.props.incidentData}
                                            graph={this.props.graph}
                                            siteId={this.props.siteId}
                                            appInsights={this.props.appInsights}
                                            userPrincipalName={this.props.userPrincipalName}
                                            localeStrings={this.props.localeStrings}
                                            updateIncidentData={this.props.updateIncidentData}
                                            onEditButtonClick={this.props.onEditButtonClick}
                                            isOwner={this.props.isOwner}
                                            updateMessagebar={this.updateMessagebar}
                                        />
                                    </div>
                                    <div className='tasks-tab'>
                                        <div className="tasks-tab-heading">
                                            {this.props.localeStrings.tasksLabel}
                                            <span className="tasks-info-icon">
                                                <TooltipHost
                                                    content={this.props.localeStrings.tasksSectionInfoText}
                                                    calloutProps={{ gapSpace: 0 }}
                                                >
                                                    <Icon iconName="Info" aria-label={this.props.localeStrings.tasksSectionInfoText} />
                                                </TooltipHost>
                                            </span>
                                        </div>
                                        {this.state.showTasksLoader && (
                                            <Loader
                                                label={this.props.localeStrings.createPlanloaderMessage + " " + this.props.localeStrings.incidentCreationLoaderMessage}
                                                size="smallest"
                                                className="tasks-spinner"
                                            />
                                        )}
                                        <Tasks
                                            incidentData={this.props.incidentData}
                                            graph={this.props.graph}
                                            siteId={this.props.siteId}
                                            appInsights={this.props.appInsights}
                                            userPrincipalName={this.props.userPrincipalName}
                                            updateMessagebar={this.updateMessagebar}
                                            showTasksLoader={this.state.showTasksLoader}
                                            localeStrings={this.props.localeStrings}
                                            graphContextURL={this.props.graphContextURL}
                                            tenantID={this.props.tenantID}
                                        />
                                    </div>
                                </div>
                            </Col>
                        </Row>
                    </div >
                </div >
            </div >
        );
    }
}
