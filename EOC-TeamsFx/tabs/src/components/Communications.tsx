import { MessageBar } from '@fluentui/react';
import { Checkbox, FormTextArea, Loader } from '@fluentui/react-northstar';
import { Client } from "@microsoft/microsoft-graph-client";
import React from 'react';
import { IListItem } from "../common/CommonService";


export interface CommunicationsProps {
    graph: Client;
    incidentData: IListItem;
    setState: Function;
    cardMessage: string;
    highImportance: boolean;
    includeLink: boolean;
    validationMessage: string;
    showLoader: boolean;
    messageType: number;
    statusMessage: string;
    localeStrings: any;
}

export interface CommunicationsState { }

export default class Communications extends React.Component<CommunicationsProps, CommunicationsState> {

    render() {
        return (
            <div className='communications'>
                {this.props.messageType !== -1 &&
                    <MessageBar
                        messageBarType={this.props.messageType}
                        title={this.props.statusMessage}
                    >
                        {this.props.statusMessage}
                    </MessageBar>
                }
                <FormTextArea
                    placeholder={this.props.localeStrings.announcementMessagePlaceholder}
                    fluid={true}
                    value={this.props.cardMessage}
                    maxLength={500}
                    onChange={(_ev, data: any) => this.props.setState({ cardMessage: data.value, validationMessage: "" })}
                    title={this.props.localeStrings.announcementMessagePlaceholder}
                    resize="both"
                    className="com-text-area"
                />
                {this.props.validationMessage !== "" &&
                    <span className="textArea-validation-msg">{this.props.validationMessage}</span>
                }

                <div className="com-checkbox-wrapper">
                    <Checkbox
                        label={this.props.localeStrings.importantLabel}
                        onChange={(_ev, { checked }: any) => this.props.setState({ highImportance: checked })}
                        checked={this.props.highImportance}
                        title={this.props.localeStrings.importantCheckboxTooltip}
                    />
                    {(this.props.incidentData.bridgeLink !== undefined && this.props.incidentData.bridgeLink !== "") &&
                        <Checkbox
                            label={this.props.localeStrings.includeBridgeLinkLabel}
                            onChange={(_ev, { checked }: any) => this.props.setState({ includeLink: checked })}
                            checked={this.props.includeLink}
                            title={this.props.localeStrings.includeBridgeLinkLabel}
                            className="include-bridge-link-checkbox"
                        />
                    }
                </div>
                {this.props.showLoader &&
                    <Loader label={this.props.localeStrings.announcementSpinnerLabel} size="small" />
                }
            </div>
        );
    }
}
