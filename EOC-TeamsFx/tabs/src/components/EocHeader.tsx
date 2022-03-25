import { Callout, Link, Text } from '@fluentui/react';
import { Flex, FlexItem, Tooltip } from '@fluentui/react-northstar';
import { Component } from 'react';
import * as constants from "../common/Constants";

interface IHeaderProps {
    clickcallback: () => void; //will redirects to home
    context?: any;
    localeStrings: any;
    currentUserName: string;
}

interface HeaderState {
    isCalloutVisible: boolean;
    isDesktop: boolean;
}

export default class EocHeader extends Component<IHeaderProps, HeaderState> {
    constructor(props: any) {
        super(props);
        this.state = {
            isCalloutVisible: false,
            isDesktop: true
        };
        this.homeRedirect = this.homeRedirect.bind(this);
    }
    public async componentDidMount() {

        //Event listener for screen resizing
        window.addEventListener("resize", this.resize.bind(this));
        this.resize();
    }

    //Function for Screen Resizing
    resize = () => this.setState({ isDesktop: window.innerWidth > constants.mobileWidth })

    componentWillUnmount() {
        //Event listener for screen resizing
        window.removeEventListener("resize", this.resize.bind(this));
    }

    // redirect to home page
    public homeRedirect() {
        this.props.clickcallback();
    }

    // toggle callout visibility
    public toggleIsCalloutVisible = () => {
        this.setState({ isCalloutVisible: !this.state.isCalloutVisible });
    }

    render() {
        const buttonId = 'callout-button';
        const labelId = 'callout-label';
        const descriptionId = 'callout-description';
        return (
            <>
                <div className='eoc-header'>
                    <Flex gap="gap.small" space='between'>
                        <Flex gap="gap.small" vAlign="center">
                            <img
                                src={require("../assets/Images/AppLogo.svg").default}
                                alt="Ms Logo"
                                className="ms-logo"
                                title={this.props.localeStrings.appTitle}
                            />
                            <span className="header-text" title={this.props.localeStrings.appTitle}>{this.props.localeStrings.appTitle} <span className="header-text-preview" title={this.props.localeStrings.appTitlePreview}>{this.props.localeStrings.appTitlePreview}</span> </span>
                        </Flex>
                        <Flex gap={this.state.isDesktop ? "gap.large" : "gap.medium"} vAlign="center">
                            <FlexItem>
                                <Tooltip
                                    trigger={<img
                                        src={require("../assets/Images/InfoIcon.svg").default}
                                        alt="Info"
                                        id={buttonId}
                                        className="header-icon"
                                        onClick={this.toggleIsCalloutVisible}
                                    />}
                                    content={this.props.localeStrings.moreInfo}
                                    pointing={false}
                                />
                            </FlexItem>
                            <FlexItem>
                                <a
                                    href={constants.helpUrl}
                                    target="_blank" rel="noreferrer"
                                >
                                    <Tooltip
                                        trigger={<img
                                            src={require("../assets/Images/HelpIcon.svg").default}
                                            alt="Help"
                                            className="header-icon"
                                        />}
                                        content={this.props.localeStrings.support}
                                        pointing={false}
                                    />
                                </a>
                            </FlexItem>
                            <FlexItem>
                                <a href={constants.feedbackUrl} target="_blank" rel="noreferrer">
                                    <Tooltip
                                        trigger={<img
                                            src={require("../assets/Images/FeedbackIcon.svg").default}
                                            alt="Feedback"
                                            className='feedback-icon'
                                        />}
                                        content={{ content: this.props.localeStrings.feedback }}
                                        pointing={false}
                                    />
                                </a>
                            </FlexItem>
                        </Flex>
                    </Flex>
                </div>
                <div className="sub-header">
                    <div className='container' id="sub-heading">{this.props.localeStrings.welcome} {this.props.currentUserName}!</div>
                </div>
                {this.state.isCalloutVisible && (
                    <Callout
                        className="info-callout"
                        ariaLabelledBy={labelId}
                        ariaDescribedBy={descriptionId}
                        gapSpace={0}
                        target={`#${buttonId}`}
                        onDismiss={this.toggleIsCalloutVisible}
                        setInitialFocus
                    >
                        <Text
                            block variant="xLarge" className="info-title">
                            {this.props.localeStrings.aboutApp}
                        </Text>
                        <Text block variant="small" className="info-titlebody">
                            {this.props.localeStrings.appDescription}
                        </Text>
                        <Text block variant="xLarge" className="info-title">
                            {this.props.localeStrings.headerAdditionalResource}
                        </Text>
                        <Text block variant="small" className="info-titlebody">
                            {this.props.localeStrings.bodyAdditionalResource}
                        </Text>
                        <Text block variant="xLarge" className="info-title">
                            {this.props.localeStrings.msPublicSector}
                        </Text>
                        <Link href={constants.msPublicSectorUrl} target="_blank" className="info-link">
                            {constants.msPublicSectorUrl}
                        </Link>
                        <Text block variant="xLarge" className="info-title">
                            {this.props.localeStrings.drivingAdoption}
                        </Text>
                        <Link href={constants.drivingAdoptionUrl} target="_blank" className="info-link">
                            {constants.drivingAdoptionUrl}
                        </Link>
                        <Text block variant="small">
                            {this.props.localeStrings.currentVersion} : {constants.AppVersion}
                        </Text>
                        <Text block variant="small">
                            {this.props.localeStrings.latestVersion} : <Link href={constants.githubEocUrl} target="_blank" rel="noreferrer">{this.props.localeStrings.githubLabel}</Link>
                        </Text>
                        <Text block variant="xLarge" className="info-title">
                            ----
                        </Text>
                        <Text block variant="small">
                            {this.props.localeStrings.eocPage}
                        </Text>
                        <Text block variant="small">
                            {this.props.localeStrings.overview} <Link href={constants.m365EocUrl} target="_blank" rel="noreferrer">{this.props.localeStrings.msAdoptionHubLink}</Link>
                        </Text>
                        <Link href={constants.m365EocUrl} target="_blank" className="info-link">
                            {constants.m365EocUrl}
                        </Link>
                        <Text block variant="small">
                            {this.props.localeStrings.solutionLink} <Link href={constants.m365EocAppUrl} target="_blank" rel="noreferrer">{this.props.localeStrings.githubLabel}</Link>
                        </Text>
                        <Link href={constants.m365EocAppUrl} target="_blank" className="info-link">
                            {constants.m365EocAppUrl}
                        </Link>
                    </Callout>
                )}
            </>
        )
    }
}
