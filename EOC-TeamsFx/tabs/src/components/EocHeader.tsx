import {
    Link,
    Popover, PopoverSurface,
    PopoverTrigger,
    Text
} from "@fluentui/react-components";
import { Info24Regular, PersonFeedback24Regular, QuestionCircle24Regular } from "@fluentui/react-icons";
import { Flex, FlexItem, Tooltip } from "@fluentui/react-northstar";
import { Component } from "react";
import * as constants from "../common/Constants";


interface IHeaderProps {
    clickcallback: () => void; //will redirects to home
    context?: any;
    localeStrings: any;
    appTitle: string,
    currentUserName: string;
    currentThemeName: string;
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
        const isDarkOrContrastTheme =
            this.props.currentThemeName === constants.darkMode || this.props.currentThemeName === constants.contrastMode;
        return (
            <div>
                <div className={`eoc-header${isDarkOrContrastTheme ? " eoc-header-darkcontrast" : ""}`}>
                    <Flex gap="gap.small" space='between' >
                        <Flex gap="gap.small" vAlign="center">
                            <img
                                src={require("../assets/Images/AppLogo.svg").default}
                                alt="Ms Logo"
                                className="ms-logo"
                                title={this.props.appTitle}
                            />
                            <main aria-label={this.props.appTitle}>
                                <h1 className="header-text" title={this.props.appTitle}>
                                    {this.props.appTitle}
                                </h1>
                            </main>
                        </Flex>
                        <Flex gap={this.state.isDesktop ? "gap.large" : "gap.medium"} vAlign="center">
                            <FlexItem>
                                <Popover
                                    withArrow={true}
                                    open={this.state.isCalloutVisible}
                                    inline={true}
                                    onOpenChange={this.toggleIsCalloutVisible}
                                    positioning="below-start"
                                    size='medium'
                                    closeOnScroll={true}
                                >
                                    <PopoverTrigger disableButtonEnhancement={true}>
                                        <div
                                            onClick={this.toggleIsCalloutVisible}
                                            onKeyUp={(event) => {
                                                if (event.key === constants.enterKey)
                                                    this.toggleIsCalloutVisible();
                                            }}
                                            tabIndex={0}
                                            role="button"
                                        >
                                            <Tooltip
                                                content={this.props.localeStrings.moreInfo}
                                                pointing={false}
                                                trigger={<Info24Regular title={this.props.localeStrings.moreInfo} className="header-icon" />}
                                            />
                                            
                                        </div>
                                    </PopoverTrigger>
                                    <PopoverSurface
                                        as="div"
                                        className="info-callout"
                                    >
                                        <h2><Text block as="h2"
                                            className="info-title">
                                            {this.props.localeStrings.aboutApp}
                                        </Text></h2>
                                        <Text block className="info-titlebody">
                                            {this.props.localeStrings.appDescription}
                                        </Text>
                                        <h2><Text block as="h2" className="info-title">
                                            {this.props.localeStrings.headerAdditionalResource}
                                        </Text></h2>
                                        <Text block as="span" className="info-titlebody">
                                            {this.props.localeStrings.bodyAdditionalResource}
                                        </Text>
                                        <Link href={constants.msPublicSectorUrl} target="_blank" inline className="info-link" onKeyDown={(event) => {
                                            if (event.shiftKey && event.key === constants.tabKey)
                                                this.toggleIsCalloutVisible()
                                        }}>
                                            {this.props.localeStrings.msPublicSector}
                                        </Link>
                                        <Link href={constants.drivingAdoptionUrl} target="_blank" inline className="info-link">
                                            {this.props.localeStrings.drivingAdoption}
                                        </Link>
                                        <Text block as="span" className="info-title">
                                            ----
                                        </Text>
                                        <Text block as="span" className="info-titlecontent">
                                            {this.props.localeStrings.currentVersion} : {constants.AppVersion}
                                        </Text>

                                        <Text block as="span" className="info-titlecontent">
                                            {this.props.localeStrings.latestVersion} : <Link href={constants.githubEocUrl} target="_blank" inline rel="noreferrer">{this.props.localeStrings.githubLabel}</Link>
                                        </Text>
                                        <Text block as="span" className="info-title">
                                            ----
                                        </Text>
                                        <Text block as="span" className="info-titlecontent">
                                            {this.props.localeStrings.eocPage}
                                        </Text>
                                        <Text block as="span" className="info-titlecontent">
                                            {this.props.localeStrings.overview} <Link href={constants.m365EocUrl} target="_blank" inline rel="noreferrer">{this.props.localeStrings.msAdoptionHubLink}</Link>
                                        </Text>

                                        <Text block as="span" className="info-titlecontent">
                                            {this.props.localeStrings.solutionLink} <Link href={constants.m365EocAppUrl} target="_blank" inline rel="noreferrer" onKeyDown={(event) => {
                                                if (!event.shiftKey)
                                                    this.toggleIsCalloutVisible()
                                            }}>{this.props.localeStrings.githubLabel}</Link>
                                        </Text>
                                    </PopoverSurface>
                                </Popover>
                            </FlexItem>
                            <FlexItem>
                                <a
                                    href={constants.helpUrl}
                                    target="_blank" rel="noreferrer"
                                    tabIndex={0}
                                >
                                    <Tooltip
                                        content={this.props.localeStrings.support}
                                        pointing={false}
                                        trigger={<QuestionCircle24Regular title={this.props.localeStrings.support} className="header-icon" />}
                                    />
                                </a>
                            </FlexItem>
                            <FlexItem>
                                <a href={constants.feedbackUrl} target="_blank" tabIndex={0} rel="noreferrer">
                                    <Tooltip
                                        content={{ content: this.props.localeStrings.feedback }}
                                        pointing={false}
                                        trigger={<PersonFeedback24Regular title={this.props.localeStrings.feedback} className="header-icon" />}
                                    />
                                </a>
                            </FlexItem>
                        </Flex>
                    </Flex>
                </div >
                <div className={`sub-header${isDarkOrContrastTheme ? " sub-header-" + this.props.currentThemeName : ""}`}>
                    <h2 className='container' id="sub-heading">
                        {this.props.localeStrings.welcome} {this.props.currentUserName}!
                    </h2>
                </div>
            </div>
        )
    }
}
