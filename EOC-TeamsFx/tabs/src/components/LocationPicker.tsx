import { ComboBox, FontIcon, IconButton, Persona, PersonaSize, Text } from "@fluentui/react";
import * as React from 'react';
import CommonService from "../common/CommonService";
import * as constants from '../common/Constants';
import styles from '../scss/LocationPicker.module.scss';
import { ILocationBoxOption, ILocationPickerItem, ILocationPickerProps, ILocationPickerState, Mode } from './ILocationPicker';

export class LocationPicker extends React.Component<ILocationPickerProps, ILocationPickerState> {
    private _token: any;
    private focusRef: any = null;
    constructor(props: ILocationPickerProps) {
        super(props);
        this.getOutlookToken();
        this.focusRef = React.createRef();
        if (props.defaultValue) {
            this.state = {
                options: [],
                currentMode: Mode.editView,
                searchText: null,
                isCalloutVisible: true,
                selectedItem: props.defaultValue,
            };
        }
        else {
            this.state = {
                options: [],
                currentMode: Mode.empty,
                searchText: null,
                isCalloutVisible: true,
                selectedItem: props.defaultValue,
            };
        }
    }
    private dataService = new CommonService();

    public componentWillReceiveProps(nextProps: ILocationPickerProps): void {
        if (nextProps.defaultValue !== this.props.defaultValue) {
            if (nextProps.defaultValue) {
                this.setState({ selectedItem: nextProps.defaultValue, currentMode: Mode.editView });
            }
        }
    }

    //render the location picker controls
    public render(): JSX.Element {
        const { label } = this.props;

        return (
            <div>
                {label ? <Text>{label}</Text> : null}
                {this.getMainContent()}
            </div>
        );
    }

    //method to render combobox callout
    private onRenderOption = (item: any): JSX.Element => {
        const {
            text,
            locationItem
        } = item;
        if (locationItem.EntityType === "Custom") {
            return <Persona
                text={text}
                imageAlt={locationItem.EntityType}
                secondaryText={locationItem.DisplayName}
                size={PersonaSize.size40}
                onRenderInitials={this.customRenderInitials}
                className='location-picker-custom-item'
            />;
        }
        else
            return <Persona
                text={text}
                imageAlt={locationItem.EntityType}
                secondaryText={this.getLocationText(locationItem, "full")}
                size={PersonaSize.size40}
                onRenderInitials={this.customRenderInitials}
                className='location-picker-custom-item'
            />;
    }

    //method to display loation picker control
    private getMainContent = (): React.ReactNode => {
        const { options, selectedItem, currentMode } = this.state;
        const { className, disabled, placeholder, errorMessage } = this.props;

        switch (currentMode) {
            case Mode.empty:
                return <ComboBox
                    className={className}
                    disabled={disabled}
                    placeholder={placeholder}
                    allowFreeform={true}
                    autoComplete="on"
                    options={options}
                    onRenderOption={this.onRenderOption}
                    calloutProps={{ directionalHintFixed: true, doNotLayer: true, className: "incident-location-callout" }}
                    buttonIconProps={{ iconName: "MapPin" }}
                    useComboBoxAsMenuWidth={true}
                    openOnKeyboardFocus={true}
                    scrollSelectedToTop={true}
                    isButtonAriaHidden={true}
                    onInput={(e: any) => this.getLocations(e.target["value"])}
                    onChange={this.onChange}
                    errorMessage={errorMessage}
                    ariaLabel={"Location" + constants.requiredAriaLabel}
                    id="incident-location-listbox"
                    onMenuOpen={this.onLocationMenuOpen}
                />;
            case Mode.editView:
                if (selectedItem.EntityType === "Custom") {
                    return <div
                        ref={this.focusRef}
                        data-selection-index={0}
                        data-is-focusable={true}
                        role="listitem"
                        className={styles.pickerItemContainer}
                        onBlur={this.onBlur}
                        tabIndex={0}>
                        <Persona
                            data-is-focusable="false"
                            imageAlt={selectedItem.EntityType}
                            tabIndex={0}
                            text={selectedItem.DisplayName}
                            title="Location"
                            className={styles.persona}
                            size={PersonaSize.size40}
                            onRenderInitials={this.customRenderInitials} />
                        <IconButton
                            data-is-focusable="false"
                            tabIndex={0}
                            iconProps={{ iconName: "Cancel" }}
                            title="Clear"
                            ariaLabel="Clear"
                            disabled={disabled}
                            className={styles.closeButton}
                            onClick={this.onIconButtonClick} />
                    </div>;
                }

                return <div
                    ref={this.focusRef}
                    data-selection-index={0}
                    data-is-focusable={true}
                    role="listitem"
                    className={styles.pickerItemContainer}
                    onBlur={this.onBlur}
                    tabIndex={0}>
                    <Persona
                        data-is-focusable="false"
                        imageAlt={selectedItem.EntityType}
                        tabIndex={0}
                        text={selectedItem.DisplayName}
                        title="Location"
                        className={styles.persona}
                        secondaryText={this.getLocationText(selectedItem, "full")}
                        size={PersonaSize.size40}
                        onRenderInitials={this.customRenderInitials} />
                    {!disabled ?
                        <IconButton
                            data-is-focusable="false"
                            tabIndex={0}
                            iconProps={{ iconName: "Cancel" }}
                            title="Clear"
                            ariaLabel="Clear"
                            disabled={disabled} className={styles.closeButton}
                            onClick={this.onIconButtonClick} /> : null}
                </div>;
        }
        return null;
    }

    //on menu open, add the ariaLabel attribute to fix the position issue in iOS for accessbility
    private onLocationMenuOpen = () => {

        //adding option position information to aria attribute to fix the accessibility issue in iOS Voiceover
        if (navigator.userAgent.match(/iPhone/i)) {
            const listBoxElement: any = document.getElementById("incident-location-listbox-list")?.children;
            if (listBoxElement?.length > 0) {
                for (let i = 0; i < listBoxElement?.length; i++) {
                    const buttonId = `incident-location-listbox-list${i}`;
                    const buttonElement: any = document.getElementById(buttonId);
                    const ariaLabel = `${buttonElement.innerText} ${i + 1} of ${listBoxElement.length}`;
                    buttonElement?.setAttribute("aria-label", ariaLabel);
                }
            }
        }

    }

    //method to get location text
    private getLocationText = (item: ILocationPickerItem, mode: "full" | "street" | "noStreet"): string => {
        if (!item.Address) {
            return '';
        }

        const address = item.Address;

        switch (mode) {
            case "street":
                return address.Street || "";
            case "noStreet":
                return `${address.City ? address.City + ", " : ''}${address.State ? address.State + ", " : ""}${address.CountryOrRegion || ""}`;
        }

        return `${address.Street ? address.Street + ", " : ''}${address.City ? address.City + ", " : ""}${address.State ? address.State + ", " : ''}${address.CountryOrRegion || ""}`;
    }

    //method to remove location
    private onIconButtonClick = (): void => {
        this.setState({ currentMode: Mode.empty, selectedItem: null });
        if (this.props.onChange) {
            this.props.onChange(null);
        }
    }

    //method to call when user navigate to some other control
    private onBlur = (ev: any): void => {
        try {
            if (ev !== null && ev.relatedTarget["title"] !== "Location" && ev.relatedTarget["title"] !== "Clear") {
                this.setState({ currentMode: Mode.editView });
            }
        } catch (error: any) {
            console.log(error);
        }
    }

    //method to call on location change
    private onChange = (ev: any, option: any): void => {
        if (option !== undefined) {
            this.setState({ selectedItem: option.locationItem, currentMode: Mode.editView },
                () => {
                    if (this.focusRef.current !== null)
                        this.focusRef.current.focus();
                });

            if (this.props.onChange) {
                this.props.onChange(option.locationItem);
            }
        }
    }

    //method to render custom locations 
    private customRenderInitials(props: any): JSX.Element {
        if (props.imageAlt === "Custom")
            return <FontIcon aria-label="Poi" iconName="Poi" style={{ fontSize: "14pt" }} />;
        else
            return <FontIcon aria-label="EMI" iconName="EMI" style={{ fontSize: "14pt" }} />;
    }

    //method to get bearer token for location picker API
    private async getOutlookToken(): Promise<void> {
        try {
            const credential = this.props.teamsUserCredential;
            const token = await credential.getToken(this.props.graphBaseUrl !== constants.defaultGraphBaseURL ? constants.defaultOutlookBaseURLGCCH : constants.defaultOutlookBaseURL);
            this._token = token?.token;
        } catch (error) {
            console.error(
                constants.errorLogPrefix + "LocationPicker_getOutlookToken \n",
                JSON.stringify(error)
            );
            // Log Exception
            this.dataService.trackException(this.props.appInsights, error, constants.componentNames.IncidentDetailsComponent, 'TeamNameConfiguration_GetConfiguration', this.props.userPrincipalName);
            throw error;
        }
    }

    ///method to get locations using outlook API
    private async getLocations(searchText: any): Promise<void> {
        try {
            let locationAPIUrl = this.props.graphBaseUrl !== constants.defaultGraphBaseURL ? constants.outlookAPIFindLocationsGCCH : constants.outlookAPIFindLocations;
            const locationAPIResponse = await fetch(locationAPIUrl, {
                method: 'post',
                headers: new Headers({
                    "Content-type": "application/json",
                    "Cache-Control": "no-cache",
                    "Authorization": `Bearer ${this._token}`
                }),
                body: JSON.stringify({
                    "QueryConstraint": {
                        "Query": searchText
                    }
                })
            });
            const data = await locationAPIResponse.json() as { MeetingLocations: [{ MeetingLocation: ILocationPickerItem }] };
            const optionsForCustomRender: ILocationBoxOption[] = [];
            data.MeetingLocations.forEach((v, i) => {
                const loc: ILocationPickerItem = v["MeetingLocation"];
                optionsForCustomRender.push({ text: v.MeetingLocation["DisplayName"], key: i, locationItem: loc });
            });
            optionsForCustomRender.push({ text: 'Use this location', key: 7, locationItem: { DisplayName: searchText, EntityType: "Custom" } });
            this.setState({ options: optionsForCustomRender });
        }
        catch (error) {
            console.error(
                constants.errorLogPrefix + "LocationPicker_GetLocations \n",
                JSON.stringify(error)
            );
            // Log Exception
            this.dataService.trackException(this.props.appInsights, error, constants.componentNames.IncidentDetailsComponent, 'TeamNameConfiguration_GetConfiguration', this.props.userPrincipalName);

        }

    }
}