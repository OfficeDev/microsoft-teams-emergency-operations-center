import { IComboBoxOption } from "@fluentui/react";
import { ApplicationInsights } from '@microsoft/applicationinsights-web';
// interface for location picker
export interface ILocationBoxOption extends IComboBoxOption {
    locationItem: any;
}

export enum Mode {
    view,
    empty,
    editView,
}

interface IAddress {
    City?: string;
    CountryOrRegion?: string;
    State?: string;
    Street?: string;
}

export interface ILocationPickerItem {
    EntityType: string;
    LocationSource?: string;
    LocationUri?: string;
    UniqueId?: string;
    DisplayName: string;
    Address?: IAddress;
    Coordinates?: any; 
}

export interface ILocationPickerProps {
    className?: string;
    disabled?: boolean;
    label?: string;
    placeholder?: string;
    defaultValue?: ILocationPickerItem;
    onChange?: (newValue: any) => void;
    errorMessage?: string;
    appInsights: ApplicationInsights;
    userPrincipalName: any;
    graphBaseUrl: any;
}

export interface ILocationPickerState {
    currentMode: Mode;
    searchText: string | null;
    isCalloutVisible: boolean;
    selectedItem: any;
    options: any;
}
