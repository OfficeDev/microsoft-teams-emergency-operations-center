import { ApplicationInsights } from '@microsoft/applicationinsights-web';
import BingMapsReact from "bingmaps-react";
import React from "react";
import CommonService, { IListItem } from '../common/CommonService';
import * as constants from '../common/Constants';

export interface IMapViewerProps {
    incidentData: IListItem;
    bingMapKey: any;
    showMessageBar(message: string, type: string): void;
    userPrincipalName: any;
    localeStrings: any;
    appInsights: ApplicationInsights;
}
export interface IMapViewerState {
    mapViewerData: any[];
}

export class MapViewer extends React.Component<IMapViewerProps, IMapViewerState> {
    constructor(props: IMapViewerProps) {
        super(props);
        this.state = {
            mapViewerData: []
        }
    }

    componentWillMount(): void {
        this.formatIncidentData(this.props.incidentData);
    }

    private dataService = new CommonService();

    // method to format incident data to show on map control 
    private formatIncidentData = (data: any) => {
        try {
            //filter locations which dont have coordinates and will not be visible in bing map like custom locations
            const filteredData = data.filter((e: any) => JSON.parse(e.location).EntityType !== "Custom");
            const pushPinsWithInfoboxes: any = [];
            filteredData.forEach((item: any) => {
                if (item.location !== "" || JSON.parse(item.location).Coordinates.Latitude !== "0") {
                    let pushPinColor: string;
                    switch (item.incidentStatusObj.status) {
                        case constants.planning:
                            pushPinColor = 'gray';
                            break;
                        case constants.active:
                            pushPinColor = 'orange';
                            break;
                        case constants.closed:
                            pushPinColor = 'green';
                            break;
                        default:
                            pushPinColor = 'gray';
                            break;
                    }
                    let data = {
                        center: {
                            latitude: JSON.parse(item.location).Coordinates === undefined ? 0 : JSON.parse(item.location).Coordinates.Latitude,
                            longitude: JSON.parse(item.location).Coordinates === undefined ? 0 : JSON.parse(item.location).Coordinates.Longitude
                        },
                        options: {
                            color: pushPinColor,
                            latitude: JSON.parse(item.location).Coordinates === undefined ? 0 : JSON.parse(item.location).Coordinates.Latitude,
                            description: '<a href="' + item.teamWebURL + '" target="_blank">' + item.incidentId + ': ' + item.incidentName + '</a> <br> ' + this.props.localeStrings.status + ': ' + item.incidentStatusObj.status + '<br>' + this.props.localeStrings.location + ': ' + JSON.parse(item.location).DisplayName + '<br>' + item.incidentCommander,
                        },
                    }
                    pushPinsWithInfoboxes.push(data);
                }
            });
            this.setState({
                mapViewerData: pushPinsWithInfoboxes
            });
        } catch (error) {
            console.log(error);
            console.error(
                constants.errorLogPrefix + "_MapViewer_FormatData \n",
                JSON.stringify(error)
            );
            this.props.showMessageBar(this.props.localeStrings.genericErrorMessage + ", " + this.props.localeStrings.formatIncidentsDataFailedErrMsg, constants.messageBarType.error);
            // Log Exception
            this.dataService.trackException(this.props.appInsights, error, constants.componentNames.MapViewer, 'FormatIncidentData', this.props.userPrincipalName);
        }
    }

    //render method to return map control
    public render(): JSX.Element {
        return (
            <div className="map-viewer-component">
                <BingMapsReact
                    bingMapsKey={this.props.bingMapKey.value}
                    mapOptions={{
                        navigationBarMode: "square",
                        color: 'dark'
                    }}
                    zoom={10}
                    pushPinsWithInfoboxes={this.state.mapViewerData}
                />
            </div>
        );
    }

}