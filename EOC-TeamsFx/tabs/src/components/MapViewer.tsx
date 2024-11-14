import { ApplicationInsights } from '@microsoft/applicationinsights-web';
import React from "react";
import CommonService, { IListItem } from '../common/CommonService';
import * as constants from '../common/Constants';
import * as atlas from 'azure-maps-control';
import "azure-maps-control/dist/atlas.min.css";

export interface IMapViewerProps {
    incidentData: IListItem;
    azureMapKey: { value: string };
    showMessageBar(message: string, type: string): void;
    userPrincipalName: any;
    localeStrings: any;
    appInsights: ApplicationInsights;
    graphBaseUrl: any;
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

    //On component Load get the incidents data and generate map
    public async componentDidMount() {
        this.initMap(this.props.incidentData);
    }

    private dataService = new CommonService();

    //Get incidents data, generate map and pin points as per the icnidents data
    private initMap(incidentsData: any) {
        try {
            //Get the Azure map key configured by the user in Config Settings screen
            var azureMapKey = this.props.azureMapKey.value;
            
            //For GCCH Tenant
            if (this.props.graphBaseUrl !== constants.defaultGraphBaseURL) {
                atlas.setDomain('atlas.azure.us');
            }

            //Create azure map control with the Azure Maps subscription key
            const incidentsMap = new atlas.Map('azureMapControl', {           
                center: [0, 0],
                zoom: 0,
                language: 'en-US',
                view: 'Auto',
                authOptions: {
                    authType: atlas.AuthenticationType.subscriptionKey,
                    subscriptionKey: azureMapKey
                }
            });

            //Exclude the list of incidents which have Custom location 
            const filteredData = incidentsData.filter((e: any) => {
                let location;
                try {
                    location = JSON.parse(e.location);
                } catch (error) {
                    console.error(constants.errorLogPrefix + "_Error parsing location:", error);
                    return false;
                }
                return location && location.EntityType !== "Custom";
            });

            //Generate pins for each incident and add it to the map with colors depending on the incident status
            filteredData.forEach((item: any) => {
                let pinColor: string;
                switch (item.incidentStatusObj.status) {
                    case constants.planning:
                        pinColor = 'gray';
                        break;
                    case constants.active:
                        pinColor = 'orange';
                        break;
                    case constants.closed:
                        pinColor = 'green';
                        break;
                    default:
                        pinColor = 'gray';
                        break;
                }

                try {
                    const coordinates = JSON.parse(item.location).Coordinates;
                    //Create an HTML marker for the pins and add it to the map.
                    var pinMarker = new atlas.HtmlMarker({
                        color: pinColor,
                        text: item.incidentId,
                        position: [coordinates.Longitude, coordinates.Latitude],
                        popup: new atlas.Popup({
                            content: '<div style="padding:10px;color:grey"><a href="' + item.teamWebURL + '" target="_blank">' + item.incidentId + ': ' + item.incidentName + '</a> <br> ' + this.props.localeStrings.status + ': ' + item.incidentStatusObj.status + '<br>' + this.props.localeStrings.location + ': ' + JSON.parse(item.location).DisplayName + '<br>' + item.incidentCommander + "</div>",
                            pixelOffset: [0, -30]
                        })
                    });

                    //add the markers to thwe map
                    incidentsMap.markers.add(pinMarker);

                    //Add on mouse enter event to toggle the popup.
                    incidentsMap.events.add('mouseenter', pinMarker, () => {
                        pinMarker.togglePopup();
                    });
                }
                catch (error) {
                    console.error(constants.errorLogPrefix + "_Error parsing coordinates:", error);
                }
            });

        } catch (error) {
            console.error(
                constants.errorLogPrefix + "_MapViewer_initMap \n",
                JSON.stringify(error)
            );
            this.props.showMessageBar(this.props.localeStrings.genericErrorMessage + ", " + this.props.localeStrings.formatIncidentsDataFailedErrMsg, constants.messageBarType.error);
            // Log Exception in App insights
            this.dataService.trackException(this.props.appInsights, error, constants.componentNames.MapViewer, 'InitMap', this.props.userPrincipalName);
        }
    }

    //render method to return map control
    public render(): JSX.Element {
        return (
            <div className="map-viewer-component">
                <div id="azureMapControl" style={{ "width": "100%", "height": "100%" }} ></div>
            </div>
        );
    }
}