import React from "react";
import { PDFDownloadLink, Page, Text, View, Document } from '@react-pdf/renderer';
import { styles as PDFStyles } from '../assets/styles/IncidentHistoryPDFStyles';

export interface IIncidentHistoryPDFState {
}
export interface IIncidentHistoryPDFProps {
    localeStrings: any;
    incidentId: string;
    versionHistoryPDFData: any;
    currentThemeName: string;
    incidentVersionData: any;
}
export default class IncidentHistoryPDF extends React.PureComponent<IIncidentHistoryPDFProps, IIncidentHistoryPDFState>  {
    constructor(props: any) {
        super(props);
    }

    //Render method
    public render() {
        const pdfFileName = this.props.localeStrings.incidentHistory + " - " + this.props.incidentId + " - " + this.props.incidentVersionData[this.props.incidentVersionData.length - 1]?.incidentName;
        return (
            <PDFDownloadLink
                document={this.formatIncidentHistoryPDF(pdfFileName, this.props.versionHistoryPDFData)}
                fileName={`${pdfFileName}.pdf`}
                className="download-pdf"
            >
                {({ blob, url, loading, error }: any) => this.downloadIncidentHistoryPDF(loading, error)}
            </PDFDownloadLink>
        )
    }

    //Format Incident History PDF Document
    private formatIncidentHistoryPDF(incident: string, versionHistoryPDFData: any): JSX.Element {
        return (
            <Document>
                <Page style={PDFStyles.body} size={'A4'}>
                    <Text style={PDFStyles.mainHeading}>{incident}</Text>
                    {versionHistoryPDFData.map((item: any) => {
                        return (
                            item?.versionData.length > 0 ?
                                <View>
                                    <View style={PDFStyles.tableHeading}>
                                        <View><Text>{this.props.localeStrings.modifiedOn}: {item.modifiedOn}</Text></View>
                                        <View><Text>{this.props.localeStrings.modifiedBy}: {item.modifiedBy}</Text></View>
                                    </View>
                                    <View style={PDFStyles.table} key={item.modifiedOn + "-" + item.modifiedBy}>
                                        <View style={{ ...PDFStyles.tableRow, ...PDFStyles.tableHeaderRow }}>
                                            <View style={{ ...PDFStyles.tableCell, ...PDFStyles.tableCell1, ...PDFStyles.tableHeaderCell }}>
                                                <Text style={PDFStyles.tableCellText}>{this.props.localeStrings.field}</Text>
                                            </View>
                                            <View style={{ ...PDFStyles.tableCell, ...PDFStyles.tableCell2, ...PDFStyles.tableHeaderCell }}>
                                                <Text style={PDFStyles.tableCellText}>{this.props.localeStrings.new}</Text>
                                            </View>
                                            <View style={{ ...PDFStyles.tableCell, ...PDFStyles.tableCell3, ...PDFStyles.tableHeaderCell }}>
                                                <Text style={PDFStyles.tableCellText}>{this.props.localeStrings.old}</Text>
                                            </View>
                                        </View>
                                        {item.versionData.map((versionDataItem: any) => {
                                            return (
                                                <View style={PDFStyles.tableRow} key={item.version + "-" + versionDataItem.field} wrap={false} >
                                                    <View style={{ ...PDFStyles.tableCell, ...PDFStyles.tableCell1 }}>
                                                        <Text style={PDFStyles.tableCellText}>{versionDataItem.field}</Text>
                                                    </View>
                                                    <View style={{ ...PDFStyles.tableCell, ...PDFStyles.tableCell2 }}>
                                                        <Text style={PDFStyles.tableCellText}>{versionDataItem.newValue}</Text>
                                                    </View>
                                                    <View style={{ ...PDFStyles.tableCell, ...PDFStyles.tableCell3 }}>
                                                        <Text style={PDFStyles.tableCellText}>{versionDataItem.oldValue}</Text>
                                                    </View>
                                                </View>
                                            );
                                        })}
                                    </View>
                                </View> : <></>
                        );
                    })}

                    <Text style={PDFStyles.pageNumbers} render={({ pageNumber, totalPages }) => (
                        `${pageNumber} / ${totalPages}`
                    )} fixed />
                </Page>
            </Document>
        );
    }
    //Render download button content based on loading or error state.
    private downloadIncidentHistoryPDF(loading: any, error: any): JSX.Element {
        if (error) {
            console.error("pdf generation failed", error);
        }
        return (
            <>
                <img src={require("../assets/Images/PdfIcon.svg").default} alt="pdf-icon" className={`pdf-icon ${this.props.currentThemeName}-icon`} />
                <span className="download-text">{loading ? this.props.localeStrings.loadingLabel : this.props.localeStrings.downloadBtnLabel + " PDF"}</span>
            </>
        );
    }
}
