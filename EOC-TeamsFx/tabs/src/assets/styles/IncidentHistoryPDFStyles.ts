import { StyleSheet } from '@react-pdf/renderer';



// Create styles
export const styles: any = StyleSheet.create({
    body: {
        width: "100%",
        padding: 10
    },

    mainHeading: {
        textAlign: "center",
        fontSize: 16,
        marginTop: 5,
        marginBottom: 10,
        color: "#000"
    },

    tableHeading: {
        fontSize: 12,
        marginTop: 10,
        color: "#000",
        display: "flex",
        flexDirection: "row",
        justifyContent: "space-between",
        alignItems: "center",
        width: "auto",
    },

    table: {
        display: "flex",
        width: "auto",
        margin: "auto",
        marginTop: 5,
        marginBottom: 30,
        borderTopLeftRadius: 3,
        borderTopRightRadius: 3
    },

    tableRow: {
        flexDirection: "row",
        borderStyle: "solid",
        borderRightWidth: 1,
        borderLeftWidth: 1,
    },

    tableHeaderRow: {
        backgroundColor: "#414156",
        borderTopLeftRadius: 2,
        borderTopRightRadius: 2
    },

    tableCell: {
        borderStyle: "solid",
        borderWidth: 0.5,
        borderLeftWidth: 0,
        borderRightWidth: 1,
        paddingBottom: 5
    },

    tableCell1: {
        width: "20%"
    },

    tableCell2: {
        width: "40%"
    },

    tableCell3: {
        borderRightWidth: 0,
        width: "40%"
    },

    tableHeaderCell: {
        color: "#fff"
    },

    tableCellText: {
        marginTop: 5,
        fontSize: 10,
        textAlign: "center"
    },

    pageNumbers: {
        position: 'absolute',
        bottom: 10,
        left: 0,
        right: 0,
        textAlign: 'center',
        fontSize: 10
    },
});