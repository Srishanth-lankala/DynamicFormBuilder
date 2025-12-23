import { Document, Page, Text, View, StyleSheet, Image, Font } from '@react-pdf/renderer';
import React from 'react';

// Define hyphenation callback to break words properly
const hyphenationCallback = (word: string) => {
    if (word.length <= 5) return [word]; // Don't hyphenate short words
    const syllables = [];
    for (let i = 0; i < word.length; i += 3) {
        syllables.push(word.slice(i, i + 3));
    }
    return syllables;
};

// Register hyphenation callback
Font.registerHyphenationCallback(hyphenationCallback);

// Define styles for the PDF
const styles = StyleSheet.create({
    page: {
        paddingTop: 60,
        paddingBottom: 30,
        paddingHorizontal: 30,
        fontSize: 10,
        fontFamily: 'Helvetica',
    },
    header: {
        position: 'absolute',
        top: 15,
        left: 30,
        right: 30,
        flexDirection: 'row',
        justifyContent: 'space-between',
        alignItems: 'center',
        borderBottom: '1px solid black',
        paddingBottom: 5,
        height: 40,
    },
    logo: {
        width: 170,
        height: 30,
    },
    headerText: {
        paddingBottom: 2,
        fontSize: 10,
        fontWeight: 'normal',
    },
    headerTexttop: {
        paddingBottom: 2,
        fontSize: 10,
        fontWeight: 'normal',
        marginTop: 5
    },
    TableheaderText: {
        fontSize: 12,
        fontWeight: 'bold',
    },
    TableheaderTextReqName: {
        fontSize: 12,

    },
    footer: {
        position: 'absolute',
        bottom: 10,
        left: 30,
        right: 30,
        textAlign: 'center',
        fontSize: 8,
        color: '#666',
        paddingTop: 5,
        borderTop: '1px solid #ccc',
    },
    table: {
        marginTop: 10,
        border: '1px solid #ccc',
        marginBottom: 10,
    },
    tableRow: {
        flexDirection: 'row',
        borderBottom: '1px solid #ccc',
    },
    tableHeader: {
        backgroundColor: '#f0f0f0',
        fontWeight: 'bold',
    },
    serialNoCell: {
        width: '10%', // For Data table
        padding: 5,
        borderRight: '1px solid #ccc',
        textAlign: 'left',
    },
    dataTableKeyCell: {
        width: '30%', // Adjusted for S.No column
        padding: 5,
        borderRight: '1px solid #ccc',
        textAlign: 'left',
    },
    dataTableValueCell: {
        width: '60%', // Adjusted for S.No column
        padding: 5,
        textAlign: 'left',
    },
    tableSerialNoCell: {
        width: '5%', // For Resources Allocation and Feedback tables
        padding: 5,
        borderRight: '1px solid #ccc',
        textAlign: 'left',
    },
    tableCell6Columns: {
        width: '15.833%', // (95% / 6 columns) for Resources Allocation
        padding: 5,
        textAlign: 'left',
        borderRight: '1px solid #ccc',
    },
    tableCell4Columns: {
        width: '23.75%', // (95% / 4 columns) for Feedback
        padding: 5,
        textAlign: 'left',
        borderRight: '1px solid #ccc',
    },
    tableCell2Columns: {
        width: '95%', // (95% / 2 columns)
        padding: 5,
        textAlign: 'left',
        borderRight: '1px solid #ccc',
    },
});

// Page dimensions (A4 in points)
const PAGE_HEIGHT = 842;
const PAGE_WIDTH = 595;
const HEADER_HEIGHT = 50;
const FOOTER_HEIGHT = 30;
const CONTENT_HEIGHT = PAGE_HEIGHT - HEADER_HEIGHT - FOOTER_HEIGHT - 60;
const BASE_ROW_HEIGHT = 20;

// Helper function to estimate text height
const estimateTextHeight = (text: string, width: number, fontSize: number) => {
    const charsPerLine = width / (fontSize / 1.5);
    const lines = Math.ceil(text.length / charsPerLine);
    return lines * fontSize * 1.1;
};

// Helper function to estimate row height for data table (3 columns now: S.No, Key, Value)
const estimateDataRowHeight = (row: { serialNo: string; label: string; value: string }) => {
    const serialNoWidth = (PAGE_WIDTH - 60) * 0.10;
    const keyWidth = (PAGE_WIDTH - 60) * 0.30;
    const valueWidth = (PAGE_WIDTH - 60) * 0.60;
    const serialNoHeight = estimateTextHeight(row.serialNo, serialNoWidth, 10);
    const keyHeight = estimateTextHeight(row.label, keyWidth, 10);
    const valueHeight = estimateTextHeight(row.value || 'N/A', valueWidth, 10);
    return Math.max(BASE_ROW_HEIGHT, serialNoHeight, keyHeight, valueHeight);
};

// Helper function to estimate row height for other tables
const estimateRowHeight = (row: any, columns: number) => {
    let maxCellHeight = BASE_ROW_HEIGHT;
    Object.values(row).forEach((cell: any) => {
        const cellText = typeof cell === 'object' ? cell.name || '' : cell?.toString() || '';
        const cellWidth = (PAGE_WIDTH - 60) * (columns === 6 ? 0.15833 : 0.2375); // Adjust for S.No column
        const cellHeight = estimateTextHeight(cellText, cellWidth, 10);
        maxCellHeight = Math.max(maxCellHeight, cellHeight);
    });
    const serialNoHeight = estimateTextHeight(row.serialNo || 'S.No', (PAGE_WIDTH - 60) * 0.05, 10);
    return Math.max(maxCellHeight, serialNoHeight) + 2;
};

// Helper function to estimate table title height
const estimateTableTitleHeight = () => {
    return estimateTextHeight("Table Name: ", (PAGE_WIDTH - 60), 12) + 10; // Approximate height of TableheaderText (fontSize: 12) + margin
};

// Helper function to paginate content
const paginateContent = (dataRows: any[], tables: any[]) => {
    const pages = [];
    let currentPage: any = { content: [] };
    let currentHeight = 0;

    // Combine all content into a single stream
    const allContent: { type: string; data: any }[] = [];

    // Add Data table with S.No
    allContent.push({ type: 'dataTableHeader', data: { serialNo: 'S.No', label: "Key", value: "Value" } });
    dataRows.forEach((row, index) => {
        allContent.push({ type: 'dataTableRow', data: { ...row, serialNo: (index + 1).toString() } });
    });

    // Add other tables with S.No
    tables.forEach((table) => {
        const header = {  ...table.header };
        const rows = table.rows.slice(1);
        allContent.push({ type: 'tableHeader', data: { header, tableIndex: tables.indexOf(table), content: table.content } });
        rows.forEach((row: any) => {
            allContent.push({ type: 'tableRow', data: { row: { ...row }, tableIndex: tables.indexOf(table) } });
        });
    });

    let currentDataTable: any = null;
    let currentTables: any[] = Array(tables.length).fill(null).map(() => ({ rows: [] }));
    let serialNoCounters: number[] = Array(tables.length).fill(0); // Track S.No for each table
    let dataTableSerialNo = 0; // Track S.No for Data table

    // Paginate all content
    for (let index = 0; index < allContent.length; index++) {
        const item = allContent[index];
        let rowHeight = 0;
        let rowData = item.data;
        let tableTitleHeight = 0;

        if (item.type === 'dataTableHeader' || item.type === 'dataTableRow') {
            if (item.type === 'dataTableRow') {
                dataTableSerialNo++;
                rowData = { ...item.data, serialNo: dataTableSerialNo.toString() };
            }
            rowHeight = estimateDataRowHeight(rowData);
            tableTitleHeight = item.type === 'dataTableHeader' ? estimateTableTitleHeight() : 0;
        } else if (item.type === 'tableHeader') {
            rowHeight = estimateRowHeight(item.data.header, Object.keys(item.data.header).length - 1);
            tableTitleHeight = estimateTableTitleHeight();
        } else if (item.type === 'tableRow') {
            serialNoCounters[item.data.tableIndex]++;
            rowData = { row: { ...item.data.row, serialNo: serialNoCounters[item.data.tableIndex].toString() }, tableIndex: item.data.tableIndex };
            rowHeight = estimateRowHeight(rowData.row, Object.keys(tables[item.data.tableIndex].header).length - 1);
        }

        // Add margin between tables
        const isNewTableStart = item.type === 'tableHeader' && currentTables[item.data.tableIndex].rows.length === 0;
        const additionalMargin = isNewTableStart && currentHeight > 0 ? 10 : 0;

        // Check if adding the current item exceeds the page height
        if (currentHeight + rowHeight + additionalMargin + tableTitleHeight > CONTENT_HEIGHT) {
            // Push current page content
            if (currentDataTable && currentDataTable.rows.length > 0) {
                currentPage.content.push({ type: 'dataTable', data: currentDataTable });
            }
            currentTables.forEach((table, tableIndex) => {
                if (table.rows.length > 0) {
                    currentPage.content.push({ type: 'table', data: { ...tables[tableIndex], rows: table.rows, content: tables[tableIndex].content } });
                }
            });

            // Push the page if it has content
            if (currentPage.content.length > 0) {
                pages.push(currentPage);
            }

            // Reset for new page
            currentPage = { content: [] };
            currentHeight = 0;
            currentDataTable = null;
            dataTableSerialNo = 0; // Reset Data table S.No (not spanning pages in this case)

            // Handle tables that continue onto the new page
            currentTables.forEach((table, tableIndex) => {
                if (table.rows.length > 0 && index < allContent.length) {
                    const nextItem = allContent.slice(index).find(next => next.type === 'tableRow' && next.data.tableIndex === tableIndex);
                    if (nextItem) {
                        // Start the table on the new page with its header
                        currentTables[tableIndex] = { rows: [{ serialNo: 'S.No', ...tables[tableIndex].header }] };
                        currentHeight = estimateRowHeight(tables[tableIndex].header, Object.keys(tables[tableIndex].header).length - 1) + estimateTableTitleHeight();
                    } else {
                        currentTables[tableIndex] = { rows: [] };
                    }
                }
            });
        }

        if (item.type === 'dataTableHeader') {
            currentDataTable = { rows: [rowData] };
            currentHeight += rowHeight + tableTitleHeight;
        } else if (item.type === 'dataTableRow') {
            if (!currentDataTable) {
                currentDataTable = { rows: [{ serialNo: 'S.No', label: "Key", value: "Value" }] };
                currentHeight += estimateDataRowHeight({ serialNo: 'S.No', label: "Key", value: "Value" }) + estimateTableTitleHeight();
            }
            currentDataTable.rows.push(rowData);
            currentHeight += rowHeight;
        } else if (item.type === 'tableHeader') {
            currentTables[item.data.tableIndex].rows.push(rowData.header);
            currentHeight += rowHeight + additionalMargin + tableTitleHeight;
        } else if (item.type === 'tableRow') {
            if (currentTables[rowData.tableIndex].rows.length === 0) {
                currentTables[rowData.tableIndex].rows.push({ serialNo: 'S.No', ...tables[rowData.tableIndex].header });
                currentHeight += estimateRowHeight(tables[rowData.tableIndex].header, Object.keys(tables[rowData.tableIndex].header).length - 1) + estimateTableTitleHeight();
            }
            currentTables[rowData.tableIndex].rows.push(rowData.row);
            currentHeight += rowHeight;
        }
    }

    // Push the last page
    if (currentDataTable && currentDataTable.rows.length > 0) {
        currentPage.content.push({ type: 'dataTable', data: currentDataTable });
    }
    currentTables.forEach((table, index) => {
        if (table.rows.length > 0) {
            currentPage.content.push({ type: 'table', data: { ...tables[index], rows: table.rows, content: tables[index].content } });
        }
    });
    if (currentPage.content.length > 0) {
        pages.push(currentPage);
    }

    return pages;
};

const today = new Date();

const options: Intl.DateTimeFormatOptions = { day: '2-digit', month: 'short', year: 'numeric' };
const formatted = today.toLocaleDateString('en-GB', options);

console.log(formatted); // e.g., "20 May 2025"

// Helper function to normalize column names to title case
const toTitleCase = (str: string) => {
    return str
        .toLowerCase()
        .split(' ')
        .map(word => word.charAt(0).toUpperCase() + word.slice(1))
        .join(' ');
};

interface MyDocumentProps {
    data: any[][] | { [key: string]: any };
    loghistory: any[];
    tables: any[];
    logo: any;
    txitem: any;
}

const MyDocument: React.FC<MyDocumentProps> = ({ data, loghistory, tables, logo, txitem }) => {
    // Convert data to fields (handle both array and object formats)
    let fields: { label: string; value: string }[] = [];

    if (Array.isArray(data)) {
        fields = data.map(([label, value]) => ({
            label,
            value: value || 'N/A',
        }));
    } else if (data && typeof data === 'object') {
        fields = Object.entries(data).map(([label, value]) => ({
            label,
            value: value || 'N/A',
        }));
    } else {
        console.error('MyDocument: Expected data to be an array or object, received:', data);
        return (
            <Document>
                <Page size="A4" style={styles.page}>
                    <View style={styles.header} fixed>
                        <View style={{ flexDirection: 'column' }}>
                            <Text style={styles.headerText}>Form : {txitem.AppName}</Text>
                            <Text style={styles.headerText}>RequestID : {txitem.Title}</Text>
                            <Text style={styles.headerText}>{formatted}</Text>
                        </View>
                        {logo && <Image style={styles.logo} src={logo} />}
                    </View>
                    <View style={styles.table}>
                        <Text>Error: Invalid data format. Please provide an array of key-value pairs or an object.</Text>
                    </View>
                    <Text
                        style={styles.footer}
                        render={({ pageNumber, totalPages }) => `Page ${pageNumber} of ${totalPages} | © 2025 CLOUDANGLES`}
                        fixed
                    />
                </Page>
            </Document>
        );
    }

    // Transform tables to include header, exclude 'id', and normalize column names
    const formattedTables = tables.map((table, index) => {
        const header = table.columns.reduce((obj: any, col: any) => {
            const headerName = toTitleCase(col.headername || ''); // Normalize to title case
            obj[headerName] = headerName;
            return obj;
        }, {});
        const rows = table.rows.map((row: any) => {
            const { id, ...rest } = row;
            const correctedRow: any = {};
            Object.keys(rest).forEach((key) => {
                const newKey = toTitleCase(key); // Normalize to title case
                correctedRow[newKey] = rest[key];
            });
            return correctedRow;
        });
        return {
            ...table,
            header,
            rows: [header, ...rows],
            content: table.content || `Table ${index + 1}`, // Fallback to a generic name if content is missing
        };
    });

    // Add log history table at the end
    if (loghistory && loghistory.length > 0) {
        const logHistoryTable = {
            header: { 'Title': 'Title' },
            rows: [{ 'Title': 'Title' }, ...loghistory.map((log, index) => ({

                'Title': log.Title || 'No title'
            }))],
            content: 'Log History'
        };
        formattedTables.push(logHistoryTable);
    }

    const pages = paginateContent(fields, formattedTables);

    return (
        <Document>
            {pages.map((pageContent, pageIndex) => (
                <Page key={pageIndex} size="A4" style={styles.page}>
                    <View style={styles.header} fixed>
                        <View style={{ flexDirection: 'column' }}>
                            <Text style={styles.headerText}>Form : {txitem.AppName}</Text>
                            <Text style={styles.headerText}>RequestID : {txitem.Title}</Text>
                            <Text style={styles.headerText}>{formatted}</Text>
                        </View>
                        {logo && <Image style={styles.logo} src={logo} />}
                    </View>

                    {pageContent.content.map((contentItem: any, index: any) => {
                        if (contentItem.type === 'dataTable') {
                            return (
                                <View key={index}>
                                    <View style={{ flexDirection: 'row', justifyContent: 'space-between', marginBottom: 10 }}>
                                        <Text key={`title-${index}`} style={styles.TableheaderText}>Data Report</Text>
                                        <Text key={`author-${index}`} style={styles.TableheaderTextReqName}>Requestor: {txitem.Author.Title}</Text>
                                    </View>

                                    <Text key={`title-${index}`} style={styles.headerTexttop}>This document contains the requested information in tabular format</Text>
                                    <View key={index} style={styles.table}>
                                        {contentItem.data.rows.map((row: any, rowIndex: any) => (
                                            <View
                                                key={rowIndex}
                                                style={[styles.tableRow, rowIndex === 0 ? styles.tableHeader : {}]}
                                            >
                                                <Text style={styles.serialNoCell}>{row.serialNo}</Text>
                                                <Text style={styles.dataTableKeyCell}>{row.label}</Text>
                                                <Text style={styles.dataTableValueCell}>{row.value}</Text>
                                            </View>
                                        ))}
                                    </View>
                                </View>
                            );
                        } else if (contentItem.type === 'table') {
                            const isResourcesAllocation = Object.keys(contentItem.data.header).length === 6;
                            const isLogHistory = contentItem.data.content === 'Log History';

                            return (
                                <View key={index}>
                                    <Text key={`title-${index}`} style={styles.TableheaderText}>
                                        {contentItem.data.content}
                                    </Text>
                                    <View style={styles.table}>
                                        {contentItem.data.rows.map((row: any, rowIndex: any) => (
                                            <View
                                                key={rowIndex}
                                                style={[styles.tableRow, rowIndex === 0 ? styles.tableHeader : {}]}
                                            >
                                                {/* <Text style={styles.tableSerialNoCell}>{row.serialNo}</Text> */}
                                                {Object.entries(row).map(([key, cell]: [string, any], cellIndex: any) => {
                                                    if (key === 'serialNo') return null; // Skip S.No as it's already rendered
                                                    return (
                                                        <Text
                                                            key={cellIndex}
                                                            style={[
                                                                isResourcesAllocation ? styles.tableCell6Columns :
                                                                    isLogHistory ? styles.tableCell2Columns : styles.tableCell4Columns,
                                                                {
                                                                    borderRight:
                                                                        cellIndex === Object.keys(row).length - 1
                                                                            ? 'none'
                                                                            : '1px solid #ccc',
                                                                },
                                                            ]}
                                                        >
                                                            {typeof cell === 'object' ? cell.name || '' : cell?.toString() || ''}
                                                        </Text>
                                                    );
                                                })}
                                            </View>
                                        ))}
                                    </View>
                                </View>
                            );
                        }
                        return null;
                    })}

                    <Text
                        style={styles.footer}
                        render={({ pageNumber, totalPages }) => `Page ${pageNumber} of ${totalPages} | © 2025 CLOUDANGLES`}
                        fixed
                    />
                </Page>
            ))}
        </Document>
    );
};

export default MyDocument;