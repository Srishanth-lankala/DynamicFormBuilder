import React, { useEffect, useState } from "react";
import { DataTable } from "primereact/datatable";
import { Column } from "primereact/column";
// import { Dropdown } from "primereact/dropdown";
import { sp } from "@pnp/sp/presets/all";
import Sidebar from '../Sidebar/SideBar'
import TopNavBar from "../../CustomFormBuilder/TopBar";
import { dataservice } from "../encryptionutil";
// import Papa from "papaparse";
import { Chart } from 'primereact/chart';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faRefresh } from '@fortawesome/free-solid-svg-icons';

import { useLocation, useNavigate } from "react-router-dom";
import "./FormReports.css"
import { fetchTenantUser } from "../FetchTenantUser/fetchTenantUser";
import MyChartComponent from "./MyChartComponent";
import { Button } from "primereact/button";
import { Checkbox } from "primereact/checkbox";
// import LoadingSpinner from "../Loading";
import * as XLSX from "xlsx";
// import { saveAs } from "file-saver";
import { Toast } from 'primereact/toast';
import { useRef } from 'react';
// import * as XLSX from 'xlsx';
import { MultiSelect } from 'primereact/multiselect';
import { Calendar } from 'primereact/calendar';
import totalitemsicon from '../../../assets/Images/totalitemsicon.svg';
import excelicon from '../../../assets/Images/excelicon.svg';
import graphicon from '../../../assets/Images/graph.svg';
import tableicon from '../../../assets/Images/tableicon.svg';
import calendaricn from '../../../assets/Images/calendaricon.svg';


// interface LocationState{
//     appnames?: string[]; 
// }
const Reports: React.FC = () => {
    const location = useLocation();
    const [workflowData, setWorkflowData] = useState<any[]>([]);
    const [filteredData, setFilteredData] = useState<any[]>([]);
    const [onlyuserData, setOnlyuserData] = useState<any[]>([]);
    const [appNames, setAppNames] = useState<string[]>([]);
    const [statuses, setStatuses] = useState<string[]>([]);
    const [selectedAppNames, setSelectedAppNames] = useState<string[]>(location.state?.filters?.selectedAppNames || []);
    const [tempSelectedAppNames, setTempSelectedAppNames] = useState<string[]>(location.state?.filters?.selectedAppNames || []);


    const [selectedStatuses, setSelectedStatuses] = useState<string[]>(location.state?.filters?.selectedStatuses || []);
    const [tempSelectedStatuses, setTempSelectedStatuses] = useState<string[]>(location.state?.filters?.selectedStatuses || []);
    const [selectedFromDate, setSelectedFromDate] = useState<Date | null | undefined>(null);
    const [selectedToDate, setSelectedToDate] = useState<Date | null | undefined>(null);
    const [CurrentUserEmail, setCurrentUserEmail] = useState<string | null>(null);
    const [childData, setChildData] = useState<any[]>([]);

    const [, setShowChart] = useState<boolean>(false);
    const navigate = useNavigate();
    const [showLineChart, setShowLineChart] = useState(true);
    // const [loading, setLoading] = useState<boolean>(true)
    const allSelectedAppnames = tempSelectedAppNames.length === appNames.length;
    const allSelectedStatuses = tempSelectedStatuses.length === statuses.length;
    const toast = useRef<Toast>(null);
    const multiSelectRef = useRef<MultiSelect>(null);

    const dataserviceobj = new dataservice();

    const fetchWorkflowData = async () => {

        try {
            const Domain = await fetchTenantUser();
            const items = await sp.web.lists
                .getByTitle("WorkFlowProcessData")
                .items.select(
                    "ID",
                    "Title",
                    "Status",
                    "Created",
                    "Modified",
                    "AppName",
                    "Author/Id",
                    "Author/Title",
                    "CurApprover/Id",
                    "CurApprover/Title",
                    "CategoryType",
                    "AppCode",
                    "FormData",
                    "FinalCost",
                    "AllApprovers",
                    "CurrentAppCode",
                    "ParentAppId"
                )
                .expand("Author", 'CurApprover')
                .filter(`Domain eq '${Domain.TenantUsers}'`)
                // .orderBy("Created", false)
                .top(100)
                .getAll();
            //console.log("itemsnin Reporytttttttttttt--------------------", items)
            //console.log("domainnnnnnnnn", Domain.TenantUsers)
            setWorkflowData(items);
            setFilteredData(items);

            // const uniqueAppNames = [...new Set(items.map((item) => item.AppName.trim()))];
            const statusmasterstatuses = await sp.web.lists.getByTitle("StatusMaster").items.select("Status").top(100).getAll();
            const uniqueStatuses = [...statusmasterstatuses.map(item => item.Status.trim())]


            setStatuses(uniqueStatuses);
        } catch (error) {
            console.error("Error fetching workflow data:", error);
        }
        // finally{ setLoading(false)}
    };
    //     useEffect(() => {
    //         const state =location.state as LocationState | null;
    //   const prps:string[]=state?.appnames ?? [];
    //   console.log("stateeeeeeeeee", prps)
    //   const uniqueAppNames=prps;
    //   setAppNames(uniqueAppNames);
    //     },[])
    const fetchFormTitles = async () => {
        try {
            const Domain = await fetchTenantUser();


            // Fetch current user details
            const currentUser = await sp.web.currentUser.get();
            //console.log("Current User ID:", currentUser.Id);


            // Fetch all forms where Domain is 'TenantUsers' and FormViewer is either a person or a group
            const RAWallForms = await sp.web.lists.getByTitle("FormMaster")
                .items
                .select("ID", "Title", "FormJSON", "AppName", "AppCode", "FormType", "Domain", "FormViewer/Id", "FormViewer/Title", "VisibilityFlag", "TransactionCount")
                .expand("FormViewer")
                .filter(`Domain eq '${Domain.TenantUsers}' and ParentAppCode eq null`)
                .getAll();
            const allForms = RAWallForms?.map((item: any) => {
                return { ...item, FormJSON: dataserviceobj.decryptjson(item.FormJSON) };
            })
            //console.log("Fetched Forms:", allForms);
            //console.log("CurrentUserEmail------------- Isssssss", CurrentUserEmail)

            if (CurrentUserEmail === "BothUser Admin" || CurrentUserEmail === "OnlyAdmin") {

                const uniqueAppNames = [...new Set(allForms.map((item) => item.AppName.trim()))];
                setAppNames(uniqueAppNames);
                //console.log("fetchTenantUser------------- IN  IFFFF", Domain)
            }
            else if (CurrentUserEmail === "OnlyUser") {
                //console.log("fetchTenantUser------------- IN  ELSEEE", Domain)
                let groupIdsToCheck = new Set<number>();
                let formsWithCurrentUser: typeof allForms = [];

                // Identify records where FormViewer is the current user or a group
                allForms.forEach(form => {
                    if (form.FormViewer?.Id === currentUser.Id) {
                        formsWithCurrentUser.push(form);
                    } else if (form.FormViewer?.Id) {
                        groupIdsToCheck.add(form.FormViewer.Id);
                    }
                });

                //console.log("Group IDs to Check:", [...groupIdsToCheck]);

                // Fetch all site groups to determine which FormViewer ID belongs to a group
                const siteGroups = await sp.web.siteGroups.get();
                const siteGroupIds = new Set(siteGroups.map(group => group.Id));

                //console.log("All Site Group IDs:", [...siteGroupIds]);

                // Filter forms where FormViewer is a group
                const formsWithGroups = allForms.filter(
                    form => form.FormViewer?.Id && siteGroupIds.has(form.FormViewer.Id)
                );

                //console.log("Forms where FormViewer is a Group:", formsWithGroups);

                // Get logged-in user's group memberships
                const userGroups = await sp.web.currentUser.groups.get();
                const userGroupIds = new Set(userGroups.map(group => group.Id));
                const userGroupTitles = new Set(userGroups.map(group => group.Title));

                //console.log("User's Group IDs:", [...userGroupIds]);
                console.log("User's Group Titles:", [...userGroupTitles]);

                // Filter forms where the user is a member of the FormViewer group
                const formsWhereUserIsGroupMember = formsWithGroups.filter(
                    form => form.FormViewer?.Id && userGroupIds.has(form.FormViewer.Id)
                );

                //console.log("Forms where FormViewer is a Group AND User is a Member:", formsWhereUserIsGroupMember);

                // **New Condition: Also include forms where FormViewer Title is "All Employees"**
                const formsForAllEmployees = formsWithGroups.filter(
                    form => form.FormViewer?.Title === "All Employees"
                );

                //console.log("Forms where FormViewer is 'All Employees':", formsForAllEmployees);

                // Merge both filtered lists
                let finalFilteredForms = [...formsWhereUserIsGroupMember, ...formsWithCurrentUser, ...formsForAllEmployees];

                //console.log("Final Filtered Forms:", finalFilteredForms);
                const uniqueAppNames = [...new Set(finalFilteredForms.map((item) => item.AppName.trim()))];
                setAppNames(uniqueAppNames);



            }

        } catch (error) {
            console.error("Error fetching forms:", error);
        }
        // finally {

        // }
    };
    useEffect(() => {
        navigate(location.pathname, {
            state: {
                filters: { selectedAppNames, selectedStatuses }
            },
            replace: true
        });
    }, [selectedAppNames, selectedStatuses]);

    useEffect(() => {
        const onlyuser = async () => {
            //console.log("frooooooooooom use", CurrentUserEmail);
            type WorkFlowItem = {
                ID: number;
                Title: string;
                Status: string;
                Created: string;
                Modified: string;
                Author: { Title: string };
                CurApprover?: { Title: string; Id: number };
                AllApprovers?: string[];
                AppName: string;
                CategoryType: string
            };
            if (CurrentUserEmail === "OnlyUser") {
                const currentUser = await sp.web.currentUser.get();
                //get user created records
                const onlyusercreateddata = workflowData.filter((item) => item.Author.Id === currentUser.Id);
                //console.log("mydocuuuuuuuuu", onlyusercreateddata);
                setOnlyuserData(onlyusercreateddata);
                //console.log("filtered", onlyuserData);

                const onlyusertransactiondata: WorkFlowItem[] = workflowData.filter((item) => item.Status !== "Draft");
                //console.log("inprogress data", onlyusertransactiondata);
                let groupIdsToCheck = new Set<number>();
                let recordsWithCurrentUser: WorkFlowItem[] = [];

                // Identify records where CurApprover is a group or the current user
                onlyusertransactiondata.forEach((item) => {
                    if (item.CurApprover?.Id === currentUser.Id) {
                        recordsWithCurrentUser.push(item);
                    } else if (item.CurApprover?.Id) {
                        groupIdsToCheck.add(item.CurApprover?.Id);
                    }
                });

                //console.log("Group IDs to Check:", [...groupIdsToCheck]);

                // Fetch site groups to determine which CurApproverId belongs to a group
                const siteGroups = await sp.web.siteGroups.get();
                const siteGroupIds = new Set(siteGroups.map(group => group.Id));

                //console.log("All Site Group IDs:", [...siteGroupIds]);

                // Filter records where CurApproverId is a group
                const recordsWithGroups = onlyusertransactiondata.filter(
                    item => item.CurApprover?.Id && siteGroupIds.has(item.CurApprover?.Id)
                );

                //console.log("Records where CurApprover is a Group:", recordsWithGroups);

                // Find groups where the logged-in user is a member
                const userGroups = await sp.web.currentUser.groups.get();
                const userGroupIds = new Set(userGroups.map(group => group.Id));

                //console.log("User's Group IDs:", [...userGroupIds]);

                // Filter records where the user is a member of the CurApprover group
                const recordsWhereUserIsGroupMember = recordsWithGroups.filter(
                    item => item.CurApprover?.Id && userGroupIds.has(item.CurApprover?.Id)
                );

                //console.log("Records where CurApprover is a Group AND User is a Member:", recordsWhereUserIsGroupMember);

                // Merge both filtered lists
                const onlyusercurrentapprvdata = [...recordsWhereUserIsGroupMember, ...recordsWithCurrentUser];
                //console.log("fiiiiiiiinnnnaaaaaaa", onlyusercurrentapprvdata);

                const onlyuseractiontakendata = onlyusertransactiondata.filter((item) => item.AllApprovers?.includes(currentUser.Email));
                //console.log("actiontaken", onlyuseractiontakendata);

                const finalfiltereddata = [...onlyusercreateddata, ...onlyusercurrentapprvdata, ...onlyuseractiontakendata];
                //console.log("ffff", finalfiltereddata)
                //for creator,approver same that record occurs twice so remove such duplicates
                const uniqueById = Array.from(
                    new Map(finalfiltereddata.map(item => [item.ID, item])).values()
                );
                //console.log("hhhh", uniqueById);
                setOnlyuserData(uniqueById);
                //console.log("kkkk", onlyuserData);

            }
        }
        void onlyuser();
    }, [workflowData, CurrentUserEmail]);

    // Handle filtering logic
    useEffect(() => {
        let filtered = workflowData;
        if (CurrentUserEmail === "OnlyUser") {
            filtered = onlyuserData;
        }
        if (selectedAppNames.length > 0) {
            filtered = filtered.filter((item) => selectedAppNames.includes(item.AppName.trim()));
        }

        if (selectedStatuses.length > 0) {
            filtered = filtered.filter((item) => selectedStatuses.includes(item.Status.trim()));
        }
        if (selectedFromDate) {
            filtered = filtered.filter((item) => {
                const itemDate = new Date(item.Created); // adjust based on your date field
                //console.log("share1", itemDate)
                return itemDate >= selectedFromDate;
            });
            //console.log("fromfilteeeeeeered", filtered)
        }
        if (selectedToDate) {
            filtered = filtered.filter((item) => {
                const itemDate = new Date(item.Created);
                //console.log("share2", itemDate)
                return itemDate <= selectedToDate;
            });
            //console.log("tofilteeeeeeered", filtered)
        }
        setFilteredData(filtered)

        const parentIds = (filtered || [])
            .filter((row: any) => String(row.CurrentAppCode || "").trim() !== "")
            .map((row: any) => row.ID);

        if (!parentIds || parentIds.length === 0) {
            setChildData([]);
        } else {
            const children = (workflowData || []).filter((w: any) =>
                parentIds.includes(w.ParentAppId)
            );
            setChildData(children);
        }

    }, [selectedAppNames, selectedStatuses, workflowData, selectedFromDate, selectedToDate, onlyuserData]);



    useEffect(() => {
        //console.log("location in reports<<<<<<<<<990876rtyghutyjdgcf", location)
        const fetchData = async () => {
            try {
                await fetchWorkflowData();
            } catch (error) {
                console.error("Error in fetchData:", error);
            }
        };
        fetchFormTitles().catch((error) => {
            console.error("Error in fetchFormTitles:", error);
        });
        void fetchData();
    }, [CurrentUserEmail]);

    // const downloadCSV = () => {
    //     if (filteredData.length === 0) {
    //         console.warn("No data available for export.");
    //         return;
    //     }
    //     const allKeys = Object.keys(filteredData[0]);
    //     const selectedKeys = allKeys.slice(8);
    //     selectedKeys.push('Created By');

    //     const formattedData = filteredData.map(item => {
    //         let newItem: any = {};

    //         selectedKeys.forEach(key => {
    //             if (key.toLowerCase().includes("date") || key.toLowerCase().includes("created") || key.toLowerCase().includes("modified")) {
    //                 newItem[key] = new Date(item[key]).toLocaleDateString('en-GB', {
    //                     day: '2-digit',
    //                     month: 'short',
    //                     year: 'numeric'
    //                 });
    //             } else if (key === "Created By") {
    //                 newItem[key] = item[key] || 'Unknown';
    //             } else {
    //                 newItem[key] = item[key];
    //             }
    //         });

    //         return newItem;
    //     });

    //     const csv = Papa.unparse(formattedData);
    //     const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    //     const link = document.createElement("a");
    //     link.href = URL.createObjectURL(blob);
    //     link.setAttribute("download", "WorkFlowProcessData.csv");
    //     document.body.appendChild(link);
    //     link.click();
    //     document.body.removeChild(link);
    // };


    // Count the statuses and prepare chart data


    const countStatuses = (data: any[]) => {
        const statusCount: { [key: string]: number } = {
            "In-Progress": 0,
            "Rejected": 0,
            "Completed": 0,
            // "Draft": 0,
            "Returned": 0,
            "Cancelled": 0
        };

        data.forEach(item => {
            if (statusCount[item.Status] !== undefined) {
                statusCount[item.Status]++;
            }
        });

        return statusCount;
    };

    // Prepare chart data with counts for each status
    const statusCount = countStatuses(filteredData);
    //console.log("statuscount", statusCount)
    debugger;
    const chartData = {
        labels: ['In-Progress', 'Rejected', 'Completed', 'Returned', 'Cancelled'],
        datasets: [
            {
                data: [
                    statusCount["In-Progress"],
                    statusCount["Rejected"],
                    statusCount["Completed"],
                    // statusCount["Draft"],
                    statusCount["Returned"],
                    statusCount["Cancelled"]
                ],
                backgroundColor: [
                    '#42A5F5',
                    '#EF5350',
                    '#66BB6A',
                    '#BFBFBF',
                    '#FFA726',
                    '#B0BEC5'
                ],
                // hoverBackgroundColor: [
                //     '#FFC107',
                //     '#D32F2F',
                //     '#388E3C',
                //     '#9E9E9E',
                //     '#5E22B1',
                //     '#D32F2F'
                // ],
            },
        ],

    };

    const chartOptions = {
        plugins: {
            legend: {
                display: false
            }
        },
        cutout: '70%',
    };

    const applyFilters = () => {
        if (selectedAppNames.length > 0 || selectedStatuses.length > 0) {
            setShowChart(true);
        } else {
            setShowChart(false);
        }
    };



    const countMonthlyStatuses = (data: any[]) => {

        const monthlyStatusCount: { [key: string]: number[] } = {
            "In-Progress": Array(12).fill(0),
            "Rejected": Array(12).fill(0),
            "Completed": Array(12).fill(0),
            // "Draft": Array(12).fill(0),
            "Returned": Array(12).fill(0),
            "Cancelled": Array(12).fill(0)
        };

        data.forEach(item => {
            const createdDate = new Date(item.Created);
            const monthIndex = createdDate.getMonth(); // Get month (0 for Jan, 1 for Feb, etc.)

            if (monthlyStatusCount[item.Status]) {
                monthlyStatusCount[item.Status][monthIndex]++;
            }
        });

        return monthlyStatusCount;
    };

    const monthlyStatusData = countMonthlyStatuses(filteredData);

    const onTitleClick = (rowData: any) => {
        // setLoading(true)

        //console.log("status--------------------------------------11111111", rowData)
        if (rowData.Status === "Returned") {
            navigate(`/EditForm/${rowData.ID}`, {
                state: {
                    filters: {
                        selectedAppNames,
                        selectedStatuses
                    },
                    ...location.state,
                }
            });
        }
        else {
            navigate(`/viewform/${rowData.ID}`, {
                state: {
                    filters: {
                        selectedAppNames,
                        selectedStatuses
                    },
                    ...location.state,
                }
            });
        }
    }

    const exportToExcelMultiSheet = (sheetDataMap: Record<string, any[]>) => {
        const workbook = XLSX.utils.book_new();

        for (const [sheetName, data] of Object.entries(sheetDataMap)) {
            const worksheet = XLSX.utils.json_to_sheet(data);
            //console.log("sheetame", sheetName, sheetDataMap)
            XLSX.utils.book_append_sheet(workbook, worksheet, sheetName.substring(0, 31).replace(/[\\[\]/*:?-]/g, '')); // sheet name max length is 31
        }

        XLSX.writeFile(workbook, `FilteredData_${new Date().toISOString().slice(0, 10)}.xlsx`);
    };


    const handlePdfGenerate = async () => {
        //console.log(selectedAppNames, "AND", filteredData);

        if (!selectedAppNames || selectedAppNames.length === 0) {
            toast.current?.show({
                severity: 'error',
                summary: '',
                detail: 'Please select at least one App Name before downloading',
                life: 3000,
                style: { height: 60 }
            });
            return;
        }

        const Domain = await fetchTenantUser();

        // Get all form definitions
        const RAWformMasterItems = await sp.web.lists
            .getByTitle("FormMaster")
            .items.select("AppName", "Domain", "FormJSON", "IsParentForm", "AppCode")
            .filter(`Domain eq '${Domain.TenantUsers}'`)
            .getAll();
        const formMasterItems = RAWformMasterItems?.map((item: any) => {
            return { ...item, FormJSON: dataserviceobj.decryptjson(item.FormJSON) };
        })

        const workbookSheets: Record<string, any[]> = {};
        let appsToExport = [...selectedAppNames];
        for (const selName of selectedAppNames) {
            // always include the selected parent name
            const matchedParent = formMasterItems.find(item => item.AppName === selName);
            if (!matchedParent) continue;

            if (matchedParent.IsParentForm === true && matchedParent.AppCode) {
                try {
                    const RAWchildItems = await sp.web.lists
                        .getByTitle("FormMaster")
                        .items.select("AppName", "ParentAppCode", "FormJSON", "Domain")
                        .filter(`Domain eq '${Domain.TenantUsers}' and ParentAppCode eq '${matchedParent.AppCode}'`)
                        .getAll();
                    const childItems = RAWchildItems?.map((item: any) => {
                        return { ...item, FormJSON: dataserviceobj.decryptjson(item.FormJSON) };
                    })
                    const childAppNames = (childItems || []).map((c: any) => (c.AppName || '').toString().trim()).filter(Boolean);

                    for (const childName of childAppNames) {
                        if (!appsToExport.includes(childName)) appsToExport.push(childName);
                    }
                } catch (e) {
                    console.warn("Error fetching child items by ParentAppCode", matchedParent.AppCode, e);
                    // continue - still export parent even if fetching children failed
                }
            }
        }

        for (const appName of appsToExport) {
            const matchedForm = formMasterItems.find(item => item.AppName === appName);
            if (!matchedForm) continue;

            const formDefinition = JSON.parse(matchedForm.FormJSON); // strucuturejson of each form which contains vlaues of checkboxes, radiobuttons and radio buttons
            // const appFilteredData = filteredData.filter(item => item.AppName === appName);
            const combined = [...filteredData, ...childData];
            const appFilteredData = combined.filter(item => (item.AppName || '').toString().trim() === appName);
            const formattedData = appFilteredData.map((entry) => {
                const parsedFormData = JSON.parse(entry.FormData);
                //console.log("parsedFormData in form reports", parsedFormData) // answer json which doesnt contain actual value 
                const obj: Record<string, any> = {};

                obj["RequestId"] = entry.Title || "";
                obj["Author"] = entry.Author?.Title || "";
                obj["Status"] = entry.Status || "";
                obj["AppName"] = entry.AppName || "";
                obj["CreatedDate"] = new Date(entry.Created).toLocaleDateString("en-US", {
                    year: "numeric",
                    month: "2-digit",
                    day: "2-digit"
                }) || "";

                formDefinition.forEach((field: any) => {
                    if (field.element !== "Header" && field.element !== "LineBreak" && field.element !== "TwoColumnRow" && field.element !== "ThreeColumnRow" && field.element !== "MultiColumnRow") {
                        const fieldData = parsedFormData.find((f: any) => f.id === field.id);
                        //console.log("fieldData in form reports", fieldData)
                        const rawLabel = field.label || field.text || "Unnamed Field";
                        debugger;
                        const cleanLabel = typeof rawLabel === "string"
                            ? rawLabel.replace(/<[^>]+>/g, "").trim()
                            : "Unnamed Field";

                        let value = "";
                        if (fieldData) {
                            const fieldValue = fieldData.value;

                            if (Array.isArray(fieldValue)) {

                                if (field.element === "Checkboxes" || field.element === "RadioButtons") {
                                    value = fieldValue.map((item: any) => {
                                        return field.options.find((option: any) => option.key === item)?.text || '';
                                    }).join(", ");
                                }
                                // Handle checkbox/multiselect with objects like [{label: "Option 1"}, {label: "Option 2"}] 
                            }
                            else if (field.element === "Dropdown") {
                                value = fieldValue.trim() == "" ? "" : field.options.find((option: any) => option.value === fieldValue)?.text || '';
                            }
                            else if (typeof fieldValue === "object" && fieldValue !== null) {
                                // Handle single dropdown object
                                value = fieldValue.label || fieldValue.value || '';
                            } else {
                                // Simple string/number/etc.
                                value = fieldValue;
                            }
                        }
                        obj[cleanLabel] = value;
                    }
                });

                return obj;
            });

            workbookSheets[appName] = formattedData;
        }

        exportToExcelMultiSheet(workbookSheets);
    };

    const handleUserEmailRetrieved = (UserAUTH: string) => {
        setCurrentUserEmail(UserAUTH);
        //console.log("Logged-in USER IS from formreports  ----------------:", UserAUTH);

    };


    return (
        <div className='formreport' style={{ display: 'flex', flexDirection: 'column' }}>
            <Toast ref={toast} />
            <TopNavBar onUserTypeRetrieved={handleUserEmailRetrieved} />
            <div style={{ display: 'flex', backgroundColor: '#070D19', minWidth: '100%', maxWidth: '100%' }}>
                <Sidebar activeMenu="Generate Reports" />
                <div className="custom-container">
                    <section style={{ backgroundColor: '#070D19', width: '100%' }}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '7px', marginLeft: '15px', height: '30px' }}>
                            {/* <img onClick={() => navigate(-1)} style={{ height: 23, width: 23, cursor: "pointer" }} src={require('../../../assets/Images/previous.png')} alt="backicon" title='Back' /> */}
                            <h5 className="fo-ap" style={{ margin: '0', color: '#FFFFFF', fontSize: 16 }}>Generate Reports</h5>
                            {/* <p style={{ alignSelf: 'end', marginLeft:'auto',color:'#FFFFFF', paddingRight:'10px', marginBottom:'12px' }}>Total {Object.values(statusCount).reduce((total, count) => total + count, 0)} Items</p>  */}
                            <div style={{ display: 'flex', marginLeft: 'auto' }}>
                                <div>
                                    <img style={{ marginRight: "10px", width: '14.06px', height: '16px', color: '#FFFFFF', paddingBottom: '2px' }} src={totalitemsicon} />
                                </div>
                                <span style={{ alignSelf: 'end', marginLeft: 'auto', color: '#FFFFFF', paddingRight: '10px' }}>Total {filteredData.length} Items</span>
                            </div>
                        </div>
                    </section>
                    <div className="custom-flex-container">
                        <div className="custom-inner-row">
                            <MultiSelect
                                ref={multiSelectRef}
                                value={tempSelectedAppNames}
                                options={appNames.map((name) => ({ label: name.trim(), value: name.trim() }))}
                                onChange={(e) => {
                                    setTempSelectedAppNames(e.value);
                                    setSelectedAppNames(e.value);
                                    applyFilters();
                                }}
                                placeholder="Filter by App Name"
                                className="custom-dropdown mrgn-btm"
                                display="chip"
                                style={{ width: '180px', height: '40px', fontSize: '12px', backgroundColor: '#070D19' }}

                                panelHeaderTemplate={() => {
                                    return (
                                        <div
                                            className='dropdownheader'
                                        // style={{
                                        //     width: '100%',
                                        //     display: 'flex',
                                        //     justifyContent: 'flex-start',
                                        //     alignItems: 'center',
                                        //     padding: '8px 12px',
                                        //     borderBottom: '1px solid #ccc',
                                        //     backgroundColor: '#6A6E79',
                                        // }}
                                        >
                                            <div>
                                                <Checkbox
                                                    inputId="selectAllCheckbox"
                                                    checked={allSelectedAppnames}
                                                    onChange={(e) => {
                                                        //console.log("filtered dtat", filteredData);
                                                        if (e.checked) {
                                                            setTempSelectedAppNames(appNames);
                                                            setSelectedAppNames(appNames);
                                                        } else {
                                                            setTempSelectedAppNames([]);
                                                            setSelectedAppNames([]);
                                                        }
                                                    }
                                                    }
                                                />
                                            </div>
                                            <div>
                                                <label htmlFor="selectAllCheckbox" className="selectall"
                                                // style={{ paddingLeft: '10px', color: '#ffffff' }}
                                                >Select All</label>
                                            </div>

                                        </div>
                                    )
                                }}
                                panelFooterTemplate={() => (
                                    <div
                                        className='dropdownfooter'
                                    // style={{ textAlign: 'right', padding: '8px', backgroundColor: '#6A6E79' }}
                                    >
                                        <Button
                                            className='footerbutton'
                                            label="OK"
                                            onClick={() => {
                                                setSelectedAppNames(tempSelectedAppNames);
                                                applyFilters();
                                                multiSelectRef.current?.hide();
                                            }}
                                            size="small"
                                        />
                                    </div>
                                )}
                            />
                            <MultiSelect
                                ref={multiSelectRef}
                                value={tempSelectedStatuses}
                                options={statuses.map((status) => ({ label: status, value: status }))}
                                onChange={(e) => { setTempSelectedStatuses(e.value); applyFilters(); setSelectedStatuses(e.value) }}
                                placeholder="Filter by Status"
                                className="custom-dropdown mrgn-btm"
                                display="chip"
                                style={{ width: '180px', height: '40px', fontSize: '12px', marginLeft: '5px', backgroundColor: '#070D19' }}

                                panelHeaderTemplate={() => {
                                    return (
                                        <div
                                            className='dropdownheader'
                                        // style={{
                                        //     display: 'flex',
                                        //     justifyContent: 'flex-start',
                                        //     alignItems: 'center',
                                        //     padding: '8px 12px',
                                        //     borderBottom: '1px solid #ccc',
                                        //     backgroundColor: '#6A6E79',
                                        // }}
                                        >
                                            <div>
                                                <Checkbox
                                                    inputId="selectAllCheckbox"
                                                    checked={allSelectedStatuses}
                                                    onChange={(e) => {
                                                        if (e.checked) {
                                                            setTempSelectedStatuses(statuses);
                                                            setSelectedStatuses(statuses);
                                                        } else {
                                                            setTempSelectedStatuses([]);
                                                            setSelectedStatuses([]);
                                                        }
                                                    }

                                                    }
                                                />
                                            </div>
                                            <div>
                                                <label htmlFor="selectAllCheckbox" className="selectall"
                                                // style={{ paddingLeft: '10px', color: '#ffffff' }}
                                                >Select All</label>
                                            </div>


                                        </div>
                                    )
                                }}
                                panelFooterTemplate={() => (
                                    <div
                                        className='dropdownfooter'
                                    // style={{ textAlign: 'right', padding: '8px', backgroundColor: '#6A6E79' }}
                                    >
                                        <Button
                                            className='footerbutton'
                                            label="OK"
                                            onClick={() => {
                                                setSelectedStatuses(tempSelectedStatuses);
                                                applyFilters();
                                                multiSelectRef.current?.hide();
                                            }}
                                            size="small"
                                        />
                                    </div>
                                )}
                            />


                            <Button className="mrgn-btm" style={{ display: 'flex', flexDirection: 'row', width: '110px', height: '40px', borderRadius: '4px', borderColor: '#4A4A4A', backgroundColor: '#070D19', textAlign: "center", marginLeft: '6px' }}
                                onClick={() => {
                                    setSelectedAppNames([]);
                                    setTempSelectedAppNames([])
                                    setSelectedStatuses([]);
                                    setTempSelectedStatuses([]);
                                    setSelectedFromDate(null);
                                    setSelectedToDate(null);

                                }}
                            >
                                <FontAwesomeIcon className="resetFilter" icon={faRefresh} title="Reset"
                                    style={{ marginRight: '10px', marginLeft: '0px', width: '14.06px', height: '16px', color: '#FFFFFF' }}
                                    onClick={() => {
                                        setSelectedAppNames([]);
                                        setTempSelectedAppNames([])
                                        setSelectedStatuses([]);
                                        setTempSelectedStatuses([]);
                                        setSelectedFromDate(null);
                                        setSelectedToDate(null);

                                    }} />
                                <p style={{ fontSize: '12px', marginTop: '16px' }}>Refresh</p>
                            </Button>
                            <Button className="export-button mrgn-btm" onClick={handlePdfGenerate}>
                                <img className="export-icon" src={excelicon} alt="Excel Icon" />
                                <p className="export-text">Export to Excel</p>
                            </Button>
                            <Button className="toggle-chart-button" onClick={() => {
                                setShowLineChart(prevState => !prevState);
                                //console.log("showLineChart toggled");
                            }}>
                                <div className="toggle-chart-content">
                                    <img
                                        className="toggle-chart-icon"
                                        src={showLineChart ? graphicon : tableicon}
                                        alt="Toggle Icon"
                                    />
                                    <span className="toggle-chart-text">Switch</span>
                                </div>
                            </Button>


                            <div className="date-picker-container">

                                <div className="calendarboxdiv" >
                                    <img src={calendaricn} className="calimgicon" />

                                    <Calendar
                                        value={selectedFromDate}
                                        className="calcontrol"
                                        placeholder="From Date"
                                        //    showIcon
                                        showButtonBar
                                        onChange={(e) => {
                                            if (!e.value) {
                                                setSelectedFromDate(e.value);
                                                //console.log("afterclosing", e.value)
                                                return;
                                            }
                                            const fromdate = new Date(e.value);
                                            fromdate.setHours(0, 0, 0, 0); //by default e.value is 0,0,0,0 anyway. set date to beginning of day 
                                            setSelectedFromDate(fromdate)
                                            //console.log("firstcal", e.value)
                                        }}
                                    />
                                </div>

                                <div className="calendarboxdiv" style={{ marginLeft: '10px' }}>
                                    <img src={calendaricn} className="calimgicon" />

                                    <Calendar
                                        value={selectedToDate}
                                        className="calcontrol"
                                        placeholder="To Date"
                                        // showIcon
                                        showButtonBar
                                        minDate={selectedFromDate || undefined}
                                        onChange={(e) => {
                                            if (!e.value) {
                                                setSelectedToDate(e.value)
                                                return;
                                            }
                                            const todate = new Date(e.value);
                                            todate.setHours(23, 59, 59, 999);
                                            setSelectedToDate(todate)
                                            //console.log("seccal", e.value)
                                        }}
                                    />
                                </div>

                            </div>
                        </div>
                        <div className="custom-card flex justify-content-center mb-4">
                            {!showLineChart ? (
                                <div className="chart-container">
                                    <MyChartComponent
                                        months={["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]}
                                        inProgressData={monthlyStatusData["In-Progress"]}
                                        rejectedData={monthlyStatusData["Rejected"]}
                                        completedData={monthlyStatusData["Completed"]}
                                        returnedData={monthlyStatusData["Returned"]}
                                        cancelledData={monthlyStatusData["Cancelled"]}
                                    />

                                    <Chart
                                        type="doughnut"
                                        data={chartData}
                                        options={chartOptions}
                                        className="doughnut-chart"
                                    />
                                </div>

                            ) :
                                (
                                    <div style={{ height: '100%', width: '100%' }}>
                                        <DataTable value={filteredData}
                                            paginator
                                            scrollable
                                            scrollHeight="420px"
                                            paginatorTemplate="CurrentPageReport FirstPageLink PrevPageLink PageLinks NextPageLink LastPageLink RowsPerPageDropdown"
                                            currentPageReportTemplate="Showing {first} to {last} of {totalRecords}"
                                            sortField="Created"
                                            sortOrder={-1}
                                            rows={10}
                                            rowsPerPageOptions={[5, 10, 15]}
                                            dataKey="ID"
                                            emptyMessage="No transactions found."
                                            onRowClick={(e) => onTitleClick(e.data)}>
                                            <Column field="Title" header="RequestID" sortable />
                                            <Column field="Status" header="Status" sortable />
                                            <Column field='AppName' header="Form Name" sortable
                                                body={(rowData) => (
                                                    <span title={rowData.AppName}>
                                                        {rowData.AppName}
                                                    </span>
                                                )} />
                                            <Column field="CategoryType" header="Request For" sortable
                                                body={(rowData) => (
                                                    <span title={rowData.CategoryType}>
                                                        {rowData.CategoryType}
                                                    </span>
                                                )}
                                            />

                                            <Column field="FinalCost" header="Final Cost" sortable />
                                            <Column
                                                field="Author"
                                                header="Created By"
                                                sortable
                                                body={(rowData) => {
                                                    if (rowData.Author && rowData.Author.Title) {
                                                        const title = rowData.Author.Title.split('|')[0];
                                                        return title || 'N/A';
                                                    }
                                                    return 'N/A';
                                                }}
                                            />
                                            <Column
                                                field="CurApprover"
                                                header="Current Approver"
                                                sortable
                                                body={(rowData) => {
                                                    if (rowData.CurApprover && rowData.CurApprover.Title) {
                                                        const approverTitle = rowData.CurApprover.Title.split('|')[0];
                                                        return approverTitle || ' ';
                                                    }
                                                    return ' '; // If CurApprover or Title is not available, return 'N/A'
                                                }}
                                            />

                                            <Column
                                                field="Created"
                                                header="Created Date & Time"
                                                sortable
                                                body={(rowData) => {
                                                    const date = new Date(rowData.Created);
                                                    return date.toLocaleDateString('en-GB', {
                                                        day: '2-digit',
                                                        month: 'short',
                                                        year: 'numeric',
                                                        hour: '2-digit',
                                                        minute: '2-digit',
                                                        hour12: false
                                                    });
                                                }} />
                                        </DataTable>
                                    </div>
                                )
                            }
                        </div>
                        {/* </div> */}
                    </div>
                    {/* )} */}
                    {/* <DataTable value={filteredData}
                    paginator

                    paginatorTemplate="CurrentPageReport FirstPageLink PrevPageLink PageLinks NextPageLink LastPageLink RowsPerPageDropdown"
                    currentPageReportTemplate="Showing {first} to {last} of {totalRecords}"
                    rows={10}
                    rowsPerPageOptions={[10, 15, 20]}
                    dataKey="ID"
                    emptyMessage="No transactions found."
                    onRowClick={(e) => onTitleClick(e.data)}>
                    <Column field="Title" header="RequestID" sortable />
                    <Column field="Status" header="Status" sortable />
                    <Column field='AppName' header="Form Name" sortable />
                    <Column field="CategoryType" header="Request For" sortable />
                    <Column field="FinalCost" header="Final Cost" sortable />
                    <Column
                        field="Author"
                        header="Created By"
                        sortable
                        body={(rowData) => {
                            if (rowData.Author && rowData.Author.Title) {
                                const title = rowData.Author.Title.split('|')[0];
                                return title || 'N/A';
                            }
                            return 'N/A';
                        }}
                    />
                    <Column
                        field="CurApprover"
                        header="Current Approver"
                        sortable
                        body={(rowData) => {
                            if (rowData.CurApprover && rowData.CurApprover.Title) {
                                const approverTitle = rowData.CurApprover.Title.split('|')[0];
                                return approverTitle || ' ';
                            }
                            return ' '; // If CurApprover or Title is not available, return 'N/A'
                        }}
                    />

                    <Column
                        field="Created"
                        header="Created Date & Time"
                        sortable
                        body={(rowData) => {
                            const date = new Date(rowData.Created);
                            return date.toLocaleDateString('en-GB', {
                                day: '2-digit',
                                month: 'short',
                                year: 'numeric',
                                hour: '2-digit',
                                minute: '2-digit',
                                hour12: false
                            });
                        }} />
                </DataTable> */}
                    {/* </div> */}
                </div>
            </div>
            {/* </div> */}
        </div>
    );
};

export default Reports;