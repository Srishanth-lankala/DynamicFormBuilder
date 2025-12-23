import * as React from 'react';
import { sp } from "@pnp/sp/presets/all";
import TopNavBar from '../TopBar';
import { useLocation, useNavigate } from 'react-router-dom';
import { useEffect, useRef, useState } from 'react';
import { Chart } from 'primereact/chart';
import './FormDashBoard.css'
import { Column } from 'primereact/column';
import { DataTable } from 'primereact/datatable';
import { mySiteUrl } from '../ConfigURL/All_URLs';
import { Button } from 'primereact/button';
import { Calendar } from 'primereact/calendar';
import calendaricn from '../../../assets/Images/calendaricon.svg';
import excelicon from '../../../assets/Images/excelicon.svg';
// import { Link as RouterLink } from 'react-router-dom';
// import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
// import { faRepeat } from '@fortawesome/free-solid-svg-icons';
import LoadingSpinner from '../Loading';
import { fetchTenantUser } from '../FetchTenantUser/fetchTenantUser';
import SideBar from '../Sidebar/SideBar';
import { dataservice } from '../encryptionutil';
// import magnify from '../../../assets/Images/magnifying-glass.png';
import graphicon from '../../../assets/Images/graph.svg';
import tableicon from '../../../assets/Images/tableicon.svg';
import searchIcon from '../../../assets/Images/searchicon.svg';
import TotalRequests from '../../../assets/Images/Staticcardicons/TotalRequests.svg';
import PendingRequests from '../../../assets/Images/Staticcardicons/PendingRequests.svg';
import CompletedRequests from '../../../assets/Images/Staticcardicons/CompletedRequests.svg';
import InprogressRequests from '../../../assets/Images/Staticcardicons/InprogressRequests.svg';

import * as XLSX from "xlsx";
import { Toast } from 'primereact/toast';

interface FormType {
    CategoryType: any;
    AppName: any;
    Title: string;
    Status: string;
    Id: number;
    Created: string;
}
const FormDashBoard = () => {
    const location = useLocation();
    const [appcode, setAppcode] = useState<string>(location.state?.appcode);
    const [appname, setAppname] = useState<string>(location.state?.appname);
    // const firstWord = appname.split(' ')[0]
    const [timesheetbool, setTimesheetbool] = useState<boolean>(false);
    const [formTypes, setFormTypes] = useState<FormType[]>([]);
    const [formmasterdata, setFormmasterdata] = useState<any>([]);
    const [chartData, setChartData] = useState({});
    const [chartOptions, setChartOptions] = useState({});
    // const [approvePercentage, setApprovePercentage] = useState<number>(0);
    // const [totalApprovalDuration, setTotalApprovalDuration] = useState<number>(0);
    const [inProgressCount, setInProgressCount] = useState<number>(0)
    const [aprovedCount, setAprovedCount] = useState<number>(0)
    const [rejectedCount, setRejectedCount] = useState<number>(0)
    const [filteredData, setFilteredData] = useState<any[]>([])
    const [datesfilteredData, setdatesFilteredData] = useState<any[]>([])
    const [displayTable, setDisplayTable] = useState(true)
    const [displayFilterTable, setDisplayFilterTable] = useState(false)
    const [sortDirection, setSortDirection] = useState<'asc' | 'desc'>('asc');
    const [countOfRecords, setCountOfRecords] = useState<number>(0);
    const [showLineChart, setShowLineChart] = useState(true);
    const [CurrentUserEmail, setCurrentUserEmail] = useState<string | null>(null);
    const [domain, setDomain] = useState<any>()
    const navigate = useNavigate();
    const [loading, setLoading] = useState<boolean>(true);
    const [selectedFromDate, setSelectedFromDate] = useState<Date | null | undefined>(null);
    const [selectedToDate, setSelectedToDate] = useState<Date | null | undefined>(null);
    const toast = useRef<Toast>(null);

    // const [curruseradmin,setcurruseradmin] = useState<boolean>(false)

    const dataserviceobj = new dataservice();


    sp.setup({
        sp: {
            baseUrl: mySiteUrl,
        },
    });

    const handleFilter = (status: string) => {
        if (status === 'All') {
            setFilteredData(formTypes); // Reset to show all data
            setdatesFilteredData(formTypes);
        } else {
            setFilteredData(formTypes.filter((item) => item.Status === status));
            setdatesFilteredData(formTypes.filter((item) => item.Status === status));
        }
    };

    const formatDateHeader = (date: Date) =>
        date.toLocaleDateString("en-GB", { day: "2-digit", month: "short" });


    //one tab for each author excel function

    // const exportTimesheetToExcel = (
    //     filteredRecords: any[],
    //     fromDate: Date,
    //     toDate: Date
    // ) => {
    //     const dynamicDateHeaders: string[] = [];
    //     const currentDate = new Date(fromDate);

    //     while (currentDate <= toDate) {
    //         dynamicDateHeaders.push(formatDateHeader(new Date(currentDate)));
    //         currentDate.setDate(currentDate.getDate() + 1);
    //     }

    //     const sheetDataMap: Record<string, any[]> = {}; // key: sheet name (author), value: row data

    //     filteredRecords.forEach((record) => {
    //         const authorName = record?.Author?.Title ?? "Unknown";

    //         let timesheetJson: any[] = [];
    //         try {
    //             timesheetJson = JSON.parse(record?.TimeSheetJSON || "[]");
    //         } catch {
    //             return; // skip malformed record
    //         }

    //         const userRows: any[] = [];

    //         timesheetJson.forEach((sheet) => {
    //             const columns = sheet?.columns || [];
    //             const rows = sheet?.rows || [];

    //             const staticHeaders = columns
    //                 .filter((col: any) => !col.isdynamic)
    //                 .map((col: any) => col.headername);

    //             rows.forEach((row: any) => {
    //                 const excelRow: any = { Author: authorName };

    //                 // Static fields
    //                 staticHeaders.forEach((header: string) => {
    //                     excelRow[header] = row?.[header] ?? "";
    //                 });

    //                 // Dynamic date fields
    //                 dynamicDateHeaders.forEach((dateHeader) => {
    //                     excelRow[dateHeader] = row?.[dateHeader] ?? "";
    //                 });

    //                 userRows.push(excelRow);
    //             });
    //         });

    //         // Append to the author's sheet
    //         if (!sheetDataMap[authorName]) {
    //             sheetDataMap[authorName] = [];
    //         }
    //         sheetDataMap[authorName].push(...userRows);
    //     });

    //     const workbook = XLSX.utils.book_new();

    //     for (const [sheetName, data] of Object.entries(sheetDataMap)) {
    //         const worksheet = XLSX.utils.json_to_sheet(data);
    //         const safeSheetName = sheetName.substring(0, 31).replace(/[\\[\]/*:?:]/g, '');
    //         XLSX.utils.book_append_sheet(workbook, worksheet, safeSheetName);
    //     }

    //     const fileName = `FilteredData_${new Date().toISOString().slice(0, 10)}.xlsx`;
    //     XLSX.writeFile(workbook, fileName);
    // }
    function toLabelValueStrings(items: any[]) {
        return items.reduce((acc, { label, value }) => {
            // Step 1 – normalize the incoming value
            const normalized =
                Array.isArray(value) ? value.join(', ') : String(value);

            // Step 2 – collect duplicates
            if (acc[label]) {
                // Append with a comma-space separator
                acc[label] += `, ${normalized}`;
            } else {
                acc[label] = normalized;
            }

            return acc;
        }, {});
    }
    const stripHtmlTags = (label: string): string =>
        label?.replace(/<\/?[^>]+(>|$)/g, "")?.trim() || "";

    function mergeFormStructureAndAnswers(structure: any[], answers: any[]) {
        return answers
            .map(answer => {
                const match = structure.find(item => item.id === answer.id);
                if (!match || match.element === "FileUpload") return null; // Exclude FileUpload

                const { id, element, label, options } = match;
                const cleanLabel = stripHtmlTags(label || match?.text);

                let value = answer.value;

                // Handle Dropdown (only if non-empty string)
                if (
                    element === "Dropdown" &&
                    typeof value === "string" &&
                    value.trim() !== ""
                ) {
                    const selectedOption = options?.find((opt: any) => opt.value === value);
                    value = selectedOption?.text || value;
                }

                // Handle Checkboxes & RadioButtons
                if (
                    ["Checkboxes", "RadioButtons"].includes(element) &&
                    Array.isArray(value)
                ) {
                    const optionsMap = new Map<string, string>();
                    options?.forEach((opt: any) => optionsMap.set(opt.key, opt.text));
                    value = value.map((val: string) => optionsMap.get(val) || val);
                }

                return {
                    id,
                    element,
                    label: cleanLabel,
                    value
                };
            })
            .filter(Boolean); // Remove nulls (e.g., FileUpload)
    }

    const exportTimesheetToExcel = (
        filteredRecords: any[],
        fromDate: Date,
        toDate: Date
    ) => {
        // Generate dynamic date headers like "07 Jul", "08 Jul", etc.
        const dynamicDateHeaders: string[] = [];
        const currentDate = new Date(fromDate);
        while (currentDate <= toDate) {
            dynamicDateHeaders.push(formatDateHeader(new Date(currentDate)));
            currentDate.setDate(currentDate.getDate() + 1);
        }

        const allRows: any[] = [];
        const FormJSONselected = JSON.parse(formmasterdata.find((item: any) => item.AppCode === appcode)?.FormJSON || "[]");
        filteredRecords.forEach((record) => {
            const authorName = record?.Author?.Title ?? "Unknown";
            let timesheetJson: any[] = [];
            try {
                timesheetJson = JSON.parse(record?.TimeSheetJSON || "[]");
            } catch {
                console.warn("Invalid TimeSheetJSON for record", record);
                return;
            }
            const mergedData = mergeFormStructureAndAnswers(FormJSONselected, JSON.parse(record?.FormData || "[]"));
            const labelValueStrings = toLabelValueStrings(mergedData);
            
            timesheetJson.forEach((sheet) => {
                const columns = sheet?.columns || [];
                const rows = sheet?.rows || [];
                const staticHeaders = columns
                    .filter((col: any) => !col.isdynamic)
                    .map((col: any) => col.headername);

                rows.forEach((row: any) => {
                    const excelRow: any = {
                        Name: authorName,
                        ...labelValueStrings
                    };


                    // Add static fields
                    staticHeaders.forEach((header: any) => {
                        excelRow[header] = row?.[header] ?? "";
                    });

                    // Add dynamic date fields
                    dynamicDateHeaders.forEach((dateHeader) => {
                        excelRow[dateHeader] = row?.[dateHeader] ?? "";
                    });
                    allRows.push(excelRow);
                });
            });
        });

        // Generate Excel workbook with a single sheet
        const workbook = XLSX.utils.book_new();
        const worksheet = XLSX.utils.json_to_sheet(allRows);
        XLSX.utils.book_append_sheet(workbook, worksheet, "Timesheet");

        // Download Excel file
        const fileName = `FilteredData_${new Date().toISOString().slice(0, 10)}.xlsx`;
        XLSX.writeFile(workbook, fileName);
    }



    const onTitleClick = (rowData: any) => {
        debugger;
        navigate(`/viewform/${rowData.ID}`);
    };

    const fetchFormTitles = async () => {
        try {
            const Domain = await fetchTenantUser();


            // Fetch current user details
            const currentUser = await sp.web.currentUser.get();
            //console.log("Current User ID:", currentUser.Id);


            // Fetch all forms where Domain is 'TenantUsers' and FormViewer is either a person or a group
            const RAWallForms = await sp.web.lists.getByTitle("FormMaster")
                .items
                .select("ID", "Title", "FormJSON", "AppName", "AppCode", "FormType", "Domain", "FormViewer/Id", "FormViewer/Title", "VisibilityFlag", "TransactionCount", "TimeSheetControl")
                .expand("FormViewer")
                .filter(`Domain eq '${Domain.TenantUsers}' and VisibilityFlag eq 1 and ChildOrder eq null`)
                .getAll();
                
            const allForms = RAWallForms?.map((item:any)=>{
                    return { ...item, FormJSON: dataserviceobj.decryptjson(item.FormJSON)};
                })

            //console.log("Fetched Forms:", allForms);
            //console.log("CurrentUserEmail------------- Isssssss", CurrentUserEmail)

            if (CurrentUserEmail === "BothUser Admin" || CurrentUserEmail === "OnlyAdmin") {

                // const uniqueAppNames = [...new Set(allForms.map((item) => item.AppName.trim()))];
                setFormmasterdata(allForms);
                setTimesheetbool(allForms.find((item: any) => item.AppCode === appcode)?.TimeSheetControl || false)
                //console.log("fetchTenantUser------------- IN  IFFFF", Domain)
            }
            else {
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
                // const userGroupTitles = new Set(userGroups.map(group => group.Title));

                //console.log("User's Group IDs:", [...userGroupIds]);
                // console.log("User's Group Titles:", [...userGroupTitles]);

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
                // const uniqueAppNames = [...new Set(finalFilteredForms.map((item) => item.AppName.trim()))];
                setFormmasterdata(finalFilteredForms);
                setTimesheetbool(finalFilteredForms.find((item: any) => item.AppCode === appcode)?.TimeSheetControl || false)
            }

        } catch (error) {
            console.error("Error fetching forms:", error);
        }
        // finally {

        // }
    };
    useEffect(() => {
        const firstitemappcode = formmasterdata.map((item: any) => {
            return item.AppCode;
        });
        const firstitemappname = formmasterdata.map((item: any) => {
            return item.AppName;
        });
        //    console.log("appppp",firstitem)
        if (location.state?.appcode && location.state?.appname) return;

        setAppcode(firstitemappcode[0]);
        setAppname(firstitemappname[0]);
    }, [formmasterdata])

    useEffect(() => {
        // console.log("location in reports<<<<<<<<<990876rtyghutyjdgcf", location)

        fetchFormTitles().catch((error) => {
            console.error("Error in fetchFormTitles:", error);
        });

    }, [CurrentUserEmail, domain]);

    useEffect(() => {
        setTimesheetbool(formmasterdata.find((item: any) => item.AppCode === appcode)?.TimeSheetControl || false)
    }, [appcode, appname]);

    const fetchFormTypes = async () => {
        if (!CurrentUserEmail || !appcode) return;
        try {
            let types = [];

            const currentUser = await sp.web.currentUser.get();
            const currentUserEmail = currentUser.Email;
            //console.log(currentUser)
            if (CurrentUserEmail === "OnlyUser") {
                const items = await sp.web.lists
                    .getByTitle("WorkFlowProcessData")
                    .items
                    .orderBy("Created", false)
                    .filter(`AppCode eq '${appcode}' and Domain eq '${domain.TenantUsers}'`)
                    .select(
                        "Id", "Title", "Status", "Created", "Modified", "AppName", "CategoryType", "FormData",
                        "Author/Title", "CurApprover/Title", "Author/Id", "AllApprovers", "MinDateArray", "MaxDateArray", "TimeSheetJSON", "Author/EMail"
                    )
                    .expand("Author", "CurApprover")
                    .get();

                types = items.filter(
                    item =>
                        item.Author?.Id === currentUser.Id ||
                        item.AllApprovers?.includes(currentUserEmail) ||
                        item.CurApprover?.Title === currentUser.Title,
                );


            }
            else {
                types = await sp.web.lists
                    .getByTitle("WorkFlowProcessData")
                    .items
                    .orderBy("Modified", false)
                    .filter(`AppCode eq '${appcode}' and Domain eq '${domain.TenantUsers}'`)
                    .select("Id", "Title", "Status", "Created", "Modified", "AppName", "CategoryType", "FormData", "Author/Title", "CurApprover/Title", "Author/Id", "Author/EMail", "MinDateArray", "MaxDateArray", "TimeSheetJSON")
                    .expand("Author", "CurApprover")
                    .get();
                //console.log("in formdashboard else")
            }
            //console.log("typeeeeeeeeee-----------", types, CurrentUserEmail, appcode)
            const inProgressCount = types.filter(item => item.Status === 'In-Progress').length;
            const approvedCount = types.filter(item => item.Status === 'Completed').length;
            const rejectedCount = types.filter(item => item.Status === 'Rejected').length;
            const totalCount = types.length;
            //console.log(inProgressCount, approvedCount, rejectedCount)
            setInProgressCount(inProgressCount)
            setAprovedCount(approvedCount)
            setRejectedCount(rejectedCount)
            setCountOfRecords(totalCount)
            //console.log("rejectedCount---->", types)


            //console.log("Fetched Form Types:", types);
            setFormTypes(types);
            setFilteredData(types)
            setdatesFilteredData(types)
            const totalDurationInMilliseconds = types
                .filter(item => item.Status === 'Completed')
                .reduce((acc, item) => {
                    const createdDate = new Date(item.Created);
                    const modifiedDate = new Date(item.Modified);
                    const durationInMilliseconds = Math.abs(modifiedDate.getTime() - createdDate.getTime());
                    return acc + durationInMilliseconds; // Accumulate total duration in milliseconds
                }, 0);

            let totalDurationInHours = totalDurationInMilliseconds / (1000 * 60 * 60);
            totalDurationInHours = Math.floor(totalDurationInHours);
            // setTotalApprovalDuration(totalDurationInHours);

            //console.log("Total Approval Duration in Hours:", totalDurationInHours);
            const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
            const monthData = months.reduce((acc, month) => {
                acc[month] = { inProgress: 0, approved: 0, rejected: 0, total: 0 };
                return acc;
            }, {} as { [key: string]: { inProgress: number; approved: number; rejected: number; total: number } });
            types.forEach(item => {
                const createdDate = new Date(item.Created);
                const createdMonth = months[createdDate.getMonth()];
                if (monthData[createdMonth]) {
                    monthData[createdMonth].total++;
                    if (item.Status === 'In-Progress') {
                        monthData[createdMonth].inProgress++;
                    } else if (item.Status === 'Completed') {
                        monthData[createdMonth].approved++;
                    }
                    else if (item.Status === 'Rejected') {
                        monthData[createdMonth].rejected++;
                    }
                }
            });
            const inProgressData = months.map(month => monthData[month].inProgress);
            const approvedData = months.map(month => monthData[month].approved);
            const totalData = months.map(month => monthData[month].total);
            const rejectedData = months.map(month => monthData[month].rejected);
            const data = {
                labels: months,
                datasets: [
                    {
                        label: 'In-Progress',
                        data: inProgressData,
                        fill: false,
                        backgroundColor: "#42A5F5",
                    },
                    {
                        label: 'Rejected',
                        data: rejectedData,
                        fill: false,
                        backgroundColor: "#EF5350"
                    },
                    {
                        label: 'Completed',
                        data: approvedData,
                        fill: false,
                        backgroundColor: "#66BB6A"
                    },
                    {
                        label: 'Total Records',
                        data: totalData,
                        fill: false,
                        backgroundColor: "#FFA726",
                    },
                ]
            };
            const options = {
                responsive: true,
                // maintainAspectRatio:false,
                plugins: {
                    legend: {
                        position: 'right',
                        labels: {
                            color: 'white',
                            padding: 20,
                            font: {
                                size: 12,
                                weight: 300
                            }
                        }
                    },
                    tooltip: {
                        bodyFont: {
                            size: 12
                        },
                        titleFont: {
                            size: 12
                        }
                    }
                },
                scales: {
                    x: {
                        stacked: true,
                        ticks: {
                            color: '#ffffff',
                            font: {
                                size: 12
                            }
                        },
                        grid: {
                            color: '#444444'
                        }
                    },
                    y: {
                        stacked: true,
                        suggestedMax: 90,
                        ticks: {
                            color: '#ffffff',
                            font: {
                                size: 12
                            }
                        },
                        grid: {
                            color: '#444444'
                        }
                    }
                }
            };

            setChartData(data);
            setChartOptions(options);
        } catch (error) {
            console.error("Error fetching form types:", error);
        } finally { setLoading(false) }
    };


    const currentdomainfetch = async () => {
        const domain = await fetchTenantUser();
        setDomain(domain)
    }

    useEffect(() => {
        void currentdomainfetch()
        // fetchFormTypes().catch(error => {
        //     console.error("Error in fetchFormTypes:", error);

        // });

    }, [CurrentUserEmail]);
    useEffect(() => {
        fetchFormTypes().catch(error => {
            console.error("Error in fetchFormTypes:", error);

        });
    }, [domain, appcode]);

    // function filterRecordsByDateOverlap(
    //     records: any[],
    //     selectedFromDate: string,
    //     selectedToDate: string
    // ): any[] {
    //     const userFrom = new Date(selectedFromDate);
    //     const userTo = new Date(selectedToDate);

    //     return records.filter(record => {
    //         const { FromDateArray, ToDateArray } = record;

    //         for (let i = 0; i < FromDateArray.length; i++) {
    //             const recordFrom = new Date(FromDateArray[i]);
    //             const recordTo = new Date(ToDateArray[i]);

    //             const isOverlap =
    //                 userFrom <= recordTo && userTo >= recordFrom;

    //             if (isOverlap) return true;
    //         }

    //         return false;
    //     });
    // }

    function filterRecordsByDateOverlap(
        records: any[],
        selectedFromDate: Date,
        selectedToDate: Date
    ): any[] {
        //console.log("records passed into function", records)
        return records.filter(record => {
            let FromDateArray: string[] = [];
            let ToDateArray: string[] = [];

            try {
                FromDateArray = JSON.parse(record.MinDateArray || '[]');
                ToDateArray = JSON.parse(record.MaxDateArray || '[]');
            } catch (err) {
                console.warn("Invalid date array format in record:", record);
                return false;
            }

            for (let i = 0; i < FromDateArray.length; i++) {
                const recordFrom = new Date(FromDateArray[i]);
                const recordTo = new Date(ToDateArray[i]);

                const isOverlap =
                    selectedFromDate <= recordTo && selectedToDate >= recordFrom;

                if (isOverlap) return true;
            }

            return false;
        });
    }


    const handleSearch = (searchTerm: string) => {
        debugger
        if (!searchTerm) {
            setFilteredData(formTypes); // Reset to full data
            return;
        }

        const filtered = formTypes.filter((item) =>
            item.Title?.toLowerCase().includes(searchTerm.toLowerCase()) ||
            item.Status?.toLowerCase().includes(searchTerm.toLowerCase()) ||
            item.AppName?.toLowerCase().includes(searchTerm.toLowerCase()) ||
            item.CategoryType?.toLowerCase().includes(searchTerm.toLowerCase())
        );
        setFilteredData(filtered);
    };
    const handleUserEmailRetrieved = (UserAUTH: string) => {
        setCurrentUserEmail(UserAUTH);
        //console.log("Logged-in USER IS   ---------------- !!!!!!!!!!!! :", UserAUTH);
    };
    return (
        <div style={{ display: 'flex', flexDirection: "column" }}>
            <Toast ref={toast} />
            <TopNavBar onUserTypeRetrieved={handleUserEmailRetrieved} />
            <div style={{ display: 'flex', backgroundColor: '#070D19' }}>
                <SideBar activeMenu="Form DashBoard" />
                <div style={{ overflowY: "auto", width: "100%" }}>
                    <div className='draft_header_flex align-items-center'>
                        <div className="headerContainer">
                            <div>
                                <h6 className="app_title">{appname}</h6>
                            </div>
                        </div>

                        <div className="form-dropdown-container">
                            <select
                                value={appcode}
                                onChange={(e) => {
                                    const selectedAppCode = e.target.value;
                                    const selectedItem = formmasterdata.find((item: { AppCode: string; }) => item.AppCode === selectedAppCode);
                                    if (selectedItem) {
                                        setAppcode(selectedItem.AppCode);
                                        setAppname(selectedItem.AppName);
                                    }
                                }}
                                className="customSingleDropdown"
                            >
                                <option value="">Select Form</option>
                                {formmasterdata.map((item: { AppCode: string | number | readonly string[] | undefined; AppName: boolean | React.ReactChild | React.ReactFragment | React.ReactPortal | null | undefined; }, index: React.Key | null | undefined) => (
                                    <option key={index} value={item.AppCode}>
                                        {item.AppName}
                                    </option>
                                ))}
                            </select>



                        </div>
                        <div style={{ display: 'flex', paddingRight: '22px' }}>
                            <div className="searchContainer">

                                <input placeholder='Search...'
                                    className="searchBox"
                                    onChange={(e: React.ChangeEvent<HTMLInputElement>) => handleSearch(e.target.value)}
                                />
                                <img src={searchIcon} alt="Search Icon" className="searchIcon" />
                            </div>
                            {
                                timesheetbool && (<div className="date-picker-container">

                                    <div className="calendarboxdivinfrmdsbrd" >
                                        <img src={calendaricn} className="calimgicon" />

                                        <Calendar
                                            value={selectedFromDate}
                                            className="newcalcontrol"
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
                                                if (fromdate && selectedToDate) {

                                                    let recc = filterRecordsByDateOverlap(datesfilteredData, fromdate, selectedToDate);
                                                    setFilteredData(recc)
                                                    //console.log("records filtered", recc)
                                                }
                                            }}
                                        />
                                    </div>

                                    <div className="calendarboxdivinfrmdsbrd" style={{ marginLeft: '10px' }}>
                                        <img src={calendaricn} className="calimgicon" />

                                        <Calendar
                                            value={selectedToDate}
                                            className="newcalcontrol"
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
                                                if (todate && selectedFromDate) {
                                                    let recc = filterRecordsByDateOverlap(datesfilteredData, selectedFromDate, todate);
                                                    setFilteredData(recc)
                                                    //console.log("records filtered", recc)
                                                }
                                            }}
                                        />
                                    </div>
                                    <Button className="new-export-button"
                                        onClick={() => {
                                            if (selectedToDate && selectedFromDate) {
                                                if (filteredData.length > 0) {

                                                    exportTimesheetToExcel(filteredData, selectedFromDate, selectedToDate)
                                                }
                                                else {
                                                    toast.current?.clear();
                                                    toast.current?.show({
                                                        severity: 'error',
                                                        summary: '',
                                                        detail: 'No records found to export',
                                                        life: 3000,
                                                        style: { height: 60 }
                                                    });
                                                }
                                            }
                                            else {
                                                toast.current?.clear();
                                                toast.current?.show({
                                                    severity: 'error',
                                                    summary: '',
                                                    detail: 'Please select from date and to date to download excel',
                                                    life: 3000,
                                                    style: { height: 60 }
                                                });
                                            }
                                        }}>
                                        <img className="export-icon" src={excelicon} alt="Excel Icon" />
                                        <p className="export-text">Export</p>
                                    </Button>

                                </div>)
                            }
                            <Button className="switch_button" onClick={() => {
                                setShowLineChart(prevState => !prevState);
                                //console.log("showLineChart toggled");
                            }}>
                                <div className="switch_inner">
                                    {showLineChart ? (
                                        <img className="switch_icon" src={graphicon} alt="Graph Icon" />
                                    ) : (
                                        <img className="switch_icon" src={tableicon} alt="Table Icon" />
                                    )}
                                    <span className="switch_text">Switch</span>
                                </div>
                            </Button>

                        </div>
                    </div>
                    <div style={{ display: 'flex' }}>
                        <div className="form-sidebar">
                            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
                                <h5 className="form-sidebar-title">Select Form</h5>
                                <button
                                    style={{
                                        background: 'none',
                                        border: 'none',
                                        cursor: 'pointer',
                                        color: '#B0D136',
                                        fontSize: '1rem',
                                        padding: '0 10px'
                                    }}
                                    onClick={() => {
                                        // Simply reverse the array
                                        setFormmasterdata([...formmasterdata].reverse());
                                        // Toggle the sort direction for UI feedback
                                        setSortDirection(sortDirection === 'asc' ? 'desc' : 'asc');
                                    }}
                                    title={`Reverse order (${sortDirection === 'asc' ? 'Newest first' : 'Original order'})`}
                                >
                                    {sortDirection === 'asc' ? '▼' : '▲'}
                                </button>
                            </div>
                            <ul className="form-list">
                                {formmasterdata.map((item: any, index: any) => (
                                    <li
                                        key={index}
                                        className={`form-list-item ${appcode === item.AppCode ? 'selected' : ''}`}
                                        onClick={() => {
                                            setAppcode(item.AppCode);
                                            setAppname(item.AppName);
                                        }}
                                    >
                                        {item.AppName}
                                    </li>
                                ))}
                            </ul>
                        </div>
                        <div style={{ display: 'flex', flexDirection: 'column', width: '100%', overflow: 'auto', height: 'calc(-46px - 44px + 100vh)' }}>
                            <div className='cards-container'>
                                <div className='card-item'>
                                    <div className='staticdivcard' onClick={() => { handleFilter('All'); setDisplayFilterTable(true); setDisplayTable(false); }}>
                                        <div className='rounddivicon'>
                                            <img src={TotalRequests} className='icon-style-formdashboard' />
                                        </div>

                                        <div className='d-flex flex-column'>
                                            <span className='text-frmdashbrd'>All Requests</span>
                                            <span style={{
                                                fontSize: '13px',
                                                fontWeight: 400,
                                                color: '#B0D136',
                                            }}>
                                                {countOfRecords < 10 ? `0${countOfRecords}` : countOfRecords}
                                            </span>
                                        </div>
                                    </div>
                                </div>
                                <div className='card-item' onClick={() => { handleFilter('Rejected'); setDisplayFilterTable(true); setDisplayTable(false); }}>
                                    <div className='staticdivcard' >
                                        <div className='rounddivicon'>
                                            <img src={PendingRequests} className='icon-style-formdashboard' />
                                        </div>

                                        <div className='d-flex flex-column'>
                                            <span className='text-frmdashbrd'>Rejected Requests</span>
                                            <span style={{
                                                fontSize: '13px',
                                                fontWeight: 400,
                                                color: '#EF5350',
                                            }}>
                                                {rejectedCount < 10 ? `0${rejectedCount}` : rejectedCount}
                                            </span>
                                        </div>
                                    </div>
                                </div>
                                <div className='card-item' onClick={() => { handleFilter('In-Progress'); setDisplayFilterTable(true); setDisplayTable(false) }}>
                                    <div className='staticdivcard'>
                                        <div className='rounddivicon'>
                                            <img src={InprogressRequests} className='icon-style-formdashboard' />
                                        </div>

                                        <div className='d-flex flex-column'>
                                            <span className='text-frmdashbrd'>In-Progress Requests</span>
                                            <span style={{
                                                fontSize: '13px',
                                                fontWeight: 400,
                                                color: '#42A5F5',
                                            }}>
                                                {inProgressCount < 10 ? `0${inProgressCount}` : inProgressCount}
                                            </span>
                                        </div>
                                    </div>
                                </div>
                                <div className='card-item' onClick={() => { handleFilter('Completed'); setDisplayFilterTable(true); setDisplayTable(false) }}>
                                    <div className='staticdivcard'>
                                        <div className='rounddivicon'>
                                            <img src={CompletedRequests} className='icon-style-formdashboard' />
                                        </div>
                                        <div className='d-flex flex-column'>
                                            <span className='text-frmdashbrd'>Completed Requests</span>
                                            <span style={{
                                                fontSize: '13px',
                                                fontWeight: 400,
                                                color: '#66BB6A',
                                            }}>
                                                {aprovedCount < 10 ? `0${aprovedCount}` : aprovedCount}
                                            </span>
                                        </div>
                                    </div>
                                </div>
                            </div>

                            {appcode ? (
                                <div >
                                    {loading ? (<LoadingSpinner />) : (<></>)}

                                    {showLineChart ?

                                        <div className='datatable-container'>
                                            {displayTable &&
                                                <DataTable
                                                    responsiveLayout="scroll"
                                                    paginator
                                                    paginatorTemplate="CurrentPageReport FirstPageLink PrevPageLink PageLinks NextPageLink LastPageLink RowsPerPageDropdown"
                                                    currentPageReportTemplate="Showing {first} to {last} of {totalRecords}"
                                                    rows={6}
                                                    onRowClick={(e) => onTitleClick(e.data)}
                                                    // rowsPerPageOptions={[10, 20, 50]}
                                                    dataKey="ID"
                                                    emptyMessage="No transactions found."
                                                    value={filteredData}
                                                // className="statusView-main-table"

                                                >
                                                    <Column field="Title" header="RequestID" sortable />
                                                    <Column field="Status" header="Status" sortable />
                                                    <Column field="CategoryType" header="Request For" sortable />

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
                                            }
                                            {displayFilterTable &&
                                                <DataTable
                                                    value={filteredData}
                                                    responsiveLayout="scroll"
                                                    paginator
                                                    paginatorTemplate="CurrentPageReport FirstPageLink PrevPageLink PageLinks NextPageLink LastPageLink RowsPerPageDropdown"
                                                    currentPageReportTemplate="Showing {first} to {last} of {totalRecords}"
                                                    rows={6}
                                                    onRowClick={(e) => onTitleClick(e.data)}
                                                    dataKey="ID"
                                                    emptyMessage="No transactions found."
                                                // className="statusView-main-table"
                                                >
                                                    <Column field="Title" header="RequestID" sortable />
                                                    <Column field="Status" header="Status" sortable />
                                                    <Column field='AppName' header="Form Name" sortable />
                                                    <Column field="CategoryType" header="Request For" sortable />

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
                                            }

                                        </div> :
                                        <div className='GraphContainer' >
                                            {formTypes.length > 0 ? (
                                                <div className="" style={{ maxWidth: '1000px', width: '100%', backgroundColor: '#070D19', padding: "20px" }}>
                                                    <Chart type="bar" data={chartData} options={chartOptions} style={{ width: '750px', height: '360px', boxShadow: '0px 0px 4px 2px #87878740', border: '1px solid #4A4A4A', padding: '10px 0px 10px 10px' }} />
                                                </div>
                                            ) : (
                                                <div className="" style={{ maxWidth: '1000px', width: '100%', backgroundColor: '#070D19', paddingLeft: '14px' }}>
                                                    <Chart
                                                        type="bar"
                                                        data={chartData}
                                                        options={chartOptions}
                                                        className="customBarChart"
                                                    />

                                                </div>
                                            )}
                                        </div>}

                                </div>
                            ) : (
                                ""
                            )}
                        </div>
                    </div>
                </div>
            </div>
        </div>
    );
};

export default FormDashBoard;