import * as React from 'react';
import { useState, useEffect } from 'react';
import { DataTable } from 'primereact/datatable';
import { Column } from 'primereact/column';
import { useNavigate } from 'react-router-dom';
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import 'primereact/resources/themes/saga-blue/theme.css';
import 'primereact/resources/primereact.min.css';
import TopNavBar from './TopBar';
// import Sidebar from './SideNavBar';
import LoadingSpinner from './Loading';
import { mySiteUrl } from './ConfigURL/All_URLs';
// import  ColorSwitch  from '../CustomFormBuilder/Toggle'
import { fetchTenantUser } from './FetchTenantUser/fetchTenantUser';
import SideBar from './Sidebar/SideBar';
import searchicn from '../../assets/Images/searchicon.svg';
import { Button } from "primereact/button";
import '../CustomFormBuilder/ConfigureWorkflow/CustomForm.css';
import './Inprocess.css';


const InProcess: React.FC = () => {
    const [transactions, setTransactions] = useState<any[]>([]);
    const [loading, setLoading] = useState<boolean>(true);
    const [filteredData, setFilteredData] = useState<any[]>([]);
    const [Toggleron, setToggleron] = useState<boolean>(true);
    const [searchTerm, setSearchTerm] = useState<string>(''); // Add this state
    const [currentuser, setcurrentuser] = useState<any>([]);
    const [Useringroups, setUseringroups] = useState<number[]>([]);
    const navigate = useNavigate();
    sp.setup({
        sp: {
            baseUrl: mySiteUrl,
        },
    });

    useEffect(() => {

        const fetchData = async () => {
            try {
                setLoading(true);
                const currentUser = await sp.web.currentUser.get();
                //console.log("Current User ID:", currentUser.Id, currentUser);
                setcurrentuser(currentUser);
                type WorkFlowItem = {
                    ID: number;
                    Title: string;
                    Status: string;
                    Created: string;
                    Modified: string;
                    Author: { Title: string };
                    CurApprover?: { Title: string };
                    CurApproverId?: number;
                    AppName: string;
                    Domain: string;
                    CategoryType: string;
                    AllApprovers: String;
                    DestinationQueue: String;
                    CurrentQueue: String;
                };
                const Domain = await fetchTenantUser();
                const allItems: WorkFlowItem[] = await sp.web.lists.getByTitle("WorkFlowProcessData")
                    .items
                    .orderBy("Modified", false) // Order by Modified (Descending)
                    .select("ID", "Title", "Status", "Created", "Modified", "Author/Title", "CurApprover/Title", "CurApproverId", "AppName", "CategoryType", "Domain", "AllApprovers", "DestinationQueue", "CurrentQueue", "CurrentAppCode")
                    .expand("Author", "CurApprover")
                    .filter(`Domain eq '${Domain.TenantUsers}' and Status ne 'Draft' and Status ne 'Submitted' and ParentAppCode eq null`)
                    .getAll();

                //console.log("Fetched Items:", allItems);
                let finalFilteredRecords: WorkFlowItem[];

                if (Toggleron) {  // Fetch inprocess data
                    setLoading(true);
                    debugger;
                    // console.log("toggler is ON")
                    let groupIdsToCheck = new Set<number>();
                    let recordsWithCurrentUser: WorkFlowItem[] = [];

                    // Identify records where CurApprover is a group or the current user
                    allItems.forEach(item => {
                        if (item.CurApproverId === currentUser.Id) {
                            recordsWithCurrentUser.push(item);
                        } else if (item.CurApproverId) {
                            groupIdsToCheck.add(item.CurApproverId);
                        }
                    });

                    //console.log("Group IDs to Check:", [...groupIdsToCheck]);

                    // Fetch site groups to determine which CurApproverId belongs to a group
                    const siteGroups = await sp.web.siteGroups.get();
                    const siteGroupIds = new Set(siteGroups.map(group => group.Id));

                    //console.log("All Site Group IDs:", [...siteGroupIds]);

                    // Filter records where CurApproverId is a group
                    const recordsWithGroups = allItems.filter(
                        item => item.CurApproverId && siteGroupIds.has(item.CurApproverId)
                    );

                    //console.log("Records where CurApprover is a Group:", recordsWithGroups);

                    // Find groups where the logged-in user is a member
                    const userGroups = await sp.web.currentUser.groups.get();
                    const userGroupIds = new Set(userGroups.map(group => group.Id));
                    setUseringroups([...userGroupIds]);

                    //console.log("User's Group IDs:", [...userGroupIds]);

                    // Filter records where the user is a member of the CurApprover group
                    const recordsWhereUserIsGroupMember = recordsWithGroups.filter(
                        item => item.CurApproverId && userGroupIds.has(item.CurApproverId)
                    );

                    //console.log("Records where CurApprover is a Group AND User is a Member:", recordsWhereUserIsGroupMember);

                    // Merge both filtered lists
                    finalFilteredRecords = [...recordsWhereUserIsGroupMember, ...recordsWithCurrentUser];

                    // Sort by "Modified" in descending order
                    finalFilteredRecords.sort((a, b) => new Date(b.Modified).getTime() - new Date(a.Modified).getTime());

                    //console.log("Final Sorted Records:", finalFilteredRecords);



                }
                else { //Fetch the records on which the user took some action return, approve, reject
                    // setLoading(true);
                    //console.log("toggler OFF")
                    finalFilteredRecords = allItems.filter(item => item.AllApprovers?.includes(currentUser.Email));
                }
                setTransactions(finalFilteredRecords);
                setFilteredData(finalFilteredRecords)
            }
            catch (error) {
                console.error("Error fetching data from SharePoint list:", error);
            } finally {
                setLoading(false);
            }
        };

        fetchData().catch((error) => {
            console.error("Fetch data failed:", error);
            setLoading(false);
        });
    }, [Toggleron]);

    const onTitleClick = (rowData: any) => {
    const currentAppCode = rowData.CurrentAppCode || "";
      
        if ((rowData.Status === "In-Progress" || rowData.Status === "Returned")
            && rowData.DestinationQueue === "0" && rowData.CurrentQueue !== "0"
            && (rowData.CurApproverId === currentuser.Id || Useringroups.includes(rowData.CurApproverId))
            && rowData.CurrentAppCode) {
            navigate(`/CForm/${rowData.ID}`);
        } 
        else if (rowData.Status === 'Returned' &&rowData.DestinationQueue === "0" && currentAppCode.slice(-3) === "-1C" && rowData.CurApproverId === currentuser.Id) {
            navigate(`/EditForm/${rowData.ID}`);
        }
        else {
            navigate(`/viewform/${rowData.ID}`);
        }
    };

    const handleSearch = (searchTerm: string) => {
        setSearchTerm(searchTerm); // Update the search term state
        if (searchTerm.trim().length === 0) {
            setFilteredData(transactions); // Reset to full data
            return;
        }

        const filtered = transactions.filter((item) =>
            item.Title?.toLowerCase().includes(searchTerm.toLowerCase()) ||
            item.Status?.toLowerCase().includes(searchTerm.toLowerCase()) ||
            item.AppName?.toLowerCase().includes(searchTerm.toLowerCase()) ||
            item.CategoryType?.toLowerCase().includes(searchTerm.toLowerCase())
        );
        setFilteredData(filtered);
    };



    return (
        <div style={{ display: 'flex', flexDirection: 'column' }}>
            <TopNavBar />
            <div style={{ display: 'flex' }}>
                <SideBar activeMenu="Pending Approvals" />
                <div className='Myrequestscroll' style={{ height: "calc(-46px + 100vh)", width: "100%", overflowY: "auto", backgroundColor: '#070D19' }}>
                    <div className="pendingheader">
                        <div className='draft_header_flex' style={{ height: '52px' }}>
                            {/* <section> <h5 className="fo-ap" style={{ width: "100%", margin: '20px 15px' }}>Pending Approvals</h5> </section> */}
                            <section style={{ width: '500px' }}>
                                <div style={{ display: 'flex', alignItems: 'center', gap: '7px', margin: '20px 0px 0px 0px' }}>
                                    <Button
                                        //    style={{borderRadius:'8px 8px 0px 0px', height:'35px',minWidth:'185px',justifyContent:'center',alignItems:'center',background:Toggleron ? '#0E8EBE':'#212B46',border:'0px',boxShadow:'0px'}}
                                        className={Toggleron ? 'tglenabled tglenabled1' : 'tgldisabled tgldisabled1'}
                                        onClick={() => { setToggleron(true); handleSearch('') }}
                                    >
                                        Pending Approvals
                                    </Button>
                                    <Button
                                        //    style={{borderRadius:'8px 8px 0px 0px', height:'35px',minWidth:'185px',justifyContent:'center',alignItems:'center',background:Toggleron ? '#212B46':'#0E8EBE',border:'0px',boxShadow:'0px'}}
                                        className={Toggleron ? 'tgldisabled' : 'tglenabled'}
                                        onClick={() => { setToggleron(false); handleSearch('') }}
                                    >
                                        My Action Items
                                    </Button>
                                    {/* <h5 className="fo-ap" style={{ margin: '0',color:'#FFFFFF' }}>{Toggleron?"Pending Approvals":"My Action Items"}</h5> */}
                                </div>
                            </section>
                            {/* <div className="form-check form-switch" style={{ marginLeft: "210px", marginTop: '20px' }}>
                            <input
                                style={{
                                    width: 40,
                                    height: 20,
                                    borderRadius: 20,
                                    border: '0.5px solid rgba(0, 0, 0, 0.66)',
                                    appearance: 'none',
                                    WebkitAppearance: 'none',
                                    backgroundColor: '#EE691E', // Light Orange or Light Violet
                                    position: 'relative',
                                    cursor: 'pointer'
                                }}
                                className="form-check-input"
                                type="checkbox"
                                id="finalapprove"
                                onChange={() => { setToggleron(!Toggleron); handleSearch('') }} // Reset search when toggling
                            />
                            <label className="form-check-label" style={{ marginLeft: 10, marginTop: '3px' }} >
                                {Toggleron ? 'Action Pending' : 'Action Taken'}
                            </label>
                        </div> */}
                            <section style={{ alignContent: 'center' }}>
                                <div className="searchContainerInprocess">
                                    <input placeholder='Search...'
                                        className="searchBox"
                                        onChange={(e: React.ChangeEvent<HTMLInputElement>) => { handleSearch(e.target.value.trim()) }} // Call the search function on input change
                                        // style={{ marginTop: '-25px' }}
                                        value={searchTerm} // Bind the input value to the search term state
                                    />
                                    <img src={searchicn} alt="Search Icon" className="searchIcon" />
                                </div>
                            </section>
                        </div>
                        <div className='tabline'></div>
                    </div>


                    {loading ? (
                        <LoadingSpinner />
                    ) : (
                        <DataTable
                            className="hover-table"
                            value={filteredData}
                            responsiveLayout="scroll"
                            paginator
                            paginatorTemplate="CurrentPageReport FirstPageLink PrevPageLink PageLinks NextPageLink LastPageLink RowsPerPageDropdown"
                            currentPageReportTemplate="Showing {first} to {last} of {totalRecords}"
                            rows={10}
                            sortField='Modified'
                            sortOrder={-1}
                            rowsPerPageOptions={[10, 15, 20]}
                            dataKey="ID"
                            emptyMessage="No transactions found."
                            onRowClick={(e) => onTitleClick(e.data)}
                            rowClassName={(rowData, options) => {
                                const index = transactions.findIndex(item => item.ID === rowData.ID); // Find the row index
                                return index % 2 === 0 ? 'Table_tr_even' : 'Table_tr_odd'; // Apply CSS class
                            }}
                        >
                            <Column field="Title" header="RequestID" sortable />
                            <Column field="Status" header="Status" sortable />
                            <Column field="AppName" header="Form Name" sortable
                                body={(rowData) => (
                                    <span title={rowData.AppName}>
                                        {rowData.AppName}
                                    </span>
                                )}
                            />
                            <Column field="CategoryType" header="Request For" sortable
                                body={(rowData) => (
                                    <span title={rowData.CategoryType}>
                                        {rowData.CategoryType}
                                    </span>
                                )}
                            />

                            {/* Created By Column */}
                            <Column
                                field="Author"
                                header="Created By"
                                sortable
                                body={(rowData) => {
                                    if (rowData.Author && rowData.Author.Title) {
                                        const title = rowData.Author.Title.split('|')[0];
                                        return title || '';
                                    }
                                    return '';
                                }}
                            />

                            {/* Current Approver Column */}
                            <Column
                                field="CurApprover"
                                header="Current Approver"
                                sortable
                                body={(rowData) => {
                                    if (rowData.CurApprover && rowData.CurApprover.Title) {
                                        const approverTitle = rowData.CurApprover.Title.split('|')[0];
                                        return approverTitle || '';
                                    }
                                    return ''; // If CurApprover or Title is not available, return 'N/A'
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
                                }}
                            />
                            <Column
                                field="Modified"
                                header="Modified Date & Time"
                                sortable
                                body={(rowData) => {
                                    const date = new Date(rowData.Modified);
                                    return date.toLocaleDateString('en-GB', {
                                        day: '2-digit',
                                        month: 'short',
                                        year: 'numeric',
                                        hour: '2-digit',
                                        minute: '2-digit',
                                        hour12: false
                                    });
                                }}
                            />
                        </DataTable>



                    )}
                </div>
            </div>

        </div>
    );
};

export default InProcess;