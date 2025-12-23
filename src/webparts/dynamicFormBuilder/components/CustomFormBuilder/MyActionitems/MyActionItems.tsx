import * as React from 'react';
import { useState, useEffect } from 'react';
import { DataTable } from 'primereact/datatable';
import { Column } from 'primereact/column';
import { useNavigate } from 'react-router-dom';
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { fetchTenantUser } from '../FetchTenantUser/fetchTenantUser';
import 'primereact/resources/themes/saga-blue/theme.css';
import 'primereact/resources/primereact.min.css';
import TopNavBar from '../TopBar';
// import Sidebar from './SideNavBar';
import LoadingSpinner from '../Loading';
import { mySiteUrl } from '../ConfigURL/All_URLs';

const InProcess: React.FC = () => {
    const [transactions, setTransactions] = useState<any[]>([]);
    const [loading, setLoading] = useState<boolean>(true);
    const [filteredData, setFilteredData] = useState<any[]>([])
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
                console.log("Current User ID:", currentUser.Id);
                const Domain = await fetchTenantUser()
                console.log('Domain: ',Domain.TenantUsers)
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
                    CategoryType: string;
                    AllApprovers : string
                };

                const allItems: WorkFlowItem[] = await sp.web.lists.getByTitle("WorkFlowProcessData")
                    .items
                    .orderBy("Created", false) // Order by Modified (Descending)
                    .select("ID", "Title", "Status","AllApprovers", "Created", "Modified", "Author/Title",'Domain', "CurApprover/Title", "CurApproverId", "AppName", "CategoryType")
                    .expand("Author", "CurApprover")
                    .filter(`Domain eq '${Domain.TenantUsers}' and Author/Id eq ${currentUser.Id}`)
                    .getAll();
                
                const allItems_1: WorkFlowItem[] = await sp.web.lists.getByTitle("WorkFlowProcessData")
                    .items
                    .orderBy("Created", false) // Order by Modified (Descending)
                    .select("ID", "Title", "Status", "Created","AllApprovers", "Modified", "Author/Title",'Domain', "CurApprover/Title", "CurApproverId", "AppName", "CategoryType")
                    .expand("Author", "CurApprover")
                    .filter(`Domain eq '${Domain.TenantUsers}'`)
                    .getAll();
                    const filteredItems = allItems_1.filter(item => item.AllApprovers?.includes(currentUser.Email));
                console.log('Domain items:',allItems_1)
                console.log("Fetched Items user created:", allItems);
                console.log('user in allapprovers : ',filteredItems);

                const finalFilteredRecords = [...filteredItems, ...allItems]
                console.log('all the items : ',finalFilteredRecords)
                
           
                setTransactions(finalFilteredRecords);
                setFilteredData(finalFilteredRecords)

            } catch (error) {
                console.error("Error fetching data from SharePoint list:", error);
            } finally {
                setLoading(false);
            }
        };
        fetchData().catch((error) => {
            console.error("Fetch data failed:", error);
            setLoading(false);
        });
    }, []);

    const onTitleClick = (rowData: any) => {
        debugger;
        navigate(`/viewform/${rowData.ID}`);
    };

    const handleSearch = (searchTerm: string) => {
        if (!searchTerm) {
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
                {/* <Sidebar /> */}
                <div style={{ height: "100vh", overflowY: "scroll", width: "100%" }}>

                    <div className='draft_header_flex'>
                        {/* <section> <h5 className="fo-ap" style={{ width: "100%", margin: '20px 15px' }}>Pending Approvals</h5> </section> */}
                        <section>
                            <div style={{ display: 'flex', alignItems: 'center', gap: '7px', margin: '20px 15px' }}>
                                <img onClick={() => navigate(-1)} style={{ height: 18, width: 18, cursor: "pointer" }} src={require('../../../assets/Images/previous.png')} alt="backicon" title='Back' />
                                <h5 className="fo-ap" style={{ margin: '0' ,fontSize:16}}>My Action Items</h5> </div>
                        </section>
                        <section > 
                        <div className="search_container">
                        <img  src={require("../../../assets/Images/magnifying-glass.png")} alt="Search Icon" className="search_icon" />
                              <input placeholder='Search...'
                            className="search_box"
                            onChange={(e: React.ChangeEvent<HTMLInputElement>) => handleSearch(e.target.value)}
                        />
                        </div>
                        </section>

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
                            <Column field="AppName" header="Form Name" sortable />
                            <Column field="CategoryType" header="Request For" sortable />

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