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
import LoadingSpinner from './Loading';
import { mySiteUrl } from './ConfigURL/All_URLs';
import '../CustomFormBuilder/ConfigureWorkflow/CustomForm.css';
import SideBar from './Sidebar/SideBar';
import searchIcon from '../../assets/Images/searchicon.svg';
import './Draft.css';
import { fetchTenantUser } from './FetchTenantUser/fetchTenantUser';

const Draft: React.FC = () => {
    const [transactions, setTransactions] = useState<any[]>([]);
    const [loading, setLoading] = useState<boolean>(true);
    const [filteredData, setFilteredData] = useState<any[]>([])
    const [currentUserId, setCurrentUserId] = useState<number | null>(null);
    const navigate = useNavigate();

    sp.setup({
        sp: {
            baseUrl: mySiteUrl,
        },
    });
    useEffect(() => {
        debugger;
        const fetchData = async () => {
            try {
                const Domain = await fetchTenantUser();
                const currentUser = await sp.web.currentUser.get();
                setCurrentUserId(currentUser.Id);


                const items = await sp.web.lists.getByTitle("WorkFlowProcessData").items.orderBy("Modified", false)
                    .filter(`AuthorId eq ${currentUser.Id} and Domain eq '${Domain.TenantUsers}' and ParentAppCode eq null`)
                    .select("ID", "Title", "Status", "Created", "Modified", "Author/Title", "CurApprover/Title", "CurApproverId","AppName", "CategoryType","CurrentAppCode")
                    .expand("Author", "CurApprover")
                    .getAll();


                debugger;
                setTransactions(items);
                setFilteredData(items)
                //console.log("In-Progress Data Table items: ", items);
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
        debugger; // Check rowData when clicking a title
        const matchedTransaction = transactions.find(
            (t: any) => t.ID === rowData.ID
        );
        // If it does NOT include '-1C' (case-insensitive)
        const currentAppCode = matchedTransaction.CurrentAppCode || "";
        const has1C = currentAppCode.slice(-3)==="-1C";

      if (rowData.Status?.toLowerCase() === 'returned' && !has1C &&rowData.CurApproverId ===currentUserId) {
            navigate(`/CForm/${rowData.ID}`);
        }
        else if (rowData.Status?.toLowerCase() === 'returned' && !has1C && rowData.CurApproverId !== currentUserId && (matchedTransaction?.CurrentAppCode)) {
            navigate(`/ViewForm/${rowData.ID}`);
        }
        else {
            navigate(`/EditForm/${rowData.ID}`);
        }
        
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

                <SideBar activeMenu="My Requests" />
                <div className='Myrequestscroll' style={{ height: "calc(-46px + 100vh)", width: "100%", overflowY: "auto", backgroundColor: '#070D19' }}>

                    <div className='draftheader'>
                        {/* <section> <h5 className="fo-ap" style={{ width: "100%", margin: '20px 15px' }}>My Documents</h5> </section> */}
                        <section>
                            <div style={{ display: 'flex', alignItems: 'center', gap: '7px', margin: '20px 15px' }}>
                                {/* <img onClick={() => navigate(-1)} style={{ height: 23, width: 23, cursor: "pointer" }} src={require('../../assets/Images/previous.png')} alt="backicon" title='Back' /> */}
                                <h5 className="fo-ap" style={{ margin: '0', color: '#FFFFFF', fontSize: 16 }}>My Requests</h5>
                            </div>
                        </section>
                        <section >
                            <div className="searchContainer" style={{ margin: '14px 10px', height: '50%' }}>

                                <input placeholder='Search...'
                                    className="searchBox"
                                    onChange={(e: React.ChangeEvent<HTMLInputElement>) => handleSearch(e.target.value)}
                                />
                                <img src={searchIcon} alt="Search Icon" className="searchIcon" />
                            </div>
                        </section>


                    </div>

                    {loading ? (
                        <LoadingSpinner />
                    ) : (
                        <DataTable className="hover-table"

                            value={filteredData}
                            responsiveLayout="scroll"
                            paginator
                            sortField='Modified'
                            sortOrder={-1}
                            paginatorTemplate="CurrentPageReport FirstPageLink PrevPageLink PageLinks NextPageLink LastPageLink RowsPerPageDropdown"
                            currentPageReportTemplate="Showing {first} to {last} of {totalRecords}"
                            rows={10}
                            font-size='16px'
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
                            <Column field='AppName' header="Form Name" sortable
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
                                }} />
                        </DataTable>
                    )}
                </div>
            </div>

        </div>
    );
};

export default Draft;