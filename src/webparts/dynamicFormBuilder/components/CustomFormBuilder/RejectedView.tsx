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
import Sidebar from './SideNavBar';

const RejectedView: React.FC = () => {
    const [transactions, setTransactions] = useState<any[]>([]);
    const [loading, setLoading] = useState<boolean>(true);
    const navigate = useNavigate();

    useEffect(() => {
        debugger; // Debugger to inspect when component loads
        // Fetch data from the "WorkFlowProcessData" SharePoint list where Status is 'In-Progress'
        const fetchData = async () => {
            try {
                const currentUser = await sp.web.currentUser.get();

                console.log(currentUser)
                const items = await sp.web.lists.getByTitle("WorkFlowProcessData").items.filter(`(Status eq 'Rejected') and (AuthorId eq ${currentUser.Id})`).get();
                debugger; // Inspect items retrieved
                setTransactions(items);
                
                console.log("In-Progress Data Table items: ", items);
            } catch (error) {
                console.error("Error fetching data from SharePoint list:", error);
            } finally {
                setLoading(false); // Ensure loading is set to false
            }
        };
        
        fetchData().catch((error) => {
            console.error("Fetch data failed:", error);
            setLoading(false);
        });
    }, []);

    const onTitleClick = (rowData: any) => {
        debugger; // Check rowData when clicking a title
        navigate(`/RejectedForm/${rowData.ID}`);
    };

    return (
        <div style={{ display: 'flex', flexDirection: 'column' }}>
            <TopNavBar />
            <div style={{ display: 'flex' }}>
                <Sidebar />
                <div style={{ width: "71vw" }}>
            <p className="fo-ap" style={{ width: "100%" }}>Pending Approval</p>
            {loading ? (
                <p>Loading...</p>
            ) : (
                <DataTable
                    value={transactions}
                    responsiveLayout="scroll"
                    paginator
                    paginatorTemplate="CurrentPageReport FirstPageLink PrevPageLink PageLinks NextPageLink LastPageLink RowsPerPageDropdown"
                    currentPageReportTemplate="Showing {first} to {last} of {totalRecords}"
                    rows={20}
                    rowsPerPageOptions={[10, 20, 50]}
                    dataKey="ID"
                    emptyMessage="No transactions found."
                    onRowClick={(e) => onTitleClick(e.data)} // Event for row click
                >
                    <Column field="Title" header="RequestID" sortable />
                    <Column field="Status" header="Status" sortable />
                </DataTable>
            )}
        </div>
                
            </div>

        </div>
        
    );
};

export default RejectedView;
{/*  */}