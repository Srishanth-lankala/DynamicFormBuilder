import * as React from "react";
import { useState, useEffect } from "react";
import { useNavigate } from "react-router-dom";
import { DataTable } from "primereact/datatable";
import { Column } from "primereact/column";
import { sp } from "@pnp/sp/presets/all";
import TopNavBar from "../TopBar";
import Sidebar from "../SideNavBar";
import { fetchTenantUser } from "../FetchTenantUser/fetchTenantUser";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "primereact/resources/themes/saga-blue/theme.css";
import "primereact/resources/primereact.min.css";
import { mySiteUrl } from "../ConfigURL/All_URLs";
import "./byStatus.css"
import { useGlobalState } from "../GlobalVariable/GlobalStateContext";
const ViewByStatus: React.FC = () => {
  const [groupedData, setGroupedData] = useState<any[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [expandedRows, setExpandedRows] = useState<any>(null);
  const navigate = useNavigate();
  const { globalVariable } = useGlobalState();
  const parsedGlobalVariable = globalVariable ? JSON.parse(globalVariable) : null;
const [CurrentUserEmail, setCurrentUserEmail] = useState<string | null>(null);
  sp.setup({
    sp: {
      baseUrl: mySiteUrl,
    },
  });

  const fetchData = async () => {
    setLoading(true);
    try {
      let items: any[] = [];
         const currentUser = await sp.web.currentUser.get();
        //  const userId = currentUser.Id;
        const Domain = await fetchTenantUser();
        let filteredItems : any[];
        items = await sp.web.lists
           .getByTitle("WorkFlowProcessData")
           .items.select("*", "CurApprover/Title", "Author/Title", "AppCode")
           .expand("CurApprover", "Author")
           .filter(`Domain eq '${Domain.TenantUsers}' and  AppName eq '${parsedGlobalVariable.appname}'`)
           .getAll();
         if(CurrentUserEmail === "OnlyUser"){
           filteredItems = items.filter(item => item.AllApprovers?.includes(currentUser.Email)||item.Author?.Title === currentUser.Title || item.CurApprover?.Title === currentUser.Title);
         }
         else{filteredItems = items}

      const statuses = ["Completed", "Rejected", "Returned", "In-Progress","Draft","Cancelled"];
      const grouped = statuses.map((status) => ({
        status,
        items: filteredItems.filter((item) => item.Status === status),
      }));

      setGroupedData(grouped);
    } catch (error) {
      console.error("Error fetching data from SharePoint list:", error);
      setGroupedData([]);
    } finally {
      setLoading(false);
    }
  };



  useEffect(() => {
  
    void fetchData();
  },  [parsedGlobalVariable.appcode, CurrentUserEmail]);

  const onRowToggle = (e: any) => {
    console.log("eeeeeeeeeee", e);
    setExpandedRows((prev: any) => (prev === e.data ? null : e.data));
  };

  const onTitleClick = (rowData: any) => {
    console.log("status--------------------------------------11111111", rowData)
    if(rowData.Status==="Returned"){
      navigate(`/EditForm/${rowData.ID}`);
    }
    else{
      navigate(`/viewform/${rowData.ID}`);
    }
  }

  const rowExpansionTemplate = (data: any) => {
    return (
      <DataTable
        value={data.items}
        responsiveLayout="scroll"
        emptyMessage="No data found."
        onRowClick={(e) => onTitleClick(e.data)}
        className="statusView-expanded-table"
      >
        <Column field="Title" header="RequestID" sortable />
        <Column field="AppName" header="Form Name" sortable />
         <Column field="CategoryType" header="Request For" sortable />
        <Column
          field="CurApprover.Title"
          header="Current Approver"
          body={(rowData) => rowData.CurApprover?.Title.split('|')[0] || ""}
          sortable
        />
        <Column
          field="Author.Title"
          header="Author"
          body={(rowData) => rowData.Author?.Title.split('|')[0] || ""}
          sortable
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
    );
  };
  const handleUserEmailRetrieved = (UserAUTH: string) => {
    setCurrentUserEmail(UserAUTH);
    console.log("Logged-in USER IS   ---------------- !!!!!!!!!!!! :", UserAUTH);
};

  return (
    <div className="statusView-page">
      <TopNavBar onUserTypeRetrieved={handleUserEmailRetrieved} />
      <div className="statusView-main-container">
        <Sidebar />
        <div className="statusView-content-container">
          {/* <p className="statusView-page-title">
            View By Status
          </p> */}
           <div style={{display:'flex',alignItems:'center',gap:'7px'}}>
          <img  onClick={()=>navigate(-1)}style={{ height:23, width:23,cursor:"pointer" }}src={require('../../../assets/Images/previous.png')} alt="backicon"  title='Back'/>

          <h4 style={{marginTop:5}}> View By Status</h4>
          </div>
          {loading ? (
            <p>Loading...</p>
          ) : (
            <DataTable
              value={groupedData}
              responsiveLayout="scroll"
              dataKey="status"
              expandedRows={expandedRows}
              onRowToggle={onRowToggle}
              rowExpansionTemplate={rowExpansionTemplate}
              emptyMessage="No data found."
              paginator
              rows={6}
              rowsPerPageOptions={[5, 10, 20]}
              className="statusView-main-table"
            >
              <Column expander style={{ width: "3em" }} />
              <Column field="status" header="Status" />
              <Column header="Count" body={(data) => data.items.length} />
            </DataTable>
          )}
        </div>
      </div>
    </div>
  );
};

export default ViewByStatus;