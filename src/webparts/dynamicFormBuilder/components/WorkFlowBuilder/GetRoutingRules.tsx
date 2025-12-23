import { sp } from "@pnp/sp/presets/all";
import * as React from "react";
import { useState, useEffect } from "react";
import { IDynamicFormBuilderProps } from "../IDynamicFormBuilderProps";
import Workflow from "./Workflow";
// import SideNavBar from "../CustomFormBuilder/SideNavBar";
import TopNavBar from "../CustomFormBuilder/TopBar";
import { mySiteUrl } from "../CustomFormBuilder/ConfigURL/All_URLs";
import { useGlobalState } from "../CustomFormBuilder/GlobalVariable/GlobalStateContext";
import SideBar from "../CustomFormBuilder/Sidebar/SideBar";

const GetRoutingRules: React.FC<IDynamicFormBuilderProps> = (props) => {
  const [MappingMasterState, SetMappingMasterState] = useState<any[]>([]);
  const [CurrentMappingMasterState, SetCurrentMappingMasterState] = useState<any[]>([]);
  const [CurrentRoutingRulesState, SetCurrentRoutingRulesState] = useState<any[]>([]);
  const [isReplacementComplete, setIsReplacementComplete] = useState(false);
  const { globalVariable } = useGlobalState();
  const parsedGlobalVariable = globalVariable ? JSON.parse(globalVariable) : null;

  // Setup SP context
  sp.setup({
    sp: {
      baseUrl: mySiteUrl,
    },
  });

  const fetchMappingMasterData = async () => {
    try {
      const MappingMaster = await sp.web.lists
        .getByTitle("MappingMaster")
        .items.select("Role", "Level", "AppCode", "AppName", "ID")
        .getAll();
      SetMappingMasterState(MappingMaster);
    } catch (error) {
      console.error("Error fetching MappingMaster data:", error);
    }
  };

  const fetchCurrentMappingMaster = async (code: string) => {
    try {
      const currentMappingMaster = MappingMasterState.filter((ele) => ele.AppCode === code);
      SetCurrentMappingMasterState(currentMappingMaster);
    } catch (error) {
      console.error("Error fetching current MappingMaster data:", error);
    }
  };

  const fetchCurrentRoutingRules = async (code: string) => {
    try {
      const RoutingRules = await sp.web.lists
        .getByTitle("RoutingRules")
        .items.filter(`AppCode eq '${code}'`)
        .select("id", "Actions", "CurrentQueue", "DestinationQueue", "Status")
        .getAll();
      SetCurrentRoutingRulesState(RoutingRules);
    } catch (error) {
      console.error("Error fetching RoutingRules:", error);
    }
  };

  const handleButtonClick = async (code: string) => {
    setIsReplacementComplete(false); // Reset the replacement status
    try {
      await Promise.all([fetchCurrentMappingMaster(code), fetchCurrentRoutingRules(code)]);
      setIsReplacementComplete(true); // Mark replacement as complete
    } catch (error) {
      console.error("Error in completing tasks:", error);
    }
  };

  useEffect(() => {
    sp.setup({
      spfxContext: props.context as any,
    });

    // Fetch MappingMaster data on component mount
    void fetchMappingMasterData();
  }, [props.context]);

  useEffect(() => {
    if (parsedGlobalVariable?.appcode && MappingMasterState.length > 0) {
      // Call handleButtonClick only once when both appcode and MappingMasterState are ready
      void handleButtonClick(parsedGlobalVariable.appcode);
    }
  }, [parsedGlobalVariable?.appcode, MappingMasterState]); // Trigger only when appcode and MappingMasterState are ready

  return (
    <div style={{ display: "flex", flexDirection: "column" }}>
      <TopNavBar />
      <div style={{ display: "flex" }}>
        <SideBar activeMenu=""/>
        <div style={{ height: "100vh", overflowY: "scroll", width: "100%",backgroundColor:'#070d19' }}>
          <h4 className='FormName' style={{color:'white' ,fontSize:'20px', fontWeight:400}}>{parsedGlobalVariable?.appname}</h4>
          
          {/* Render workflow when replacement is complete */}
          {isReplacementComplete && (
            <Workflow
              context={props.context}
              RR={CurrentRoutingRulesState}
              MM={CurrentMappingMasterState}
            />
          )}
        </div>
      </div>
    </div>
  );
};

export default GetRoutingRules;