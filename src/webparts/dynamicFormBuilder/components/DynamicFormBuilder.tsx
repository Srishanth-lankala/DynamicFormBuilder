import React from 'react';
import { HashRouter, Route, Routes } from 'react-router-dom';
import { GlobalStateProvider } from './CustomFormBuilder/GlobalVariable/GlobalStateContext';
import { SPComponentLoader } from "@microsoft/sp-loader";
// import GetJsonFromSP from './CustomFormBuilder/GetJsonFromSP';
import CreateForm from './CustomFormBuilder/CreateForm';
import ConfigureWorkflow from './CustomFormBuilder/ConfigureWorkflow/ConfigureWorkflow';
import InProcess from './CustomFormBuilder/InProcessView';
import ViewForm from './CustomFormBuilder/ViewForm';
import Draft from './CustomFormBuilder/Draft';
import RejectedForm from './CustomFormBuilder/RejectedForm';
import RejectedView from './CustomFormBuilder/RejectedView';
import MappingMasterComponent from './CustomFormBuilder/MappingMasterComponent';
import FormDashBoard from './CustomFormBuilder/FormDashBoard/FormDashBoard';
import MyActionItems from './CustomFormBuilder/MyActionitems/MyActionItems';
import ViewByDate from './CustomFormBuilder/View/byDate';
import ViewByStatus from './CustomFormBuilder/View/byStatus';
import GetRoutingRules from './WorkFlowBuilder/GetRoutingRules';
import { IDynamicFormBuilderProps } from './IDynamicFormBuilderProps';
// import FormComponent from './CustomFormBuilder/FormComponent';
import App from './CustomFormBuilder/JSONForm/JsonForm';
import Layout from "./Layout";
import Reports from './CustomFormBuilder/AllFormReports/FormReports';
import GlobalDashboard from './CustomFormBuilder/Globaldashboard/GlobalDashboard';
import '../components/styles/theme.css'
import '../components/styles/common.css'
import FormComponent from './CustomFormBuilder/FormComponent';
import MyZone from './CustomFormBuilder/MyZone/MyZone';
import EditTransactionForm from './CustomFormBuilder/EditTransactionForm';
import ChildForm from './CustomFormBuilder/ChildForm/ChildForm';
import TemplateMapper from './CustomFormBuilder/InvoiceGenerator/Templatemappeing';
import InvoiceGenerator from './CustomFormBuilder/InvoiceGenerator/InvoiceGenerator';
import { FORM_ANSWERS, TABLE_ANSWERS, TIMESHEET_ANSWERS } from './CustomFormBuilder/InvoiceGenerator/mockData';
import formSchema from './CustomFormBuilder/InvoiceGenerator/formjson.json';
export default class DynamicFormBuilder extends React.Component<
  IDynamicFormBuilderProps,
  { showFormBuilder: boolean; loading: boolean }
> {
  constructor(props: IDynamicFormBuilderProps) {
    super(props);
    this.state = {
      showFormBuilder: false,
      loading: true,
    };
  }

  componentDidMount() {
    // Load external CSS
    SPComponentLoader.loadCss("https://use.fontawesome.com/releases/v5.13.0/css/all.css");
    SPComponentLoader.loadCss("https://stackpath.bootstrapcdn.com/bootstrap/4.4.1/css/bootstrap.min.css");

    // Hide SharePoint elements
    const styleElement = document.createElement("style");
    styleElement.innerHTML = `
      [data-automation-id="CanvasZone-SectionContainer"] { padding: 0px !important; min-width: 100% !important; }
      [data-automation-id="CanvasZone"] { margin-left: 0px !important; }
      #spCommandBar, #spSiteHeader, #SuiteNavWrapper { display: none !important; }
    `;
    document.head.appendChild(styleElement);

    document.getElementById("spSiteHeader")?.setAttribute("style", "display:none");
    document.getElementById("spLeftNav")?.setAttribute("style", "display:none");
    document.getElementById("spCommandBar")?.setAttribute("style", "display:none");

    const appBarElement = document.querySelector("#sp-appBar") as HTMLElement;
    if (appBarElement) {
      appBarElement.style.display = 'none';
    }
    const element = document.querySelector("#workbenchPageContent") as HTMLElement;
    if (element) {
      element.style.maxWidth = "100%";
    }

    const noscrollelement = document.querySelector('[data-automation-id="contentScrollRegion"]') as HTMLElement;
    console.log(noscrollelement);
    if (noscrollelement) {
      noscrollelement.style.overflow = "hidden";
    }



    setTimeout(() => {
      this.setState({ loading: false });
    }, 800);
  }

  public render(): React.ReactElement<IDynamicFormBuilderProps> {
    if (this.state.loading) {
      return (
        <div style={{
          height: '100vh',
          width: '100%',
          display: 'flex',
          justifyContent: 'center',
          alignItems: 'center',
          backgroundColor: '#f4f4f4',
          zIndex: 9999,
          position: 'fixed',
          top: 0,
          left: 0
        }}>
          <div className="spinner-border text-primary" role="status" style={{ width: '3rem', height: '3rem' }}>
            <span className="sr-only">Loading...</span>
          </div>
        </div>
      );
    }

    return (
      <GlobalStateProvider>
        <HashRouter>
          <Layout>
            <div className="Dashboard">
              <Routes>
                <Route path="/LoadForm" element={<GlobalDashboard context={this.props.context} />} />
                <Route path="/FormComponent" element={<FormComponent context={this.props.context} />} />
                <Route path="/CreateForm" element={<CreateForm context={this.props.context} />} />
                <Route path="/ConfigureWorkflow" element={<ConfigureWorkflow context={this.props.context} data={[]} />} />
                <Route path="/InProcess" element={<InProcess />} />
                <Route path="/viewform/:id" element={<ViewForm context={this.props.context} />} />
                <Route path="/RejectedView" element={<RejectedView />} />
                <Route path="/RejectedForm/:id" element={<RejectedForm />} />
                <Route path="/Draft" element={<Draft />} />
                <Route path="/MyActionItems" element={<MyActionItems />} />
                <Route path="/App" element={<App />} />
                <Route path="/EditForm/:id" element={<EditTransactionForm context={this.props.context} />} />
                <Route path="/mapping-master/:appCode" element={<MappingMasterComponent />} />
                <Route path="/FormDashBoard" element={<FormDashBoard />} />
                <Route path="/ViewByDate" element={<ViewByDate />} />
                <Route path="/ViewByStatus" element={<ViewByStatus />} />
                <Route path="/ReactDiagram" element={<GetRoutingRules context={this.props.context} />} />
                <Route path="/Reports" element={<Reports />} />
                <Route path="*" element={<GlobalDashboard context={this.props.context} />} />
                <Route path="/MyZone" element={<MyZone context={this.props.context} />} />
                <Route path="/CForm/:id" element={<ChildForm context={this.props.context} />} />
                <Route path="/TemplateMapper" element={<TemplateMapper />} />
                <Route path="/InvoiceGenerator" element={<InvoiceGenerator
                  answerData={FORM_ANSWERS}
                  tableData={TABLE_ANSWERS}
                  timesheetData={TIMESHEET_ANSWERS}
                  formSchema={formSchema}
                />} />

              </Routes>
            </div>
          </Layout>
        </HashRouter>
      </GlobalStateProvider>
    );
  }
}
