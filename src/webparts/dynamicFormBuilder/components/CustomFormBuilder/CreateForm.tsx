import * as React from "react";
import "react-form-builder2/dist/app.css";
import type { IDynamicFormBuilderProps } from '../IDynamicFormBuilderProps';
import { SPComponentLoader } from "@microsoft/sp-loader";
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
// import Sidebar from "./SideNavBar";
import TopNavBar from "./TopBar";
import "./Popup.css"
import withNavigate from './withNavigate';
import { Toast } from 'primereact/toast';
import { fetchTenantUser } from "./FetchTenantUser/fetchTenantUser";
import LoadingSpinner from './Loading';
import { dataservice } from "./encryptionutil";
import { fetchFormCount } from "./CustomHooks/TransactionCount";

export interface ICreateFormState {
  dataserviceobj: dataservice,
  messagesFromIframe: any[];
  postingbool: boolean;
  Editpostingbool: boolean;
  title: string;
  formPrefix: string,
  formType: string;
  formTypeOptions: string[]; // New state variable for options
  showPopup: boolean,
  popupError: boolean,
  formAccessOptions: string[],
  ApprovalRequired: string,
  ApprovalOptions: string[],
  accessOption: string,
  Domain: any,
  storedData: any,
  isSameDomain: string,
  loading: boolean,
  formDescription: string,
  isMultiForm: boolean,
  multiFormOptions: any[],
  selectedMultiFormAppCode: string
}

export interface IFrameMessage {
  type: string;
  payload: object;
  posting: boolean;
  Editposting: boolean;
  title: string;
  Domain: string;
}

class CreateForm extends React.Component<IDynamicFormBuilderProps & { navigate: Function }, ICreateFormState> {
  toast = React.createRef<Toast>();
  iframeRef: React.RefObject<HTMLIFrameElement>;
  constructor(props: IDynamicFormBuilderProps & { navigate: Function }) {
    super(props);
    this.iframeRef = React.createRef<HTMLIFrameElement>();
    this.state = {
      dataserviceobj: new dataservice(),
      messagesFromIframe: [],
      postingbool: false,
      Editpostingbool: false,
      title: "",
      formPrefix: "",
      formType: "IT", // Default form type
      formTypeOptions: [],// Initialize as an empty array
      showPopup: false,
      popupError: false,
      formAccessOptions: ["Yes", "No"],
      ApprovalOptions: ["Yes", "No"],
      ApprovalRequired: "Yes",
      accessOption: "No",
      Domain: "",
      loading: true,
      storedData: {
        isEdit: false,
        AppCode: "",
        AppName: "",
        FormType: "",
        FormJSON: "",
        Description: ""
      },
      isSameDomain: '',
      formDescription: '',
      isMultiForm: false,
      multiFormOptions: [],
      selectedMultiFormAppCode: ''
    };
  }

  async componentDidMount() {

    SPComponentLoader.loadCss("https://use.fontawesome.com/releases/v5.13.0/css/all.css");
    SPComponentLoader.loadCss("https://stackpath.bootstrapcdn.com/bootstrap/4.4.1/css/bootstrap.min.css");
    window.addEventListener('message', this.handleIframeMessage);
    sp.setup({ spfxContext: this.props.context as any });


    // get tenant domain name to filter form 
    const Domain = await fetchTenantUser();
    this.setState({ Domain: Domain.TenantUsers })


    // Fetch form type options from the FormType list
    try {
      const items = await sp.web.lists.getByTitle("FormType").items.select("Title").get();
      const options = items.map(item => item.Title);
      this.setState({ formTypeOptions: options });
    } catch (error) {
      this.handleError(error, "Error fetching form type options");
    }
  }

  // sendDataToIframe = () => {
  //   const storedData = JSON.parse(sessionStorage.getItem("EditFormData") || "{}")
  //   this.setState({ storedData: storedData })
  //   if (this.iframeRef.current && this.iframeRef.current.contentWindow && storedData.FormJSON) {
  //     // const parsedData = JSON.parse(storedData.FormJSON);
  //     const SendingData = {
  //       type: "EditForm_FROM_SPFX",
  //       payload: storedData

  //     };
  //     // Send the data to the iframe
  //     this.iframeRef.current.contentWindow.postMessage(
  //       SendingData, // Data being sent
  //       "*" // Target origin (Use specific origin for security in production)
  //     );
  //   }
  //   this.setState({ loading: false })
  // }

  sendDataToIframe = () => {
    if (!this.iframeRef.current || !this.iframeRef.current.contentWindow) return;

    // ---- Read sessionStorage ----
    const isMultiForm = sessionStorage.getItem("isMultiForm") === "true";
    const multiformedit = JSON.parse(sessionStorage.getItem("MultiFormEditData") || "[]");
    const allForms = JSON.parse(sessionStorage.getItem("FormMasterData") || "[]");
    const newFormFlag = sessionStorage.getItem("isNewForm") === "true"; // <-- NEW FLAG
    this.setState({ isMultiForm: isMultiForm });

    //console.log("SendDataToIframe - isMultiForm:", isMultiForm);
    //console.log("SendDataToIframe - MultiFormEditData exists:", !!multiformedit);
    //console.log("SendDataToIframe - isNewForm:", newFormFlag);

    const allformsdecrypted = allForms.map((item: any) => {
      //console.log("jsonincreate", item.FormJSON, typeof item.FormJSON);
      item.FormJSON = this.state.dataserviceobj.decryptjson(item.FormJSON);
      return item;
    })
    // ---- Always send master data ----
    this.iframeRef.current.contentWindow.postMessage(
      { type: "FormMasterData_FROM_SPFX", payload: allformsdecrypted },
      "*"
    );

    // ---- Multi-form logic ----
    if (isMultiForm) {
      if (!newFormFlag) {
        // Edit existing multi-form
        //console.log("Sending existing multi-form data:", multiformedit);
        //console.log("multiformedit parent form data >>>>>>>>>>>>>>>>>", multiformedit.parentForm);
        multiformedit.parentForm.FormJSON = JSON.parse(this.state.dataserviceobj.decryptjson(multiformedit.parentForm.FormJSON));
        multiformedit.childForms = multiformedit.childForms.map((item: any) => {
          //console.log("insidechild", item.FormJSON, typeof item.FormJSON)
          return { ...item, FormJSON: this.state.dataserviceobj.decryptjson(item.FormJSON) };
        })
        this.setState({ storedData: multiformedit.parentForm })

        //console.log("multipayload", multiformedit);

        this.iframeRef.current.contentWindow.postMessage(
          { type: "EditFormMulti_FROM_SPFX", payload: multiformedit },
          "*"
        );
      } else {
        // New multi-form -> send empty array
        const defaultChildForms = {
          parentForm: { id: 1, name: 'Parent Form', data: [] },
          childForms: [
            { id: 2, name: 'Child Form 1', data: [] },
            { id: 3, name: 'Child Form 2', data: [] }
          ]
        };
        //console.log("New multi-form: sending empty payload", defaultChildForms);
        this.iframeRef.current.contentWindow.postMessage(
          { type: "EditFormMulti_FROM_SPFX", payload: defaultChildForms },
          "*"
        );
      }
    } else {
      // ---- Single form logic ----
      const Formmasteritem = JSON.parse(sessionStorage.getItem("EditFormData") || "{}");

      Formmasteritem.FormJSON = Formmasteritem.FormJSON ? JSON.parse(this.state.dataserviceobj.decryptjson(Formmasteritem.FormJSON)) : ""

      //console.log("FormMasterItem", Formmasteritem);
      const storedData = Formmasteritem;
      this.setState({ storedData: storedData })

      if (storedData && storedData.FormJSON) {
        this.iframeRef.current.contentWindow.postMessage(
          { type: "EditForm_FROM_SPFX", payload: storedData },
          "*"
        );
      }
    }

    // ---- Send multi/single form flag ----
    this.iframeRef.current.contentWindow.postMessage(
      { type: "IS_MULTI_FORM_FROM_SPFX", payload: isMultiForm },
      "*"
    );

    this.setState({ loading: false });
  };



  componentWillUnmount() {
    window.removeEventListener('message', this.handleIframeMessage);
  }

  handleIframeMessage = (event: MessageEvent): void => {
    try {
      if (event.data.type === 'FROM_IFRAME') {
        const jsonData: IFrameMessage[] = event.data.payload;
        this.setState({
          messagesFromIframe: jsonData,
          postingbool: event.data.posting,
          Editpostingbool: event.data.Editposting,
          title: event.data.title
        });
      } else if (event.data.type === 'REQUEST_FORM_JSON') {
        // Handle request from iframe for specific form JSON
        //console.log("Received REQUEST_FORM_JSON from iframe:", event.data.payload);
        void this.fetchAndSendFormJSON(event.data.payload.formId);
      } else {
        throw new Error("Invalid message type or payload");
      }
    } catch (error) {
      // this.handleError(error, "Error processing iframe message");
      console.error("Error processing iframe message : ", error)
    }
  };

  fetchAndSendFormJSON = async (formId: string) => {
    try {
      //console.log("📥 Fetching form JSON for formId:", formId);

      // First check if form exists in multiFormOptions (already loaded)
      let selectedForm = this.state.multiFormOptions.find((form: any) =>
        form.AppCode === formId || form.ID === formId || form.AppName === formId
      );

      // If not found in multiFormOptions, fetch from database
      if (!selectedForm) {
        //console.log("🔍 Not in cache, fetching from FormMaster database...");

        // Check if formId is numeric (ID) or string (AppCode/AppName)
        const isNumeric = !isNaN(Number(formId));
        let filterQuery = '';

        if (isNumeric) {
          filterQuery = `ID eq ${formId}`;
        } else {
          filterQuery = `AppCode eq '${formId}' or AppName eq '${formId}'`;
        }

        const formItems = await sp.web.lists
          .getByTitle("FormMaster")
          .items
          .filter(filterQuery)
          .select("ID", "AppCode", "AppName", "FormJSON", "FormType", "Description")
          .top(1)
          .get();
        formItems[0].FormJSON = this.state.dataserviceobj.decryptjson(formItems[0].FormJSON);


        if (formItems.length > 0) {
          selectedForm = formItems[0];
          //console.log("✅ Found form in database:", selectedForm.AppName);
        }
      } else {
        //console.log("✅ Found form in cache:", selectedForm.AppName);
      }

      if (selectedForm && this.iframeRef.current && this.iframeRef.current.contentWindow) {
        const formData = {
          type: "SELECTED_FORM_JSON_FROM_SPFX",
          payload: {
            AppCode: selectedForm.AppCode,
            AppName: selectedForm.AppName,
            FormJSON: JSON.parse(selectedForm.FormJSON),
            FormType: selectedForm.FormType,
            Description: selectedForm.Description || ""
          }
        };

        console.log("📤 Sending form to iframe:", 
          {
          AppName: selectedForm.AppName,
          AppCode: selectedForm.AppCode,
          FormJSONLength: formData.payload.FormJSON.length
        });

        this.iframeRef.current.contentWindow.postMessage(
          formData,
          "*" // TODO: replace "*" with exact iframe origin
        );
      } else {
        console.warn("❌ Form not found for formId:", formId);
      }
    } catch (error) {
      console.error("❌ Error fetching and sending form JSON:", error);
    }
  };

  // handleMultiFormSelection = async (event: { target: { value: string } }) => {
  //   const selectedAppCode = event.target.value;
  //   this.setState({ selectedMultiFormAppCode: selectedAppCode });

  //   if (!selectedAppCode) return; // Don't process empty selection

  //   try {
  //     // Fetch fresh data from FormMaster database
  //     console.log("Fetching FormJSON from FormMaster for AppCode:", selectedAppCode);

  //     const formItems = await sp.web.lists
  //       .getByTitle("FormMaster")
  //       .items
  //       .filter(`AppCode eq '${selectedAppCode}'`)
  //       .select("ID", "AppCode", "AppName", "FormJSON", "FormType", "Description")
  //       .top(1)
  //       .get();

  //     if (formItems.length > 0 && this.iframeRef.current && this.iframeRef.current.contentWindow) {
  //       const selectedForm = formItems[0];
  //       const formData = {
  //         type: "SELECTED_FORM_JSON_FROM_SPFX",
  //         payload: {
  //           AppCode: selectedForm.AppCode,
  //           AppName: selectedForm.AppName,
  //           FormJSON: JSON.parse(selectedForm.FormJSON),
  //           FormType: selectedForm.FormType,
  //           Description: selectedForm.Description || ""
  //         }
  //       };

  //       console.log("Fetched FormJSON from FormMaster:", selectedForm.FormJSON);
  //       console.log("Sending selected form to iframe:", formData);

  //       // Send the selected form JSON to iframe
  //       this.iframeRef.current.contentWindow.postMessage(
  //         formData,
  //         "*" // TODO: replace "*" with exact iframe origin
  //       );
  //     } else {
  //       console.warn("Form not found in FormMaster for AppCode:", selectedAppCode);
  //     }
  //   } catch (error) {
  //     console.error("Error handling multi-form selection:", error);
  //   }
  // };


  private generateUniqueAppCode(appName: string): string {
    const letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
    const randomLetters = Array.from({ length: 3 }, () => letters[Math.floor(Math.random() * letters.length)]).join("");
    const randomDigits = Math.floor(1000 + Math.random() * 9000).toString();
    const namePrefix = appName.slice(0, 3).toUpperCase().padEnd(3, 'X');

    const combinedString = randomDigits + randomLetters;
    const shuffledString = combinedString.split('').sort(() => Math.random() - 0.5).join('');

    return `${namePrefix}-${shuffledString}`;
  }

  private async createListItem(data: any, appName: string): Promise<void> {
    try {
      //console.log("datain cretelistitem", data)
      const appCode = this.generateUniqueAppCode(appName);
      if (this.state.formType === "") {
        // alert("Please Select Form Type");
        this.toast.current?.show({
          severity: 'error',
          summary: "",
          detail: 'Please Select Form Type',
          life: 1000,
        });
      } else {
        //-------------if immediate manager approval required,adding two records to mapping master----------------------
        const encrypteddata = this.state.dataserviceobj.encryptjson(data);
        if (this.state.accessOption === "Yes") {

          const groupName = "All Employees";
          const group = await sp.web.siteGroups.getByName(groupName).get();
          await sp.web.lists.getByTitle("FormMaster").items.add({
            AppName: appName.trim(),
            AppCode: appCode,
            FormJSON: encrypteddata,
            FormType: this.state.formType,
            Everyone: this.state.accessOption,
            Domain: this.state.Domain,
            TransactionCount: 0,
            FormPrefix: this.state.formPrefix,
            Description: this.state.formDescription,
            FormViewerId: group.Id,
            TimeSheetControl: data?.some((item: any) => item?.text === "Timesheet") ?? false
          });

          const item1 = {
            Role: "Creator",
            Level: 0,
            Users_x002f_GroupsId: group.Id,
            AppName: appName.trim(),
            AppCode: appCode,
          };

          const item2 = {
            Role: "Approver",
            Level: 1,
            AppName: appName.trim(),
            AppCode: appCode,
          };
          Promise.all([
            sp.web.lists.getByTitle("MappingMaster").items.add(item1),
            sp.web.lists.getByTitle("MappingMaster").items.add(item2),
          ])
            .then(() => {
              // Both items added successfully
            })
            .catch((err) => {
              console.error("Error adding one or both items", err);
            });


        } //-------------if immediate manager approval required,adding two records to mapping master----------------------
        else {
          await sp.web.lists.getByTitle("FormMaster").items.add({
            AppName: appName.trim(),
            AppCode: appCode,
            FormJSON: encrypteddata,
            FormType: this.state.formType,
            Everyone: this.state.accessOption,
            Domain: this.state.Domain,
            TransactionCount: 0,
            isWorkflowRequired: this.state.ApprovalRequired,
            FormPrefix: this.state.formPrefix,
            Description: this.state.formDescription,
            TimeSheetControl: data?.some((item: any) => item?.text === "Timesheet") ?? false
          });

        }


        // alert("");
        this.toast.current?.show({
          severity: 'success',
          summary: "",
          detail: 'Item created successfully!',
          life: 1000,
        });
      }
    } catch (error) {
      this.handleError(error, "Error creating list item");
    }
  }



  private handleError(error: unknown, context: string): void {
    if (error instanceof Error) {
      console.error(`${context}:`, error);
      alert(`${context}: ${error.message}`);
    } else {
      console.error(`${context}: An unknown error occurred`);
      alert(`${context}: An unknown error occurred`);
    }
  }
  showPopup = () => {
    this.setState({ showPopup: true });
  };
  closePopup = () => {
    this.setState({
      showPopup: false,
      title: '',
      formType: '',
      popupError: false,
      postingbool: false,
      accessOption: 'No',
      ApprovalRequired: ''
    });

  };
  closeUpdatePopup = () => {
    this.setState({
      Editpostingbool: false,
    });
  };

  updateAppCodeVersion = (appCode: string): string => {
    // Check if AppCode already has a version suffix (e.g., "/1", "/2", etc.)
    const match = appCode.match(/\/(\d+)$/);

    if (match) {
      // Extract and increment version number
      const currentVersion = parseInt(match[1], 10);
      return `${appCode.replace(/\/\d+$/, '')}/${currentVersion + 1}`;
    } else {
      // If no version exists, start with "/1"
      return `${appCode}/1`;
    }
  };
  stripHtmlTags = (label: any) => { return label?.replace(/<\/?[^>]+(>|$)/g, "")?.trim() || ""; }
  areLabelsUnique = (jsonArray: any) => {
    const labels: any = [];

    // Iterate through each object in the array
    for (const item of jsonArray) {
      // Check if the object has a 'label' property
      if (item?.label) {
        // If label already exists in our array, return false (not unique)
        if (labels.includes(this.stripHtmlTags(item.label))) {
          return false;
        }
        // Otherwise add it to our tracking array
        labels.push(this.stripHtmlTags(item.label));
      }
    }

    // If we made it through all objects without duplicates, return true
    return true;
  }

  handleUpdate = async () => {
    try {
      const FormMasterItem = await sp.web.lists.getByTitle("FormMaster").items
        .filter(`AppCode eq '${this.state.storedData.AppCode}'`)
        .top(1)
        .get();
      if (FormMasterItem.length === 0) {
        console.error("No matching FormMasterItem found.");
        return;
      }
      FormMasterItem[0].FormJSON = this.state.dataserviceobj.decryptjson(FormMasterItem[0].FormJSON)
      const FormMasterItemId = FormMasterItem[0].ID;
      const WorkflowProcessItems = await sp.web.lists.getByTitle("WorkFlowProcessData").items
        .filter(`AppCode eq '${this.state.storedData.AppCode}'`)
        .select("ID")
        .top(1) // Only fetch one item if exists
        .get();
      if (FormMasterItem[0].FormJSON === JSON.stringify(this.state.messagesFromIframe)) {
        this.toast.current?.clear()
        this.toast.current?.show({
          severity: 'error',
          summary: "No Updates",
          detail: `Please update the form`,
          life: 2000,
          style: { height: 70 }
        });
        return;
      }
      const isUnique = this.areLabelsUnique(this.state.messagesFromIframe);
      if (!isUnique) {
        this.toast.current?.clear()
        this.toast.current?.show({
          severity: 'error',
          summary: 'Duplicate Field Names',
          detail: 'All fields in the form should have unique names',
          life: 3000,
          style: { height: 80 }
        });
        return;
      } //console.log("inupdate", this.state.messagesFromIframe)
      const encryptedmessagesfromiframe = this.state.dataserviceobj.encryptjson(this.state.messagesFromIframe);
      if (WorkflowProcessItems.length > 0) {
        const OldFormItem = FormMasterItem[0]
        const NewAppCode = this.updateAppCodeVersion(this.state.storedData.AppCode);
        await sp.web.lists.getByTitle("FormMaster").items.add({
          AppName: OldFormItem.AppName,
          AppCode: NewAppCode,
          FormJSON: encryptedmessagesfromiframe,
          FormType: OldFormItem.FormType,
          Domain: OldFormItem.Domain,
          FormPrefix: OldFormItem.FormPrefix,
          Description: OldFormItem.Description,
          TransactionCount: OldFormItem.TransactionCount,
          FormViewerId: OldFormItem.FormViewerId,
          DynamicStatuses: OldFormItem.DynamicStatuses,
          SLA: OldFormItem.SLA,
          SLAinhours: OldFormItem.SLAinhours,
          isWorkflowRequired: OldFormItem.isWorkflowRequired,
          TimeSheetControl: this.state.messagesFromIframe?.some((item: any) => item?.text === "Timesheet") ?? false

        });
        try {
          await sp.web.lists.getByTitle("FormMaster").items.getById(FormMasterItemId).update({
            VisibilityFlag: false
          });
        } catch (error) {
          console.error("Error updating VisibilityFlag:", error);
        }

        this.toast.current?.show({
          severity: 'success',
          summary: "",
          detail: `Form updated succesfully`,
          life: 1000,
        });
        this.closeUpdatePopup()
        setTimeout(() => {
          this.props.navigate('/LoadForm');
        }, 1000);
        const MappingMaster = await sp.web.lists.getByTitle("MappingMaster").items
          .filter(`AppCode eq '${this.state.storedData.AppCode}'`)
          .get()

        if (MappingMaster.length > 0) {
          try {
            await Promise.all(
              MappingMaster.map(item =>
                sp.web.lists.getByTitle("MappingMaster").items.getById(item.Id).update({ AppCode: NewAppCode })
              )
            );
          } catch (error) {
            console.error("Error updating MappingMaster:", error);
          }
        }


        const RoutingRules = await sp.web.lists.getByTitle("RoutingRules").items
          .filter(`AppCode eq '${this.state.storedData.AppCode}'`)
          .getAll()

        if (RoutingRules.length > 0) {
          try {
            await Promise.all(
              RoutingRules.map(item =>
                sp.web.lists.getByTitle("RoutingRules").items.getById(item.Id).update({ AppCode: NewAppCode })
              )
            );
          } catch (error) {
            console.error("Error updating RoutingRules:", error);
          }
        }

      }
      else {
        try {
          if (FormMasterItemId) {
            await sp.web.lists.getByTitle("FormMaster").items.getById(FormMasterItemId).update({
              FormJSON: encryptedmessagesfromiframe,
              TimeSheetControl: this.state.messagesFromIframe?.some((item: any) => item?.text === "Timesheet") ?? false
            });
            this.toast.current?.show({
              severity: 'success',
              summary: "Form Update",
              detail: `Form updated succesfully`,
              life: 1000,
            });
            this.closeUpdatePopup()
            setTimeout(() => {
              this.props.navigate('/LoadForm');
            }, 1000);
            try {
              const conditionListItems = await sp.web.lists.getByTitle("ConditionsList").items
                .filter(`AppCode eq '${this.state.storedData.AppCode}'`)
                .get();

              if (conditionListItems.length > 0) {
                await Promise.all(
                  conditionListItems.map(item =>
                    sp.web.lists.getByTitle("ConditionsList").items.getById(item.Id).delete()
                  )
                );
              }
            } catch (error) {
              console.error("Error deleting records from ConditionsList:", error);
            }

          }
        } catch (error) {
          console.error("Error in Form Update FormMasterItem", error)
        }
      }
    } catch (error) {
      console.error("Error in handleUpdate:", error);
    }
  }
  handlemultiUpdate = async () => {
    debugger;
    try {
      const multiFormData: any = this.state.messagesFromIframe;
      //console.log("🛠️ Starting Multi-Form Update", multiFormData);

      if (!multiFormData?.isMultiForm ||
        !multiFormData?.parentForm ||
        !Array.isArray(multiFormData?.childForms)) {
        console.error("Invalid multiFormData structure:", multiFormData);
        return;
      }
      //console.log("🛠️ Starting Multi-Form Update", multiFormData);

      const parentForm = multiFormData.parentForm;
      const childForms = multiFormData.childForms;

      if (!parentForm || !childForms) {
        // console.error("Invalid multiFormData structure:", multiFormData);
        return;
      }

      // ✅ Fetch Parent FormMaster item using storedData.AppCode (the old one)
      const { storedData } = this.state;
      const oldParentAppCode = storedData.AppCode;
      const FormMaster = sp.web.lists.getByTitle("FormMaster");

      const existingParent = await FormMaster.items
        .filter(`AppCode eq '${oldParentAppCode}'and IsParentForm eq 1`)
        .top(1)
        .get();

      if (existingParent.length === 0) {
        // console.error("❌ No matching parent form found for update.");
        return;
      }

      const oldParent = existingParent[0];
      const newParentAppCode = this.updateAppCodeVersion(oldParentAppCode);

      console.log("🔁 Updating Multi-Form Parent:", {
        oldParentAppCode,
        newParentAppCode
      });

      // ✅ Create new Parent version
      await FormMaster.items.add({
        AppName: oldParent.AppName,
        AppCode: newParentAppCode,
        FormJSON: this.state.dataserviceobj.encryptjson(parentForm.data),
        FormType: oldParent.FormType,
        Domain: oldParent.Domain,
        FormPrefix: oldParent.FormPrefix,
        Description: oldParent.Description,
        IsParentForm: true,
        TransactionCount: oldParent.TransactionCount,
        TimeSheetControl:
          parentForm.data?.some((item: any) => item?.text === "Timesheet") ?? false
      });

      // Hide old parent version
      await FormMaster.items.getById(oldParent.ID).update({
        VisibilityFlag: false
      });

      // ✅ Update or Add Child Forms
      for (const [index, childForm] of childForms.entries()) {
        const childOrder = childForm.order || index + 1;
        const oldChildAppCode = `${oldParentAppCode}-${childOrder}C`;
        const newChildAppCode = `${newParentAppCode}-${childOrder}C`;

        const existingChild = await FormMaster.items
          .filter(`ParentAppCode eq '${oldParentAppCode}' and ChildOrder eq ${childOrder}`)
          .top(1)
          .get();

        if (existingChild.length > 0) {
          // 🔁 Create new version of child
          const oldChild = existingChild[0];
          await FormMaster.items.add({
            AppName: oldChild.AppName,
            AppCode: newChildAppCode,
            FormJSON: this.state.dataserviceobj.encryptjson(childForm.data),
            FormType: oldChild.FormType,
            Domain: oldChild.Domain,
            FormPrefix: oldChild.FormPrefix,
            Description: oldChild.Description,
            IsParentForm: false,
            ParentAppCode: newParentAppCode,
            ChildOrder: childOrder,
            TimeSheetControl:
              childForm.data?.some((item: any) => item?.text === "Timesheet") ??
              false,
          });

          // Hide old version
          // await FormMaster.items.getById(oldChild.ID).update({
          //   VisibilityFlag: false
          // });

          console.log(`🔁 Updated child ${childOrder} (${oldChildAppCode})`);
        } else {
          // ➕ New child added by user
          //console.log(`🆕 Creating new child ${childOrder}`);
          await FormMaster.items.add({
            AppName: childForm.name,
            AppCode: newChildAppCode,
            FormJSON: this.state.dataserviceobj.encryptjson(childForm.data),
            FormType: oldParent.FormType,
            Domain: oldParent.Domain,
            FormPrefix: oldParent.FormPrefix,
            Description: oldParent.Description,
            IsParentForm: false,
            ParentAppCode: newParentAppCode,
            ChildOrder: childOrder,
            TimeSheetControl:
              childForm.data?.some((item: any) => item?.text === "Timesheet") ??
              false
          });
        }
      }

      this.toast.current?.show({
        severity: "success",
        summary: "Multi-Form Updated",
        detail: "All parent and child forms updated successfully.",
        life: 2500
      });

      this.closeUpdatePopup();
      setTimeout(() => this.props.navigate("/LoadForm"), 1000);
    } catch (error) {
      console.error("❌ Error in handleMultiFormUpdate:", error);
      this.toast.current?.show({
        severity: "error",
        summary: "Update Failed",
        detail: "Error updating multi-form data.",
        life: 3000
      });
    }
  };

  // Helper function to normalize the form name (remove spaces and convert to lowercase)
  private normalizeFormName(formName: string): string {
    return formName.replace(/\s+/g, '').toLowerCase(); // Remove spaces and convert to lowercase
  }

  handleSave = async () => {
    const {formcount , limitReached}=await fetchFormCount(this.state.Domain);
    if(limitReached){
      this.toast.current?.clear()
      this.toast.current?.show({
        severity: 'warn',
        summary: 'Freemium subscription limit reached',
        detail: `Form creation limit ${formcount} exceeded. Please upgrade your license.`,
        life: 3000,
        style: { height: 90 }
      });
      return;
    }
    if (this.state.isSameDomain === "NO") {
      this.setState({ accessOption: "No" })
    }

    const { title, formType, formPrefix, messagesFromIframe, accessOption } = this.state;

    // ✅ Check if multi-form data
    const isMultiFormData = messagesFromIframe &&
      typeof messagesFromIframe === 'object' &&
      !Array.isArray(messagesFromIframe) &&
      Object.prototype.hasOwnProperty.call(messagesFromIframe, 'isMultiForm') &&
      (messagesFromIframe as any).isMultiForm === true;

    if (isMultiFormData) {
      // ✅ Handle Multi-Form Save
      await this.handleMultiFormSave(messagesFromIframe as any);
      return;
    }

    // ✅ Single Form Save Logic (existing)
    const isUnique = this.areLabelsUnique(messagesFromIframe as any[]);
    if (!isUnique) {
      this.toast.current?.clear()
      this.toast.current?.show({
        severity: 'error',
        summary: 'Duplicate Field Names',
        detail: 'All fields in the form should have unique names',
        life: 3000,
        style: { height: 80 }
      });
      return;
    }
    if (title && title.trim() !== "" && formType && messagesFromIframe.length && accessOption.length > 0 && formPrefix.length && formPrefix.length && !(/[\\\][/*:?-]/.test(title))) {
      this.setState({ popupError: false });

      const normalizedTitle = this.normalizeFormName(title.trim());
      this.setState({ loading: true });

      if (formPrefix.length < 2 || formPrefix.length > 3) {
        this.toast.current?.show({
          severity: 'warn',
          summary: 'Invalid Form Prefix',
          detail: 'Form prefix should be 2 or 3 characters only.',
          life: 3000,
          style: { height: 90 }
        });
        return;
      }

      try {
        const existingForms = await sp.web.lists
          .getByTitle("FormMaster")
          .items.select("AppName", "Domain", "FormPrefix")
          .get();

        const matchedForms = existingForms.filter(item =>
          this.normalizeFormName(item.AppName) === normalizedTitle
        );

        const formExistsWithSameDomain = matchedForms.some(item =>
          item.Domain?.trim().toLowerCase() === this.state.Domain?.trim().toLowerCase()
        );

        const matchedFormPrefix = existingForms.filter(item =>
          item.FormPrefix === formPrefix
        );

        const formPrefixExistsWithSameDomain = matchedFormPrefix.some(item =>
          item.Domain?.trim().toLowerCase() === this.state.Domain?.trim().toLowerCase()
        );

        if (formExistsWithSameDomain) {
          this.toast.current?.show({
            severity: 'error',
            summary: 'Form Name Already Exists',
            detail: 'A form with this name already exists, Please choose a different name.',
            life: 3000,
            style: { height: 90 }
          });
        }
        else if (formPrefixExistsWithSameDomain) {
          this.toast.current?.show({
            severity: 'error',
            summary: 'Form Prefix Already Exists',
            detail: 'A form with this prefix already exists, Please choose a different form prefix.',
            life: 3000,
            style: { height: 90 }
          });
        }
        else {
          await this.createListItem(messagesFromIframe, title);
          this.setState({
            showPopup: false,
            postingbool: true,
            title: '',
            formType: '',
            accessOption: 'No',
            popupError: false,
            formDescription: '',
            ApprovalRequired: ''
          });

          this.closePopup();

          setTimeout(() => {
            this.props.navigate('/ConfigureWorkflow');
          }, 2000);
        }
      } catch (error) {
        this.handleError(error, "Error checking if form name exists");
      }

    } else {
      this.setState({ popupError: true });
    }
    this.setState({ loading: false })
  };

  // ✅ New method for Multi-Form Save
  handleMultiFormSave = async (multiFormData: any) => {
    try {
      debugger;
      const { title, formType, formPrefix, formDescription } = this.state;
      // Validate required fields
          if (!title || !formType || !formPrefix) {
            this.toast.current?.show({
              severity: 'error',
              summary: 'Missing Required Fields',
              detail: 'Please fill in Form Name, Category, and Prefix.',
              life: 3000,
            });
            return;
          }

          if (formPrefix.length < 2 || formPrefix.length > 3) {
            this.toast.current?.show({
              severity: 'warn',
              summary: 'Invalid Form Prefix',
              detail: 'Form prefix should be 2 or 3 characters only.',
              life: 3000,
            });
            return;
          }

      const normalizedTitle = this.normalizeFormName(title.trim());

        try {
          const existingForms = await sp.web.lists
            .getByTitle("FormMaster")
            .items.select("AppName", "Domain", "FormPrefix")
            .get();

          const matchedForms = existingForms.filter(item =>
            this.normalizeFormName(item.AppName) === normalizedTitle
          );

          const formExistsWithSameDomain = matchedForms.some(item =>
            item.Domain?.trim().toLowerCase() === this.state.Domain?.trim().toLowerCase()
          );

          const matchedFormPrefix = existingForms.filter(item =>
            item.FormPrefix === formPrefix
          );

          const formPrefixExistsWithSameDomain = matchedFormPrefix.some(item =>
            item.Domain?.trim().toLowerCase() === this.state.Domain?.trim().toLowerCase()
          );
          if (formExistsWithSameDomain) {
          this.toast.current?.show({
            severity: 'error',
            summary: 'Form Name Already Exists',
            detail: 'A form with this name already exists, Please choose a different name.',
            life: 3000,
            style: { height: 90 }
          });
          return;          
        }
        else if (formPrefixExistsWithSameDomain) {
          this.toast.current?.show({
            severity: 'error',
            summary: 'Form Prefix Already Exists',
            detail: 'A form with this prefix already exists, Please choose a different form prefix.',
            life: 3000,
            style: { height: 90 }
          });
          return;
        }

      }catch (error) {
        this.handleError(error, "Error checking if form name exists");
      }
          
          this.setState({ loading: true });

          // Generate Parent AppCode
          const parentAppCode = this.generateUniqueAppCode(title);
          //console.log("📝 Creating Multi-Form - Parent AppCode:", parentAppCode);

          // ✅ Create Parent Form
          await sp.web.lists.getByTitle("FormMaster").items.add({
            AppName: title.trim(),
            AppCode: parentAppCode,
            FormJSON: this.state.dataserviceobj.encryptjson(multiFormData.parentForm.data),
            FormType: formType,
            Domain: this.state.Domain,
            FormPrefix: formPrefix,
            Description: formDescription,
            IsParentForm: true,  // ✅ Mark as parent
            TransactionCount: 0,
            TimeSheetControl: multiFormData.parentForm.data?.some((item: any) => item?.text === "Timesheet") ?? false
          });

          //console.log("✅ Parent form created successfully");
          //console.log("creatingchildfrommulti", multiFormData)
          // ✅ Create Child Forms
          for (const [index, childForm] of multiFormData.childForms.entries()) {
            try {
              const childOrder = childForm.order || index + 1;
              const childAppCode = `${parentAppCode}-${childOrder}C`;
              const templateAppCode = childForm.TEMPLATEAPPCODE || "";

              console.log(`📝 Creating child form ${index + 1}:`, {
                name: childForm.name,
                childAppCode: childAppCode,
                parentAppCode: parentAppCode,
                childOrder: childForm.order,
                templateAppCode: templateAppCode
              });

              const childData: any = {
                AppName: childForm.name,
                AppCode: childAppCode,
                FormJSON: this.state.dataserviceobj.encryptjson(childForm.data),
                FormType: formType,
                Domain: this.state.Domain,
                FormPrefix: formPrefix,
                Description: formDescription || "",
                IsParentForm: false,
                ParentAppCode: parentAppCode,  // ✅ Link to parent
                ChildOrder: childForm.order,    // ✅ Order from array (1, 2, 3...)
                TransactionCount: 0,
                TimeSheetControl: childForm.data?.some((item: any) => item?.text === "Timesheet") ?? false
              };

              debugger;

              // ✅ Only add TemplateAppCode if it has a value
              if (templateAppCode && templateAppCode.trim() !== "") {
                childData.TemplateAppCode = templateAppCode;
                //console.log(`✅ Adding TemplateAppCode: ${templateAppCode}`);
              }
              debugger;

              //console.log(`💾 Child form data being saved:`, childData);

              await sp.web.lists.getByTitle("FormMaster").items.add(childData);

              //console.log(`✅ Child form ${index + 1} saved successfully - ${childForm.name} (AppCode: ${childAppCode}, Template: ${templateAppCode})`);
            } catch (childError) {
              console.error(`❌ Error saving child form ${index + 1}:`, childError);
              throw childError;
            }
          }

          this.toast.current?.show({
            severity: 'success',
            summary: 'Multi-Form Created',
            detail: `Parent form and ${multiFormData.childForms.length} child forms created successfully!`,
            life: 3000,
          });

          this.setState({
            showPopup: false,
            postingbool: false,
            title: '',
            formType: '',
            formPrefix: '',
            formDescription: '',
            loading: false
          });

          this.closePopup();

          setTimeout(() => {
            this.props.navigate('/ConfigureWorkflow');
          }, 2000);

        } catch (error) {
          console.error("❌ Error saving multi-form:", error);
          this.toast.current?.show({
            severity: 'error',
            summary: 'Save Failed',
            detail: 'Error saving multi-form data.',
            life: 3000,
          });
          this.setState({ loading: false });
        }
      };


      handleTextInputChange = (event: { target: { value: any; }; }) => {
        this.setState({ title: event.target.value });
      };
      handleTextInputChangedescription = (event: { target: { value: any; }; }) => {
        this.setState({ formDescription: event.target.value });
      };
      handleFormPrefix = (event: { target: { value: any } }) => {
        this.setState({ formPrefix: event.target.value.toUpperCase() });
      };
      handleWorkflowChange = (event: { target: { value: any; }; }) => {
        this.setState({
          ApprovalRequired: event.target.value,
          accessOption: 'No'
        });
      };

      handleDropdownChange = (event: { target: { value: any; }; }) => {
        this.setState({ formType: event.target.value });
      };

      handleEveryOne = (event: { target: { value: any; }; }) => {
        this.setState({ accessOption: event.target.value });
      };

      handleUserDomain = (UserDomain: string) => {
        this.setState({ isSameDomain: UserDomain })
      };


  public render(): React.ReactElement<IDynamicFormBuilderProps> {
    try {
      return (

        <div style={{ display: 'flex', flexDirection: 'column' }}>
          <Toast ref={this.toast} />

          {
            this.state.postingbool && (
              <div className="custom-dark-modal-overlay" >
                <div className="custom-dark-modal" style={{ minWidth: '367px' }} >

                  {/* <div className="savepopup modal-content" > */}
                  {/* <div className="d-flex justify-content-between"> */}
                  {/* <h5 className="modal-title" style={{ fontWeight: 'bold' }}>Save Form</h5> */}
                  <h2 className="addsteptitle">Save Form</h2>
                  {/* <button type="button" className="close" onClick={this.closePopup}>
                        <span>&times;</span>
                      </button>
                    </div> */}
                  <div className="form-group" style={{ padding: '0px 20px' }}>
                    <div className="form-control-group">
                      <label className="new-label">
                        Form Name <span style={{ color: 'red' }}>*</span>
                      </label>
                      <input
                        type="text"
                        className=" formcontrol-extra form-control"
                        value={this.state.title}

                        onChange={this.handleTextInputChange}
                        placeholder="Enter your form name (e.g. Travel Request) "
                        required
                      />
                    </div>

                    <div className="form-control-group">
                      <label className="new-label">
                        Form Description <span style={{ color: 'red' }}>*</span>
                      </label>
                      <input
                        type="text"
                        className="form-control formcontrol-extra"
                        value={this.state.formDescription}

                        onChange={this.handleTextInputChangedescription}
                        required
                      />
                    </div>


                    <div className="form-control-group">
                      <label className="new-label">
                        Form Title Prefix <span style={{ color: 'red' }}>*</span>
                      </label>
                      <input
                        type="text"
                        className="form-control formcontrol-extra"
                        value={this.state.formPrefix}

                        onChange={this.handleFormPrefix}
                        placeholder="Enter 2–3 letter prefix (e.g. TR, ITE )"
                        required
                      />
                    </div>

                    <div className="form-control-group">
                      <label className="new-label">
                        Form Category <span style={{ color: 'red' }}>*</span>
                      </label>
                      <select
                        className="form-select formcontrol-extra"
                        value={this.state.formType}
                        onChange={this.handleDropdownChange}

                        required
                      >
                        <option value="">Select</option>
                        {this.state.formTypeOptions.map((option, index) => (
                          <option key={index} value={option}>{option}</option>
                        ))}
                      </select>
                    </div>
                    {/* ✅ Hide approval options for multi-form */}
                    {!this.state.isMultiForm && (
                      <>
                        <div className="form-control-group">
                          <label className="new-label">
                            Is Approval Required ? <span style={{ color: 'red' }}>*</span>
                          </label>
                          <select
                            className="form-select formcontrol-extra"
                            value={this.state.ApprovalRequired}
                            onChange={this.handleWorkflowChange}
                            required
                          >
                            <option value="">Select</option>
                            {this.state.ApprovalOptions.map((option, index) => (
                              <option key={index} value={option}>{option}</option>
                            ))}
                          </select>
                        </div>

                        {/* ----------------------------- View IMMEDIATE MANAGER  ----------------------------------------------- */}
                        {this.state.ApprovalRequired === "Yes" &&
                          (this.state.isSameDomain === "YES" ?
                            <div className="form-control-group">
                              <label className="new-label">
                                Immediate Manager Approval Required <span style={{ color: 'red' }}>*</span>
                              </label>
                              <select
                                className="form-select formcontrol-extra"
                                value={this.state.accessOption}
                                onChange={this.handleEveryOne}
                                required
                              >
                                <option value="">Select</option>
                                {this.state.formAccessOptions.map((option, index) => (
                                  <option key={index} value={option}>{option}</option>
                                ))}
                              </select>
                            </div> :
                            <div className="form-control-group">
                              <label className="new-label">
                                Immediate Manager Approval Required <span style={{ color: 'red' }}>*</span>
                              </label>
                              <select
                                className="form-select formcontrol-extra"
                                value={this.state.accessOption}
                                disabled
                                style={{
                                  border: '1px solid #6A6E79',
                                  backgroundColor: '#2B2D32'
                                }}
                              >
                                <option value="No">No</option>
                              </select>
                            </div>
                          )}
                      </>
                    )}



                    {this.state.popupError && (
                      <div style={{ marginLeft: '-5', width: '100%' }} className="alert alert-danger">
                        {this.state.messagesFromIframe.length > 0 ? ((/[\\\][/*:?-]/.test(this.state.title)) ? "No Special Characters allowed in FormName" : "Please fill the required fields") : "Please drag atleast one control in dropzone"}
                      </div>
                    )}

                  </div>
                  <div className="modal-buttons" style={{ padding: '0px 20px 20px 20px' }}>

                    <button
                      type="button"
                      className="newlogocolorbtn"

                      onClick={this.handleSave}
                    // style={{ width: 80, height: 30, alignItems: 'center', paddingTop: 4 }}
                    >
                      Save
                    </button>
                    <button
                      type="button"
                      className="newblackcolorbtn"
                      onClick={this.closePopup}
                    // style={{ backgroundColor: 'gray', width: 80, height: 30, alignItems: 'center', color: 'white', paddingTop: 1 }}
                    >
                      Cancel
                    </button>
                  </div>
                  {/* </div> */}
                </div>
              </div>
            )
          }

          {
            this.state.Editpostingbool && (
              <div className="custom-dark-modal-overlay">
                <div className="custom-dark-modal" style={{ minWidth: '367px' }}>
                  {/* <div className="savepopup modal-content"> */}
                  {/* <div className="modal-header"> */}
                  <h5 className="addsteptitle">Update Form</h5>
                  {/* <button type="button" className="close" onClick={this.closeUpdatePopup}>
                        <span>&times;</span>
                      </button>
                    </div> */}
                  <div className="form-group" style={{ padding: '0px 20px' }}>
                    <div className="form-control-group">
                      <label className="new-label">
                        Form Name <span style={{ color: 'red' }}>*</span>
                      </label>
                      <input
                        type="text"
                        className="form-control formcontrol-extra new-disabled"
                        value={this.state.storedData.AppName}

                        disabled

                      />
                    </div>
                    <div className="form-control-group">
                      <label className="new-label">
                        Form Description <span style={{ color: 'red' }}>*</span>
                      </label>
                      <input
                        type="text"
                        className="form-control formcontrol-extra new-disabled"
                        value={this.state.storedData.Description || 'No Description Available'}

                        disabled

                      />
                    </div>
                    <div className="form-control-group">
                      <label className="new-label">
                        Form Category <span style={{ color: 'red' }}>*</span>
                      </label>
                      <select
                        className="form-select formcontrol-extra new-disabled"
                        value={this.state.storedData.FormType}
                        disabled
                      >
                        <option value="">Select</option>
                        {this.state.formTypeOptions.map((option, index) => (
                          <option key={index} value={option}>{option}</option>
                        ))}
                      </select>
                    </div>
                    <div className="alert alert-warning" role="alert" style={{
                      marginBottom: 0, padding: 5, minWidth: '367px',
                      maxWidth: '367px'
                    }}>
                      <p style={{ marginBottom: 0, fontSize: '13px' }}><b>Note: </b>{this.state.isMultiForm
                        ? "This update will remove existing workflows. Please rebuild them."
                        : "Build the conditions again for this form in configure workflow."}
                      </p>
                    </div>
                  </div>
                  <div className="modal-buttons" style={{ padding: '0px 20px 20px 20px' }}>
                    <button
                      type="button"
                      className="newlogocolorbtn"
                      onClick={async () => {
                        if (this.state.isMultiForm) {
                          // console.log("Multi-form update triggered");
                          await this.handlemultiUpdate();

                        } else {
                          // console.log("Single-form update triggered");
                          await this.handleUpdate();
                        }
                      }}
                    // style={{ width: 80, height: 30, alignItems: 'center', paddingTop: 4 }}
                    >
                      Update
                    </button>
                    <button
                      type="button"
                      className="newblackcolorbtn"
                      onClick={this.closeUpdatePopup}
                    // style={{ backgroundColor: 'gray', width: 80, height: 30, alignItems: 'center', color: 'white', paddingTop: 1 }}
                    >
                      Cancel
                    </button>

                  </div>
                  {/* </div> */}
                </div>
              </div>

            )
          }

          <TopNavBar onUserDomainRetrieved={this.handleUserDomain} />
          <div style={{ display: 'flex' }}>
            {/* <Sidebar /> */}
            {this.state.loading ? (<LoadingSpinner />) : (<></>)}
            <section style={{ width: '100%' }}>
              <div>
                <iframe
                  ref={this.iframeRef}
                  // src="https://smartofficenxtbuilder.cloudangles.com/reactformbuilder/"
                  src="https://smartofficenxtbuilder.cloudangles.com/reactformbuilder-qa/"
                  // src="http://localhost:8080/"
                  title="Embedded Application"
                  style={{ border: "none", display: "block", width: "100%", height: "100vh" }}
                  onLoad={this.sendDataToIframe}
                />
              </div>
            </section>  
          </div>
        </div >
      );
    } catch (error) {
      this.handleError(error, "Error rendering component");
      return (
        <div style={{ display: 'flex', flexDirection: 'column' }}>
          <TopNavBar />
          {/* <div style={{ display: 'flex' }}>
            <Sidebar /> */}
          <div>Error rendering component.</div>
          {/* </div> */}
        </div>
      );
    }
    // finally{this.handleloading()}
  }
}
export default withNavigate(CreateForm);