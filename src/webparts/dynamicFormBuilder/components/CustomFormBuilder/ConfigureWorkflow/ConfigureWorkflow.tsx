import * as React from 'react';
// import {useState } from 'react';
import { sp } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
// import SideNavBar from './SideNavBar';
import TopNavBar from '../TopBar';
import './CustomForm.css';
import { BaseWebPartContext } from '@microsoft/sp-webpart-base';
import { MouseEventHandler, Component } from 'react';
import { Toast } from 'primereact/toast';
import '../statuspopup.css'
// import { ToggleButton } from 'primereact/togglebutton';
import Addstep from '../../../assets/Images/ConfigureWorkflowIcons/AddStep.svg';
import FinalNotification from '../../../assets/Images/ConfigureWorkflowIcons/FinalNotification.svg';
import ConfigureSLA from '../../../assets/Images/ConfigureWorkflowIcons/ConfigureSLA.svg';
import taskperformerplus from '../../../assets/Images/ConfigureWorkflowIcons/taskperformerplus.svg';
import { PeoplePicker, PrincipalType, IPeoplePickerContext } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import ConditionBuilder from './ConditionalRouting';
import { Edit, Network } from 'lucide-react';
import { fetchTenantUser } from '../FetchTenantUser/fetchTenantUser';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faUserPen, faTrash } from "@fortawesome/free-solid-svg-icons";
import LoadingSpinner from '../Loading';
// import { DefaultButton, Dialog, DialogFooter, PrimaryButton } from '@fluentui/react'
// import { fetchAllUsers } from '../ADService/ADService';
import SideBar from '../Sidebar/SideBar';
import { dataservice } from '../encryptionutil';


interface ApproverData {
    ParentAppCode: string;
    Role: string;
    User: string;
    Level: string;
    AppName: string;
    AppCode: string;
    Id: number;
    UserId?: number;
}

interface RowData {
    Id: number
    Role: string;
    User?: string;
    Group?: string
    Level: string;
    AppName: string;
    AppCode: string;
}



interface ConfigureWorkflowProps {
    context: BaseWebPartContext;
    data: RowData[];
}

interface User {
    userPrincipalName: any;
    displayName: string;
    mail: string;
}

interface ConfigureWorkflowState {
    approverData: ApproverData[];
    TempapproverData: ApproverData[];
    showDialog: boolean;
    showDialog2: boolean;
    // selectedUser: string;
    formMasterItems: { ID: number;AppName: string; AppCode: string; FormJSON: string; VisibilityFlag: boolean, IsParentForm: boolean, ParentAppCode: string, isWorkflowRequired: string, ChildOrder: number, TemplateAppCode?: string }[];
    selectedAppCode: string;
    roleOptions: string[];
    myOrganizationUsers: User[]
    selectedRole: string;
    level: number;
    formName: string;
    showAppSelectionError: boolean;
    showNotif: boolean;
    canEdit: boolean;
    isModalOpen: boolean;
    conditionalRow: any;
    CurrentformJSON: any;
    databool: boolean;
    recievedjson: any;
    loading: boolean;
    ChildappID: number;
    // conditionUser: string;
    ConditionsList: any;
    selectedUserObject?: any;
    // isGroup:boolean;
    selectedUser: string | null; // Allow null values
    isGroup: boolean;
    editconditionJson: any;
    showExistingWorkflowDialog: boolean;
    selectedExistingForm: string;
    EditMe: boolean;
    showEditDialog: boolean;
    selectedRow: RowData | null;
    NotifyButton: boolean,
    notifyperson: string,
    isSameDomain: string,
    notifyIDs: number[],
    notifyPrevIDs: number[],
    notifyEmailsTitles: string[],
    FinalApproveYesNo: string,
    isWorkflowRequired: string,
    searchText: string;
    filteredUsers: User[];
    showDropdown: boolean;
    selectedUserEmail?: string;
    selectedUsers: User[],

    duplicateCreatorId: number,
    isdynamicstatusPopupOpen: boolean;
    statusrow: any;
    newStatus: string[];
    FormItem: any;
    liststatuses: string[],
    isSLAPopupOpen: boolean;
    SLAFormMasterItem: any;
    SLAinhours: number;
    isSLAEnabled: boolean;
    domainfield: string;
    externalGroups: string[];
    lastroles: any;
    isParent: boolean;
    Childappcode: string;
    issaveall: boolean;
    saveallpopup: boolean;
    dataserviceobj: dataservice;
}

class ConfigureWorkflow extends Component<ConfigureWorkflowProps, ConfigureWorkflowState> {
    toast = React.createRef<Toast>();

    constructor(props: ConfigureWorkflowProps) {
        super(props);
        this.handleChildData = this.handleChildData.bind(this);
        this.openEditDialog = this.openEditDialog.bind(this);
        sp.setup({
            spfxContext: this.props.context as any,
        });
        this.state = {
            selectedAppCode: '',
            showDialog2: false,
            isModalOpen: false,
            showAppSelectionError: false,
            showNotif: false,
            selectedRole: '',
            ChildappID: 0,
            level: 0,
            selectedUser: null as any | null,
            formName: '',
            formMasterItems: [],
            approverData: [],
            TempapproverData: [],
            showDialog: false,
            roleOptions: ['Role1', 'Role2', 'Role3'],
            myOrganizationUsers: [],
            canEdit: false,
            conditionalRow: {},
            CurrentformJSON: [],
            databool: false,
            recievedjson: [],
            // conditionUser: '',
            ConditionsList: [],
            selectedUserObject: null,
            isGroup: false,
            editconditionJson: {},
            showExistingWorkflowDialog: false,
            selectedExistingForm: '',
            EditMe: false,
            selectedRow: null,
            showEditDialog: false,
            loading: true,
            NotifyButton: false,
            notifyperson: '',
            isSameDomain: '',
            notifyIDs: [],
            notifyPrevIDs: [],
            notifyEmailsTitles: [],
            FinalApproveYesNo: '',

            searchText: '',
            filteredUsers: [],
            showDropdown: false,
            selectedUserEmail: '',
            selectedUsers: [],
            duplicateCreatorId: 0,
            isdynamicstatusPopupOpen: false,
            statusrow: {},
            newStatus: [],
            FormItem: {},
            liststatuses: [],
            isSLAPopupOpen: false,
            SLAFormMasterItem: {},
            SLAinhours: 0,
            isSLAEnabled: false,
            domainfield: "",
            externalGroups: [],
            lastroles: '',
            isWorkflowRequired: 'Yes',
            isParent: false,
            Childappcode: '',
            issaveall: true,
            saveallpopup: false,
            dataserviceobj: new dataservice(),
        };
    }
    openModal = () => {
        this.setState({ isModalOpen: true });
    };
    databoolchange = () => {
        //console.log("✅ Current notifyIDs:", this.state.notifyIDs);

        const notifyIDs = this.state.notifyIDs;
        // 🧩 Validation: ensure "Route to" field is not empty
        if (!notifyIDs[0]) {
            this.toast?.current?.clear();
            this.toast?.current?.show({
                severity: 'warn',
                detail: "Please select at least one person in 'Route to' field.",
                life: 2500
            });
            return; // ❌ stop execution
        }

        // ✅ Proceed if valid
        this.setState({ databool: true });
    }
    fetchStatuses = async (): Promise<void> => {
        try {
            const appcodetoFetch = this.state.isParent ? this.state.Childappcode : this.state.selectedAppCode;

            const RAWitems = await sp.web.lists.getByTitle('FormMaster')
                .items.filter(`AppCode eq '${appcodetoFetch}'`)
                .select("Id,DynamicStatuses,IsParentForm,FormJSON")
                .top(1)
                .get();
            const items = RAWitems?.map((item:any)=>{
                    return { ...item, FormJSON: this.state.dataserviceobj.decryptjson(item.FormJSON)};
                })
            //console.log("get from form master", items);
            this.setState({ FormItem: items[0] });
            const listStatuses = JSON.parse(items[0].DynamicStatuses);
            this.setState({ liststatuses: listStatuses || [] })
            this.setState({ newStatus: listStatuses || [] });
            if (!listStatuses || listStatuses.length === 0) {
                this.addStatusField();
            }
        } catch (error) {
            console.error("Error fetching statuses:", error);
        }
    };

    areArraysEqualIgnoringSpaces(arr1: string[], arr2: string[]) {
        if (arr1.length !== arr2.length) return false;

        return arr1.every((val, index) => val.trim() === arr2[index].trim());
    }

    dynamicstatusopenPopup = async (data: any) => {
        //console.log("data", data)
        await this.fetchStatuses()
        this.setState({ statusrow: data },
            // ()=>void this.fetchStatuses()
        );
        this.setState({ isdynamicstatusPopupOpen: true })
    };

    dynamicstatusclosePopup = (): void => {
        this.setState({ isdynamicstatusPopupOpen: false });
    };
    addStatusField = (): void => {
        this.setState(prevState => ({
            newStatus: [...prevState.newStatus, '']
        }), () => {
            //console.log("addstatusfield=====================================", this.state.newStatus);
        });
    };
    removeStatusField = (index: number): void => {
        this.setState(prevState => ({
            newStatus: prevState.newStatus.filter((_, i) => i !== index)
        }), () => console.log("wghjasdhngfssdbvvvvvvvvvvv", this.state.newStatus));
    };
    updateStatusField = (index: number, value: string): void => {
        const updatedStatuses = [...this.state.newStatus];
        updatedStatuses[index] = value;
        this.setState({ newStatus: updatedStatuses }, () => {
            //console.log("updated statuses ========================= ", this.state.newStatus);
        });
    };

    saveStatuses = async (): Promise<void> => {
        try {
            if (this.state.newStatus.some(str => str.trim() === "") || this.state.newStatus.length === 0) {
                this.toast.current?.clear();
                this.toast.current?.show({
                    severity: 'error',
                    summary: '',
                    detail: 'Status cannot be empty',
                    life: 3000
                });
                return;
            }
            await sp.web.lists.getByTitle('FormMaster').items.getById(this.state.ChildappID).update({
                DynamicStatuses: JSON.stringify(this.state.newStatus)
            });
            this.toast.current?.clear();
            this.toast.current?.show({
                severity: 'success',
                summary: 'Success',
                detail: 'Statuses saved successfully',
                life: 3000
            });
            this.dynamicstatusclosePopup();
            void this.fetchStatuses();
        } catch (error) {
            console.error("Error saving statuses:", error);
            this.toast.current?.clear();
            this.toast.current?.show({
                severity: 'error',
                summary: 'Error',
                detail: 'Error saving statuses',
                life: 3000
            });
        }
    };
    updateStatuses = async (): Promise<void> => {
        try {
            if (this.state.newStatus.some(str => str.trim() === "")) {
                this.toast.current?.clear();
                this.toast.current?.show({
                    severity: 'error',
                    summary: '',
                    detail: 'Status cannot be empty',
                    life: 3000
                });
                return;
            }
            else if (this.areArraysEqualIgnoringSpaces(this.state.liststatuses, this.state.newStatus)) {
                this.toast.current?.clear();
                this.toast.current?.show({
                    severity: 'warn',
                    summary: '',
                    detail: 'No changes are made',
                    life: 3000
                });
                return;
            }
            if (this.state.FormItem?.Id) {
                await sp.web.lists.getByTitle('FormMaster').items.getById(this.state.ChildappID).update({
                    DynamicStatuses: JSON.stringify(this.state.newStatus)
                });
                this.toast.current?.clear();
                this.toast.current?.show({
                    severity: 'success',
                    summary: 'Success',
                    detail: 'Statuses updated successfully',
                    life: 3000
                });
                this.dynamicstatusclosePopup();
                void this.fetchStatuses();
            }
        } catch (error) {
            console.error("Error updating statuses:", error);
            this.toast.current?.show({
                severity: 'error',
                summary: 'Error',
                detail: 'Error updating statuses',
                life: 3000
            });
        }
    };

    handleChildData = (data: any) => {
        debugger
        //console.log("---childdata in parent---", data);

        // Avoid multiple executions using conditional checks
        if (JSON.stringify(data) !== JSON.stringify(this.state.recievedjson)) {
            this.setState({ recievedjson: data }, async () => {
                //console.log("---this.state.recievedjson----", this.state.recievedjson);

                // Call AddConditions once the state is updated
                await this.AddConditions();
            });
        } else {
            this.setState({ databool: false })
            //console.log("Duplicate data, skipping AddConditions");
        }
    };

    private getPeoplePickerContext(): IPeoplePickerContext {
        return {
            absoluteUrl: this.props.context.pageContext.web.absoluteUrl,
            msGraphClientFactory: this.props.context.msGraphClientFactory as any,
            spHttpClient: this.props.context.spHttpClient as any,
        };
    }

    async fetchConditionsList() {
        try {
            const ConditionalListData = await sp.web.lists
                .getByTitle('ConditionsList')
                .items.select('Role', 'PersonOrGroup/Title', 'PersonOrGroup/Id', 'FinalApprove', 'Level', 'AppName', 'AppCode', 'Id', 'Created', 'ConditionJson')
                .expand('PersonOrGroup')
                .get();

            this.setState({ ConditionsList: ConditionalListData })
        } catch (error) {
            console.error('Error fetching conditions list:', error);
        }
    }

    SLApopupopen = async () => {
        let Apptoshow = (this.state.isParent) ? this.state.Childappcode : this.state.selectedAppCode;
        const RAWFormMaster = await sp.web.lists.getByTitle("FormMaster")
            .items.filter(`AppCode eq '${Apptoshow}'`)
            .select("Id,SLAinhours,SLA,FormJSON")
            .top(1)
            .get()
        const FormMaster = RAWFormMaster?.map((item:any)=>{
            return { ...item, FormJSON: this.state.dataserviceobj.decryptjson(item.FormJSON)};
        }) 
        this.setState({ SLAFormMasterItem: FormMaster[0], SLAinhours: FormMaster[0].SLAinhours, isSLAEnabled: FormMaster[0].SLA })
        this.setState({ isSLAPopupOpen: true })
    }
    SLApopupclose = () => {
        this.setState({ isSLAPopupOpen: false })
    }
    saveSLA = async () => {
        if (this.state.SLAinhours === this.state.SLAFormMasterItem.SLAinhours && this.state.isSLAEnabled === this.state.SLAFormMasterItem.SLA) {
            this.toast?.current?.clear();
            this.toast?.current?.show({
                severity: 'warn',
                detail: 'No changes made',
                life: 2000
            })
        }
        else if (this.state.SLAinhours === 0 || !this.state.SLAinhours) {
            this.toast?.current?.clear();
            this.toast?.current?.show({
                severity: 'warn',
                detail: 'SLA cannot be empty or zero',
                life: 2000
            })
        }
        else if (this.state.SLAinhours < 0) {
            this.toast?.current?.clear();
            this.toast?.current?.show({
                severity: 'warn',
                detail: 'SLA cannot be negative',
                life: 2000
            })
        }
        else {
            try {
                await sp.web.lists.getByTitle("FormMaster").items.getById(Number(this.state.SLAFormMasterItem.Id)).update({
                    SLAinhours: this.state.SLAinhours,
                    SLA: this.state.isSLAEnabled
                })
                this.setState({ isSLAPopupOpen: false })
                this.toast?.current?.clear();
                this.toast?.current?.show({
                    severity: 'success',
                    detail: 'SLA configured successfully for selected form',
                    life: 2000
                })
            }
            catch (err) {
                console.log("error in sla configuration", err)
            }
        }
    }
    fetchExternalGroups = async () => {
        let extgrps: any[] = []
        // console.log("domainandissame", domainfield, isSameDomain)
        if (this.state.isSameDomain && this.state.isSameDomain === "YES") {
            // console.log("domainfieldinIFFF", domainfield)
            extgrps = await sp.web.lists
                .getByTitle("ExternalGroupsList")
                .items.select("Domain,GroupName/Title").expand("GroupName").filter(`Domain ne '${this.state.domainfield}'`).get()
        }
        else if (this.state.isSameDomain) {
            // console.log("INELSE", domainfield)
            extgrps = await sp.web.lists
                .getByTitle("ExternalGroupsList")
                .items.select("Domain,GroupName/Title").expand("GroupName").filter(`Domain eq '${this.state.domainfield}'`).get()
        }
        const titles: string[] = extgrps.map(item => item.GroupName?.Title);
        this.setState({ externalGroups: titles });
        // console.log("Ext", externalGroups)
        // console.log("external groups", titles)
    };

    async componentDidMount() {
        // const currentUser = await sp.web.currentUser();
        // console.log("current loggedIn user detail --------------------", currentUser)

        this.setState({ canEdit: false });
        try {
            // Fetch approver data
            const data = await sp.web.lists
                .getByTitle('MappingMaster')
                .items.select('Role', 'Users_x002f_Groups/Title', 'Level', 'AppName', 'AppCode', 'Id', 'Created', 'ParentAppCode')
                .expand('Users_x002f_Groups')
                .getAll();

            await this.fetchConditionsList()

            const approverData = data.map((item: any) => ({
                Role: item.Role,
                User: item.Users_x002f_Groups ? item.Users_x002f_Groups.Title : '',
                Level: item.Level,
                AppName: item.AppName,
                AppCode: item.AppCode,
                Id: item.Id,
                ParentAppCode: item.ParentAppCode,
            }));
            const count = approverData.filter(data => data.ParentAppCode);
            //console.log("count in mount :", count)

            const data1 = await sp.web.lists
                .getByTitle('Temporary_Mapping_master')
                .items.select('Role', 'Users_x002f_Groups/Title', 'Users_x002f_Groups/Id', 'Level', 'AppName', 'AppCode', 'Id', 'Created', 'ParentAppCode')
                .expand('Users_x002f_Groups')
                .getAll();

            await this.fetchConditionsList()


            const TempapproverData1 = data1.map((item: any) => ({
                Role: item.Role,
                User: item.Users_x002f_Groups ? item.Users_x002f_Groups.Title : '',
                Level: item.Level,
                AppName: item.AppName,
                AppCode: item.AppCode,
                Id: item.Id,
                ParentAppCode: item.ParentAppCode,
                UserId: item.Users_x002f_Groups?.Id
            }));

            const TempapproverData = [...TempapproverData1, ...count];
            // this.setState({TempapproverData : TempapproverData})
            //console.log(`Coumponenetdidmount approverdata count : ${approverData.length};tempappdata count : ${TempapproverData1.length}`)
            //console.log("TempapproverData from componentdidmount :", TempapproverData)

            // if ( count.length > 0 ){                
            //     this.setState({TempapproverData :  [...this.state.TempapproverData, ...count]});
            //     console.log(`Updated tempapprover data : ${this.state.TempapproverData}`)
            // }

            const Domain = await fetchTenantUser()
            debugger;

            const RAWformMasterData = await sp.web.lists
                .getByTitle('FormMaster')
                .items.select('AppName', 'AppCode', 'Created', 'FormJSON', 'Domain', 'VisibilityFlag', 'IsParentForm', 'ParentAppCode', 'isWorkflowRequired', 'ChildOrder', 'TemplateAppCode','ID')
                .filter(`Domain eq '${Domain.TenantUsers}' and VisibilityFlag eq 1 `)
                .getAll();

            const formMasterData = RAWformMasterData?.map((item:any)=>{
                    return { ...item, FormJSON: this.state.dataserviceobj.decryptjson(item.FormJSON)};
                })

            this.setState({
                approverData,
                TempapproverData,
                formMasterItems: formMasterData,
                selectedAppCode: '', // Default to no selection
                domainfield: Domain.TenantUsers,
            }, () => this.fetchExternalGroups());



        } catch (error: unknown) {
            if (error instanceof Error) {
                console.error('Error loading data:', error);
                alert(`Error loading data: ${error.message}`);
            } else {
                console.error('An unknown error occurred:', error);
                this.toast.current?.show({
                    severity: 'error',
                    summary: '',
                    detail: 'An unknown error occurred.',
                    life: 2000,
                });
            }
        } finally { this.setState({ loading: false }) }
    }
    notifyToClick = async () => {
        this.setState({ notifyEmailsTitles: [], selectedUsers: [] });

        this.setState({ notifyEmailsTitles: [] });
        let Apptoshow = (this.state.isParent) ? this.state.Childappcode : this.state.selectedAppCode;
        //console.log("fghjikgfgh===========appcode", Apptoshow)
        const items = await sp.web.lists.getByTitle("FormMaster")
            .items
            .filter(`AppCode eq '${Apptoshow}'`)
            .select("*", "Id, NotifyTo/Id, NotifyTo/Title")
            .expand("NotifyTo")
            .get();

        const notifyToIds = items[0]?.NotifyToId || [];

        this.setState({
            notifyIDs: notifyToIds,
            notifyPrevIDs: notifyToIds
        }, () => {
            //console.log("state previous ids ", this.state.notifyIDs, this.state.notifyPrevIDs)
        });

        const usersArray: User[] = [];
        const notifyEmailsTitles: string[] = [];

        for (const id of notifyToIds) {
            try {
                const user = await sp.web.siteUsers.getById(id).select("Id", "Email", "Title", "LoginName").get();

                usersArray.push({
                    userPrincipalName: user.LoginName,      // or user.UserPrincipalName if available
                    displayName: user.Title,
                    mail: user.Email
                });

                notifyEmailsTitles.push(user.Title);

            } catch (error) {
                //console.log(`ID ${id} is not a user, treating as group.`);
                const group = await sp.web.siteGroups.getById(id).select("Id", "Title").get();

                // Optionally: if you want to add groups as well
                notifyEmailsTitles.push(group.Title);
            }
        }

        this.setState({
            notifyEmailsTitles,
            selectedUsers: usersArray,
            showDialog2: true,
            searchText: ""
        });
    };
    editiconClick = async () => {
        debugger;
        this.setState({ notifyEmailsTitles: [] })
        if (this.state.editconditionJson) {
            this.setState({
                FinalApproveYesNo: this.state.editconditionJson.FinalApprove || "No"
            });
        } else {
            this.setState({
                FinalApproveYesNo: "No"
            });
        }
        //console.log("EITTTEDITTT>>>>>>>>>>>>>>>>>>>>>>>>>>>", this.state.editconditionJson)
        this.setState({ notifyIDs: [this.state.editconditionJson?.PersonOrGroup?.Id], notifyPrevIDs: [this.state.editconditionJson?.PersonOrGroup?.Id] })

        if (this.state.editconditionJson.PersonOrGroup?.Id) {
            const usersArray: string[] = [];
            try {
                const user = await sp.web.siteUsers.getById(this.state.editconditionJson.PersonOrGroup?.Id).select("Id", "Email").get();
                if (user?.Email) {
                    usersArray.push(user.Email);
                }
            } catch (error) {
                //console.log(`ID ${this.state.editconditionJson.PersonOrGroup?.Id} is not a user, treating as group.`);
                const group = await sp.web.siteGroups.getById(this.state.editconditionJson.PersonOrGroup?.Id).select("Id", "Title").get();
                if (group?.Title) {
                    usersArray.push(group.Title);
                }
            }
            this.setState({ notifyEmailsTitles: usersArray })
            this.setState({ searchText: usersArray[0] || "" })       // POPULATING USER NAME IN EDIT CONFITION 
        }
    }
    onNotify: MouseEventHandler<HTMLButtonElement> = async () => {
        if (!this.state.selectedAppCode) {
            this.setState({ showNotif: true })


            this.toast.current?.show({
                severity: 'error',
                summary: '',
                detail: 'Please select a form',
                life: 2000,
            });
        } else {
            try {
                const items = await sp.web.lists.getByTitle("FormMaster")
                    .items
                    .filter(`AppCode eq '${this.state.selectedAppCode}'`)
                    .select("Id", "NotifyTo/Id", "NotifyTo/Title")
                    .expand("NotifyTo")
                    .top(1)
                    .get();
                //console.log("==========Items========", items)
                const isGroup = await sp.web.siteGroups
                    .getById(items[0].NotifyTo[0].Id)
                    .get()
                    .then(() => true)
                    .catch(() => false);
                //console.log("group id", items[0].NotifyTo[0].Id)
                //console.log("isgroup", isGroup)
                if (isGroup) {
                    const group = await sp.web.siteGroups.getById(items[0].NotifyTo[0].Id).get();
                    this.setState({ notifyperson: group.Title })
                }
                else {
                    const user = await sp.web.getUserById(items[0].NotifyTo[0].Id).select("Title", "Email").get();
                    console.log("Notify To:", user); this.setState({ notifyperson: user.Email })
                }
            } catch (error) {
                console.error("Error fetching NotifyTo field:", error);
            }
            this.setState({ showDialog2: true, showNotif: false });
        }
    };
    // onApproverClick: MouseEventHandler<HTMLButtonElement> = () => {

    //     this.setState({ searchText: "" });


    //     if (!this.state.selectedAppCode) {
    //         this.setState({ showAppSelectionError: true });

    //         this.toast.current?.show({
    //             severity: 'error',
    //             summary: '',
    //             detail: 'Please select a Form Name.',
    //             life: 2000,
    //         });
    //     } else {
    //         this.setState({ showDialog: true, showAppSelectionError: false });
    //     }
    // };

    onApproverClick: MouseEventHandler<HTMLButtonElement> = () => {

        debugger;

        this.setState({ searchText: "" });
        const appcode = (!this.state.Childappcode) ? this.state.selectedAppCode : this.state.Childappcode;

        //console.log("Add button clicked ----- this.state.selectedAppCode", appcode);
        if (!appcode) {
            this.setState({ showAppSelectionError: true });

            this.toast.current?.show({
                severity: 'error',
                summary: '',
                detail: 'Please select a Form Name.',
                life: 2000,
            });
        } else {
            this.setState({ showDialog: true, showAppSelectionError: false });
        }

        if (appcode) {
            const selectedAppCode = appcode;
            const selectedApp = this.state.formMasterItems.find((item) => item.AppCode === selectedAppCode);

            if (selectedApp) {
                debugger
                //console.log(`App Name: ${selectedApp.AppName}, App Code: ${selectedApp.AppCode}, from onapproverclick`);
                this.setState({ formName: selectedApp.AppName })
                const dataforapproverData = (this.state.Childappcode && this.state.issaveall) ? this.state.TempapproverData : this.state.approverData
                //console.log("on approver click: ", this.state.Childappcode && !this.state.issaveall)
                // const filteredData = dataforapproverData.filter(data => data.AppCode === AppCode);

                const appData = dataforapproverData.filter((approver) => approver.AppCode === selectedAppCode);

                if (appData.length === 0) {

                    this.setState({
                        roleOptions: ['Creator'],
                        selectedRole: 'Creator',
                        level: 0,
                    });
                } else {

                    const levels = appData.map((approver) => parseInt(approver.Level, 10));
                    const maxLevel = Math.max(...levels);
                    const lastRole = appData.find((approver) => parseInt(approver.Level, 10) === maxLevel)?.Role;
                    const isTaskPerformerexist = appData.some(data => data.Role === "Task Performer");
                    debugger;
                    this.setState({ lastroles: lastRole })
                    //console.log("appproval required", selectedApp.isWorkflowRequired);
                    if (lastRole === 'Creator' && selectedApp.isWorkflowRequired === "No") {
                        this.setState({
                            roleOptions: ['Receiver'],
                            selectedRole: 'Receiver',
                            level: maxLevel + 1,
                        });
                    }
                    else if (lastRole === 'Creator' || lastRole === 'Reviewer' || lastRole === 'Approver' || lastRole === 'Task Performer' && selectedApp.isWorkflowRequired !== "No") {
                        this.setState({
                            roleOptions: isTaskPerformerexist ? ['Reviewer', 'Approver'] : ['Reviewer', 'Approver', 'Task Performer'],
                            selectedRole: 'Reviewer',
                            level: maxLevel + 1,
                        });
                    }
                }
            } else {
                //console.log('App Code does not exist in FormMaster list');


                this.setState({
                    roleOptions: ['Creator'],
                    selectedRole: 'Creator',
                    level: 0,
                });
            }
        }
    };

    closeDialog: MouseEventHandler<HTMLButtonElement> = () => {
        this.setState({ showDialog: false, selectedUser: '' });
        //console.log("this.state.canEdit-------", this.state.canEdit)
    };
    conditionCloseDialog: MouseEventHandler<HTMLButtonElement> = () => {
        this.setState({
            isModalOpen: false, selectedUser: '', notifyIDs: [],
            notifyEmailsTitles: [],
            notifyPrevIDs: [], recievedjson: []
        });
    };
    deleteCondition = async () => {
        this.setState({ loading: true })
        try {
            await sp.web.lists.getByTitle("ConditionsList").items.getById(this.state.editconditionJson.Id).delete();
            //console.log("this condition deleted", this.state.editconditionJson)

            this.toast.current?.show({
                severity: 'success',
                summary: "",
                detail: `Condition deleted successfully!`,
                life: 2000,
            });
            this.setState({
                isModalOpen: false, selectedUser: '', notifyIDs: [],
                notifyEmailsTitles: [],
                notifyPrevIDs: [], recievedjson: []
            });
            await this.fetchConditionsList();
        } catch (error) {
            console.error("Error deleting item:", error);
        }
        this.setState({ loading: false })

    }
    onDeleteApprover = (id: number) => {
        //console.log(`Delete approver with Id: ${id}`);
    };
    onChange = async (items: any[]) => {
        const Ids = items.map((ele) => {
            return parseInt(ele?.id)
        })
        this.setState({ notifyIDs: Ids }, () => { console.log("------notifyids-------", this.state.notifyIDs) })
    }
    private onAddApprover = async () => {
        debugger;
        //console.log('Starting to add new approver...', this.state.canEdit);
        //console.log('All details...', this.state.isGroup, this.state.selectedUser, this.state.selectedUserObject);
        const listname = (this.state.Childappcode && this.state.issaveall) ? "Temporary_Mapping_master" : 'MappingMaster'

        try {
            const emailOrGroupName = this.state.isGroup
                ? this.state.selectedUserObject?.text?.trim()
                : this.state.selectedUser?.trim();
            //console.log("emailOrGroupName-------------", emailOrGroupName)
            if (!emailOrGroupName) {
                this.toast.current?.show({
                    severity: 'error',
                    summary: '',
                    detail: 'Please select a valid user or group.',
                    life: 2000,
                });
                return;
            }
            //console.log(`added in to ${listname}`)

            let userId: number | undefined;
            if (this.state.isGroup) {
                const selectedUserObject = this.state.selectedUserObject;
                if (!selectedUserObject || !selectedUserObject.id) {
                    this.toast.current?.show({
                        severity: 'error',
                        summary: '',
                        detail: 'Invalid group selection. Please select a valid group.',
                        life: 2000,
                    });
                    return;
                }
                userId = selectedUserObject.id;
            } else {
                const user = await sp.web.siteUsers
                    .filter(`Email eq '${emailOrGroupName}'`)
                    .select('Id,Email')
                    .get();

                if (user.length === 0) {
                    this.toast.current?.show({
                        severity: 'error',
                        summary: '',
                        detail: 'User not found.',
                        life: 2000,
                    });
                    this.setState({ loading: false })
                    return;
                }
                userId = user[0].Id;
            }

            // Proceed with the rest of the logic
            const { selectedRole, selectedAppCode, canEdit, level, formMasterItems } = this.state;

            const existingItem = await sp.web.lists
                .getByTitle('MappingMaster')
                .items.filter(`AppCode eq '${selectedAppCode}' and Users_x002f_GroupsId eq ${userId}`)
                .select('Role')
                .getAll();
            if (existingItem.length <= 0 && this.state.Childappcode.length > 0) {
                const existingItem = await sp.web.lists
                    .getByTitle('Temporary_Mapping_master')
                    .items.filter(`AppCode eq '${selectedAppCode}' and Users_x002f_GroupsId eq ${userId}`)
                    .select('Role')
                    .getAll();
                console.log("child flow: ", existingItem)
            }
            //console.log("exisitngitems: ", existingItem)
            if (existingItem.length > 0) {
                const currentRole = existingItem[0].Role;
                this.toast.current?.show({
                    severity: 'error',
                    summary: '',
                    detail: `User or group already exists with the role ${currentRole}.`,
                    life: 2000,
                });

                return;
            }

            const editing = canEdit ? 'true' : 'false';
            let appcode = (this.state.Childappcode) ? this.state.Childappcode : selectedAppCode
            await sp.web.lists.getByTitle(listname).items.add({
                Role: selectedRole,
                Level: level,
                AppName: formMasterItems.find((item) => item.AppCode === appcode)?.AppName || '',
                AppCode: appcode,
                Users_x002f_GroupsId: userId,
                ParentAppCode: (this.state.Childappcode) ? selectedAppCode : '',
                CanEdit: editing,
            });
            if (selectedRole === "Creator") {
                const FormMasterRecord = await sp.web.lists.getByTitle('FormMaster')
                    .items.filter(`AppCode eq '${selectedAppCode}'`)
                    .top(1)
                    .get();
                if (FormMasterRecord.length > 0) {
                    const id = FormMasterRecord[0].Id;
                    await sp.web.lists.getByTitle('FormMaster').items.getById(id).update(
                        {
                            FormViewerId: userId
                        }
                    )
                }
            }

            this.toast.current?.show({
                severity: 'success',
                summary: '',
                detail: `New ${selectedRole} added successfully!`,
                life: 2000,
            });

            this.setState({
                showDialog: false,
                selectedUser: '',
                searchText: ""
            });

            await this.Refreshdata();
        } catch (error: unknown) {
            if (error instanceof Error) {
                console.error('Error adding new approver:', error);
                alert(`Error adding new approver: ${error.message}`);
            } else {
                console.error('An unknown error occurred:', error);
                alert('An unknown error occurred.');
            }
        }
    };
    AddConditions = async () => {
        debugger
        //  this.setState({ databool: true }); // Set loading state
        // const notifyIDs = this.state.notifyIDs;
        // if (!notifyIDs[0] || notifyIDs.length === 0) {
        //     this.toast?.current?.clear();
        //     this.toast?.current?.show({
        //         severity: 'warn',
        //         detail: "Please select at least one person in 'Route to' field.",
        //         life: 2500
        //     });
        //     return; // ❌ stop execution
        // }
        const payload = {
            Role: this.state.conditionalRow.Role,
            Level: this.state.conditionalRow.Level,
            AppName: this.state.conditionalRow.AppName,
            AppCode: this.state.conditionalRow.AppCode,
            ConditionJson: JSON.stringify(this.state.recievedjson), // Ensure JSON is a string
            FinalApprove: this.state.FinalApproveYesNo
        };

        //console.log('Payload being sent:', payload);

        try {

            if (this.state.recievedjson.conditions.length === 0) {
                this.toast.current?.show({
                    severity: 'error',
                    summary: '',
                    detail: 'Error: Conditions cannot be empty!',
                    life: 2000,
                });
                return;

            }
            else {
                for (const condition of this.state.recievedjson.conditions) {
                    if (!condition.field || !condition.operator || !condition.value) {
                        this.toast.current?.show({
                            severity: 'error',
                            summary: '',
                            detail: `Error: ${!condition.field ? "Field," : ""}  ${!condition.operator ? "Operator," : ""}  ${!condition.value ? "Value" : ""} cannot be empty in conditions!`,
                            life: 2000,
                        });
                        return;
                    }
                }
            }
            if (this.state.notifyIDs.length < 1) {
                this.toast.current?.show({
                    severity: 'error',
                    summary: '',
                    detail: 'Please select a valid user or group.',
                    life: 2000,
                });
                this.setState({ recievedjson: [] })
                return;
            } else if (this.CompareIdArray(this.state.notifyPrevIDs, this.state.notifyIDs) && (this.state.editconditionJson?.ConditionJson === payload.ConditionJson) && (this.state.editconditionJson?.FinalApprove === this.state.FinalApproveYesNo)) {

                this.toast.current?.show({
                    severity: 'error',
                    summary: '',
                    detail: 'No updates found ',
                    life: 2000,
                });
                this.setState({ recievedjson: [] })
                return;

            }
            debugger

            if (this.state.editconditionJson) {
                //console.log("----edit---this.state.editconditionJson-------------", this.state.editconditionJson)
                await sp.web.lists.getByTitle("ConditionsList").items.getById(this.state.editconditionJson.Id).update({
                    Role: this.state.conditionalRow.Role,
                    Level: this.state.conditionalRow.Level,
                    AppName: this.state.conditionalRow.AppName,
                    AppCode: this.state.conditionalRow.AppCode,
                    ConditionJson: JSON.stringify(this.state.recievedjson),
                    PersonOrGroupId: this.state.notifyIDs[0],
                    FinalApprove: this.state.FinalApproveYesNo
                });

            }
            else {
                //console.log("----plus---this.state.editconditionJson-------------", this.state.editconditionJson)
                debugger
                await sp.web.lists.getByTitle('ConditionsList').items.add({
                    Role: this.state.conditionalRow.Role,
                    Level: this.state.conditionalRow.Level,
                    AppName: this.state.conditionalRow.AppName,
                    AppCode: this.state.conditionalRow.AppCode,
                    ConditionJson: JSON.stringify(this.state.recievedjson),
                    PersonOrGroupId: this.state.notifyIDs[0],
                    FinalApprove: this.state.FinalApproveYesNo
                });
            }
            //console.log('Condition saved successfully');

            this.toast.current?.show({
                severity: 'success',
                summary: "",
                detail: `New Condition added successfully!`,
                life: 2000,
            });
            this.setState({
                isModalOpen: false,
                selectedUser: '',
                recievedjson: [],
                notifyIDs: [],
                notifyEmailsTitles: [],
                notifyPrevIDs: []
            });
            // window.location.reload()
            await this.fetchConditionsList();
        } catch (error) {
            console.error('Error saving condition:', error);
        } finally {
            this.setState({ databool: false }); // Reset loading state
        }

    };

    onChildSelectChange = async (AppCode: any) => {

        //console.log("Childformopening", AppCode);
        this.setState({ loading: true })
        this.setState({ Childappcode: AppCode })
        // this.setState({ ChildappID : 0 })
        const selectedAppCode = AppCode;
        console.log("formmasteritems", this.state.formMasterItems)
        const x = this.state.formMasterItems
        // console.log("x", x)
        let dataforapproverData = this.state.TempapproverData;
        const filteredData = dataforapproverData.filter(data => data.AppCode === AppCode);
        
        //console.log(`Childon selelct : approverdata count : ${this.state.approverData.length};tempappdata count : ${this.state.TempapproverData.length}`)
        // console.log('filterddsta in child app select change:', formmaster);

        // ✅ Check if child form has TemplateAppCode
        const selectedChildForm = x.find(item => item.AppCode === AppCode);
        console.log("selectedChildForm",selectedChildForm,selectedChildForm?.ID);
        this.setState({ ChildappID: selectedChildForm?.ID || this.state.FormItem.Id })
        console.log('Selected child form details:', selectedChildForm,AppCode);
        const templateAppCode = selectedChildForm?.TemplateAppCode;

        //console.log(`TemplateAppCode for selected child: ${templateAppCode}`);

        // ✅ If TemplateAppCode exists and no workflow configured yet, fetch from MappingMaster
        if (templateAppCode && filteredData.length === 0) {
            //console.log(`Fetching workflow from MappingMaster for TemplateAppCode: ${templateAppCode}`);
            try {
                const templateWorkflow = await sp.web.lists.getByTitle("MappingMaster")
                    .items.filter(`AppCode eq '${templateAppCode}'`)
                    .select('Role', 'Users_x002f_Groups/Title', 'Users_x002f_Groups/Id', 'Level', 'AppName', 'AppCode', 'Id')
                    .expand('Users_x002f_Groups')
                    .orderBy("Level", true)
                    .get();

                if (templateWorkflow && templateWorkflow.length > 0) {
                    //console.log(`Found ${templateWorkflow.length} workflow steps for template`);

                    // Check if workflow already exists in Temporary_Mapping_master
                    const existingWorkflow = await sp.web.lists.getByTitle('Temporary_Mapping_master')
                        .items.filter(`AppCode eq '${selectedAppCode}'`)
                        .get();

                    if (existingWorkflow.length === 0) {
                        // Add to Temporary_Mapping_master with new AppCode
                        const addPromises = templateWorkflow.map(async (workflow: any) => {
                            const userId = workflow.Users_x002f_Groups?.Id;
                            const newItem = {
                                AppCode: selectedAppCode,
                                AppName: selectedChildForm.AppName,
                                Role: workflow.Role || "Creator",
                                Users_x002f_GroupsId: userId ? parseInt(userId) : null,
                                Level: workflow.Level !== undefined ? workflow.Level : 0,
                                ParentAppCode: this.state.selectedAppCode
                            };
                            //console.log("Adding workflow step to Temporary_Mapping_master:", newItem);
                            return sp.web.lists.getByTitle('Temporary_Mapping_master').items.add(newItem);
                        });

                        await Promise.all(addPromises);
                    } else {
                        console.log(`Workflow already exists for ${selectedAppCode} in Temporary_Mapping_master`);
                    }

                    // Refresh TempapproverData
                    const updatedTempData = await sp.web.lists
                        .getByTitle('Temporary_Mapping_master')
                        .items.select('Role', 'Users_x002f_Groups/Title', 'Users_x002f_Groups/Id', 'Level', 'AppName', 'AppCode', 'Id', 'Created', 'ParentAppCode')
                        .expand('Users_x002f_Groups')
                        .getAll();

                    const updatedTempapproverData = updatedTempData.map((item: any) => ({
                        Role: item.Role,
                        User: item.Users_x002f_Groups ? item.Users_x002f_Groups.Title : '',
                        Level: item.Level,
                        AppName: item.AppName,
                        AppCode: item.AppCode,
                        Id: item.Id,
                        ParentAppCode: item.ParentAppCode,
                        UserId: item.Users_x002f_Groups?.Id
                    }));

                    this.setState({ TempapproverData: updatedTempapproverData });

                    this.toast.current?.show({
                        severity: 'success',
                        summary: 'Workflow Added',
                        detail: `Workflow from template form added successfully for ${selectedChildForm.AppName}`,
                        life: 3000,
                    });

                    // Update filteredData to avoid showing dialog
                    dataforapproverData = updatedTempapproverData;
                } else {
                    console.log("No workflow found for TemplateAppCode");
                    this.setState({ showExistingWorkflowDialog: true, showEditDialog: false });
                }
            } catch (error) {
                console.error("Error fetching template workflow:", error);
                this.setState({ showExistingWorkflowDialog: true, showEditDialog: false });
            }
        } else if (filteredData.length === 0) {
            // Show the dialog if no data is available and no template
            //console.log("child1 null")
            this.setState({ showExistingWorkflowDialog: true, showEditDialog: false });
        }


        const yy = x.find((ele) => { return ele.AppCode === AppCode })?.FormJSON //aa particle app form structre json
        if (yy) {
            try {
                const jsonObject = JSON.parse(yy); // Safely parse the JSON string
                const filtereddata = jsonObject.filter((ele: any) => {
                    if (ele.element === "TextInput" || ele.element === "NumberInput" || ele.element === "Dropdown") {
                        return ele
                    }
                })
                this.setState({ CurrentformJSON: filtereddata })
                //console.log("Parsed JSON object:", jsonObject);
                //console.log("filtered parsed", filtereddata)
            } catch (error) {
                console.error("Error parsing JSON:", error);
            }
        } else {
            console.warn("No matching object found or yy is undefined");
        }
        //console.log(`App Code selected: ${AppCode}`);


        const selectedApp = this.state.formMasterItems.find((item) => item.AppCode === AppCode);
        this.setState({
            isWorkflowRequired: selectedApp?.isWorkflowRequired ?? 'Yes'
        })

        if (selectedApp) {
            //console.log(`App Name: ${selectedApp.AppName}, App Code: ${selectedApp.AppCode}`);
            this.setState({ formName: selectedApp.AppName })
            const dataforapproverData = (this.state.Childappcode) ? this.state.TempapproverData : this.state.approverData
            const appData = dataforapproverData.filter((approver) => approver.AppCode === selectedAppCode);


            if (appData.length === 0) {

                this.setState({
                    roleOptions: ['Creator'],
                    selectedRole: 'Creator',
                    level: 0,
                });
            } else {

                const levels = appData.map((approver) => parseInt(approver.Level, 10));
                const maxLevel = Math.max(...levels);
                const lastRole = appData.find((approver) => parseInt(approver.Level, 10) === maxLevel)?.Role;
                this.setState({ lastroles: lastRole })
                const isTaskPerformerexist = appData.some(data => data.Role === "Task Performer");
                if (lastRole === 'Creator' || lastRole === 'Reviewer' || lastRole === 'Approver' || lastRole === 'Task Performer') {
                    this.setState({
                        roleOptions: isTaskPerformerexist ? ['Reviewer', 'Approver'] : ['Reviewer', 'Approver', 'Task Performer'],
                        selectedRole: 'Reviewer',
                        level: maxLevel + 1,
                    });
                }

            }
        } else {
            console.log('App Code does not exist in FormMaster list');


            this.setState({
                roleOptions: ['Creator'],
                selectedRole: 'Creator',
                level: 0,
            });
        }
        // this.setState({ selectedAppCode });

        this.setState({ isParent: true })
        this.setState({ loading: false })
    };

    onAppSelectChange = async (AppCode: any) => {
        debugger
        // console.log("event---------------------", event.target.value)
        this.setState({ formName: '' })
        this.setState({ NotifyButton: true })
        this.setState({ EditMe: false });
        this.setState({ notifyperson: '' })
        this.setState({ isParent: false })
        this.setState({ Childappcode: '' })
        this.setState({ loading: true })
        const selectedAppCode = AppCode;
        //console.log("formmasteritems", this.state.formMasterItems)
        this.setState({ selectedAppCode });
        const x = this.state.formMasterItems
        //console.log("x", x)
        const selectedFormArray = x.filter(data => data.AppCode === selectedAppCode);
        const isparentform = selectedFormArray.length > 0 ? selectedFormArray[0].IsParentForm : false;
        const isparentState: boolean = (isparentform === true) ? true : false;
        //console.log("isparent:", isparentState);
        this.setState({ isParent: isparentState });
        const toshowsaveall = this.state.approverData.filter(data => data.ParentAppCode === selectedAppCode);
        (toshowsaveall.length > 0) ? this.setState({ issaveall: false }) : this.setState({ issaveall: true })
        const dataforapproverData = (isparentState) ? this.state.TempapproverData : this.state.approverData
        const filteredData = dataforapproverData.filter(data => data.AppCode === selectedAppCode);
        //console.log('filterddsta in app select change:', filteredData);

        // Disabling popup on change from non approval required form - Add existing workflow popup
        const filterData = this.state.approverData.filter(data => data.AppCode === selectedAppCode);
        const selectedApp = this.state.formMasterItems.find((item) => item.AppCode === selectedAppCode);
        this.setState({
            isWorkflowRequired: selectedApp?.isWorkflowRequired ?? 'Yes'
        })
        this.closeExistingWorkflowDialog();
        if (filterData.length === 0 && selectedApp?.isWorkflowRequired !== 'No' && !isparentState) {
            // Show the dialog if no data is available
            this.setState({ showExistingWorkflowDialog: true, showEditDialog: false });
        }


        if (isparentState) {
            const firstchild = x.filter(data => data.ParentAppCode === selectedAppCode);
            //console.log("Child forms :", firstchild);
            let firstchildcode = (firstchild && firstchild.length > 0) ? firstchild[0].AppCode : "";
            //console.log("firstchild from the slectedapps :", firstchildcode);
            this.setState({ Childappcode: firstchildcode });
            this.setState({formName: firstchild[0]?.AppName})
            if (firstchildcode) {
                await this.onChildSelectChange(firstchildcode);
            }
            // return ;
        }else{this.setState({ formName: selectedApp?.AppName || '' })}

        if (filteredData.length === 0 && !isparentState) {
            // Show the dialog if no data is available
            //console.log("parent null")
            this.setState({ showExistingWorkflowDialog: true, showEditDialog: false });
        }

        if (!isparentState) {
            const yy = x.find((ele) => { return ele.AppCode === selectedAppCode })?.FormJSON //aa particle app form structre json
        if (yy) {
            try {
                const jsonObject = JSON.parse(yy); // Safely parse the JSON string
                const filtereddata = jsonObject.filter((ele: any) => {
                    if (ele.element === "TextInput" || ele.element === "NumberInput" || ele.element === "Dropdown") {
                        return ele
                    }
                })
                this.setState({ CurrentformJSON: filtereddata })
                //console.log("Parsed JSON object:", jsonObject);
                //console.log("filtered parsed", filtereddata)
            } catch (error) {
                console.error("Error parsing JSON:", error);
            }
        } else {
            console.warn("No matching object found or yy is undefined");
        }
        //console.log(`App Code selected: ${selectedAppCode}`);


            const selectedApp = this.state.formMasterItems.find((item) => item.AppCode === selectedAppCode);
            this.setState({
                isWorkflowRequired: selectedApp?.isWorkflowRequired ?? 'Yes'
            })

        if (selectedApp) {
            //console.log(`App Name: ${selectedApp.AppName}, App Code: ${selectedApp.AppCode}`);
            // if(!this.state.isParent){this.setState({ formName: selectedApp.AppName });}

            const appData = this.state.approverData.filter((approver) => approver.AppCode === selectedAppCode);

            if (appData.length === 0) {

                this.setState({
                    roleOptions: ['Creator'],
                    selectedRole: 'Creator',
                    level: 0,
                });
                //console.log(this.state.lastroles)
                this.setState({ lastroles: 'Creator' });
                //console.log(this.state.lastroles)
            } else {

                const levels = appData.map((approver) => parseInt(approver.Level, 10));
                const maxLevel = Math.max(...levels);
                const lastRole = appData.find((approver) => parseInt(approver.Level, 10) === maxLevel)?.Role;
                    this.setState({ lastroles: lastRole })
                const isTaskPerformerexist = appData.some(data => data.Role === "Task Performer");
                if (lastRole === 'Creator' && selectedApp.isWorkflowRequired === "No") {
                    this.setState({
                        roleOptions: ['Receiver'],
                        selectedRole: 'Receiver',
                        level: maxLevel + 1,
                    });
                }
                else if (lastRole === 'Creator' || lastRole === 'Reviewer' || lastRole === 'Approver' || lastRole === 'Task Performer' && selectedApp.isWorkflowRequired !== "No") {
                    this.setState({
                        roleOptions: isTaskPerformerexist ? ['Reviewer', 'Approver'] : ['Reviewer', 'Approver', 'Task Performer'],
                        selectedRole: 'Reviewer',
                        level: maxLevel + 1,
                    });
                }
                this.setState({ lastroles: lastRole })
                // else if (lastRole === 'Reviewer') {
                //     this.setState({
                //         roleOptions: ['Reviewer', 'Approver'],
                //         selectedRole: 'Reviewer',
                //         level: maxLevel + 1,
                //     });
                // } else if (lastRole === 'Approver') {
                //     this.setState({
                //         roleOptions: ['Approver'],
                //         selectedRole: 'Approver',
                //         level: maxLevel + 1,
                //     });
                // }
            }
        } else {
            //console.log('App Code does not exist in FormMaster list');


            this.setState({
                roleOptions: ['Creator'],
                selectedRole: 'Creator',
                level: 0,
            });
        }


        this.setState({ selectedAppCode });
        }
        // else{
        //     let Childformslist = this.state.formMasterItems.filter(item => item.VisibilityFlag && item.ParentAppCode === selectedAppCode);
        // }
        this.setState({ loading: false })
    };


    onExistingFormSelectChange = async (e: React.ChangeEvent<HTMLSelectElement>) => {
        const selectedExistingForm = e.target.value;
        this.setState({ selectedExistingForm });
        this.setState({ loading: true })
        if (selectedExistingForm) {
            try {
                // Fetch the data from the SharePoint list based on the selected form's AppCode
                const data = await sp.web.lists.getByTitle("MappingMaster")
                    .items.filter(`AppCode eq '${selectedExistingForm}'`)  // Filter by selected AppCode
                    .get();

                //console.log("Selected form data:", data);  // Log the data to the console

                // Optionally, set the fetched data to the state if needed
                this.setState({ approverData: data });
            } catch (error) {
                console.error("Error fetching data:", error);
            }
        }
        this.setState({ loading: false })
    };




    addExistingWorkflow = async () => {
        debugger
        this.setState({ loading: true })
        const { selectedExistingForm, selectedAppCode, formName, Childappcode, isParent } = this.state;
        const listname = (this.state.Childappcode) ? 'Temporary_Mapping_master' : 'MappingMaster';
        if (selectedExistingForm) {
            try {
                const selectedFormDetails = await sp.web.lists.getByTitle("MappingMaster")
                    .items.filter(`AppCode eq '${selectedExistingForm}'`)
                    .orderBy("Level", true)
                    .get();

                //console.log("Fetched form details including Level 0:", selectedFormDetails);
 
                if (selectedFormDetails && selectedFormDetails.length > 0) {
                    const sortedFormDetails = selectedFormDetails.sort((a, b) => Number(a.Level) - Number(b.Level));
                    //console.log("Sorted Form Details:", sortedFormDetails);
                    const duplicatePromises = sortedFormDetails
                        .filter(formDetail => formDetail.Users_x002f_GroupsId) // ✅ Only include items with valid Users/Groups
                        .map(async (formDetail) => {
                            const userId = formDetail.Users_x002f_GroupsId;
                        const newItem = {
                                AppCode: (isParent) ? Childappcode : selectedAppCode,
                            AppName: formName,
                            Role: formDetail.Role || "Default Role",
                                Users_x002f_GroupsId: parseInt(userId),
                            Level: formDetail.Level !== undefined ? formDetail.Level : "Default Level",
                                ParentAppCode: (isParent) ? selectedAppCode : '',
                        };
                        //console.log("Adding Item:", newItem);
                        if (formDetail.Level === 0) {
                                this.setState({ duplicateCreatorId: userId })
                        }
                            return sp.web.lists.getByTitle(listname).items.add(newItem);
                    });

                    try {
                        await Promise.all(duplicatePromises);
                        this.toast.current?.show({
                            severity: 'success',
                            summary: 'Success',
                            detail: 'Added successfully in ascending order!',
                            life: 3000,
                        });
                        if (this.state.duplicateCreatorId) {
                            const NewFormMasterItem = await sp.web.lists.getByTitle("FormMaster").items
                                .filter(`AppCode eq '${selectedAppCode}'`)
                                .top(1)
                                .get()
                            await sp.web.lists.getByTitle('FormMaster').items.getById(NewFormMasterItem[0].Id).update(
                                {
                                    FormViewerId: this.state.duplicateCreatorId
                                }
                            )
                        }
                        await this.Refreshdata();
                        // Close the dialog and update state
                        this.closeExistingWorkflowDialog();
                        this.setState({ loading: false })
                    } catch (error) {
                        console.error("Error adding workflow(s):", error);
                        this.setState({ loading: false })

                        // Show error toast
                        this.toast.current?.show({
                            severity: 'error',
                            summary: 'Error',
                            detail: 'Failed to add workflow(s). Please try again.',
                            life: 3000,
                        });
                    }
                }

                else {
                    this.setState({ loading: false })
                    console.warn("No data found for the selected form.");
                    this.toast.current?.show({
                        severity: 'warn',
                        summary: 'No Data',
                        detail: 'No details found for the selected form.',
                        life: 3000,
                    });
                }
            } catch (error: unknown) {
                this.setState({ loading: false })
                console.error("Error adding workflow(s):", error);

                const errorMessage =
                    error instanceof Error ? error.message : JSON.stringify(error);

                this.toast.current?.show({
                    severity: 'error',
                    summary: 'Error',
                    detail: `Failed to add workflow(s). Details: ${errorMessage}`,
                    life: 3000,
                });
            }
        } else {
            this.setState({ loading: false })
            console.warn("No form selected.");
            this.toast.current?.show({
                severity: 'warn',
                summary: 'Warning',
                detail: 'Please select a form to add.',
                life: 3000,
            });
        }
        
    };
    closeExistingWorkflowDialog = () => {
        this.setState({ showExistingWorkflowDialog: false, selectedExistingForm: '' });
    };

    openEditDialog = async (rowData: RowData) => {
        //console.log("rrrrrrrrrrrrr", rowData)

        const WorkflowItems = await sp.web.lists
            .getByTitle("WorkFlowProcessData")
            .items
            .filter(`AppCode eq '${rowData.AppCode}' and (Status eq 'Draft' or Status eq 'In-Progress' or Status eq 'Returned')`)
            .top(1)
            .get();

        if (WorkflowItems.length > 0) {
            console.warn("Already records exist in Draft/Returned/In-Progress, so cannot edit.");
            this.toast.current?.clear();
            this.toast.current?.show({
                severity: 'error',
                summary: 'Form Exists!',
                detail: 'This form has existing records in Draft/In-Progress/Returned state. Complete it before editing!',
                life: 3000,
                style: { height: 100 }
            });


            return;
        }

        this.setState({
            showEditDialog: true,
            selectedRow: rowData,
            selectedUser: rowData.User ?? rowData.Group ?? null,
            searchText: rowData.User ?? rowData.Group ?? ""
        });
    };
    closeEditDialog = () => {
        this.setState({ showEditDialog: false, selectedRow: null });
    };

    updateUser = async () => {
        const {
            selectedUser,
            selectedUserObject,
            selectedRow,
            isGroup,
        } = this.state;

        this.setState({ loading: true });

        try {
            const emailOrGroupName = isGroup
                ? selectedUserObject?.text?.trim()
                : selectedUser?.trim();

            if (!emailOrGroupName) {
                this.toast.current?.show({
                    severity: 'error',
                    summary: '',
                    detail: 'Please select a valid user or group.',
                    life: 2000,
                });
                this.setState({ loading: false });
                return;
            }

            if (!selectedRow) {
                this.toast.current?.show({
                    severity: 'warn',
                    summary: '',
                    detail: 'Please select a valid row!',
                    life: 2000,
                });
                this.setState({ loading: false });
                return;
            }

            const existingValue = selectedRow.User ?? selectedRow.Group ?? "";
            if (existingValue.trim() === emailOrGroupName) {
                this.toast.current?.show({
                    severity: 'info',
                    summary: '',
                    detail: 'No updates found.',
                    life: 2000,
                });
                this.setState({ loading: false });
                return;
            }

            // 🔍 Get user ID
            let userId: number | undefined;
            if (isGroup) {
                if (!selectedUserObject?.id) {
                    this.toast.current?.show({
                        severity: 'error',
                        summary: '',
                        detail: 'Invalid group selection. Please select a valid group.',
                        life: 2000,
                    });
                    this.setState({ loading: false });
                    return;
                }
                userId = selectedUserObject.id;
            } else {
                const user = await sp.web.siteUsers
                    .filter(`Email eq '${emailOrGroupName}'`)
                    .select('Id,Email')
                    .get();

                if (user.length === 0) {
                    this.toast.current?.show({
                        severity: 'error',
                        summary: '',
                        detail: 'User not found.',
                        life: 2000,
                    });
                    this.setState({ loading: false });
                    return;
                }

                userId = user[0].Id;
            }

            // ✅ Determine which list to update based on whether item has UserId (temp list) or not (finalized list)
            const isTemporary = !!(selectedRow as any).UserId;
            const listName = isTemporary ? "Temporary_Mapping_master" : "MappingMaster";

            // ✅ Additional Creator conflict check - check both lists
            const existingCreatorsMaster = await sp.web.lists.getByTitle("MappingMaster")
                .items
                .filter(`AppCode eq '${selectedRow.AppCode}' and Role eq 'Creator' and Users_x002f_GroupsId eq ${userId}`)
                .top(1)
                .get();

            const existingCreatorsTemp = await sp.web.lists.getByTitle("Temporary_Mapping_master")
                .items
                .filter(`AppCode eq '${selectedRow.AppCode}' and Role eq 'Creator' and Users_x002f_GroupsId eq ${userId}`)
                .top(1)
                .get();

            if (existingCreatorsMaster.length > 0 || existingCreatorsTemp.length > 0) {
                this.toast.current?.show({
                    severity: 'warn',
                    summary: '',
                    detail: 'This user is already assigned as a Creator for this application.',
                    life: 3000,
                });
                this.setState({ loading: false });
                return;
            }

            const list = sp.web.lists.getByTitle(listName);

            await list.items.getById(selectedRow.Id).update({
                Users_x002f_GroupsId: null,
            });

            await list.items.getById(selectedRow.Id).update({
                Users_x002f_GroupsId: userId,
            });

            // ✅ Update FormMaster if user is Creator
            if (selectedRow.Role === "Creator") {
                const App = (this.state.isParent && selectedRow.AppCode.endsWith("-1C")) ? this.state.selectedAppCode : selectedRow.AppCode;
                const FormMasterRecord = await sp.web.lists
                    .getByTitle('FormMaster')
                    .items
                    .filter(`AppCode eq '${App}'`)
                    .top(1)
                    .get();

                if (FormMasterRecord.length > 0) {
                    const id = FormMasterRecord[0].Id;
                    await sp.web.lists.getByTitle('FormMaster')
                        .items.getById(id)
                        .update({ FormViewerId: userId });
                }
            }

            this.toast.current?.show({
                severity: 'success',
                summary: '',
                detail: 'User updated successfully!',
                life: 2000,
            });

            this.setState({
                showEditDialog: false,
                searchText: "",
            });

            await this.Refreshdata();

        } catch (error) {
            console.error("Error updating user:", error);
            this.toast.current?.show({
                severity: 'error',
                summary: '',
                detail: 'Failed to update user!',
                life: 2000,
            });
        }

        this.setState({ loading: false });
    };
    CompareIdArray = (prev: number[], current: number[]) => {
        if (prev.length !== current.length) {
            return false
        }
        else {
            return prev.sort((a, b) => a - b).toString() === current.sort((a, b) => a - b).toString();
        }
    }

    handleNotifyAddUser = async (appCode: string) => {
        debugger
        this.setState({ loading: true })
        const listItem = await sp.web.lists.getByTitle("FormMaster")
            .items.filter(`AppCode eq '${appCode}'`)
            .select("*", "Id,NotifyTo/Id,NotifyTo/Title")
            .expand("NotifyTo")
            .get();
        //console.log("==========item", listItem)
        //console.log("==========this.state.notifyids", this.state.notifyIDs)

        const itemId = listItem[0].Id;

        try {
            if (this.state.notifyIDs.length < 1) {
                this.toast.current?.show({
                    severity: 'error',
                    summary: '',
                    detail: 'Please select a valid user or group.',
                    life: 2000,
                });
                return;
            } else if (this.CompareIdArray(this.state.notifyPrevIDs, this.state.notifyIDs)) {

                this.toast.current?.show({
                    severity: 'error',
                    summary: '',
                    detail: 'No updates found ',
                    life: 2000,
                });
                return;

            }
            await sp.web.lists.getByTitle("FormMaster").items.getById(itemId).update({

                NotifyToId: {
                    results: this.state.notifyIDs
                }
            });
            this.toast.current?.show({
                severity: 'success',
                summary: '',
                detail: 'User successfully added to the notify list!',
                life: 2000,
            });
            const listItem = await sp.web.lists.getByTitle("FormMaster")
                .items.filter(`AppCode eq '${appCode}'`)
                .select("*", "Id,NotifyTo/Id,NotifyTo/Title")
                .expand("NotifyTo")
                .get();
            console.log("==========itemagain", listItem)

            setTimeout(() => {
                this.setState({ showDialog2: false })
                // this.setState({ NotifyButton: false })
                this.setState({ searchText: "" })
            }, 1000);
            // await this.Refreshdata();
        } catch (error) {
            console.error("Error updating list:", error);
            alert("Error updating the record. Check console for details.");
        }
        this.setState({ loading: false })
    }
    handleAddUser = async (appCode: string) => {
        debugger;
        this.setState({ loading: true })
        // Get selected user email or group name
        const emailOrGroupName = this.state.isGroup
            ? this.state.selectedUserObject?.text?.trim()
            : this.state.selectedUser?.trim();

        //console.log("emailOrGroupName-------------", emailOrGroupName);

        if (!emailOrGroupName) {
            this.toast.current?.show({
                severity: 'error',
                summary: '',
                detail: 'Please select a valid user or group.',
                life: 2000,
            });
            return;
        }

        let userId: number | undefined;

        try {
            if (this.state.isGroup) {
                // Handle group selection
                const selectedUserObject = this.state.selectedUserObject;
                if (!selectedUserObject || !selectedUserObject.id) {
                    this.toast.current?.show({
                        severity: 'error',
                        summary: '',
                        detail: 'Invalid group selection. Please select a valid group.',
                        life: 2000,
                    });
                    return;
                }
                userId = selectedUserObject.id;
            } else {
                // Fetch user ID from SharePoint
                const user = await sp.web.siteUsers
                    .filter(`Email eq '${emailOrGroupName}'`)
                    .select('Id,Email')
                    .get();

                if (user.length === 0) {
                    this.toast.current?.show({
                        severity: 'error',
                        summary: '',
                        detail: 'User not found.',
                        life: 2000,
                    });
                    this.setState({ loading: false })
                    return;
                }
                userId = user[0].Id;
                //console.log('added user:............', user[0])
                //console.log('userId', userId)
                //console.log("Selected user: ", this.state.selectedUser)
            }
            // Step 1: Fetch data from SharePoint list filtered by AppCode
            const list = sp.web.lists.getByTitle("FormMaster");
            const items = await list.items.filter(`AppCode eq '${appCode}'`).get();

            if (items.length === 0) {
                // alert("No matching record found for this AppCode.");
                this.toast.current?.show({
                    severity: 'warn',
                    summary: '',
                    detail: 'No matching record found for this AppCode.',
                    life: 2000,
                });
                return;
            }

            // Step 2: Update the first matched record
            const itemId = items[0].Id;
            //console.log(itemId)
            //console.log(Event)
            try {
                await list.items.getById(itemId).update({
                    NotifyToId: { results: [userId] }
                });
            } catch (updateError) {
                console.error("Error updating item:", updateError);
            }


            // alert("User successfully added to the notify list!");
            this.toast.current?.show({
                severity: 'success',
                summary: '',
                detail: 'User successfully added to the notify list!',
                life: 2000,
            });


            setTimeout(() => {
                this.setState({ showDialog2: false })
                // this.setState({ NotifyButton: false })
            }, 1000);
            await this.Refreshdata();
        } catch (error) {
            console.error("Error updating list:", error);
            alert("Error updating the record. Check console for details.");
        }
        this.setState({ loading: false })
    };

    onsaveall = async () => {
        this.setState({ loading: true })
        debugger;
        const templedata = this.state.TempapproverData.filter(data =>
            data.ParentAppCode === this.state.selectedAppCode &&
            data.UserId // Only items with UserId (from Temporary_Mapping_master)
            )
            const parentschildren = this.state.formMasterItems.filter(item => item.ParentAppCode === this.state.selectedAppCode);
            for (const parent of parentschildren) {
                const count = templedata.filter((t) => t.AppCode === parent.AppCode).length;
                if (count < 2) {
                    this.toast.current?.show({
                        severity: "error",
                        summary: "",
                        detail: "Each form must have at least two workflow roles defined.",
                        life: 2000,
                    });
                    this.setState({ loading: false });
                    return; // stop further execution of onsaveall
                }
            }

        const promises = templedata.map((item: any) =>
            sp.web.lists.getByTitle("MappingMaster").items.add({
                Role: item.Role,
                Level: Number(item.Level),
                Users_x002f_GroupsId: item.UserId,
                AppName: item.AppName.trim(),
                AppCode: item.AppCode,
                ParentAppCode: item.ParentAppCode,

            })
        );

        await Promise.all(promises); // wait for all inserts
        //console.log("saved plz check the list.");
        templedata.map(item => console.log("user: ", item));
        //console.log("Deleting the items >>--> (-_*)")
        const promiseDelete = templedata.map((item) =>
            sp.web.lists.getByTitle("Temporary_Mapping_Master").items.getById(item.Id).delete()
        )
        try {
            await Promise.all(promiseDelete);
            //console.log("Items Deleted Successfully (x_x)")
        }
        catch (error) { console.error("Error deleting temporary items:", error) }
        await this.Refreshdata();
        this.setState({ issaveall: false });
        this.setState({ saveallpopup: false });
        this.setState({ loading: false });
    }

    deleteWorkflowRole = async (data: RowData) => {
        this.setState({ loading: true });
        try {
            const deletedLevel = Number(data.Level);
            const appCode = data.AppCode;

            //console.log(`Deleting role at level ${deletedLevel} for AppCode: ${appCode}`);

            // Delete the item from Temporary_Mapping_master
            await sp.web.lists.getByTitle("Temporary_Mapping_master")
                .items.getById(data.Id)
                .delete();

            //console.log(`Deleted item with ID: ${data.Id}`);

            // Fetch all remaining items for this AppCode
            const remainingItems = await sp.web.lists.getByTitle("Temporary_Mapping_master")
                .items.filter(`AppCode eq '${appCode}'`)
                .select('Id', 'Level', 'Role', 'Users_x002f_GroupsId', 'AppName', 'AppCode', 'ParentAppCode')
                .orderBy('Level', true)
                .get();

            //console.log('Remaining items:', remainingItems);

            // Re-level items that were after the deleted level
            const updatePromises = remainingItems
                .filter((item: any) => Number(item.Level) > deletedLevel)
                .map((item: any) => {
                    const newLevel = Number(item.Level) - 1;
                    //console.log(`Updating item ${item.Id}: Level ${item.Level} -> ${newLevel}`);
                    return sp.web.lists.getByTitle("Temporary_Mapping_master")
                        .items.getById(item.Id)
                        .update({ Level: newLevel });
                });

            if (updatePromises.length > 0) {
                await Promise.all(updatePromises);
                //console.log('Successfully re-leveled remaining items');
            }

            // Refresh ONLY Temporary_Mapping_master data (don't touch MappingMaster)
            const updatedTempData = await sp.web.lists
                .getByTitle('Temporary_Mapping_master')
                .items.select('Role', 'Users_x002f_Groups/Title', 'Users_x002f_Groups/Id', 'Level', 'AppName', 'AppCode', 'Id', 'Created', 'ParentAppCode')
                .expand('Users_x002f_Groups')
                .getAll();

            const updatedTempapproverData = updatedTempData.map((item: any) => ({
                Role: item.Role,
                User: item.Users_x002f_Groups ? item.Users_x002f_Groups.Title : '',
                Level: item.Level,
                AppName: item.AppName,
                AppCode: item.AppCode,
                Id: item.Id,
                ParentAppCode: item.ParentAppCode,
                UserId: item.Users_x002f_Groups?.Id
            }));

            // Include items from MappingMaster that have ParentAppCode (multi-form scenario)
            const approverDataWithParent = this.state.approverData.filter(data => data.ParentAppCode);
            const combinedTempData = [...updatedTempapproverData, ...approverDataWithParent];

            this.setState({ TempapproverData: combinedTempData });

            this.toast.current?.show({
                severity: 'success',
                summary: 'Deleted',
                detail: `${data.Role} at level ${deletedLevel} deleted successfully`,
                life: 3000,
            });

        } catch (error) {
            console.error("Error deleting workflow role:", error);
            this.toast.current?.show({
                severity: 'error',
                summary: 'Error',
                detail: 'Failed to delete workflow role',
                life: 3000,
            });
        } finally {
            this.setState({ loading: false });
            void this.Refreshdata();
        }
    };


    // async Refreshdata() {
    //     this.setState({ loading: true })
    //     const data = await sp.web.lists
    //         .getByTitle('MappingMaster')
    //         .items.select('Role', 'Users_x002f_Groups/Title', 'Level', 'AppName', 'AppCode', 'Id', 'Created')
    //         .expand('Users_x002f_Groups')
    //         .getAll();
    //     const approverData1 = data.map((item: any) => ({
    //         Role: item.Role,
    //         User: item.Users_x002f_Groups ? item.Users_x002f_Groups.Title : '',
    //         Level: item.Level,
    //         AppName: item.AppName,
    //         AppCode: item.AppCode,
    //         Id: item.Id,
    //     }))
    //     this.setState({ approverData: approverData1 })
    //     const appData = this.state.approverData.filter((approver) => approver.AppCode === this.state.selectedAppCode);

    //     if (appData.length === 0) {

    //         this.setState({
    //             roleOptions: ['Creator'],
    //             selectedRole: 'Creator',
    //             level: 0,
    //         });
    //     }
    //     else {

    //         const levels = appData.map((approver) => parseInt(approver.Level, 10));
    //         const maxLevel = Math.max(...levels);
    //         const lastRole = appData.find((approver) => parseInt(approver.Level, 10) === maxLevel)?.Role;
    //         const isTaskPerformerexist = appData.some(data => data.Role === "Task Performer");
    //         if (lastRole === 'Creator' || lastRole === 'Reviewer' || lastRole === 'Approver' || lastRole === 'Task Performer') {
    //             this.setState({
    //                 roleOptions: isTaskPerformerexist ? ['Reviewer', 'Approver'] : ['Reviewer', 'Approver', 'Task Performer'],
    //                 selectedRole: 'Reviewer',
    //                 level: maxLevel + 1,
    //             });
    //         }
    //         // console.log("approver data232",approverData1)
    //     }
    //     this.setState({ loading: false })
    // }

    async Refreshdata() {
        this.setState({ loading: true })
        const data = await sp.web.lists
            .getByTitle('MappingMaster')
            .items.select('Role', 'Users_x002f_Groups/Title', 'Level', 'AppName', 'AppCode', 'Id', 'Created', 'ParentAppCode')
            .expand('Users_x002f_Groups')
            .getAll();


        const Tempdata = await sp.web.lists
            .getByTitle('Temporary_Mapping_master')
            .items.select('Role', 'Users_x002f_Groups/Title', 'Users_x002f_Groups/Id', 'Level', 'AppName', 'AppCode', 'Id', 'Created', 'ParentAppCode')
            .expand('Users_x002f_Groups')
            .getAll();
        // await this.fetchConditionsList()
        //console.log("tempdata :", Tempdata.length);

        const approverData1 = data.map((item: any) => ({
            Role: item.Role,
            User: item.Users_x002f_Groups ? item.Users_x002f_Groups.Title : '',
            Level: item.Level,
            AppName: item.AppName,
            AppCode: item.AppCode,
            Id: item.Id,
            ParentAppCode: item.ParentAppCode
        }))

        const TempapproverData1 = Tempdata.map((item: any) => ({
            Role: item.Role,
            User: item.Users_x002f_Groups ? item.Users_x002f_Groups.Title : '',
            Level: item.Level,
            AppName: item.AppName,
            AppCode: item.AppCode,
            Id: item.Id,
            ParentAppCode: item.ParentAppCode,
            UserId: item.Users_x002f_Groups?.Id
        }))
        const count = approverData1.filter(data => data.ParentAppCode);

        //console.log("count after refresh : ", count);
        const TempapproverData = [...TempapproverData1, ...count]
        //console.log('Refreash in ni count: ', TempapproverData);

        this.setState({ TempapproverData })
        //console.log(`upon refresh; approverdata count : ${this.state.approverData.length};tempappdata count : ${TempapproverData.length}`)

        this.setState({ approverData: approverData1 });
        let displayapp = (this.state.Childappcode) ? this.state.Childappcode : this.state.selectedAppCode;
        const dataforapproverData = (this.state.Childappcode && !this.state.issaveall) ? TempapproverData : approverData1
        const appData = dataforapproverData.filter((approver) => approver.AppCode === displayapp);
        debugger;

        if (appData.length === 0) {

            this.setState({
                roleOptions: ['Creator'],
                selectedRole: 'Creator',
                level: 0,
            });
            this.setState({ lastroles: 'Creator' });
        }
        else {

            const levels = appData.map((approver) => parseInt(approver.Level, 10));
            const maxLevel = Math.max(...levels);
            const lastRole = appData.find((approver) => parseInt(approver.Level, 10) === maxLevel)?.Role;
            const isTaskPerformerexist = appData.some(data => data.Role === "Task Performer");
            this.setState({ lastroles: lastRole })
            if (lastRole === 'Creator' && this.state.isWorkflowRequired === "No") {
                this.setState({
                    roleOptions: ['Reciever'],
                    selectedRole: 'Reciever',
                    level: maxLevel + 1,
                });
            }
            if (lastRole === 'Creator' || lastRole === 'Reviewer' || lastRole === 'Approver' || lastRole === 'Task Performer') {
                this.setState({
                    roleOptions: isTaskPerformerexist ? ['Reviewer', 'Approver'] : ['Reviewer', 'Approver', 'Task Performer'],
                    selectedRole: 'Reviewer',
                    level: maxLevel + 1,
                });
            }
            // console.log("approver data232",approverData1)
        }
        this.setState({ loading: false })
    }

    handleUserDomain = (UserDomain: string) => {
        this.setState({ isSameDomain: UserDomain })
        //console.log("Logged-in USER DOMAIN IS   ----------------:", UserDomain);

    };

    // DROPDOWN FILTER FOR EXTERNAL USERS -----------------------------------------------------------------------

    // INTERNAL USERS  !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    async getPersonData(e: any) {
        if (Array.isArray(e)) {
            // Use `Promise.all` to await all promises in parallel
            await Promise.all(e.map(item => this.identifyEntity(item)));
        } else {
            await this.identifyEntity(e); // Await single promise
        }
    }

    async identifyEntity(entity: any) {
        try {
            //console.log("Identifying Entity:", entity);

            const isGroup = await sp.web.siteGroups
                .getById(entity.id)
                .get()
                .then(() => true)
                .catch(() => false);

            if (isGroup) {
                //console.log("Group Selected:", entity);

                this.setState({
                    isGroup: true,
                    selectedUser: entity.text,
                    selectedUserObject: entity,
                });
            } else {
                //console.log("Individual User Selected:", entity.secondaryText);

                this.setState({
                    isGroup: false,
                    selectedUser: entity.secondaryText,
                    selectedUserObject: entity,
                });
            }
        } catch (error) {
            console.error("Error identifying entity:", error);
        }
    }

    toggleDropdown = () => {
        this.setState(prev => ({ showDropdown: !prev.showDropdown }));
    };

    render() {
        // const { selectedUsers, showDropdown, filteredUsers } = this.state;
        let Apptoshow = (this.state.isParent) ? this.state.Childappcode : this.state.selectedAppCode;
        const dataforapproverData = (this.state.Childappcode && this.state.issaveall) ? this.state.TempapproverData : this.state.approverData
        const filteredApproverData = Apptoshow
            ?
            dataforapproverData
                .filter(data => data.AppCode === Apptoshow)
                .reduce((acc: ApproverData[], curr: ApproverData) => {
                    if (!acc.some(item => item.Level === curr.Level)) {
                        acc.push(curr);
                    }
                    return acc;
                }, [])
                .sort((a, b) => Number(a.Level) - Number(b.Level))
            : [];


        let UniqueFormItems = this.state.formMasterItems.filter(item => item.VisibilityFlag && !item.ParentAppCode);
        // let addexisitingforms = UniqueFormItems.filter(item => !item.IsParentForm);
        let Childformslist = this.state.formMasterItems.filter(item => item.VisibilityFlag && item.ParentAppCode === this.state.selectedAppCode);
        let singleformmasteritems = this.state.formMasterItems.filter(item => !item.ChildOrder && item.IsParentForm !== true && item.isWorkflowRequired !== "No");

        return (
            <div style={{ display: 'flex', flexDirection: 'column' }}>
                <TopNavBar onUserDomainRetrieved={this.handleUserDomain} />
                <div style={{ display: 'flex' }}>
                    <SideBar activeMenu="Create Workflow" />
                    {this.state.loading ? (<LoadingSpinner />) : (<></>)}
                    <div className='Table'>
                        <div className='d-flex justify-content-between ' style={{ margin: '10px 14px' }}>
                            <div className='d-flex align-items-center'>
                                <h6 className='white-title'>Configure WorkFlow</h6>
                                
                                {this.state.isParent && (
                                    <div style={{ marginLeft: '20px' }}>
                                        
                                        <select
                                            className="child-dropdown"
                                            value={this.state.Childappcode || ''}
                                            onChange={(e) => this.onChildSelectChange(e.target.value)}>
                                            {/* <option value="" disabled>Select a form</option> */}
                                            {Childformslist.map((data: { AppName: any; AppCode: any }, index: number) => (
                                                <option key={index} value={data.AppCode}>
                                                    {data.AppName}
                                                </option>
                                            ))}
                                        </select>
                                        </div>
                                    
                                )}
                                
                            </div>
                            <Toast ref={this.toast} />
                            <div style={{ display: 'flex', gap: '15px' }}>
                                {this.state.NotifyButton && (
                                    <>
                                        <button
                                            // style={{ cursor: 'pointer' }}
                                            disabled={this.state.isWorkflowRequired === 'No'}
                                            style={{
                                                cursor: this.state.isWorkflowRequired === 'No' ? 'not-allowed' : 'pointer',
                                                opacity: this.state.isWorkflowRequired === 'No' ? 0.5 : 1,
                                            }}
                                            onClick={this.SLApopupopen}
                                            className="newblackcolorbtn"
                                        > <img className="buttonicon notify" src={ConfigureSLA} alt="FN" />
                                            Configure SLA
                                        </button>
                                        <button
                                            style={{ cursor: 'pointer' }}
                                            onClick={this.notifyToClick}
                                            className="newblackcolorbtn"
                                        >
                                            <img className="buttonicon notify" src={FinalNotification} alt="FN" /> Final Notification
                                        </button>
                                    </>
                                )

                                }

                                <button
                                    disabled={this.state.lastroles === 'Receiver' && this.state.isWorkflowRequired==="No"}
                                    style={{
                                        cursor: this.state.lastroles === 'Receiver'&& this.state.isWorkflowRequired==="No" ? 'not-allowed' : 'pointer',
                                        opacity: this.state.lastroles === 'Receiver' && this.state.isWorkflowRequired==="No"? 0.5 : 1,
                                    }}
                                    onClick={this.onApproverClick}
                                    className="newlogocolorbtn"
                                >
                                    <img className="buttonicon" src={Addstep} alt="AS" />  Add Step
                                </button>

                                {(this.state.isParent && this.state.issaveall) && (
                                    <button
                                        className="newlogocolorbtn"
                                        onClick={() => { this.setState({ saveallpopup: true }); }}>
                                        Finalize workflow
                                    </button>)
                                }


                            </div>
                        </div>
                        
                        

                            <div style={{ display: 'flex', paddingLeft: 2 }}>
                            <div className="form-sidebar">
                                <h5 className="form-sidebar-title">Select Form</h5>

                                <ul className="form-list">
                                    {UniqueFormItems.map((item, index) => (
                                        <li
                                            key={index}
                                            className={`form-list-item ${this.state.selectedAppCode === item.AppCode ? 'selected' : ''}`}
                                            onClick={() => this.onAppSelectChange(item.AppCode)}
                                        >
                                            {item.AppName}
                                        </li>
                                    ))}
                                </ul>
                            </div>
                                    


                            <div style={{ width: '100%', margin: '0px 2px' }}>
                                {/* <>{this.state.isParent && ( */}
                        {/* <div className="horizontal-button-scroll"> */}
                                        {/* {Childformslist.map((data: { AppName: {} | null | undefined; AppCode: {} | null | undefined; }, index: any) => ( */}
                                            {/* <button */}
                                                {/* className={`child-list-item ${this.state.Childappcode === data.AppCode ? 'selected' : ''}`} */}
                                                {/* onClick={() => this.onChildSelectChange(data.AppCode)}> */}
                                                {/* {data.AppName} */}
                                            {/* </button> */}
                                        {/* ))}</div>)}</> */}
                                <table className="ta-table1">
                                    
                                    <thead style={{ background: 'none', backgroundColor: 'var(--color-navigation)', fontSize: 13 }}>
                                        <tr>
                                            <th>Role</th>
                                            <th>User</th>
                                            <th>Level</th>
                                            <th>App Name</th>
                                            <th> Actions</th>
                                            {/* <th></th> */}
                                            {/* {filteredApproverData.length > 1 && <th>Conditions</th>} */}
                                        </tr>
                                    </thead>
                                    <tbody style={{ fontSize: 12 }}>
                                        {filteredApproverData.length === 0 ? (
                                            <tr>
                                                <td colSpan={5} className='plsselect-form'>
                                                    Please select form
                                                </td>
                                            </tr>
                                        ) : (

                                            filteredApproverData.map((data, idx) => (
                                                <tr key={idx} className="Table_tr">
                                                    <td>{data.Role} {data.Role === "Task Performer" ? <img src={taskperformerplus} onClick={() => {
                                                        void this.dynamicstatusopenPopup(data);
                                                    }} className='dynaStatusbtn' alt="Plusicon" /> : ""}</td>
                                                    {/* <td>{data.User.split('|')[0]}</td> */}
                                                    <td>{data.User ? data.User.split('|')[0] : "Immediate Manager"}</td>
                                                    <td>{data.Level}</td>
                                                    <td>{data.AppName}</td>
                                                    <td style={{ display: 'flex', alignItems: 'center', gap: '10px', justifyContent: 'center' }}>
                                                        <FontAwesomeIcon
                                                            icon={faUserPen}
                                                            onClick={() => this.openEditDialog(data)}
                                                            style={{ cursor: "pointer", height: 12, width: 15 }}
                                                        />

                                                        {data.Role !== "Creator" &&
                                                            <button
                                                                title={this.state.ConditionsList.find((ele: any) => (ele.AppCode === data.AppCode && ele.Level === data.Level)) ? 'Edit Condition' : 'Add Condition'}
                                                                className='btn'
                                                                style={{
                                                                    width: 25,
                                                                    height: 25,
                                                                    display: "flex",
                                                                    justifyContent: "center",
                                                                    alignItems: "center",
                                                                    padding: 0,
                                                                    backgroundColor: 'var(--color-logo)'
                                                                }}
                                                                onClick={() => {
                                                                    this.openModal();
                                                                    this.setState({
                                                                        conditionalRow: data,
                                                                        editconditionJson: this.state.ConditionsList.find((ele: any) => (ele.AppCode === data.AppCode && ele.Level === data.Level))
                                                                    }, () => {
                                                                        //console.log("NotifyIDS and emails", this.state.notifyIDs, this.state.notifyEmailsTitles, this.state.notifyPrevIDs);
                                                                        void this.editiconClick();
                                                                    });
                                                                }}
                                                            >
                                                                {this.state.ConditionsList.find((ele: any) => (ele.AppCode === data.AppCode && ele.Level === data.Level)) ?
                                                                    <Edit style={{ marginLeft: 0, color: "white", width: 13, height: 15 }} /> :
                                                                    <Network style={{ marginLeft: 0, color: "white", width: 13, height: 15 }} />}
                                                            </button>
                                                        }

                                                        {/* Delete button - only show for non-Creator roles in Temporary_Mapping_master */}
                                                        {Number(data.Level) !== 0 && this.state.Childappcode && this.state.issaveall && (
                                                            <FontAwesomeIcon
                                                                icon={faTrash}
                                                                onClick={() => this.deleteWorkflowRole(data)}
                                                                title="Delete this workflow step"
                                                                style={{
                                                                    cursor: "pointer",
                                                                    height: 12,
                                                                    width: 15,
                                                                    color: '#dc3545'
                                                                }}
                                                            />
                                                        )}
                                                    </td>

                                                </tr>
                                            ))
                                        )}
                                    </tbody>
                                </table>
                            </div>
                        </div>
                        {this.state.showEditDialog && (
                            <div className="custom-dark-modal-overlay">
                                <div className="custom-dark-modal" >
                                    <h2 className="addsteptitle">Update Existing User/Group</h2>
                                    <div className="form-group" style={{ padding: '0px 20px 10px 10px' }}>
                                        <div className="form-control-group">
                                            <label>Update User/Group:</label>

                                            <PeoplePicker
                                                context={this.getPeoplePickerContext()}
                                                personSelectionLimit={1}
                                                groupName={""}
                                                required={true}
                                                onChange={(e: any) => this.getPersonData(e)}
                                                showHiddenInUI={false}
                                                principalTypes={[PrincipalType.User, PrincipalType.SharePointGroup]} // ✅ Allow users & groups
                                                resolveDelay={1000}
                                                placeholder="Search..."
                                                defaultSelectedUsers={[
                                                    this.state.selectedRow?.User || this.state.selectedRow?.Group || "", // Ensure correct default selection
                                                ]}
                                                allowUnvalidated={true}
                                                ensureUser={true}
                                                resultFilter={(results: any[]) =>
                                                    results.filter(persona => {
                                                        if (this.state.isSameDomain && this.state.isSameDomain === "NO") {
                                                            let email = persona?.loginName || '';
                                                            //console.log("persona in if ", persona)
                                                            if (!email.includes('#ext#')) {
                                                                const x = this.state.externalGroups.find(group => group === persona?.text);
                                                                return persona?.text.includes(x);
                                                            }
                                                            return email.includes(`${this.state.domainfield}`);
                                                        } else if (this.state.isSameDomain && !persona?.imageUrl) {
                                                            //console.log("persona in else if ", persona)
                                                            return !this.state.externalGroups.includes(persona?.text);
                                                        }
                                                        else if (this.state.isSameDomain) {
                                                            //console.log("persona in else ", persona)
                                                            let email = persona?.loginName || '';
                                                            return !email.includes('#ext#');
                                                        }
                                                    })}

                                            />

                                        </div>


                                        <div className="modal-buttons">

                                            <button
                                                className="newlogocolorbtn"
                                                onClick={this.updateUser}
                                            >
                                                Update
                                            </button>
                                            <button
                                                className="newblackcolorbtn"
                                                onClick={this.closeEditDialog}
                                            >
                                                Cancel
                                            </button>

                                        </div>
                                    </div>
                                </div>
                            </div>

                        )}


                        {this.state.isModalOpen && (
                            <div className="custom-dark-modal-overlay">
                                <div className="custom-dark-modal" style={{ minWidth: '800px', maxWidth: '900px' }}>
                                    <h3 className="addsteptitle" style={{ padding: '10px 20px' }}>
                                        Build your condition
                                    </h3>
                                    <ConditionBuilder
                                        // conditionalRow={this.state.conditionalRow}
                                        filteredjson={this.state.CurrentformJSON}
                                        onDataSend={this.handleChildData}
                                        databool={this.state.databool}

                                        {...(this.state.editconditionJson?.ConditionJson
                                            ? { initialConditions: JSON.parse(this.state.editconditionJson.ConditionJson) }
                                            : {})}
                                    />

                                    <div className="row g-3 mx-2">
                                        <div className="col-md-6" style={{ marginTop: 0 }}>
                                            <label className="form-label" style={{ fontSize: 13, fontWeight: 400 }}>Route to:</label>

                                            <PeoplePicker
                                                styles={{ root: { borderRadius: 'inherit' } }}
                                                context={this.getPeoplePickerContext()}
                                                personSelectionLimit={1}
                                                groupName={""}
                                                required={true}
                                                onChange={this.onChange}
                                                showHiddenInUI={false}
                                                principalTypes={[PrincipalType.User, PrincipalType.SharePointGroup]}
                                                resolveDelay={1000}
                                                placeholder="Search..."
                                                defaultSelectedUsers={this.state.notifyEmailsTitles}
                                                allowUnvalidated={true}
                                                ensureUser={true}
                                                resultFilter={(results: any[]) =>
                                                    results.filter(persona => {
                                                        if (this.state.isSameDomain && this.state.isSameDomain === "NO") {
                                                            let email = persona?.loginName || '';
                                                            //console.log("persona in if ", persona)
                                                            if (!email.includes('#ext#')) {
                                                                const x = this.state.externalGroups.find(group => group === persona?.text);
                                                                return persona?.text.includes(x);
                                                            }
                                                            return email.includes(`${this.state.domainfield}`);
                                                        } else if (this.state.isSameDomain && !persona?.imageUrl) {
                                                            //console.log("persona in else if ", persona)
                                                            return !this.state.externalGroups.includes(persona?.text);
                                                        }
                                                        else if (this.state.isSameDomain) {
                                                            //console.log("persona in else ", persona)
                                                            let email = persona?.loginName || '';
                                                            return !email.includes('#ext#');
                                                        }
                                                    })}
                                            />

                                        </div>

                                        {(!this.state.isParent) && <div className="col-md-6" style={{ marginTop: 0 }}>
                                            <label htmlFor="finalapprove" className="form-label " style={{ fontSize: 13, fontWeight: 400 }}>
                                                Should flow end with this approval
                                            </label>
                                            <div className="form-check form-switch" style={{ margin: "5px 0px 0px 15px" }}>
                                                <input
                                                    style={{ width: 35, height: 18 }}
                                                    className="form-check-input"
                                                    type="checkbox"
                                                    id="finalapprove"
                                                    checked={this.state.FinalApproveYesNo === 'Yes'}
                                                    onChange={(e) => this.setState({ FinalApproveYesNo: e.target.checked ? 'Yes' : 'No' })}
                                                />
                                                <label className="form-check-label" style={{ marginLeft: 10, marginTop: 3 }} htmlFor="finalapprove">
                                                    {this.state.FinalApproveYesNo === 'Yes' ? 'Yes' : 'No'}
                                                </label>
                                            </div>
                                        </div>}
                                    </div>
                                    <div className="modal-buttons" style={{ margin: "30px 20px 20px 22px" }}>
                                        <button className="newlogocolorbtn" onClick={this.databoolchange}>{!this.state.editconditionJson ? "Save" : "Update"}</button>
                                        {this.state.editconditionJson &&
                                            <button className="newlogocolorbtn" onClick={this.deleteCondition}>Delete</button>}
                                        <button className="newblackcolorbtn" onClick={this.conditionCloseDialog}>Close</button>
                                    </div>
                                </div>
                            </div>
                        )}

                        {this.state.showDialog && (
                            <div className="modal-overlay">
                                <div>
                                    {this.state.showAppSelectionError ? (
                                        <div className="error-alert">
                                            <div className="error-message">
                                                <strong>Please select a form name.</strong>
                                            </div>
                                            <div className="modal-buttons">
                                                <button className="Add_btn" onClick={this.closeDialog}>Cancel</button>
                                            </div>
                                        </div>
                                    ) : (

                                        <>
                                            <div className="custom-dark-modal-overlay">
                                                <div className="custom-dark-modal">
                                                    <h2 className="addsteptitle">Add Step</h2>
                                                    <div className="form-group" style={{ padding: '0px 15px' }}>
                                                        {/* Role dropdown */}
                                                        <div className="form-control-group">
                                                            <label htmlFor="roleSelect">Role:</label>
                                                            <select

                                                                value={this.state.selectedRole}
                                                                className="user-search-input"

                                                                onChange={(e) => this.setState({ selectedRole: e.target.value })}
                                                            >
                                                                {this.state.roleOptions.map((role, index) => (
                                                                    <option key={index} value={role}>
                                                                        {role}
                                                                    </option>
                                                                ))}
                                                            </select>
                                                        </div>

                                                        {/* Level */}
                                                        <div className="form-control-group">
                                                            <label>Level: {this.state.level}</label>
                                                        </div>

                                                        {/* People Picker or input */}
                                                        <div className="form-control-group">
                                                            <label>User/Group:</label>

                                                            <PeoplePicker
                                                                context={this.getPeoplePickerContext()}
                                                                personSelectionLimit={1}
                                                                groupName={""}
                                                                required={true}
                                                                onChange={(e: any) => this.getPersonData(e)}
                                                                showHiddenInUI={false}
                                                                principalTypes={[PrincipalType.User, PrincipalType.SharePointGroup]}
                                                                resolveDelay={1000}
                                                                placeholder="Select user or group"
                                                                allowUnvalidated={true}
                                                                ensureUser={true}
                                                                resultFilter={(results: any[]) =>
                                                                    results.filter(persona => {
                                                                        if (this.state.isSameDomain && this.state.isSameDomain === "NO") {
                                                                            let email = persona?.loginName || '';
                                                                            //console.log("persona in if ", persona)
                                                                            if (!email.includes('#ext#')) {
                                                                                const x = this.state.externalGroups.find(group => group === persona?.text);
                                                                                return persona?.text.includes(x);
                                                                            }
                                                                            return email.includes(`${this.state.domainfield}`);
                                                                        } else if (this.state.isSameDomain && !persona?.imageUrl) {
                                                                            //console.log("persona in else if ", persona)
                                                                            return !this.state.externalGroups.includes(persona?.text);
                                                                        }
                                                                        else if (this.state.isSameDomain) {
                                                                            //console.log("persona in else ", persona)
                                                                            let email = persona?.loginName || '';
                                                                            return !email.includes('#ext#');
                                                                        }
                                                                    })}
                                                            />

                                                        </div>

                                                        {/* Buttons */}
                                                        <div className="modal-buttons">
                                                            <button className="newlogocolorbtn" onClick={this.onAddApprover}>Save</button>
                                                            <button className="newblackcolorbtn" onClick={this.closeDialog}>Cancel</button>
                                                        </div>
                                                    </div>
                                                </div>
                                            </div>
                                        </>

                                    )}
                                </div>
                            </div>
                        )}
                        {this.state.saveallpopup && (
                            <div className="custom-dark-modal-overlay">
                                <div className="custom-dark-modal">
                                    <h2 className='addsteptitle' style={{ fontSize: '1rem' }}>Do you want to Finalize Workflow</h2>
                                    <div className="form-group" style={{ padding: '0px 20px' }}>
                                        <div className="form-control-group">
                                            <p>Once finalized the workflow can no longer be edited, Except for the Users in the Workflow</p>
                                        </div>
                                    </div>
                                    <div className="modal-buttons" style={{ padding: '10px 20px 20px 20px' }}>
                                        <div style={{ width: '50%' }} className='modal-buttons'>
                                            <button
                                                className=" newlogocolorbtn"
                                                onClick={() => {
                                                    //console.log("saveallpopup clicked");
                                                    void this.onsaveall();
                                                }}>          Yes
                                            </button>
                                            <button
                                                className=" newlogocolorbtn"
                                                onClick={() => {
                                                    this.setState({ saveallpopup: false });
                                                }}>          No
                                            </button>
                                        </div>
                                    </div>
                                </div>
                            </div>)

                        }
                        {this.state.showExistingWorkflowDialog && this.state.isWorkflowRequired !== 'No' && (
                            <div className="custom-dark-modal-overlay">
                                <div className="custom-dark-modal">
                                    <h2 className='addsteptitle' style={{ fontSize: '1rem' }}>
                                        Add Existing Workflow to "{this.state.formName}":
                                    </h2>
                                    <div className="form-group" style={{ padding: '0px 20px' }}>
                                        <div className="form-control-group">
                                            <label>Select an existing workflow to add:</label>
                                            <select
                                                className='form-select'
                                                style={{ fontSize: 12, marginTop: 5, padding: '6px 2rem 6px 8px' }}
                                                value={this.state.selectedExistingForm}
                                                onChange={this.onExistingFormSelectChange}
                                            >
                                                <option value="">Select Existing Workflow</option>
                                                {singleformmasteritems.map((item, index) => (
                                                    <option key={index} value={item.AppCode}>
                                                        {item.AppName}
                                                    </option>
                                                ))}
                                            </select>
                                        </div>
                                    </div>
                                    <div className="modal-buttons" style={{ padding: '10px 20px 20px 20px' }}>
                                        <div style={{ width: '50%' }} className='modal-buttons'>
                                            <button
                                                className=" newlogocolorbtn"
                                                onClick={this.addExistingWorkflow}
                                            >
                                                Add
                                            </button>
                                            <button
                                                className=" newlogocolorbtn"
                                                onClick={() => {
                                                    this.closeExistingWorkflowDialog();
                                                    this.setState({ showDialog: true });
                                                }}
                                            >
                                                Skip
                                            </button>
                                        </div>
                                        <div style={{ width: '50%', display: 'flex', alignSelf: 'center', justifyContent: 'flex-end' }} className='modal-buttons'>
                                            <button className="newblackcolorbtn" onClick={() => this.closeExistingWorkflowDialog()} >Cancel</button>
                                        </div>

                                    </div>
                                </div>
                            </div>
                        )}
                        {this.state.showDialog2 && (
                            <div className="custom-dark-modal-overlay">
                                <div className="custom-dark-modal">
                                    <h2 className="addsteptitle">Select final notification recipients</h2>
                                    <div className="form-group" style={{ padding: '0px 20px 10px 20px' }}>
                                        <div className="form-control-group" style={{ marginBottom: '2px' }}>
                                            <label  >Users/Groups:</label>
                                        </div>

                                        <PeoplePicker
                                            context={this.getPeoplePickerContext()}
                                            personSelectionLimit={5}
                                            groupName={""}
                                            required={true}
                                            onChange={this.onChange} // Capture selected user
                                            showHiddenInUI={false}
                                            principalTypes={[PrincipalType.User, PrincipalType.SharePointGroup]}
                                            resolveDelay={1000}
                                            placeholder="Select Users or Groups"
                                            allowUnvalidated={true}
                                            ensureUser={true}
                                            defaultSelectedUsers={this.state.notifyEmailsTitles}
                                            resultFilter={(results: any[]) =>
                                                results.filter(persona => {
                                                    if (this.state.isSameDomain && this.state.isSameDomain === "NO") {
                                                        let email = persona?.loginName || '';
                                                        //console.log("persona in if ", persona)
                                                        if (!email.includes('#ext#')) {
                                                            const x = this.state.externalGroups.find(group => group === persona?.text);
                                                            return persona?.text.includes(x);
                                                        }
                                                        return email.includes(`${this.state.domainfield}`);
                                                    } else if (this.state.isSameDomain && !persona?.imageUrl) {
                                                        //console.log("persona in else if ", persona)
                                                        return !this.state.externalGroups.includes(persona?.text);
                                                    }
                                                    else if (this.state.isSameDomain) {
                                                        //console.log("persona in else ", persona)
                                                        let email = persona?.loginName || '';
                                                        return !email.includes('#ext#');
                                                    }
                                                })}
                                        />

                                        <div className="modal-buttons" style={{ marginTop: "20px" }} >
                                            <button
                                                className="newlogocolorbtn"
                                                // style={{margin:'0 0 20px -10px' , fontSize:'15px', width:'50px'}}
                                                onClick={() => this.handleNotifyAddUser(Apptoshow)}

                                            >
                                                Add
                                            </button>
                                            <button
                                                className="newblackcolorbtn"
                                                onClick={() => {
                                                    this.closeExistingWorkflowDialog();
                                                    this.setState({ showDialog2: false });
                                                }}
                                            >
                                                Cancel
                                            </button>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        )}
                        {this.state.isdynamicstatusPopupOpen &&
                            <div className="custom-dark-modal-overlay">
                                <div className="custom-dark-modal">
                                    <h2 className="addsteptitle">Add Dynamic Statuses</h2>
                                    <div className="form-group" style={{ padding: '20px 20px 10px 20px' }}>
                                        {this.state.newStatus.map((status, index) => (
                                            <div key={index} className="status-input-group" >
                                                <input
                                                    type="text"
                                                    value={status}
                                                    placeholder="Enter status"
                                                    onChange={(e) => this.updateStatusField(index, e.target.value)}
                                                    className="status-input"
                                                />
                                                <button className="status-delete-btn" onClick={() => this.removeStatusField(index)}>✕</button>
                                            </div>
                                        ))}
                                        <button className="newlogocolorbtn" onClick={this.addStatusField}>Add Option</button>

                                    </div>
                                    <div className="modal-buttons" style={{ padding: '0px 20px 20px 20px' }}>
                                        {this.state.liststatuses.length > 0 ? (
                                            <button className="newlogocolorbtn" onClick={this.updateStatuses}>Update</button>
                                        ) : (
                                            <button className="newlogocolorbtn" onClick={this.saveStatuses}>Save</button>
                                        )}
                                        <button className="newblackcolorbtn" onClick={this.dynamicstatusclosePopup}>Cancel</button>
                                    </div>
                                </div>
                            </div>
                        }
                        {this.state.isSLAPopupOpen &&
                            <div className="custom-dark-modal-overlay">
                                <div className="custom-dark-modal" style={{ minWidth: 300 }} >
                                    <h2 className="addsteptitle">Configure SLA</h2>
                                    <div className="form-group" style={{ padding: '0px 20px 10px 20px' }}>
                                        <div className="form-control-group" style={{ flexDirection: 'row', marginBottom: '-20px' }}>
                                            <label htmlFor="slaenabled" style={{ paddingTop: 4 }} className="form-label">
                                                SLA:
                                            </label>
                                            <div className="form-check form-switch" style={{ marginLeft: "22px" }}>
                                                <input
                                                    style={{ width: 30, height: 16 }}
                                                    className="form-check-input"
                                                    type="checkbox"
                                                    id="slaenabled"
                                                    checked={this.state.isSLAEnabled}
                                                    onChange={(e) => this.setState({ isSLAEnabled: e.target.checked ? true : false }, () => console.log("wwwwwwwww", this.state.isSLAEnabled))}
                                                />
                                                <label className="form-check-label" style={{ paddingLeft: 2 }} htmlFor="slaenabled">
                                                    {this.state.isSLAEnabled ? "On" : "Off"}
                                                </label>
                                            </div>
                                        </div><br />
                                        <div className="form-group">
                                            <label>Time in hrs:</label>
                                            <input
                                                type="number"
                                                min={0}
                                                style={{
                                                    width: 200,
                                                    backgroundColor: this.state.isSLAEnabled ? "white" : "grey",
                                                    color: this.state.isSLAEnabled ? "black" : "#6c757d",
                                                    height: 25
                                                }}
                                                step="1"
                                                value={this.state.SLAinhours}
                                                onChange={(e) => {
                                                    const value = e.target.value;

                                                    // Prevent empty input or decimals, allow only whole numbers >= 1
                                                    if (/^\d+$/.test(value)) {
                                                        const intValue = Number(value);
                                                        this.setState({ SLAinhours: intValue });
                                                    }
                                                }}
                                                className={`form-control ${this.state.isSLAEnabled ? "formcontrol-extra" : "formcontrol-disabled"}`}
                                                disabled={!this.state.isSLAEnabled}
                                            />
                                        </div>
                                        <div className="modal-buttons" style={{ marginTop: 20 }}>
                                            <button className="newlogocolorbtn" onClick={this.saveSLA}>Save</button>
                                            <button className="newblackcolorbtn" onClick={this.SLApopupclose}>Cancel</button>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        }
                    </div>
                </div>
            </div >
        );
    }
}

export default ConfigureWorkflow;