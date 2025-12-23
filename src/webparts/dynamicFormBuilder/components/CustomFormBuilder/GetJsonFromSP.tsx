import * as React from 'react';
import { useState, useEffect } from 'react';
import { sp } from "@pnp/sp/presets/all";
import FormComponent from './FormComponent';
import { useGlobalState } from './GlobalVariable/GlobalStateContext';
// import { IDynamicFormBuilderProps } from '../IDynamicFormBuilderProps';
import './GetData.css';
import { Link as RouterLink, useNavigate } from 'react-router-dom';
// import SideNavBar from './SideNavBar';
import TopNavBar from './TopBar';
import LoadingSpinner from './Loading';
import { mySiteUrl } from './ConfigURL/All_URLs';
import { Icon } from '@fluentui/react/lib/Icon';
import { fetchTenantUser } from './FetchTenantUser/fetchTenantUser';
import { Toast } from 'primereact/toast';
import { useRef } from 'react';
import { myDomain } from './ConfigURL/All_URLs';
import SideBar from './Sidebar/SideBar';
import { dataservice } from './encryptionutil';

interface FormItem {
    AppName: string;
    Author: {
        Title: string;
    };
    FormJSON: string;
    AppCode: string;
    FormType: string; // Ensure this column is present in your SharePoint list
    VisibilityFlag: boolean;
    TransactionCount: number;
    Description: string;

}

interface FormType {
    Title: string; // Assuming 'Title' is the column holding the form type name
}

const GetData: React.FC<{ context: any }> = ({ context }) => {
    const [formItems, setFormItems] = useState<FormItem[]>([]); // List of form items
    const [formTypes, setFormTypes] = useState<FormType[]>([]); // List of form types
    //  const [selectedFormItem, setSelectedFormItem] = useState<FormItem | null>(null); // Selected form item
    // const [selectedFormJSON, setSelectedFormJSON] = useState<any[]>([]); // Form JSON data for selected form
    // const [selectedFormTitle, setSelectedFormTitle] = useState<string | null>(null); // Selected form title
    // const [selectedFormAppCode, setSelectedFormAppCode] = useState<string | null>(null); // Selected form AppCode
    const [error, setError] = useState<string | null>(null);
    const [isLoading, setIsLoading] = useState<boolean>(false);
    const [isFormSelected, setIsFormSelected] = useState<boolean>(false);
    const [selectedCategory, setSelectedCategory] = useState<string>('All');
    const [, setHoverData] = useState<string | null>(null)
    const [draftCount, setDraftCount] = useState(0);
    const [inProgressCount, setInProgressCount] = useState(0);
    // const [, setCountOfForms] = useState(0)
    // const [mappingCount, setMappingCount] = useState(0)
    const { setGlobalVariable } = useGlobalState();
    const [CurrentUserEmail, setCurrentUserEmail] = useState<string | null>(null);
    const [isAdmin, SetisAdmin] = useState<boolean>(false)
    const [deleteflag, setdeleteflag] = useState<boolean>(false)
    const [deleteitem, setdeleteitem] = useState<FormItem>()
    const [createFormToasterr, setCreateFormToasterr] = useState<boolean>();
    const [, setIsSameDomain] = useState("");
    const [activeMenu, setActiveMenu] = useState<string | null>("Home");



    const toast = useRef<Toast>(null);

    const navigate = useNavigate();


    sp.setup({
        sp: {
            baseUrl: mySiteUrl,
        },
    });

    const dataserviceobj = new dataservice();

    const fetchFormTitles = async () => {
        try {
            const Domain = await fetchTenantUser();
            setIsLoading(true);
            setError(null);

            // Fetch current user details
            const currentUser = await sp.web.currentUser.get();
            console.log("Current User ID:", currentUser.Id);


            // Fetch all forms where Domain is 'TenantUsers' and FormViewer is either a person or a group
            const RAWallForms = await sp.web.lists.getByTitle("FormMaster")
                .items
                .select("ID", "Title", "FormJSON", "AppName", "AppCode", "FormType", "Domain", "FormViewer/Id", "FormViewer/Title", "VisibilityFlag", "TransactionCount", "Description")
                .expand("FormViewer")
                .filter(`Domain eq '${Domain.TenantUsers}'`)
                .get();
            const allForms = RAWallForms?.map((item:any)=>{
                    return {...item,FormJSON: dataserviceobj.decryptjson(item.FormJSON)}
                })
            console.log("Fetched Forms:", allForms);
            console.log("CurrentUserEmail------------- Isssssss", CurrentUserEmail)

            if (CurrentUserEmail === "BothUser Admin" || CurrentUserEmail === "OnlyAdmin") {
                setFormItems(allForms);

                console.log("fetchTenantUser------------- IN  IFFFF", Domain)
            }
            else {
                console.log("fetchTenantUser------------- IN  ELSEEE", Domain)
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

                console.log("Group IDs to Check:", [...groupIdsToCheck]);

                // Fetch all site groups to determine which FormViewer ID belongs to a group
                const siteGroups = await sp.web.siteGroups.get();
                const siteGroupIds = new Set(siteGroups.map(group => group.Id));

                console.log("All Site Group IDs:", [...siteGroupIds]);

                // Filter forms where FormViewer is a group
                const formsWithGroups = allForms.filter(
                    form => form.FormViewer?.Id && siteGroupIds.has(form.FormViewer.Id)
                );

                console.log("Forms where FormViewer is a Group:", formsWithGroups);

                // Get logged-in user's group memberships
                const userGroups = await sp.web.currentUser.groups.get();
                const userGroupIds = new Set(userGroups.map(group => group.Id));
                const userGroupTitles = new Set(userGroups.map(group => group.Title));

                console.log("User's Group IDs:", [...userGroupIds]);
                console.log("User's Group Titles:", [...userGroupTitles]);

                // Filter forms where the user is a member of the FormViewer group
                const formsWhereUserIsGroupMember = formsWithGroups.filter(
                    form => form.FormViewer?.Id && userGroupIds.has(form.FormViewer.Id)
                );

                console.log("Forms where FormViewer is a Group AND User is a Member:", formsWhereUserIsGroupMember);

                // **New Condition: Also include forms where FormViewer Title is "All Employees"**
                const formsForAllEmployees = formsWithGroups.filter(
                    form => form.FormViewer?.Title === "All Employees"
                );

                console.log("Forms where FormViewer is 'All Employees':", formsForAllEmployees);

                // Merge both filtered lists
                let finalFilteredForms = [...formsWhereUserIsGroupMember, ...formsWithCurrentUser, ...formsForAllEmployees];

                console.log("Final Filtered Forms:", finalFilteredForms);


                setFormItems(finalFilteredForms);

            }

        } catch (error) {
            console.error("Error fetching forms:", error);
        } finally {
            setIsLoading(false);
        }
    };



    const fetchFormTypes = async () => {
        try {
            const types = await sp.web.lists.getByTitle("FormType").items
                .select("Title")
                .get();
            setFormTypes(types);
        } catch (error) {
            console.error("Error fetching form types:", error);
        }
    };


    const getFormDataByTitle = async (title: string) => {


        try {
            const selectedForm = formItems.find(item => item.AppCode === title);

            if (selectedForm) {
                //  setSelectedFormItem(selectedForm);
                // setSelectedFormJSON(JSON.parse(selectedForm.FormJSON));
                // setSelectedFormTitle(selectedForm.AppName);
                // setSelectedFormAppCode(selectedForm.AppCode);
                setIsFormSelected(true);
            }
        } catch (error) {
            console.error("Error getting form data by title:", error);
            setError("An error occurred while loading the form.");
        }
    };
    const isAdminUser = async () => {
        const currentUser = await sp.web.currentUser.get();
        const currentUserId = currentUser.Id
        const AuthlistAdminUsers = await sp.web.lists.getByTitle("AuthList").items
            .filter(`AdminUser eq 'Yes' and AuthName/Id eq ${currentUserId}`)
            .select("Title", "AuthName/Id", "AuthName/Title", "AdminUser")
            .expand("AuthName")
            .getAll()
        if (AuthlistAdminUsers.length > 0) {
            SetisAdmin(true)

        } else {
            SetisAdmin(false)

        }
    }

    async function getStatusCounts() {
        try {



            const currentUser = await sp.web.currentUser.get();
            const draftCount = await sp.web.lists.getByTitle('WorkFlowProcessData')
                .items.filter(`AuthorId eq ${currentUser.Id}`)
                .getAll();



            setDraftCount(draftCount.length);

            // console.log("In-Progress count:", inProgressCount.length);
            // console.log("Draft count:", draftCount.length);
        } catch (error) {
            console.error("Error fetching counts:", error);
        }
    }
    getStatusCounts().catch((error) => {
        console.error("Unhandled error:", error);
    });

    const fetchMappingMasterData = async (appCode: string) => {
        try {
            const mappingItems = await sp.web.lists.getByTitle("MappingMaster").items
                .filter(`AppCode eq '${appCode}'`)
                .select("Users_x002f_Groups/Title", "Level", "Role")
                .expand("Users_x002f_Groups")
                .get();

            if (mappingItems.length > 0) {
                const formattedData = mappingItems.map(item => ({
                    Title: item.Users_x002f_Groups?.Title,
                    Level: item.Level,
                    Role: item.Role,
                }));

                // Update hoverData with formatted information as a string
                setHoverData(
                    `${formattedData}`
                );
                console.log(`Filtered data for AppCode '${appCode}':`, formattedData[0].Title);
            } else {
                console.log(`No data found in MappingMaster list for AppCode '${appCode}'.`);
                setHoverData(null); // Reset hoverData if no items found
            }
        } catch (error) {
            console.error("Error fetching data from MappingMaster list:", error);
            setHoverData(null); // Reset hoverData on error
        }
    };


    // Navigate to the MappingMaster page
    const navigateToMappingMaster = (appCode: string | null, appName: string, authorName: string) => {
        if (appCode) {
            navigate(`/mapping-master/${encodeURIComponent(appCode)}?appName=${appName}&authorName=${authorName}`);
            // console.log((`/mapping-master/${appCode}?appName=${appName}&authorName=${authorName}`))
        }
    };

    const getFilteredForms = () => {
        console.log("formItems----------------------------------", formItems)
        if (selectedCategory === 'All') {
            debugger;
            return formItems.filter(item => item.VisibilityFlag === true);

        }
        console.log("formItems", formItems)
        return formItems.filter(item => item.FormType === selectedCategory && item.VisibilityFlag === true);
    };

    const openDeletePopup = async (item: any) => {
        const WorkflowItems = await sp.web.lists
            .getByTitle("WorkFlowProcessData")
            .items
            .filter(`AppName eq '${item.AppName}'`)
            .top(1)
            .get();
        console.log("WorkflowItems in delete", WorkflowItems)
        toast.current?.clear();
        if (WorkflowItems.length === 0) {
            setdeleteflag(true)
            setdeleteitem(item)
        }
        else {

            toast.current?.show({
                severity: 'error',
                summary: 'Cannot Delete',
                detail: 'This form has existing records',
                life: 3000,
                style: { height: 70 }
            });
        }
    }
    const closeDeletePopup = () => {
        setdeleteflag(false)
    }

    const handleDeleteClick = async (item: any) => {
        console.log("deleteFormItem", item);


        try {


            await sp.web.lists.getByTitle("FormMaster").items.getById(Number(item.Id)).delete();
            await fetchFormTitles()
            getFilteredForms()
            toast.current?.clear();
            toast.current?.show({
                severity: 'success',
                summary: `Form: ${item.AppName}`,
                detail: 'Form and all related records deleted successfully',
                life: 3000,
                style: { height: 100 }
            });

            const [mappingMasterItems, routingRulesItems] = await Promise.all([
                sp.web.lists.getByTitle("MappingMaster")
                    .items.filter(`AppCode eq '${item.AppCode}'`).get(),
                sp.web.lists.getByTitle("RoutingRules")
                    .items.filter(`AppCode eq '${item.AppCode}'`).get()
            ]);

            const deleteOperations = [];

            if (mappingMasterItems.length > 0) {
                deleteOperations.push(...mappingMasterItems.map(i =>
                    sp.web.lists.getByTitle("MappingMaster").items.getById(i.Id).delete()
                ));
            }

            if (routingRulesItems.length > 0) {
                deleteOperations.push(...routingRulesItems.map(i =>
                    sp.web.lists.getByTitle("RoutingRules").items.getById(i.Id).delete()
                ));
            }

            console.log("delllllllllllllllll", deleteOperations)
            // Execute all deletions
            await Promise.all(deleteOperations);

        } catch (error) {
            console.error("form or routing or mapping delete failed", error)
            toast.current?.clear();
            toast.current?.show({
                severity: 'error',
                summary: `Form: ${item.AppName}`,
                detail: 'Form or related records delete failed',
                life: 3000,
            });

        }
        closeDeletePopup()


    }

    const handleEditClick = async (item: FormItem) => {
        console.log("FormItem", item);

        try {
            const WorkflowItems = await sp.web.lists
                .getByTitle("WorkFlowProcessData")
                .items
                .filter(`AppCode eq '${item.AppCode}' and (Status eq 'Draft' or Status eq 'In-Progress' or Status eq 'Returned')`)
                .top(1)
                .get();

            if (WorkflowItems.length > 0) {
                console.warn("Already records exist in Draft/Returned/In-Progress, so cannot edit.");
                toast.current?.clear();
                toast.current?.show({
                    severity: 'error',
                    summary: 'Form Exists!',
                    detail: 'This form has existing records in Draft/In-Progress/Returned state. Complete it before editing!',
                    life: 3000,
                    style: { height: 100 }
                });


                return;
            }

            // Prepare form data for editing
            const EditFormData = {
                isEdit: true,
                AppCode: item.AppCode,
                AppName: item.AppName,
                FormType: item.FormType,
                FormJSON: dataserviceobj.encryptjson(item.FormJSON),
                Description: item.Description,
            };

            sessionStorage.setItem("EditFormData", JSON.stringify(EditFormData));
            console.log("EditFormData:", EditFormData);

            // Navigate to the CreateForm page
            navigate('/CreateForm');
        } catch (error) {
            console.error("Error handling item click:", error);
        }
    };

    // const goBack = () => {
    //     setIsFormSelected(false);
    //     // setSelectedFormJSON([]);
    //     // setSelectedFormTitle(null);
    //     // setSelectedFormAppCode(null);
    // };

    const fetchData = async () => {
        try {

            const currentUser = await sp.web.currentUser.get();
            console.log("Current User ID:", currentUser.Id);

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
            };

            const allItems: WorkFlowItem[] = await sp.web.lists
                .getByTitle("WorkFlowProcessData")
                .items.orderBy("Modified", false) // Order by Modified (Descending)
                .select(
                    "ID",
                    "Title",
                    "Status",
                    "Created",
                    "Modified",
                    "Author/Title",
                    "CurApprover/Title",
                    "CurApproverId",
                    "AppName"
                )
                .expand("Author", "CurApprover")
                .filter(`Status eq 'In-Progress'`)
                .getAll();

            console.log("Fetched Items:", allItems);

            let groupIdsToCheck = new Set<number>();
            let recordsWithCurrentUser: WorkFlowItem[] = [];

            allItems.forEach((item) => {
                if (item.CurApproverId === currentUser.Id) {
                    recordsWithCurrentUser.push(item);
                } else if (item.CurApproverId) {
                    groupIdsToCheck.add(item.CurApproverId);
                }
            });

            console.log("Group IDs to Check:", [...groupIdsToCheck]);

            const siteGroups = await sp.web.siteGroups.get();
            const siteGroupIds = new Set(siteGroups.map((group) => group.Id));

            console.log("All Site Group IDs:", [...siteGroupIds]);

            const recordsWithGroups = allItems.filter(
                (item) => item.CurApproverId && siteGroupIds.has(item.CurApproverId)
            );

            console.log("Records where CurApprover is a Group:", recordsWithGroups);

            const userGroups = await sp.web.currentUser.groups.get();
            const userGroupIds = new Set(userGroups.map((group) => group.Id));

            console.log("User's Group IDs:", [...userGroupIds]);

            const recordsWhereUserIsGroupMember = recordsWithGroups.filter(
                (item) => item.CurApproverId && userGroupIds.has(item.CurApproverId)
            );

            console.log(
                "Records where CurApprover is a Group AND User is a Member:",
                recordsWhereUserIsGroupMember
            );

            let finalFilteredRecords = [
                ...recordsWhereUserIsGroupMember,
                ...recordsWithCurrentUser,
            ];

            finalFilteredRecords.sort(
                (a, b) =>
                    new Date(b.Modified).getTime() - new Date(a.Modified).getTime()
            );

            console.log("Final Sorted Records:", finalFilteredRecords);

            setInProgressCount(finalFilteredRecords.length);

        } catch (error) {
            console.error("Error fetching data from SharePoint list:", error);
        }
    };

    useEffect(() => {
        sp.setup({
            spfxContext: context as any,  // Use 'context' directly, not 'props.context'
        });
        fetchFormTitles().catch((error) => {
            console.error("Error in fetchFormTitles:", error);
        });
        fetchFormTypes().catch((error) => {
            console.error("Error in fetchFormTypes:", error);
        });
        void fetchData()
        isAdminUser().catch((error) => {
            console.error("Error in isAdminUser:", error);
        });
    }, [context, CurrentUserEmail]);
    const Domainfilter = async () => {
        const x = await fetchTenantUser();
        // setFetchDomain(x.TenantUsers);
        if (myDomain !== x.TenantUsers) {
            const domainlistItem = await sp.web.lists.getByTitle("DomainList")
                .items.select("Id", "LicenseType")
                .filter(`Tenant eq '${x.TenantUsers}'`).top(1).get();
            const licenseListItem = await sp.web.lists.getByTitle("LicenseList")
                .items.select("Id", "FormCount", "TransactionsPerForm")
                .filter(`LicenseType eq '${domainlistItem[0].LicenseType}'`).top(1).get();
            const formMasterItem = await sp.web.lists.getByTitle("FormMaster")
                .items.select("Id", "TransactionCount")
                .filter(`Domain eq '${x.TenantUsers}' and VisibilityFlag eq true`).top(5000).getAll();
            formMasterItem.length >= licenseListItem[0].FormCount ? setCreateFormToasterr(true) : setCreateFormToasterr(false)
        }
        else {
            setCreateFormToasterr(false)
        }


    }
    useEffect(() => {
        void Domainfilter();
    }, [])
    const handleItemClick = async (item: FormItem) => {
        try {

            await getFormDataByTitle(item.AppCode);


            setGlobalVariable(JSON.stringify({ appcode: item.AppCode, appname: item.AppName }));
        } catch (error) {
            console.error("Error handling item click:", error);
        }
    };

    const handleUserEmailRetrieved = (UserAUTH: string) => {
        setCurrentUserEmail(UserAUTH);
        console.log("Logged-in USER IS   ----------------:", UserAUTH);
    };
    const handleUserDomain = (UserDomain: string) => {
        setIsSameDomain(UserDomain)

    };

    return (

        <div style={{ display: 'flex', flexDirection: "column", overflowY: 'scroll' }}>
            <TopNavBar onUserTypeRetrieved={handleUserEmailRetrieved} onUserDomainRetrieved={handleUserDomain} />

            <div style={{ display: 'flex' }}>
                {!isFormSelected && <SideBar activeMenu={activeMenu} onItemClick={(menu) => setActiveMenu(menu)} />}
                <Toast ref={toast} />
                <div style={{ width: '100%' }}>
                    {!isFormSelected ? (
                        <>
                            <div style={{ margin: "20px", flex: 1 }}>
                                <div className='allRecardsCards'>


                                    {CurrentUserEmail === "OnlyUser" ? (
                                        <div className="disabledLink hideInMobile">
                                            <div className="Recard-Cards" title='You dont have a Admin Access'>
                                                <div className="Create-Card_Disable">
                                                    <h6>Create Form</h6>
                                                    <div className="Icon-Count">
                                                        <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" fill="none">
                                                            <line x1="12" y1="5" x2="12" y2="19" stroke="gray"></line>
                                                            <line x1="5" y1="12" x2="19" y2="12" stroke="gray"></line>
                                                        </svg>

                                                        {/* <p>{countOfForms}</p> */}
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                    ) :

                                        (<RouterLink to="/CreateForm" className='RoutLink'
                                            onClick={(event) => {
                                                if (!createFormToasterr) {
                                                    const EditFormData = {
                                                        isEdit: false,
                                                        AppCode: "",
                                                        AppName: "",
                                                        FormType: "",
                                                        FormJSON: "",
                                                        Description: "",
                                                    };

                                                    sessionStorage.setItem("EditFormData", JSON.stringify(EditFormData)); // Store JSON as string
                                                }
                                                else {
                                                    event.preventDefault();
                                                    toast.current?.clear();
                                                    toast.current?.show({
                                                        severity: 'warn',
                                                        summary: 'Freemium subscription limit has reached ',
                                                        detail: `Form creation limit Reached. Please upgrade your license`,
                                                        life: 10000,
                                                        style: { height: 100 }
                                                    });
                                                }

                                            }}
                                        >
                                            <div className='Recard-Cards hideInMobile'>
                                                <div className='Create-Card'>
                                                    <h6>Create Form</h6>
                                                    <div className='Icon-Count'>
                                                        <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="feather feather-plus">
                                                            <line x1="12" y1="5" x2="12" y2="19"></line>
                                                            <line x1="5" y1="12" x2="19" y2="12"></line>
                                                        </svg>
                                                        {/* <p>{countOfForms}</p> */}
                                                    </div>
                                                </div>
                                            </div>
                                        </RouterLink>
                                        )}


                                    {CurrentUserEmail === "OnlyUser" ? (
                                        <div className="disabledLink hideInMobile">
                                            <div className="Recard-Cards" title='You dont have a Admin Access'>
                                                <div className="Create-Card_Disable">
                                                    <h6>Create WorkFlow</h6>
                                                    <div className="Icon-Count2">
                                                        <svg xmlns="http://www.w3.org/2000/svg" className="icon icon-tabler icon-tabler-settings" width="24" height="24" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor" fill="none" stroke-linecap="round" stroke-linejoin="round">
                                                            <path stroke="none" d="M0 0h24v24H0z" fill="none" />
                                                            <path d="M12 8a4 4 0 1 0 0 8a4 4 0 1 0 0 -8" stroke="gray" />
                                                            <path d="M4.6 9l-.6 -2.3l2.3 -.6l1.4 2.2a7.8 7.8 0 0 0 -1 1.6l-2.1 -.9z" stroke="gray" />
                                                            <path d="M19.4 9l.6 -2.3l-2.3 -.6l-1.4 2.2a7.8 7.8 0 0 0 1 1.6l2.1 -.9z" stroke="gray" />
                                                            <path d="M12 4.5a7.5 7.5 0 0 1 3.8 1" stroke="gray" />
                                                            <path d="M12 19.5a7.5 7.5 0 0 1 -3.8 -1" stroke="gray" />
                                                            <path d="M9 12l-2 1" stroke="gray" />
                                                            <path d="M15 12l2 -1" stroke="gray" />
                                                        </svg>
                                                        {/* <p>{mappingCount}</p> */}
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                    ) : (
                                        <RouterLink to="/ConfigureWorkflow" className="RoutLink">
                                            <div className="Recard-Cards hideInMobile">
                                                <div className="Create-Card2">
                                                    <h6>Create WorkFlow</h6>
                                                    <div className="Icon-Count2">
                                                        <svg xmlns="http://www.w3.org/2000/svg" className="icon icon-tabler icon-tabler-settings" width="24" height="24" viewBox="0 0 24 24" stroke-width="1.5" stroke="currentColor" fill="none" stroke-linecap="round" stroke-linejoin="round">
                                                            <path stroke="none" d="M0 0h24v24H0z" fill="none" />
                                                            <path d="M12 8a4 4 0 1 0 0 8a4 4 0 1 0 0 -8" />
                                                            <path d="M4.6 9l-.6 -2.3l2.3 -.6l1.4 2.2a7.8 7.8 0 0 0 -1 1.6l-2.1 -.9z" />
                                                            <path d="M19.4 9l.6 -2.3l-2.3 -.6l-1.4 2.2a7.8 7.8 0 0 0 1 1.6l2.1 -.9z" />
                                                            <path d="M12 4.5a7.5 7.5 0 0 1 3.8 1" />
                                                            <path d="M12 19.5a7.5 7.5 0 0 1 -3.8 -1" />
                                                            <path d="M9 12l-2 1" />
                                                            <path d="M15 12l2 -1" />
                                                        </svg>
                                                        {/* <p>{mappingCount}</p> */}
                                                    </div>
                                                </div>
                                            </div>
                                        </RouterLink>
                                    )}

                                    {CurrentUserEmail !== "OnlyAdmin" ? (
                                        <RouterLink to="/InProcess" className='RoutLink'>
                                            <div className='Recard-Cards'>
                                                <div className='Create-Card'>
                                                    <h6>Pending Approvals</h6>
                                                    <div className='Icon-Count'>
                                                        <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="feather feather-file-text">
                                                            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                                                            <polyline points="14 2 14 8 20 8"></polyline>
                                                            <line x1="16" y1="13" x2="8" y2="13"></line>
                                                            <line x1="16" y1="17" x2="8" y2="17"></line>
                                                            <line x1="10" y1="9" x2="8" y2="9"></line>
                                                        </svg>
                                                        <p>{inProgressCount}</p>
                                                    </div>
                                                </div>
                                            </div>
                                        </RouterLink>
                                    ) : (
                                        <div className="disabledLink hideInMobile">
                                            <div className="Recard-Cards" title='You dont have a Admin Access'>
                                                <div className="Create-Card_Disable">
                                                    <h6 style={{ marginLeft: "2px" }}>Pending Approvals</h6>
                                                    <div className='Icon-Count2'>
                                                        <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="feather feather-file-text">
                                                            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" stroke="gray"></path>
                                                            <polyline points="14 2 14 8 20 8" stroke="gray"></polyline>
                                                            <line x1="16" y1="13" x2="8" y2="13" stroke="gray"></line>
                                                            <line x1="16" y1="17" x2="8" y2="17" stroke="gray"></line>
                                                            <line x1="10" y1="9" x2="8" y2="9" stroke="gray"></line>
                                                        </svg>
                                                        <p>{inProgressCount}</p>
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                    )}

                                    {CurrentUserEmail !== "OnlyAdmin" ? (
                                        <RouterLink to="/Draft" className='RoutLink'>
                                            <div className='Recard-Cards'>
                                                <div className='Create-Card2'>
                                                    <h6>My Documents</h6>
                                                    <div className='Icon-Count-Draft'>
                                                        <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="feather feather-file-text">
                                                            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                                                            <polyline points="14 2 14 8 20 8"></polyline>
                                                            <line x1="16" y1="13" x2="8" y2="13"></line>
                                                            <line x1="16" y1="17" x2="8" y2="17"></line>
                                                            <line x1="10" y1="9" x2="8" y2="9"></line>
                                                        </svg>
                                                        <p>{draftCount}</p>
                                                    </div>
                                                </div>
                                            </div>
                                        </RouterLink>
                                    ) : (
                                        <div className="disabledLink hideInMobile">
                                            <div className="Recard-Cards" title='You dont have a Admin Access'>
                                                <div className="Create-Card_Disable">
                                                    <h6>My Documents</h6>
                                                    <div className='Icon-Count2'>
                                                        <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="gray" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" className="feather feather-file-text">
                                                            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" stroke="gray"></path>
                                                            <polyline points="14 2 14 8 20 8" stroke="gray"></polyline>
                                                            <line x1="16" y1="13" x2="8" y2="13" stroke="gray"></line>
                                                            <line x1="16" y1="17" x2="8" y2="17" stroke="gray"></line>
                                                            <line x1="10" y1="9" x2="8" y2="9" stroke="gray"></line>
                                                        </svg>

                                                        <p>{draftCount}</p>
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                    )}

                                    <RouterLink to="/Reports" className='RoutLink'>
                                        <div className='Recard-Cards'>
                                            <div className='Create-Card'>
                                                <h6>Generate Reports</h6>
                                                <div className='Icon-Count'>
                                                    <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="feather feather-file-text" xmlns="http://www.w3.org/2000/svg">
                                                        <path d="m20 8-6-6H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8zM9 19H7v-9h2v9zm4 0h-2v-6h2v6zm4 0h-2v-3h2v3zM14 9h-1V4l5 5h-4z"></path></svg>

                                                </div>
                                            </div>
                                        </div>
                                    </RouterLink>

                                </div>
                            </div>
                            <div className='Search-container'>
                                <div style={{ display: 'none' }}>
                                    <input className='searchBar' type='search' placeholder="Ask me anything?" />
                                    <img className='mic' src={require("../../assets/Images/mic.png")} />
                                </div>
                            </div>
                            <div className='AvailableFormsContainer'>
                                <div>
                                    <ul className='Select_Btn'>
                                        <li
                                            style={{ marginLeft: '-40px', }}
                                            className={selectedCategory === 'All' ? 'onSelect' : 'onOthers'}
                                            onClick={() => setSelectedCategory('All')}
                                        >
                                            All
                                        </li>
                                        {formTypes.map((type) => (
                                            <li
                                                key={type.Title}
                                                className={selectedCategory === type.Title ? 'onSelect' : 'onOthers'}
                                                onClick={() => setSelectedCategory(type.Title)}
                                            >
                                                {type.Title}
                                            </li>
                                        ))}
                                    </ul>
                                </div>
                                {error && <div style={{ color: "red", margin: "10px" }}>{error}</div>}
                                {isLoading ? (
                                    <LoadingSpinner />
                                ) : (
                                    <div className="card-container" style={{ height: '440px', width: '100%' }}>

                                        {console.log("getFilteredForms in getjosnfromSP", getFilteredForms())}
                                        {getFilteredForms().length > 0 ? (
                                            getFilteredForms().map((item) => (
                                                <div className="card" key={item.AppName}>
                                                    <div style={{ display: 'flex', justifyContent: 'right', marginBottom: '-16px' }}>
                                                    </div>
                                                    <RouterLink
                                                        to={{
                                                            pathname: '/FormDashBoard',
                                                        }}
                                                        state={{ appcode: item.AppCode, appname: item.AppName }}
                                                        className='RoutLink'
                                                        onClick={() => handleItemClick(item)}
                                                    >
                                                        <div className='Card_Icon_Heading'>
                                                            <svg xmlns="http://www.w3.org/2000/svg" className="form-pen-icon" viewBox="0 0 24 24">
                                                                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                                                                <polyline points="14 2 14 8 20 8"></polyline>
                                                                <line x1="16" y1="13" x2="8" y2="13"></line>
                                                                <line x1="16" y1="17" x2="8" y2="17"></line>
                                                                <line x1="10" y1="9" x2="8" y2="9"></line>
                                                                <path d="M16 21l2-2a1 1 0 0 0 0-1.41l-1.5-1.5a1 1 0 0 0-1.41 0l-2 2"></path>
                                                                <line x1="17" y1="16" x2="21" y2="20"></line>
                                                            </svg>
                                                            <h5 className='CardHeading'>{item.AppName}</h5>


                                                        </div>
                                                    </RouterLink>

                                                    <div className='ButtonContainer'>
                                                        <RouterLink
                                                            to={{
                                                                pathname: '/FormComponent',
                                                            }}
                                                            state={{ selectedformdetails:item }}
                                                            className='RoutLink'
                                                            onClick={async () => {
                                                                const domain = await fetchTenantUser();
                                                                await handleItemClick(item);
                                                                const domainName = domain.TenantUsers
                                                                if (myDomain === domainName) {
                                                                    await getFormDataByTitle(item.AppCode)
                                                                    console.log("I AM IN IFFFFFFFFFFFFF")
                                                                }
                                                                else {
                                                                    console.log("I AM IN ELSEEEEEEEEEEEEE")
                                                                    const domainlistItem = await sp.web.lists.getByTitle("DomainList")
                                                                        .items.select("Id", "LicenseType")
                                                                        .filter(`Tenant eq '${domainName}'`).top(1).get();
                                                                    const licenseListItem = await sp.web.lists.getByTitle("LicenseList")
                                                                        .items.select("Id", "FormCount", "TransactionsPerForm")
                                                                        .filter(`LicenseType eq '${domainlistItem[0].LicenseType}'`).top(1).get();
                                                                    const formMasterItem = await sp.web.lists.getByTitle("FormMaster")
                                                                        .items.select("Id", "TransactionCount")
                                                                        .filter(`AppCode eq '${item.AppCode}'`).top(1).get();
                                                                    if (formMasterItem[0].TransactionCount >= licenseListItem[0].TransactionsPerForm) {
                                                                        toast.current?.clear();
                                                                        toast.current?.show({
                                                                            severity: 'warn',
                                                                            summary: 'Freemium subscription limit has reached ',
                                                                            detail: `${item.AppName}: Reached ${formMasterItem[0].TransactionCount} transactions limit. Please upgrade your license`,
                                                                            life: 10000,
                                                                            style: { height: 100 }
                                                                        });
                                                                    }
                                                                    else {
                                                                        await getFormDataByTitle(item.AppCode)
                                                                    }

                                                                }
                                                            }}
                                                        >
                                                            <button className='CreateNew_Btn'
                                                            >Create New +</button>
                                                        </RouterLink>
                                                        {/* <button onClick={updateGlobalState}>click</button> */}
                                                        <div style={{ display: 'flex', justifyContent: 'center', width: '35%', gap: '10px' }}>
                                                            <div className="info-icon-container">
                                                                <Icon
                                                                    iconName="Info"
                                                                    className="info-icon"
                                                                    onClick={() => navigateToMappingMaster(item.AppCode, item.AppName, item.Author?.Title)}
                                                                    onMouseEnter={() => fetchMappingMasterData(item.AppCode)}
                                                                />
                                                                <div className="tooltip">WorkFlow Process</div>
                                                            </div>
                                                            {isAdmin &&
                                                                (<><div className="info-icon-containerr">
                                                                    <Icon
                                                                        iconName="Edit"
                                                                        className="edit-icon"
                                                                        onClick={() => handleEditClick(item)}
                                                                    />


                                                                    <div className="Edit">Edit Form</div> {/* Custom tooltip */}
                                                                </div>
                                                                    <div className="info-icon-containerr">
                                                                        <Icon
                                                                            iconName="Delete"
                                                                            className="edit-icon"
                                                                            onClick={() => openDeletePopup(item)}
                                                                        />


                                                                        <div className="Edit" style={{ width: 65 }}>Delete Form</div> {/* Custom tooltip */}
                                                                    </div></>)}
                                                            {item.AppCode.match(/\/\d+$/) && (
                                                                <div>
                                                                    <span className="new-version-badge" title='New Version of Form'>V{item.AppCode.split('/')[1]}</span>
                                                                </div>
                                                            )}


                                                        </div>

                                                    </div>
                                                    {/* <div className='HoverCard'>
                                                        {hoverData ? (
                                                            <p>{hoverData[1]}</p>
                                                        ) : (
                                                            ""
                                                        )}
                                                    </div> */}

                                                </div>
                                            ))
                                        ) : (
                                            // <div>No forms available in this category.</div>
                                            ""
                                        )}
                                    </div>
                                )}
                            </div>
                        </>
                    ) : (
                        <div style={{ display: 'flex', flexDirection: "column" }}>
                            {/* <TopNavBar /> */}
                            <div style={{ display: 'flex' }}>
                                {/* <SideNavBar /> */}
                                <SideBar activeMenu={""}></SideBar>
                                <div style={{ width: '100%' }}>
                                    <FormComponent
                                        context={context}
                                    />
                                </div>
                            </div>
                        </div>
                    )}

                    <div>
                        {deleteflag && <div className="modal-overlay">
                            <div className="popupsmall" style={{ minHeight: 180 }}>
                                <h5 style={{ fontSize: '1.25rem' }}>{`Are you sure you want to delete ${deleteitem?.AppName} form?`}</h5>
                                <div className="modal-buttonss" style={{ gap: '5%', marginTop: '30px' }}>
                                    <div>
                                        <button
                                            className="Add_btn"
                                            style={{ marginRight: '20px' }}
                                            onClick={() => handleDeleteClick(deleteitem)}
                                        >
                                            Yes
                                        </button>
                                        <button
                                            className="Add_btn"
                                            onClick={closeDeletePopup}
                                        >
                                            No
                                        </button>
                                    </div>
                                </div>
                            </div>
                        </div>}
                    </div>
                </div>
            </div>
        </div>
    );
};

export default GetData;
