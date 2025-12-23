import * as React from 'react';
import { useState, useEffect, useRef } from 'react';
import { sp } from "@pnp/sp/presets/all";

// import './GetData.css';
import { Link as RouterLink, useNavigate } from 'react-router-dom';
// import SideNavBar from './SideNavBar';
// import TopNavBar from './TopBar';
import { dataservice } from '../encryptionutil';
import LoadingSpinner from '../Loading';
import { mySiteUrl } from '../ConfigURL/All_URLs';
// Update the import path and extension as per your project structure, for example:
import TotalRequests from '../../../assets/Images/Staticcardicons/TotalRequests.svg';
import PendingRequests from '../../../assets/Images/Staticcardicons/PendingRequests.svg';
import CompletedRequests from '../../../assets/Images/Staticcardicons/CompletedRequests.svg';
import threedots from '../../../assets/Images/threedots.svg';
import FormIcon from '../../../assets/Images/Staticcardicons/FormIcon.svg';
import Editinthreedots from '../../../assets/Images/Staticcardicons/Editinthreedots.svg';
import Deleteinthreedots from '../../../assets/Images/Staticcardicons/Deleteinthreedots.svg';
import Workflowinthreedots from '../../../assets/Images/Staticcardicons/Workflowinthreedots.svg';
// Or, if it's a TypeScript/JS file:
// import { PendingRequests, TotalRequests } from '../../../assets/Images/Staticcardicons';
import { fetchTenantUser } from '../FetchTenantUser/fetchTenantUser';
import { useGlobalState } from '../GlobalVariable/GlobalStateContext';
// import { Toast } from 'primereact/toast';
// import { useRef } from 'react';
import { myDomain } from '../ConfigURL/All_URLs';
import SideBar from '../Sidebar/SideBar';
import TopNavBar from '../TopBar';
import './GlobalDashboard.css';
import { Toast } from 'primereact/toast';
import { fetchWorkflowLimit } from '../CustomHooks/TransactionCount';

interface FormItem {
    ID: number;
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
    Published: boolean;
    IsParentForm: boolean;
}
interface Tile {
    icon: string;
    title: string;
    count: number;
    countcolor: string; // Optional: if you want to set a specific color for the count
}
interface FormType {
    Title: string; // Form category in global dashboard
}
const GlobalDashboard: React.FC<{ context: any }> = ({ context }) => {
    const [formItems, setFormItems] = useState<FormItem[]>([]); // List of form items from formmaster list
    const [formTypes, setFormTypes] = useState<FormType[]>([]); // List of form category  from formtype list
    const [CurrentUserEmail, setCurrentUserEmail] = useState<string | null>(null);
    const [isSameDomain, setIsSameDomain] = useState("");
    const [statictiles, setStatictiles] = useState<Tile[]>([]); // List of static tiles
    const [selectedCategory, setSelectedCategory] = useState<string>('All');
    const [isLoading, setIsLoading] = useState<boolean>(false);
    const [selectedFormItem, setSelectedFormItem] = useState<FormItem | null>(null); // Selected form item
    const [deleteflag, setdeleteflag] = useState<boolean>(false)
    const [deleteitem, setdeleteitem] = useState<FormItem>()
    const [UnPublishflag, setUnPublishflag] = useState<boolean>(false)
    const [Publishflag, setPublishflag] = useState<boolean>(false)
    const [Publishitem, setPublishitem] = useState<FormItem>()
    const { setGlobalVariable } = useGlobalState();
    const toast = useRef<Toast>(null);
    const menuRef = useRef<HTMLDivElement>(null); // Ref for the popup menu
    const navigate = useNavigate();

    sp.setup({
        sp: {
            baseUrl: mySiteUrl,
        },
    });

    const dataserviceobj = new dataservice();

    const getFormDataByTitle = async (title: string) => {
        try {
            const selectedForm = formItems.find(item => item.AppCode === title);
            if (selectedForm) {
                setSelectedFormItem(selectedForm);
                
            }
        } catch (error) {
            console.error("Error getting form data by title:", error);
            // setError("An error occurred while loading the form.");
        }
    };
    const handleItemClick = async (item: FormItem) => {
        try {

            await getFormDataByTitle(item.AppCode);


            setGlobalVariable(JSON.stringify({ appcode: item.AppCode, appname: item.AppName }));
        } catch (error) {
            console.error("Error handling item click:", error);
        }
    };

    const fetchFormTitles = async () => {
        try {
            const Domain = await fetchTenantUser();
            setIsLoading(true);
            const currentUser = await sp.web.currentUser.get();
            //console.log("Current User ID:", currentUser.Id);


            // Fetch all forms where Domain is 'TenantUsers' and FormViewer is either a person or a group
            const RAWallForms = await sp.web.lists.getByTitle("FormMaster")
                .items
                .select("ID", "Title", "FormJSON", "AppName", "AppCode", "FormType", "Domain", "FormViewer/Id", "FormViewer/Title", "VisibilityFlag", "TransactionCount", "Description", "Published", "IsParentForm")
                .expand("FormViewer")
                .filter(`Domain eq '${Domain.TenantUsers}' and ChildOrder eq null`)
                .getAll();
            //console.log("Fetched Forms:", RAWallForms);
            //console.log("CurrentUserEmail------------- Isssssss", CurrentUserEmail)
            const allForms = RAWallForms?.map((item:any)=>{
                return {...item,FormJSON: dataserviceobj.decryptjson(item.FormJSON)}
            })
            if (CurrentUserEmail === "BothUser Admin" || CurrentUserEmail === "OnlyAdmin") {
                setFormItems(allForms);

                //console.log("fetchTenantUser------------- IN  IFFFF", Domain)
            }
            else if (CurrentUserEmail === "OnlyUser") {
                //console.log("fetchTenantUser------------- IN  ELSEEE", Domain)
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

                //console.log("Group IDs to Check:", [...groupIdsToCheck]);

                // Fetch all site groups to determine which FormViewer ID belongs to a group
                const siteGroups = await sp.web.siteGroups.get();
                const siteGroupIds = new Set(siteGroups.map(group => group.Id));

                //console.log("All Site Group IDs:", [...siteGroupIds]);

                // Filter forms where FormViewer is a group
                const formsWithGroups = allForms.filter(
                    form => form.FormViewer?.Id && siteGroupIds.has(form.FormViewer.Id)
                );

                //console.log("Forms where FormViewer is a Group:", formsWithGroups);

                // Get logged-in user's group memberships
                const userGroups = await sp.web.currentUser.groups.get();
                const userGroupIds = new Set(userGroups.map(group => group.Id));
                const userGroupTitles = new Set(userGroups.map(group => group.Title));

                //console.log("User's Group IDs:", [...userGroupIds]);
                console.log("User's Group Titles:", [...userGroupTitles]);

                // Filter forms where the user is a member of the FormViewer group
                const formsWhereUserIsGroupMember = formsWithGroups.filter(
                    form => form.FormViewer?.Id && userGroupIds.has(form.FormViewer.Id)
                );

                //console.log("Forms where FormViewer is a Group AND User is a Member:", formsWhereUserIsGroupMember);

                // **New Condition: Also include forms where FormViewer Title is "All Employees"**
                const formsForAllEmployees = formsWithGroups.filter(
                    form => form.FormViewer?.Title === "All Employees"
                );

                //console.log("Forms where FormViewer is 'All Employees':", formsForAllEmployees);

                // Merge both filtered lists
                let finalFilteredForms = [...formsWhereUserIsGroupMember, ...formsWithCurrentUser, ...formsForAllEmployees];

                //console.log("Final Filtered Forms:", finalFilteredForms);


                setFormItems(finalFilteredForms);

            }
            else {
                // console.log("Error fetching forms");
            }

        } catch (error) {
            console.error("Error fetching forms:", error);
        } finally {
            setIsLoading(false);
        }
    };
    const staticcardcount = async () => {

        try {

            const currentUser = await sp.web.currentUser.get();
            const response = await sp.web.lists.getByTitle('WorkFlowProcessData')
                .items.filter(`AuthorId eq ${currentUser.Id}and (Status ne 'Draft')`)
                .getAll();
            const inprogresscount = response.filter(item => item.Status === 'In-Progress').length;
            const completedcount = response.filter(item => item.Status === 'Completed').length;
            const totalcount = response.length;
            setStatictiles([
                { icon: TotalRequests, title: "Total Requests", count: totalcount, countcolor: "#B0D136" },
                { icon: PendingRequests, title: "Pending Requests", count: inprogresscount, countcolor: "#EF1B2E" },
                { icon: CompletedRequests, title: "Completed Requests", count: completedcount, countcolor: "#79D48D" }
            ]);
        }
        catch (error) {
            console.error("Error fetching static card count:", error);
        }
    }

    const getFilteredForms = () => {
        // Add loading check to prevent running with incomplete data
        if (isLoading || !formItems) {
            return [];
        }

        //console.log("Full formItems before filtering:", formItems);

        // Filter by visibility first
        const visibleItems = formItems.filter(item => item.VisibilityFlag === true);
        //console.log("Visible items:", visibleItems);

        if (selectedCategory === 'All') {
            return visibleItems;
        }

        // Then filter by category
        const categoryItems = visibleItems.filter(item => item.FormType === selectedCategory);
        //console.log("Category filtered items:", categoryItems);

        return categoryItems;
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

    useEffect(() => {
        fetchFormTitles().catch((error) => {
            console.error("Error in fetchFormTitles:", error);
        });
        fetchFormTypes().catch((error) => {
            console.error("Error in fetchFormTypes:", error);
        });
        staticcardcount().catch((error) => {
            console.error("Error in staticcardcount:", error);
        });
    }, [context, CurrentUserEmail]);

    // Click-outside handler
    useEffect(() => {
        const handleClickOutside = (event: MouseEvent) => {
            if (menuRef.current && !menuRef.current.contains(event.target as Node)) {
                setSelectedFormItem(null); // Close the popup
            }
        };

        if (selectedFormItem) {
            document.addEventListener('mousedown', handleClickOutside);
        }

        // Clean up the event listener on component unmount
        return () => {
            document.removeEventListener('mousedown', handleClickOutside);
        };
    }, [selectedFormItem]);

    const handleUserEmailRetrieved = (UserAUTH: string) => {
        setCurrentUserEmail(UserAUTH);
        //console.log("Logged-in USER IS   ----------------:", UserAUTH);
    };
    const handleUserDomain = (UserDomain: string) => {
        setIsSameDomain(UserDomain)
        console.log("Logged-in USER IS   ----------------:", isSameDomain);

    };

    const okpublish = async () => {
        if (!Publishitem) {
            console.error("Publishitem is undefined.");
            return;
        }
        try {
            // Update the Published status to true
            await sp.web.lists.getByTitle("FormMaster").items.getById(Publishitem.ID).update({
                Published: true
            });
            await fetchFormTitles()
            getFilteredForms()
            setPublishflag(false);
            toast.current?.clear();
            toast.current?.show({
                severity: 'success',
                summary: 'Form Published',
                detail: `Form ${Publishitem.AppName} has been published successfully.`,
                life: 3000,
                style: { height: 100 }
            });
        } catch (error) {
            console.error("Error publishing form:", error);
            toast.current?.clear();
            toast.current?.show({
                severity: 'error',
                summary: 'Publish Failed',
                detail: `Failed to publish form ${Publishitem.AppName}. Please try again.`,
                life: 3000,
            });
        }
    }

    const okunpublish = async () => {
        if (!Publishitem) {
            console.error("Publishitem is undefined.");
            return;
        }
        // Update the Published status to false
        await sp.web.lists.getByTitle("FormMaster").items.getById(Publishitem.ID).update({
            Published: false
        });
        await fetchFormTitles()
        getFilteredForms()
        setUnPublishflag(false);
        toast.current?.clear();
        toast.current?.show({
            severity: 'success',
            summary: 'Form UnPublished',
            detail: `Form ${Publishitem.AppName} has been unpublished successfully.`,
            life: 3000,
            style: { height: 100 }
        });
    }

    const cancelpublish = () => {
        setPublishflag(false);
        setUnPublishflag(false);
        setPublishitem(undefined);
    }

    const handleEditClick = async (item: FormItem) => {
        //console.log("FormItem", item);
        //   item.FormJSON=dataserviceobj.decryptjson(item.FormJSON);
        debugger;
        try {
            const Domainn = await fetchTenantUser();
            // ---- Check for Draft/In-Progress/Returned records ----
            const formMasterItems = await sp.web.lists
                        .getByTitle('FormMaster')
                        .items.select("ID", "AppCode", "AppName", "FormJSON", "IsParentForm","DynamicStatuses")
                        .filter(`ChildOrder eq null and IsParentForm ne 1 and Domain eq '${Domainn.TenantUsers}'and VisibilityFlag ne 0 and isWorkflowRequired ne 'No'`)
                        .getAll();
            const DynamicStatuses = JSON.parse(formMasterItems.find(fm => fm.AppCode === item.AppCode)?.DynamicStatuses || '[]');
            const dynamicStatusFilter =  [
                        "Status eq 'Draft'",
                        "Status eq 'In-Progress'",
                        "Status eq 'Returned'",
                        ...DynamicStatuses.map((ds:any) => `Status eq '${ds}'`)
                    ].join(" or ");
            const WorkflowItems = await sp.web.lists
                .getByTitle("WorkFlowProcessData")
                .items
                .filter(`AppCode eq '${item.AppCode}' and (${dynamicStatusFilter})`)
                .top(1)
                .get();

            if (WorkflowItems.length > 0) {
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
            //console.log("IsParentForm value:", item.IsParentForm);
            // ---- MULTI-FORM EDIT (Parent Form) ----
            if (item.IsParentForm) {
                try {
                    
                    //console.log("fetchedparent",formMasterItems);
                    sessionStorage.setItem("FormMasterData", JSON.stringify(formMasterItems));
                    //console.log("Multi Form: Set isMultiForm = true with", formMasterItems.length, "forms");
                } catch (error) {
                    console.error("Error fetching MultiForm data:", error);
                }
                try {
                    const childFormsData = await sp.web.lists
                        .getByTitle("FormMaster") // Ensure this list contains child forms
                        .items
                        .filter(`ParentAppCode eq '${item.AppCode}'`)
                        .get();

                    //console.log("Child forms fetched:", childFormsData);

                    // Parent form like single-form
                    const parentFormData = {
                        isEdit: true,
                        AppCode: item.AppCode,
                        AppName: item.AppName,
                        FormType: item.FormType,
                        FormJSON: dataserviceobj.encryptjson(item.FormJSON),
                        Description: item.Description
                    };

                    // Map child forms
                    const childForms = childFormsData.map(child => ({
                        name: child.AppName,
                        AppCode: child.AppCode,
                        FormJSON: child.FormJSON
                    }));

                    //console.log("Mapped child forms:", childForms);

                    // Store in sessionStorage
                    sessionStorage.setItem("EditFormData", JSON.stringify(parentFormData));
                    sessionStorage.setItem("isMultiForm", "true");
                    sessionStorage.setItem("isNewForm", "false");


                    if (childForms.length > 0) {
                        sessionStorage.setItem("MultiFormEditData", JSON.stringify({ parentForm: parentFormData, childForms })
                        );
                    } else {
                        console.warn("No child forms found for this parent form. Only parent will be edited.");
                    }

                    navigate('/CreateForm');
                    return; // stop further execution
                } catch (err) {
                    console.error("Error fetching multi-form data:", err);
                }
            }
            else {
                // ---- SINGLE FORM EDIT (unchanged) ----
                const EditFormData = {
                    isEdit: true,
                    AppCode: item.AppCode,
                    AppName: item.AppName,
                    FormType: item.FormType,
                    FormJSON: dataserviceobj.encryptjson(item.FormJSON),
                    Description: item.Description,
                };
                sessionStorage.setItem("isMultiForm", "false");
                sessionStorage.setItem("EditFormData", JSON.stringify(EditFormData));
                //console.log("EditFormData:", EditFormData);

                navigate('/CreateForm');
            }
        } catch (error) {
            console.error("Error handling item click:", error);
        }
    };




    // This function opens a confirmation popup when the delete button is clicked in the three-dot menu.
    // It checks if there are existing records in the WorkFlowProcessData list for the selected form.
    // If no existing records are found, it sets the delete flag and the item to be deleted.
    // If existing records are found, it shows an error message indicating that the form cannot be deleted.
    const openDeletePopup = async (item: any) => {
        const WorkflowItems = await sp.web.lists
            .getByTitle("WorkFlowProcessData")
            .items
            .filter(`AppName eq '${item.AppName}'`)
            .top(1)
            .get();
        //console.log("WorkflowItems in delete", WorkflowItems)
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

    //closes the delete confirmation popup when the user clicks "No" or after a successful deletion.
    const closeDeletePopup = () => {
        setdeleteflag(false)
    }

    // This function handles the deletion of a form when the user confirms the deletion in the popup.
    const handleDeleteClick = async (item: any) => {
        //console.log("deleteFormItem", item);


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

            //console.log("delllllllllllllllll", deleteOperations)
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

    return (
        <div className='d-flex flex-column' style={{ overflow: "hidden", backgroundColor: "#070d19" }}>
            <TopNavBar onUserTypeRetrieved={handleUserEmailRetrieved} onUserDomainRetrieved={handleUserDomain} />
            <Toast ref={toast} />
            <div className='d-flex'>
                <SideBar activeMenu="Home" />
                <div className='d-flex flex-column w-100' style={{ backgroundColor: '#070d19', overflowY: "auto", height: 'calc(-46px + 100vh)' }}>                    {/* Static Tiles Section */}
                    <div className='d-flex justify-content-evenly' style={{ padding: ' 35px 0px 25px 0px' }}>
                        {statictiles.map((tile, index) => {
                            const Icon = tile.icon; // if icon is a component
                            return (
                                <div key={index} className='staticcard'>
                                    <div className='round-icon'>
                                        <img src={Icon} alt={tile.title} className='icon-style' />
                                    </div>
                                    <div className='card-content'>
                                        <span className='card-name'>{tile.title}</span>
                                        <span style={{
                                            fontSize: '17px',
                                            fontWeight: 500,
                                            color: tile.countcolor || 'white',
                                        }}>
                                            {tile.count < 10 ? `0${tile.count}` : tile.count}
                                        </span>
                                    </div>
                                </div>
                            );
                        })}
                    </div>
                    {/* Category Header */}
                    <div className='category'>
                        <h4 style={{
                            color: 'white',
                            fontSize: '16px',
                            fontWeight: '400',
                            marginBottom: '25px'
                        }}>
                            Form Category
                        </h4>

                    </div>
                    {/* Category Tabs */}
                    <div className="categoryItems">
                        <div className='d-flex flex-column' style={{ width: '100%' }}>
                            <div className='scrollable-category'>
                                <ul className='SelectCategory'>
                                    <li
                                        className={selectedCategory === 'All' ? 'onSelectCategory' : 'onOtherCategory'}
                                        onClick={() => setSelectedCategory('All')}
                                    >
                                        All
                                    </li>
                                    {formTypes.map((type) => (
                                        <li
                                            key={type.Title}
                                            className={selectedCategory === type.Title ? 'onSelectCategory' : 'onOtherCategory'}
                                            onClick={() => setSelectedCategory(type.Title)}
                                        >
                                            {type.Title}
                                        </li>
                                    ))}
                                </ul>
                            </div>
                            <div className='tabline'></div>
                        </div>

                    </div>
                    {/* Form List Section */}
                    {isLoading ? (
                        <LoadingSpinner />
                    ) : (
                        <div className="card-container">
                            {getFilteredForms().length > 0 ? (

                                getFilteredForms().map((item) => {
                                    const isMenuOpen = selectedFormItem?.AppCode === item.AppCode; return (

                                        <div className="newcard" key={item.AppCode}>

                                            <div className='d-flex justify-content-between'>
                                                <RouterLink
                                                    to={{
                                                        pathname: '/FormDashBoard',
                                                    }}
                                                    state={{ appcode: item.AppCode, appname: item.AppName }}
                                                    className='RoutLink'
                                                    style={{ marginLeft: 0, textDecoration: 'none' }}
                                                    onClick={() => handleItemClick(item)}
                                                >

                                                    <div className='newheadericon-container'>
                                                        <div className='newicon'>
                                                            <img src={FormIcon} alt={item.AppName} className='icon-style'
                                                            // style={{padding:3}}
                                                            />
                                                        </div>
                                                        <h5 className='newcard-header'>{item.AppName}</h5>
                                                    </div>
                                                </RouterLink>


                                                <div className="newmenu" onClick={(e) => {
                                                    e.stopPropagation();
                                                    e.preventDefault();
                                                    setSelectedFormItem((prev) =>
                                                        prev?.AppCode === item.AppCode ? null : item
                                                    );
                                                }}><img src={threedots}></img> {isMenuOpen && (
                                                    <div className='card-menu-dropdown' ref={menuRef}>
                                                        {/* {(CurrentUserEmail === "BothUser Admin" || CurrentUserEmail === "OnlyAdmin") && <div className='menu-item' onClick={() => handlePublishClick(item)} >
                                                            <img src={Editinthreedots} alt="Edit" className='menu-icon' />
                                                            {item.Published === false ? "Publish" : "UnPublish"}
                                                        </div>} */}
                                                        {(CurrentUserEmail === "BothUser Admin" || CurrentUserEmail === "OnlyAdmin") && <div className='menu-item' onClick={() => handleEditClick(item)}>
                                                            <img src={Editinthreedots} alt="Edit" className='menu-icon' />
                                                            Edit
                                                        </div>}
                                                        {(CurrentUserEmail === "BothUser Admin" || CurrentUserEmail === "OnlyAdmin") && <div className='menu-item' onClick={() => openDeletePopup(item)}>
                                                            <img src={Deleteinthreedots} alt="Delete" className='menu-icon' />
                                                            Delete
                                                        </div>}
                                                        <div className='menu-item' onClick={async () => { await handleItemClick(item); navigate('/ReactDiagram') }}>
                                                            <img src={Workflowinthreedots} alt="Workflow" className='menu-icon' />
                                                            Workflow
                                                        </div>
                                                    </div>)}
                                                </div>

                                            </div>
                                            <RouterLink
                                                to={{
                                                    pathname: '/FormDashBoard',
                                                }}
                                                state={{ appcode: item.AppCode, appname: item.AppName }}
                                                className='RoutLink'
                                                style={{ marginLeft: 0, textDecoration: 'none' }}
                                                onClick={() => handleItemClick(item)}
                                            >
                                                <div
                                                    style={{
                                                        height: '48px',
                                                        display: 'flex',
                                                        alignItems: 'center',
                                                        margin: '0px 1px 8px 1px'
                                                    }}
                                                    title={item.Description || "No description available."}
                                                >
                                                    <p
                                                        className="newdescription"
                                                        style={{
                                                            margin: 0,
                                                            // overflow: 'hidden',
                                                            // whiteSpace: 'nowrap',
                                                            // textOverflow: 'ellipsis'
                                                            display: '-webkit-box',
                                                            overflow: 'hidden',
                                                            WebkitBoxOrient: 'vertical',
                                                            WebkitLineClamp: 2,
                                                        }}
                                                    >
                                                        {item.Description || "No description available."}
                                                    </p>
                                                </div>
                                            </RouterLink>
                                            <div className='ButtonContainer'>
                                                <RouterLink
                                                    to={{
                                                        pathname: `/FormComponent`,
                                                    }}
                                                    state={{ selectedformdetails: item }}
                                                    // className='RoutLink'
                                                    onClick={async (e) => {
                                                        e.preventDefault();
                                                        const domain = await fetchTenantUser();
                                                        const domainName = domain.TenantUsers
                                                        // setSelectedAppCode(item.AppCode);
                                                        if (myDomain === domainName) {
                                                            await getFormDataByTitle(item.AppCode)
                                                            navigate("/FormComponent", {
                                                                state: { selectedformdetails: item },
                                                            });
                                                            //console.log("I AM IN IFFFFFFFFFFFFF")
                                                        }
                                                        else {
                                                            const { recordscount, trnxRecLimitReached } = await fetchWorkflowLimit(domainName, item.AppCode);
                                                            if (trnxRecLimitReached) {
                                                                toast.current?.clear();
                                                                toast.current?.show({
                                                                    severity: 'warn',
                                                                    summary: 'Freemium subscription limit reached',
                                                                    detail: `${item.AppName}: You reached ${recordscount} transactions. Please upgrade your license.`,
                                                                    life: 8000,
                                                                    style: { height: 100 }
                                                                });
                                                                return;
                                                            }
                                                            await getFormDataByTitle(item.AppCode);
                                                            navigate("/FormComponent", {
                                                                state: { selectedformdetails: item },
                                                            });

                                                        }
                                                    }}
                                                >
                                                    <button
                                                        title={item.Published === false ? 'This form is not in use' : ''}
                                                        className='newbtn'
                                                        disabled={item.Published === false}
                                                        style={{ backgroundColor: item.Published === false ? 'gray' : '' }}

                                                    >Raise request</button>
                                                </RouterLink>
                                                {/* <button onClick={updateGlobalState}>click</button> */}
                                                <div style={{ display: 'flex', justifyContent: 'center', width: '35%', gap: '10px' }}>
                                                    {item.AppCode.match(/\/\d+$/) && (
                                                        <div>
                                                            <span className="newversionbadge" title='New Version of Form'>V{item.AppCode.split('/')[1]}</span>
                                                        </div>
                                                    )}
                                                </div>

                                            </div>
                                        </div>
                                    )
                                })
                            ) : (
                                // <div>No forms available in this category.</div>
                                ""
                            )}
                        </div>
                    )}

                    {/* Delete Confirmation Popup */}
                    <div>
                        {deleteflag && <div className="custom-dark-modal-overlay">
                            <div className="custom-dark-modal" style={{ minWidth: 330, maxWidth: 330, }}>
                                <h5 className="addsteptitle custom-text-style" >{`Are you sure you want to delete ${deleteitem?.AppName} form?`}</h5>
                                <div className="modal-buttons" style={{ margin: '15px 20px 25px 20px' }}>

                                    <button
                                        className="newlogocolorbtn"
                                        onClick={() => handleDeleteClick(deleteitem)}
                                    >
                                        Yes
                                    </button>
                                    <button
                                        className="newblackcolorbtn"
                                        onClick={closeDeletePopup}
                                    >
                                        No
                                    </button>
                                </div>
                            </div>
                        </div>}
                    </div>
                    <div>
                        {(Publishflag || UnPublishflag) && <div className="custom-dark-modal-overlay">
                            <div className="custom-dark-modal" style={{ minWidth: 330, maxWidth: 330, }}>
                                <h5 className="addsteptitle custom-text-style" >{`Are you sure you want to ${Publishflag ? 'publish' : 'unpublish'} ${Publishitem?.AppName} form?`}</h5>
                                <div className="modal-buttons" style={{ margin: '15px 20px 25px 20px' }}>

                                    <button
                                        className="newlogocolorbtn"
                                        onClick={() => Publishflag ? okpublish() : okunpublish()}
                                    >
                                        {Publishflag ? 'Publish' : 'Unpublish'}
                                    </button>
                                    <button
                                        className="newblackcolorbtn"
                                        onClick={cancelpublish}
                                    >
                                        Cancel
                                    </button>

                                </div>
                            </div>
                        </div>}
                    </div>
                </div>
            </div >

        </div >
    )
}
export default GlobalDashboard;