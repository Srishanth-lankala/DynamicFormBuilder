import * as React from 'react';
import { useEffect, useState } from 'react';
import { useParams, useNavigate, useLocation, To } from 'react-router-dom';
import { sp } from "@pnp/sp/presets/all";
import { ReactFormGenerator } from 'react-form-builder2';
import '../FormGenerator.css';
import TopNavBar from '../TopBar';
import { PDFDocument, rgb, degrees } from 'pdf-lib';
import { Column } from 'primereact/column';
import { DataTable } from 'primereact/datatable';
import { mySiteUrl } from "../ConfigURL/All_URLs";
import { Toast } from 'primereact/toast';
import { Accordion, AccordionTab } from 'primereact/accordion';
import LoadingSpinner from '../Loading';
import { fileFetchUrl } from '../ConfigURL/All_URLs';
import ChatBox from '../UtilityComponents/ChatBox';
import { IDynamicFormBuilderProps } from '../../IDynamicFormBuilderProps';
import { useGlobalState } from '../GlobalVariable/GlobalStateContext';
import { PeoplePicker, PrincipalType, IPeoplePickerContext } from "@pnp/spfx-controls-react/lib/PeoplePicker";
// import  from "../../assets/images/logo-smartoffice.png";
import { fetchTenantUser } from '../FetchTenantUser/fetchTenantUser';
// import { fetchAllUsers } from '../ADService/ADService';
import PasteArea from '../PasteArea';
// import { myDomain } from './ConfigURL/All_URLs';
import { Dialog } from 'primereact/dialog';
// Ensure the file exists at the correct path or update the path to the correct location
import logoimage from '../../../assets/Images/logosmartoffctrp.png';
import TableComponent, { TableData } from '../TableStructure/TableComponent';
import TimesheetComponent, { TimeSheetData } from '../Timesheetstructure/TimesheetComponent';
import ReactDOM from 'react-dom';
import MyDocument from '../MyDocument';
import { pdf } from '@react-pdf/renderer';
import SideBar from '../Sidebar/SideBar';
import { FlowData, NodeItem } from '../AccordionWorkflow/FlowJSON'
import NoAccess from '../NoAccessScreen';
import ParentWithChildren from '../ParentWithChildren/ParentWithChildren';
import { fetchToken } from '../ADService/ADService';
import { dataservice } from '../encryptionutil';

interface AuthorADdata {
    RequesterName: string;
    Email: string;
    EmployeeID: string;
    Designation: string;
    Department: string;
    ContactNo: string;
}
interface User {
    userPrincipalName: any;
    displayName: string;
    mail: string;
}
export interface CommentData {
    Comments: string;
    CommentedBy: string;
    Created: string; // or `Date` if you're converting it before setting
    Attachments?: { FileUrl: string; FileName: string }[];
}


const ChildForm: React.FC<IDynamicFormBuilderProps> = ({ context }) => {

    type Approver = {
        email: string;
        id: number
    }
    const { id } = useParams<{ id: string }>(); // Retrieve the ID from the route
    const navigate = useNavigate();
    const [datefieldids, setdatefieldids] = useState<any[]>([]);
    const [formData, setFormData] = useState<any[]>([]);
    const [ParentFormAnswerCopy, setParentFormAnswerCopy] = useState<any[]>([]);
    const [formJSON, setFormJSON] = useState<any>(null);
    const [ParentFormJSONCopy, setParentFormJSONCopy] = useState<any>(null);
    const [ChildFormJsons, setChildFormJsons] = useState<any>(null);
    const [RequesterTitle, setRequesterTitle] = useState("");
    const [RequesterEmail, setRequesterEmail] = useState("");
    const [childappselected, setchildappselected] = useState<string | null>(null);
    const [childformData, setChildFormData] = useState<any[]>([]);
    const [AllApprovers, setAllApprovers] = useState<Approver[]>([]);
    const [isButtonDisabled, setIsButtonDisabled] = useState(false);
    // const [FormMasterJson, SetFormMasterJson] = useState<any>(null)
    const [appName, setAppName] = useState<string>(""); // Track AppName separately
    const [appCode, setAppCode] = useState<string>("");
    const [loading, setLoading] = useState<boolean>(true);
    const [FormEditMode, setFormEditMode] = useState<boolean>(true);
    const [buttonView, setbuttonView] = useState(false);
    const [ResubmitButton, setResubmitButton] = useState(false);

    // const [SecChildFormRaised, setSecChildFormRaised] = useState<boolean>(false);

    // const [Cancelrequest, setCancelrequest] = useState(false);
    const [selectTab, setSelectTab] = useState("AllDetails");
    const [AuditLog, setAuditLog] = useState<any[]>([]);
    const [attachmentFiles, setAttachmentFiles] = useState<any[]>([]); // State to hold file names
    const [Categorytype, setCategorytype] = useState<string>("");
    const [allComments, setAllComments] = useState<any[]>([])
    const toast = React.useRef<Toast>(null);
    const [, setLevel] = useState('0')
    const [recentStatus, setRecentStatus] = useState("")
    const [recentTitle, setRecentTitle] = useState("")
    const [comment, setComment] = useState("")
    const [Openpop, setOpenpop] = useState<boolean>(false);
    const [fileFolder, setFileFolder] = useState("")
    const [AppCodevalue, setAppCodevalue] = useState<string>('');
    const [CurrentAppCode, setCurrentAppCode] = useState<string>('');

    const [userName, setUserName] = React.useState("");
    const [userEmail, setUserEmail] = React.useState("");
    const [appSeqNo, setAppSeqNo] = useState("")
    const [, setHighlight] = useState(false);
    const [mappingLevels, setMappingLevels] = useState<{ Level: string; Role: string; User: string }[]>([]);
    const [RequesterName, setRequesterName] = useState("")
    const [, setRequesterNamee] = useState("")
    const [NotifyRequestorEMails, setNotifyRequestorEMails] = useState<string[]>([]);
    const [domainfield, setDomainField] = useState<string>("");
    const [userId, setUserId] = useState<number>();
    const [selectedFile, setSelectedFile] = useState<File[]>([]);
    // --------- starting code for multi file attachments
    const [fileData, setFileData] = useState<Record<string, File[] | null>>({});
    const [mappedfile, setMappedfile] = useState({})
    const [authorADdata, setAuthorADdata] = useState<AuthorADdata | null>(null);
    const [Opencanpop, setOpencanpop] = useState<boolean>(false);
    const [SourceFieldJsons, SetSourceFieldJsons] = useState<any[]>([])
    const [ChildFieldJsons, SetChildFieldJsons] = useState<any[]>([])
    const [SrcFieldAnswer, SetSrcFieldAnswer] = useState<any[]>([])
    const [trnxItem, setTransactionItem] = useState<any>();
    const [hiddendata, SetHiddenData] = useState<any[]>([])
    const [isSameDomain, setIsSameDomain] = useState("")
    // const [selectedUsers, setSelectedUsers] = useState<User[]>([]);
    const [notifytorequestorID, SetnotifytorequestorID] = useState<number[]>([]);
    // const [myOrganizationUsers, setMyOrganizationUsers] = useState<User[]>([]);
    // const [searchText, setSearchText] = useState('');
    // const [filteredUsers, setFilteredUsers] = useState<User[]>([]);
    // const [showDropdown, setShowDropdown] = useState(false);
    const [preservedfilters, setPreservedFilters] = useState<any>({})
    // Comment box popup for image 
    const [selectedComment, setSelectedComment] = useState<CommentData | null>(null);
    const [isDialogVisible, setIsDialogVisible] = useState(false);
    const { setGlobalVariable } = useGlobalState();
    const [tables, setTables] = React.useState<TableData[]>([]);
    const [timesheettables, setTimesheettables] = useState<TimeSheetData[]>([]);
    const [hasBreaches, setHasBreaches] = useState<boolean>(false); // State to store breaches
    const [pdfanswerjson, setPdfAnswerJson] = useState<any>({});
    const [isreopen, setisreopen] = useState<boolean>(false)
    const [reopenbuttonview, setreopenbuttonview] = useState<boolean>(false)
    const [FlowJSON, setFlowJSON] = useState<NodeItem[]>([])
    // const [userRole, setUserRole] = useState<string>("")
    const [record, setrecord] = useState<any>()
    const [IsParentForm, setIsParentForm] = useState<boolean>(false);
    const [ChildFormDetails, SetChildFormDetails] = useState<any[]>([])
    const [filteredChildFormDetails, setFilteredChildFormDetails] = useState<any[]>([])
    const [ChildtransactionItems, setChildtransactionItems] = useState<any[]>([])
    // const [isAccessDenied, setisAccessDenied] = useState<boolean>(true)
    const [isAdmin, setIsAdmin] = useState(false);
    const location = useLocation();

    const [isAccordionVisible, setIsAccordionVisible] = useState<boolean>(false);
    const [isChildReturned, setIsChildReturned] = useState<boolean>(false);

    const [selectedChild, setSelectedChild] = useState<any>(null);

    // this multiFormSeqNo state for now using for filter the chatbox i.e. communication channel
    const [multiFormSeqNo, setMultiFormSeqNo] = useState("");

    const [selectedFormJSON, setSelectedFormJSON] = useState<any[]>([]);


    // const [EndFlow, setEndFlow] = useState<boolean>(false);

    //   const [RequesterNameee, setRequesterNameee] = useState("")
    //   const [, setRequesterTitle] = useState("");
    //   const [, setRequesterEmail] = useState("");
    const [childselected, setChildselected] = useState<string>("");
    const dataserviceobj = new dataservice();
    // const AnswerJsonSeparator = (ParentCopy: any[], Answers: any[]) => {
    //     let ParentAnswers: any[] = []
    //     let ChildAnswers: any[] = []
    //     Answers.forEach((ele) => {
    //         ParentCopy.find(elem => elem.id === ele.id) ? ParentAnswers.push(ele) : ChildAnswers.push(ele)
    //     })

    //     return [ParentAnswers, ChildAnswers]
    // }

    const generateRandomString = (length: number) => {
        const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
        return Array.from({ length }, () => characters.charAt(Math.floor(Math.random() * characters.length))).join('');
    };
    const getCurrentUserSeqNo = async () => {
        const currentUser = await sp.web.currentUser.get();
        const displayNamePrefix = currentUser.Title.substring(0, 6).toUpperCase();
        const randomString = generateRandomString(6);
        return `${displayNamePrefix}-${randomString}`;
    };

    const handleTableUpdate = (updatedTable: TableData) => {
        setTables(prevTables => {
            const updatedTables = prevTables.map(table =>
                table.id === updatedTable.id ? updatedTable : table
            );
            console.log('All tables updated:', updatedTables);
            return updatedTables;
        });
        console.log("tables", tables);
    };
    const handleRestrictionBreachUpdate = (hasBreaches: boolean) => {
        setHasBreaches(hasBreaches);
        console.log("Has Restriction Breaches for this table:", hasBreaches);
    };

    sp.setup({
        sp: {
            baseUrl: mySiteUrl,
        },
    });

    const fetchingTenantUsers = async () => {
        try {
            const domain = await fetchTenantUser();
            setDomainField(domain.TenantUsers)
            setUserId(domain.CurrentUserId)
            // const allUsers = await fetchAllUsers();
            // if (allUsers) {  // Ensure data is available before updating state
            //     console.log("Fetched Users:------------------------------------ ", allUsers);
            //     // setMyOrganizationUsers(allUsers)
            //     // setFilteredUsers(allUsers.slice(0, 10));
            // } else {
            //     console.warn("No users found or API call failed.");
            // }
        } catch (error) {
            console.error("Error fetching users:", error);
        }
    }

    useEffect(() => {
        setPreservedFilters(location.state || {})
        void fetchingTenantUsers()
        console.log("locationdetailsm in edit >>>>>>>>>>>>>>>>>>>>>>>>>>>", location)
    }, [])

    useEffect(() => {
        if (Array.isArray(formJSON)) {
            const pdfjson: any = {};
            formJSON.forEach((element: any) => {
                const match = formData.find((e: any) => e.id === element.id);
                if (!match) return;
                const elementType = element.element;
                console.log("elemenetType**********", elementType);
                let readableValue = "";
                if (Array.isArray(match.value)) {
                    readableValue = match.value.map((val: any) => {
                        const option = element.options?.find((opt: any) => opt.key === val);
                        return option ? option.text : val;
                    }).join(",");
                }
                else if (elementType === "Dropdown") {
                    const option = element.options?.find((opt: any) => opt.value === match.value);
                    readableValue = option ? option.text : match.value;
                }
                else if (elementType === "FileUpload") {
                    return;
                } else {
                    readableValue = match.value;
                }
                pdfjson[stripHtmlTags(element.label)] = readableValue;
                console.log("readableValue", readableValue)
                console.log("pdfjson", JSON.stringify(pdfjson, null, 2))

            })
            setPdfAnswerJson(pdfjson);
        }

    }, [formData, formJSON])
    const convertImageToBase64 = async (imageUrl: any) => {
        try {
            // For browser environments
            const response = await fetch(imageUrl);
            const blob = await response.blob();
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onloadend = () => resolve(reader.result);
                reader.onerror = reject;
                reader.readAsDataURL(blob);
            });
        } catch (error) {
            console.error('Error converting image to base64:', error);
            return null;
        }
    };
    const getBase64Logo = async () => {
        const imageUrl = logoimage;
        const base64 = await convertImageToBase64(imageUrl);
        console.log('Base64 image string:', base64);
        return base64;
    };
    const getloggedInUser = async () => {

        const curUser = await sp.web.currentUser();
        const domain = await fetchTenantUser();
        setDomainField(domain.TenantUsers)
        const RequesterTitle = curUser.Title.split("|")[0];
        setRequesterTitle(RequesterTitle);
        setRequesterEmail(curUser.Email)
        // setRequesterNameee("Requester: " + RequesterTitle);
    }
    useEffect(() => {
        void getloggedInUser();
        // void fetchingTenantUsers()
        // void PrefixFetch()
    }, []);



    const handleWatermarkedDownload = async () => {
        // Step 1: Create PDF from react-pdf component
        // const base64Logo = await getBase64Logo();
        const blob = await pdf(<MyDocument data={pdfanswerjson} loghistory={AuditLog} tables={tables} logo={getBase64Logo} txitem={trnxItem} />).toBlob();

        // Step 2: Load into pdf-lib
        const arrayBuffer = await blob.arrayBuffer();
        const pdfDoc = await PDFDocument.load(arrayBuffer);
        const pages = pdfDoc.getPages();

        // Step 3: Add watermark
        for (const page of pages) {
            const { width, height } = page.getSize();
            page.drawText('SMARTOFFICENXT', {
                x: (width / 2) - 220,
                y: (height / 2) - 240,
                size: 70,
                rotate: degrees(45),
                opacity: 0.3,
                color: rgb(0.75, 0.75, 0.75),
            });
        }

        // Step 4: Download the final watermarked PDF
        const finalBytes: any = await pdfDoc.save();
        const finalBlob = new Blob([finalBytes], { type: 'application/pdf' });
        const url = URL.createObjectURL(finalBlob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `SmartofficeNxt/${appName}/${trnxItem.Title}.pdf`;
        a.click();
        URL.revokeObjectURL(url);
    };
    // useEffect(() => {

    //     setFilteredUsers(myOrganizationUsers.slice(0, 10));
    // }, [myOrganizationUsers]);


    useEffect(() => {
        if (!formJSON || !Array.isArray(formJSON)) return;

        type uploadelement = {
            field_name: string;
            id: string;
            element: string;
            text: string;
            group_name: string;
            required: boolean;
            VisibilityCondition: any[];
            canHavePageBreakBefore: boolean;
            canHaveAlternateForm: boolean;
            canHaveDisplayHorizontal: boolean;
            canHaveOptionCorrect: boolean;
            canHaveOptionValue: boolean;
            canPopulateFromApi: boolean;
            label: string;
            dirty: boolean;
        };

        const fileElements = formJSON.filter(
            (item: uploadelement) => item.element === "FileUpload"
        );

        if (!fileElements.length) return;
        console.log("fileeleeemenen", fileElements)
        // const fieldName = fileElements[0].field_name;

        const observeDOM = () => {
            const observer = new MutationObserver((mutations, obs) => {


                fileElements.forEach((fileuploadelement) => {
                    const fieldName = fileuploadelement.field_name
                    const fileInput = document.querySelector(`[name="${fieldName}"]`) as HTMLInputElement;
                    if (fileInput) {
                        // Stop observing once found
                        // obs.disconnect();

                        // Set multiple attribute
                        fileInput.setAttribute("multiple", "");

                        // Add event listener
                        const handlefilChange = (e: Event) => {
                            setTimeout(() => {
                                const target = e.target as HTMLInputElement;
                                const files = target.files ? Array.from(target.files) : [];
                                const fileListContainer = fileInput.closest(".image-upload-container")?.querySelector(".file-upload-preview");
                                const clearBtn = fileListContainer?.closest(".image-upload-container")?.querySelector(".btn-file-upload-clear");
                                console.log("clearbtn", clearBtn)
                                console.log("fromobserve", files);
                                setFileData((prev) => ({
                                    ...prev,
                                    [fieldName]: files,
                                }));
                                console.log("newfiledata", fieldName)
                                if (fileListContainer) {
                                    fileListContainer.innerHTML = "";
                                    files.forEach((file) => {
                                        const fileItem = document.createElement("div");
                                        fileItem.textContent = `Name: ${file.name} — Size: ${Math.ceil(file.size / 1024)} KB`;
                                        fileListContainer.appendChild(fileItem);
                                    });
                                }
                                const handleClearClick = () => {
                                    console.log(`Cleared files for ${fieldName}`);

                                    // Clear from state
                                    setFileData((prev) => ({
                                        ...prev,
                                        [fieldName]: null,
                                    }));
                                    fileInput.value = "" //after clearing make input element value null or it wont accept same file again
                                };
                                clearBtn?.addEventListener("click", handleClearClick);
                                //    return clearBtn?.removeEventListener("click", handleClearClick);
                            }, 100)
                        };


                        // const clearBtn = fileInput.closest(".image-upload-container")?.querySelector(".btn-file-upload-clear");


                        fileInput.addEventListener("change", handlefilChange);




                        // Cleanup
                        return () => {
                            fileInput.removeEventListener("change", handlefilChange);

                        };
                    }
                })
            });

            observer.observe(document.body, {
                childList: true,
                subtree: true,
            });

            return () => observer.disconnect();
        };

        const cleanup = observeDOM();
        return cleanup;

    }, [formJSON]);



    useEffect(() => {
        console.log("filedatachanged", fileData)
        setFormData((prevData: any) =>
            prevData.map((el: any) =>
                el.name.startsWith("file_upload") && fileData[el.name]
                    ? { ...el, value: fileData[el.name] }
                    : el
            )
        );
        setParentFormAnswerCopy((prevData: any) =>
            prevData.map((el: any) =>
                el.name.startsWith("file_upload") && fileData[el.name]
                    ? { ...el, value: fileData[el.name] }
                    : el
            )
        );
    }, [fileData]); // This ensures that file uploads are correctly reflected in `formData`

    const ConditionValidator = (leftvalue: any, operator: string, rightvalue: any, elementType: string) => {   //leftvalue is user entered and right value is set by admin to check condition
        switch (operator) {
            case "EqualsTo":
                if (elementType === "Dropdown" || elementType === "TextInput" || elementType === "RadioButtons") {
                    return leftvalue.toLowerCase() === rightvalue.toLowerCase()
                }
                if (elementType === "NumberInput") {
                    return parseFloat(leftvalue) === parseFloat(rightvalue) //assuming leftvalue is number/string based on its type and right value is always string
                }
                break;
            case "Not EqualsTo":
                if (elementType === "Dropdown" || elementType === "TextInput" || elementType === "RadioButtons") {
                    return leftvalue.toLowerCase() !== rightvalue.toLowerCase()
                }
                if (elementType === "NumberInput") {
                    return parseFloat(leftvalue) !== parseFloat(rightvalue) //assuming leftvalue is number/string based on its type and right value is always string
                }
                break;
            case "Contains":
                return leftvalue.toLowerCase().includes(rightvalue.toLowerCase())
            case "StartsWith":
                return leftvalue.toLowerCase().startsWith(rightvalue.toLowerCase());
            case "GreaterThan":
                return parseFloat(leftvalue) > parseFloat(rightvalue)
            case "LessThan":
                return parseFloat(leftvalue) < parseFloat(rightvalue)
            case "LessThan or EqualsTo":
                return parseFloat(leftvalue) <= parseFloat(rightvalue)
            case "GreaterThan or EqualsTo":
                return parseFloat(leftvalue) >= parseFloat(rightvalue)
            default:
                console.log("operation failed ");
                break;
        }
    }

    function stripHtmlTags(label: any) { return label?.replace(/<\/?[^>]+(>|$)/g, "")?.trim() || ""; }

    const fetchUserInfo = async () => {
        //debuggergger
        try {
            const currentUser = await sp.web.currentUser();

            let x = currentUser.Title.split("|")[0]
            console.log("currentUser---------", currentUser)
            setUserName(x);
            setUserEmail(currentUser.Email)
        } catch (error) {
            console.error('Error fetching user data', error);
        }
    };
    function getPeoplePickerContext(): IPeoplePickerContext {
        return {
            absoluteUrl: context.pageContext.web.absoluteUrl,
            msGraphClientFactory: context.msGraphClientFactory,
            spHttpClient: context.spHttpClient,
        };
    }
    const onChange = async (items: any[]) => {
        console.log("Selected People:", items);

        try {
            const users = await Promise.all(items.map(async (useritem) => {
                const user = await sp.web.siteUsers.filter(`Email eq '${useritem.secondaryText}'`).select('Id,Email').get()
                return user[0]?.Id
            }))

            console.log("AllUserIds", users);

            const validUserIds = users.filter(id => id !== undefined);
            console.log("Selected valid User IDs:", validUserIds);
            SetnotifytorequestorID(validUserIds)

        }
        catch (err) {
            console.log(err);
        }
    };

    function convertDateFormat(date: string) {
        let [year, month, day] = date.split('-')
        return `${month}/${day}/${year}`
    }


    const validateFileSize = (file: File) => {
        const maxSizeInMB = 3;
        const fileSizeInMB = file.size / (1024 * 1024);
        return fileSizeInMB <= maxSizeInMB;
    };

    const onFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
        if (event.target.files) {
            const newFiles = Array.from(event.target.files);
            const validFiles: File[] = [];

            console.log("newFilesss-------------------", newFiles);
            console.log("selectedFile state before checking:", selectedFile);

            setSelectedFile((prevFiles) => {
                const existingFileNames = new Set(prevFiles.map(file => file.name));

                newFiles.forEach((file) => {
                    if (existingFileNames.has(file.name)) {

                        toast.current?.show({
                            severity: 'error',
                            summary: 'Duplicate File',
                            detail: 'This file already exists in the queue',
                            life: 2000
                        });
                    } else if (validateFileSize(file)) {
                        validFiles.push(file);
                    } else {
                        toast.current?.show({
                            severity: 'error',
                            summary: 'File Size Error',
                            detail: `The file "${file.name}" exceeds the allowed size (3 MB).`,
                            life: 3000,
                        });
                    }
                });

                console.log("validFiles--------------", validFiles);

                // Ensure total files do not exceed 3
                if (prevFiles.length + validFiles.length > 3) {
                    toast.current?.show({
                        severity: 'error',
                        summary: 'Upload Limit Exceeded',
                        detail: 'Total uploaded files cannot exceed 3.',
                        life: 3000,
                    });
                    return prevFiles;
                }
                return [...prevFiles, ...validFiles];
            });
        }
    };

    const handleRemoveFile = (index: number) => {
        const updatedFiles = selectedFile.filter((_, i) => i !== index);
        setSelectedFile(updatedFiles);
    };
    const funforVisibility = (data: any) => {
        // console.log("Data in funforvisibility", data)
        // console.log("sourcefieldjsons====", SourceFieldJsons)
        // console.log("srcfieldans", SrcFieldAnswer)
        let filtereddata = data.filter((item: any) => {
            return SourceFieldJsons.find(ele => ele.id === item.id)
        })
        // console.log("filtereddata", filtereddata)
        // console.log("ChildFieldJsons", ChildFieldJsons)
        let idboolobjArray = compareValues(filtereddata, SrcFieldAnswer)
        console.log("idboolobjArray", idboolobjArray)
        let updatedFormData = [...data];
        idboolobjArray.forEach((element: any) => {
            if (!element.isSame) {
                let field = SourceFieldJsons.find(ele => ele.id === element.id)
                console.log("field", field)
                ChildFieldJsons.forEach((childjson) => {
                    console.log("689745689=====", childjson?.VisibilityCondition[0]?.SourceField)
                    console.log("stripHtmlTags(field.label)", stripHtmlTags(field.label))
                    if (childjson?.VisibilityCondition[0]?.SourceField === stripHtmlTags(field.label)) {
                        let leftvalue
                        if (childjson?.VisibilityCondition[0]?.element === "RadioButtons") {
                            console.log("element.options====", field.options)
                            let optionjson = field.options.find((lfv: any) => {
                                return lfv.key === element.value[0]
                            });
                            console.log("lfvv====", optionjson)
                            leftvalue = optionjson.text
                        }
                        else if (childjson?.VisibilityCondition[0]?.element === "Dropdown") {
                            let optionjson = field.options.find((lfv: any) => {
                                return lfv.value === element.value
                            });
                            leftvalue = optionjson.text
                        }
                        else {
                            leftvalue = element.value
                        }
                        console.log("parameters condn validator", leftvalue, childjson?.VisibilityCondition[0]?.Operator, childjson?.VisibilityCondition[0]?.TargetValue, field.element)

                        let newbool = ConditionValidator(
                            leftvalue,
                            childjson?.VisibilityCondition[0]?.Operator,
                            childjson?.VisibilityCondition[0]?.TargetValue,
                            field.element)
                        console.log("newbool", newbool)
                        // if(newbool){
                        let visiblechangeelement
                        if (childjson?.parentId) {
                            if (childjson.element === "Checkboxes") {
                                const idPart = childjson?.options?.[0]?.key;

                                const matchedElement = Array.from(document.querySelectorAll('input[id]')).find(
                                    el => el.id.includes(idPart)
                                );

                                visiblechangeelement = matchedElement?.closest('.SortableItem.rfb-item')?.parentElement as HTMLElement;

                            }
                            else {
                                visiblechangeelement = document.querySelector(`[name="${childjson.field_name}"]`)?.closest('.SortableItem.rfb-item')?.parentElement as HTMLElement
                            }
                        } else {
                            if (childjson.element === "Checkboxes") {
                                const idPart = childjson?.options?.[0]?.key;

                                const matchedElement = Array.from(document.querySelectorAll('input[id]')).find(
                                    el => el.id.includes(idPart)
                                );

                                visiblechangeelement = matchedElement?.closest('.SortableItem.rfb-item') as HTMLElement;

                            } else {
                                visiblechangeelement = document.querySelector(`[name="${childjson.field_name}"]`)?.closest('.SortableItem.rfb-item') as HTMLElement
                            }
                        }
                        console.log("visiblechangeelement=====", visiblechangeelement)
                        if ((childjson?.controlVisibility && newbool) || (!childjson?.controlVisibility && !newbool)) {
                            (visiblechangeelement as HTMLElement).style.display = "block"
                            SetHiddenData(prev => prev.filter(ele => ele.id !== childjson?.id))
                        }
                        else if ((!childjson?.controlVisibility && newbool) || (childjson?.controlVisibility && !newbool)) {
                            (visiblechangeelement as HTMLElement).style.display = "none"
                            SetHiddenData(prev => [...prev, childjson])

                            // Get the field element
                            const fieldElement = document.querySelector(`[name="${childjson.field_name}"]`);
                            if (childjson.element === "Checkboxes") {
                                console.log("if entered")
                                childjson.options.forEach((checkbox: any) => {
                                    const idPart = checkbox?.key;
                                    console.log("foreach entered")
                                    const matchedElement = Array.from(document.querySelectorAll('input[type="checkbox"][id]')).find(
                                        el => el.id.includes(idPart)
                                    ) as HTMLInputElement | undefined;

                                    if (matchedElement) {
                                        console.log("false block entered")
                                        matchedElement.checked = false;
                                    }
                                });
                            }
                            if (fieldElement) {
                                // Handle different field types appropriately
                                if (childjson.element === "RadioButtons") {
                                    // For radio buttons, unselect all options
                                    document.querySelectorAll(`input[name="${childjson.field_name}"]:checked`).forEach((radio: any) => {
                                        radio.checked = false;
                                    });
                                } else if (childjson.element === "Dropdown") {
                                    // For dropdowns, reset to first option
                                    const selectElement = fieldElement as HTMLSelectElement;
                                    selectElement.selectedIndex = 0;
                                } else if (childjson.element === "FileUpload") {
                                    // For file inputs, reset the value
                                    const inputElement = fieldElement as HTMLInputElement;
                                    inputElement.value = "";
                                } else {
                                    // For text inputs, textareas, etc.
                                    const inputElement = fieldElement as HTMLInputElement;
                                    inputElement.value = "";
                                }
                            }

                            updatedFormData = updatedFormData.map((item: any) => {
                                if (item.id === childjson?.id) {
                                    if (childjson.element === "Checkboxes" || childjson.element === "RadioButtons") {
                                        return { ...item, value: [] };
                                    }
                                    else if (childjson.element === "FileUpload") {
                                        return { ...item, value: null };
                                        // Also clear from fileData state if needed
                                        // setFileData(prev => ({ ...prev, [childjson.field_name]: null }));
                                    }
                                    else {
                                        return { ...item, value: "" };
                                    }
                                }
                                filtereddata = filtereddata.map((item: any) => {
                                    if (item.id === childjson?.id) {
                                        return { ...item, value: item.value };
                                    } else {
                                        return item;
                                    }
                                })
                                SetSrcFieldAnswer(filtereddata)
                                return item;
                            })


                        }
                    }
                })
            }
        });

        SetSrcFieldAnswer(filtereddata)

    }
    useEffect(() => {

        if (!selectedFormJSON || !Array.isArray(selectedFormJSON)) return;
        type uploadelement = {
            field_name: string;
            id: string;
            element: string;
            text: string;
            group_name: string;
            required: boolean;
            VisibilityCondition: any[];
            canHavePageBreakBefore: boolean;
            canHaveAlternateForm: boolean;
            canHaveDisplayHorizontal: boolean;
            canHaveOptionCorrect: boolean;
            canHaveOptionValue: boolean;
            canPopulateFromApi: boolean;
            label: string;
            dirty: boolean;
        };

        const fileElements = selectedFormJSON.filter(
            (item: uploadelement) => item.element === "FileUpload"
        );

        if (!fileElements.length) return;
        console.log("fileeleeemenen", fileElements)
        // const fieldName = fileElements[0].field_name;

        const observeDOM = () => {
            const observer = new MutationObserver((mutations, obs) => {
                fileElements.forEach((fileuploadelement) => {
                    const fieldName = fileuploadelement.field_name
                    const fileInput = document.querySelector(`[name="${fieldName}"]`) as HTMLInputElement
                    if (fileInput) {
                        // Stop observing once found
                        obs.disconnect();

                        // Set multiple attribute
                        fileInput.setAttribute("multiple", "");

                        // Add event listener
                        const handlefilChange = (e: Event) => {
                            setTimeout(() => {
                                const target = e.target as HTMLInputElement;
                                const files = target.files ? Array.from(target.files) : [];
                                const fileListContainer = fileInput.closest(".image-upload-container")?.querySelector(".file-upload-preview");
                                const clearBtn = fileListContainer?.closest(".image-upload-container")?.querySelector(".btn-file-upload-clear");
                                console.log("clearbtn", clearBtn)
                                console.log("fromobserve", files);
                                setFileData((prev) => ({
                                    ...prev,
                                    [fieldName]: files,
                                }));
                                console.log("newfiledata", fieldName)
                                if (fileListContainer) {
                                    fileListContainer.innerHTML = "";
                                    files.forEach((file) => {
                                        const fileItem = document.createElement("div");
                                        fileItem.textContent = `Name: ${file.name} — Size: ${Math.ceil(file.size / 1024)} KB`;
                                        fileListContainer.appendChild(fileItem);
                                    });
                                }
                                const handleClearClick = () => {
                                    console.log(`Cleared files for ${fieldName}`);

                                    // Clear files from state
                                    setFileData((prev) => ({
                                        ...prev,
                                        [fieldName]: null,
                                    }));
                                    fileInput.value = "" //after clearing make input element value null or it wont accept same file again
                                };
                                clearBtn?.addEventListener("click", handleClearClick);
                            }, 100)
                        };

                        fileInput.addEventListener("change", handlefilChange);

                        // Cleanup
                        return () => {
                            fileInput.removeEventListener("change", handlefilChange);
                        };
                    }
                })
            });

            observer.observe(document.body, {
                childList: true,
                subtree: true,
            });

            return () => observer.disconnect();
        };

        const cleanup = observeDOM();
        return cleanup;

    }, [selectedFormJSON]);
    const initialfunforvisibility = async () => {
        const transactionItem = await sp.web.lists
            .getByTitle("WorkFlowProcessData")
            .items.getById(Number(id))
            .select("*", "Title, Created, Author/Title,CategoryType,NotifyRequestor/EMail,Managers/Title")
            .expand("Author", "NotifyRequestor", "Managers")
            .get();
        setTransactionItem(transactionItem)
        const RAWformmasteritem = await sp.web.lists.getByTitle("FormMaster").items.filter(`AppCode eq '${transactionItem.AppCode}'`).top(1).get();
        const formmasteritem = RAWformmasteritem?.map((item:any)=>{
            return { ...item, FormJSON: dataserviceobj.decryptjson(item.FormJSON)}
        })
        const selectedFormJSON = JSON.parse(formmasteritem[0].FormJSON)
        const dependentFields = selectedFormJSON.filter((ele: any) => { //All child fields having conditions
            return ele?.dependentVisibility
        })
        SetChildFieldJsons(dependentFields)
        const dependentShowFields = selectedFormJSON.filter((ele: any) => { //Child fields which should be hidden initially
            return ele?.dependentVisibility && ele?.controlVisibility
        })
        console.log("dependentFields", dependentFields)
        console.log("dependentShowFields", dependentShowFields)
        const SourceFieldNames = [...new Set(dependentFields.map((ele: any) => ele?.VisibilityCondition[0]?.SourceField))]
        console.log("SourceFieldNames", SourceFieldNames)

        const SourceFields = selectedFormJSON.filter((ele: any) => {
            return ele?.label && SourceFieldNames.includes(stripHtmlTags(ele?.label))
        })
        console.log("SourceFields", SourceFields)

        SetSourceFieldJsons(SourceFields)

        const transformedArray = SourceFields.map((item: any) => ({
            id: item.id,
            name: item.field_name,
            custom_name: item.field_name,
            value: item.element === "RadioButtons" ? [] : ""
        }));
        console.log("transformedarray", transformedArray)
        SetSrcFieldAnswer(transformedArray)
        funforVisibility(JSON.parse(transactionItem.FormData))
    }

    const checkDomReady = () => {
        const formContainer = document.querySelector('.react-form-builder-form'); // Adjust selector to match your form container
        if (formContainer) {
            funforVisibility(JSON.parse(trnxItem.FormData));
            //   hasRunRef.current = true; // Mark as run
        } else {
            requestAnimationFrame(checkDomReady); // Keep checking until DOM is ready
        }
    };
    useEffect(() => {
        requestAnimationFrame(checkDomReady);
    }, [trnxItem]);
    useEffect(() => {
        initialfunforvisibility().then().catch(ee => console.log("loaddataerrrorinitialllll", ee))
    }, [formJSON, id]);
    function compareValues(array1: any, array2: any) {
        return array1.map((item1: any) => {
            const item2 = array2.find((item: any) => item.id === item1.id);
            return {
                id: item1.id,
                value: item1.value,
                isSame: item2 ? JSON.stringify(item1.value) === JSON.stringify(item2.value) : false
            };
        });
    }
    const isAdminUser = async () => {
        const currentUser = await sp.web.currentUser.get();
        const authUsers = await sp.web.lists.getByTitle('AuthList').items
            .filter(`AdminUser eq 'Yes' and AuthName/Id eq ${currentUser.Id}`)
            .select('AuthName/Id')
            .expand('AuthName')
            .getAll();
        setIsAdmin(authUsers.length > 0);
    };

    // Function to check if user is part of a child form's workflow
    const isUserInWorkflow = async (childAppCode: string, userId: number): Promise<boolean> => {
        try {
            // Get all workflow mappings for this child form
            const mappings = await sp.web.lists.getByTitle("MappingMaster")
                .items.select("Users_x002f_Groups/Id")
                .expand("Users_x002f_Groups")
                .filter(`AppCode eq '${childAppCode}'`)
                .get();

            for (const mapping of mappings) {
                const groupOrUserId = mapping.Users_x002f_Groups?.Id;
                if (!groupOrUserId) continue;

                // Check if it's a direct user match
                if (groupOrUserId === userId) {
                    return true;
                }

                // Check if it's a group and user is member
                try {
                    const isGroup = await sp.web.siteGroups.getById(groupOrUserId).get().then(() => true).catch(() => false);
                    if (isGroup) {
                        const groupUsers = await sp.web.siteGroups.getById(groupOrUserId).users.get();
                        if (groupUsers.some((user: any) => user.Id === userId)) {
                            return true;
                        }
                    }
                } catch (error) {
                    // Not a group, skip
                    continue;
                }
            }
            return false;
        } catch (error) {
            console.error("Error checking user in workflow:", error);
            return false;
        }
    };

    // Function to check if all previous child form workflows are completed
    const arePreviousWorkflowsCompleted = (childFormDetails: any[], childTransactions: any[], targetChildOrder: number): boolean => {
        // Get all child forms with lower order than target
        const previousChildren = childFormDetails.filter(child => child.ChildOrder < targetChildOrder);

        console.log(`🔍 Checking if previous workflows completed for ChildOrder ${targetChildOrder}`);
        console.log(`Previous children to check:`, previousChildren.map(c => `${c.AppName} (Order: ${c.ChildOrder})`));

        // Check if all previous child forms have completed workflows
        for (const prevChild of previousChildren) {
            // Try both CurrentAppCode and AppCode for matching
            const transaction = childTransactions.find(t =>
                t.CurrentAppCode === prevChild.AppCode || t.AppCode === prevChild.AppCode
            );

            if (!transaction) {
                // No transaction exists for this child form - workflow not started/completed
                console.log(`❌ No transaction found for ${prevChild.AppName} - previous workflow not complete`);
                return false;
            }

            // Check if workflow is completed (Status should be final approval status)
            const status = (transaction.Status || transaction.RecentStatus || '').toLowerCase();
            const isCompleted = status === "approved" || status === "completed" || status === "closed";

            console.log(`📋 ${prevChild.AppName} status: "${status}" - Completed: ${isCompleted}`);

            if (!isCompleted) {
                console.log(`❌ ${prevChild.AppName} workflow not completed yet`);
                return false;
            }
        }

        console.log(`✅ All previous workflows are completed`);
        return true;
    };

    // Filter visible child forms based on user workflow participation and workflow completion
    const filterVisibleChildren = async (
        childFormDetails: any[],
        childTransactions: any[],
        isAdminUser: boolean,
        parentTransaction: any
    ): Promise<any[]> => {
        debugger;
        console.log("=== FILTERING CHILD FORMS WITH USER WORKFLOW CHECK ===");
        console.log("Total child forms:", childFormDetails.length);
        console.log("Total child transactions:", childTransactions.length);
        console.log("Is Admin User:", isAdmin);

        // If admin, show all child forms
        const initial: string[] = [];
        ChildtransactionItems.forEach((item) => {
            initial.push(item.AppCode);
        });
        if (!initial.includes(record.CurrentAppCode)) { initial.push(record.CurrentAppCode) }
        const adminapps = ChildFormDetails.filter((child) => initial.includes(child.AppCode));
        if (isAdminUser) {
            console.log("✅ Admin user - showing all child forms");
            return adminapps;
        }

        const currentUser = await sp.web.currentUser.get();
        const userId = currentUser.Id;

        const visibleChildren: any[] = [];

        // Sort child forms by ChildOrder
        const sortedChildren = [...childFormDetails].sort((a, b) => a.ChildOrder - b.ChildOrder);

        for (const child of sortedChildren) {
            console.log(`\n🔎 Checking child form: ${child.AppName} (Order: ${child.ChildOrder}, AppCode: ${child.AppCode})`);

            // Find corresponding transaction for this child (if any)
            const correspondingChildTx = childTransactions.find(t => t.AppCode === child.AppCode || t.CurrentAppCode === child.AppCode);

            if (!correspondingChildTx) {
                console.log(`❌ HIDDEN: No transaction exists yet for ${child.AppName}`);
                continue;
            }

            // Check transaction-level participation (AuthorId, CurApproverId, AllApprovers emails)
            const txAuthorId = correspondingChildTx.AuthorId || correspondingChildTx.Author?.Id;
            const txCurApproverId = correspondingChildTx.CurApproverId;
            let txAllApproversEmails: string[] = [];
            try {
                if (correspondingChildTx.AllApprovers) {
                    if (typeof correspondingChildTx.AllApprovers === 'string') {
                        txAllApproversEmails = JSON.parse(correspondingChildTx.AllApprovers);
                    } else if (Array.isArray(correspondingChildTx.AllApprovers)) {
                        txAllApproversEmails = correspondingChildTx.AllApprovers;
                    }
                }
            } catch (err) {
                console.log('Error parsing AllApprovers for transaction', err);
            }

            const currentUserObj = await sp.web.currentUser.get();
            const currentUserEmail = currentUserObj.Email;

            const isTransactionParticipant = (txAuthorId && Number(txAuthorId) === Number(userId)) || (txCurApproverId && Number(txCurApproverId) === Number(userId)) || (txAllApproversEmails && txAllApproversEmails.includes && txAllApproversEmails.includes(currentUserEmail));

            // Fallback to mapping membership if needed
            const isUserInvolved = isTransactionParticipant || await isUserInWorkflow(child.AppCode, userId);

            if (!isUserInvolved) {
                console.log(`❌ HIDDEN: User not a participant in transaction or mapping for ${child.AppName}`);
                continue;
            }

            console.log(`✓ User is participant for ${child.AppName}`);

            // Check if all previous workflows are completed
            const previousCompleted = arePreviousWorkflowsCompleted(childFormDetails, childTransactions, child.ChildOrder);

            if (!previousCompleted) {
                console.log(`❌ HIDDEN: Previous workflows not completed for ${child.AppName}`);
                continue;
            }

            // Special-case for Form 4: require parent workflow and this child transaction to be completed
            if (child.AppName && child.AppName.toString().toLowerCase().includes('form 4')) {
                const parentStatus = (parentTransaction?.Status || '').toString().toLowerCase();
                const parentCompleted = parentStatus === 'completed' || parentStatus === 'approved' || parentStatus === 'closed';
                if (!parentCompleted) {
                    console.log(`❌ HIDDEN: Parent workflow not completed for ${child.AppName}`);
                    continue;
                }

                const correspondingChildTx = childTransactions.find(t => t.AppCode === child.AppCode || t.CurrentAppCode === child.AppCode);
                const childStatus = (correspondingChildTx?.Status || '').toString().toLowerCase();
                const childCompleted = childStatus === 'completed' || childStatus === 'approved' || childStatus === 'closed';
                if (!childCompleted) {
                    console.log(`❌ HIDDEN: ${child.AppName} transaction is not completed (status: ${childStatus})`);
                    continue;
                }
                console.log(`✅ Special-case passed: Parent and ${child.AppName} workflows completed`);
            }

            console.log(`✅ VISIBLE: Showing ${child.AppName} to user`);
            visibleChildren.push(child);
        }

        return visibleChildren;
    };

    // Filter child forms when ChildFormDetails, ChildtransactionItems, or isAdmin change
    useEffect(() => {
        const applyChildFiltering = async () => {
            console.log("🔄 Child filtering useEffect triggered");
            console.log("ChildFormDetails.length:", ChildFormDetails.length);
            console.log("ChildtransactionItems.length:", ChildtransactionItems.length);
            console.log("isAdmin state:", isAdmin);

            if (ChildFormDetails.length > 0 && ChildtransactionItems.length > 0) {
                const visibleChildren = await filterVisibleChildren(
                    ChildFormDetails,
                    ChildtransactionItems,
                    // isAdmin,
                    false,
                    record
                );
                setFilteredChildFormDetails(visibleChildren);
            }
        };

        void applyChildFiltering();
    }, [ChildFormDetails, ChildtransactionItems, isAdmin]);

    useEffect(() => {
        if (filteredChildFormDetails.length > 0 && ChildtransactionItems.length > 0 &&
            !selectedChild) {
            const firstChild = filteredChildFormDetails[0];
            setSelectedChild(firstChild);

            const matchedChild = ChildtransactionItems.find(
                (item) => item.AppCode === firstChild.AppCode
            );
            debugger;

            setIsAccordionVisible(!matchedChild);
            // if (!matchedChild || matchedChild.Status === "Draft") {
            //     setFormEditMode(false);
            //     setbuttonView(true);
            // }

            console.log("Matchedchildddd", matchedChild)
            console.log("From useEffect vsisibility", isAccordionVisible)
        }
    }, [filteredChildFormDetails, ChildtransactionItems]);

    const fetchFormData = async () => {
        try {
            debugger;
            // Fetch the transaction item
            const transactionItem = await sp.web.lists
                .getByTitle("WorkFlowProcessData")
                .items.getById(Number(id))
                .select("*", "Title, Created,Domain,CurrentAppCode,AuthorId, Author/Title,CategoryType,NotifyRequestor/EMail,TableJSON,TimeSheetJSON,CurApprover/Title,CurrentAppCode")
                .expand("Author", "NotifyRequestor", "CurApprover")
                .get();
            const approvers = transactionItem.AllApprovers
                ? JSON.parse(transactionItem.AllApprovers)
                : [];

            // Ensure the parsed data is an array
            if (Array.isArray(approvers)) {
                setAllApprovers(approvers as Approver[]);
            } else {
                setAllApprovers([]);
            }

            // console.log("transactionItemtransactionItemtransactionItemtransactionItemtransactionItem----", transactionItem);
            // console.log("in function outside if, id and isaccessdenied values are ", transactionItem.ID, isAccessDenied)
            // console.log("transaction item .domain ", transactionItem.Domain)
            setrecord(transactionItem);
            // Set state for AppCode and related details
            const appCodeValue = transactionItem.AppCode;
            setAppCode(appCodeValue);
            setAppCodevalue(appCodeValue);
            if (transactionItem.CurrentAppCode) {
                setCurrentAppCode(transactionItem.CurrentAppCode)
                console.log(CurrentAppCode)
            }

            const RAWformMasterItems = await sp.web.lists
                .getByTitle("FormMaster")
                .items.filter(`AppCode eq '${appCodeValue}'`)
                .get();
            const formMasterItems = RAWformMasterItems?.map((item:any)=>{
                return { ...item, FormJSON: dataserviceobj.decryptjson(item.FormJSON)};
            })
            setIsParentForm(formMasterItems[0].IsParentForm)

            // if (formMasterItems[0].IsParentForm === true) {
            //     const ChildtransactionItems = await sp.web.lists.getByTitle("WorkFlowProcessData").items
            //         .filter(`ParentAppId eq '${transactionItem.Id}'`).get()

            //     setChildtransactionItems(ChildtransactionItems)
            // }


            const ChildtransactionItems = await sp.web.lists.getByTitle("WorkFlowProcessData").items
                .filter(`ParentAppId eq '${transactionItem.Id}'`)
                .select("*", "Title", " Created", "Domain", "CurrentAppCode", "AuthorId", "Author/Title", "CategoryType", "NotifyRequestor/EMail", "TableJSON", "TimeSheetJSON", "CurApprover/Title")
                .expand("Author", "NotifyRequestor", "CurApprover")
                .orderBy("Id", true)
                .get()

            setChildtransactionItems(ChildtransactionItems)

            const CurrentChild = ChildtransactionItems.find(
                (child) => child.AppCode === transactionItem.CurrentAppCode
            );
            // setAppSeqNo(transactionItem.SeqNo)

            // if (CurrentChild) {
            //     setCurrentChildId(CurrentChild.ID);
            //     // const appSeqNo = 
            //     // setAppSeqNo(CurrentChild.SeqNo);
            //     console.log("✅ Matching child ID:", CurrentChild.ID);
            // } else {
            //     console.log("⚠️ No matching child found for CurrentAppCode:", transactionItem.CurrentAppCode);
            // }
            // const domain = await fetchTenantUser();
            await isAdminUser()
            // if (transactionItem.Domain === domain.TenantUsers && (domain.CurrentUserId === transactionItem.AuthorId || isAdmin)) {
            //     console.log("went inside if , id is ", transactionItem.ID)
            //     setisAccessDenied(false)
            // }

            setTables(JSON.parse(transactionItem.TableJSON) || [])
            setTimesheettables(JSON.parse(transactionItem.TimeSheetJSON) || [])

            setCategorytype(transactionItem.CategoryType)
            setGlobalVariable(JSON.stringify({ appcode: transactionItem.AppCode, appname: transactionItem.AppName }));
            if (transactionItem.NotifyRequestor && transactionItem.NotifyRequestor.length > 0) {
                // 1. Set emails (leave this part as is)
                const userEmails = transactionItem.NotifyRequestor.map((user: { EMail: any }) => user?.EMail);
                console.log("===========setNotifyRequestorEMail++++++++++++", userEmails);
                setNotifyRequestorEMails(userEmails);

                // 2. Fetch full user details for dropdown badges
                const usersArray: User[] = [];

                for (const user of transactionItem.NotifyRequestor) {
                    try {
                        // First, find the user ID using email (if NotifyRequestor has ID, use that instead)
                        const siteUser = await sp.web.siteUsers.getByEmail(user?.EMail).get();

                        usersArray.push({
                            userPrincipalName: siteUser.LoginName,
                            displayName: siteUser.Title,
                            mail: siteUser.Email,
                        });
                    } catch (error) {
                        console.error("Failed to fetch full user for email:", user?.EMail, error);
                    }
                }

                // setSelectedUsers(usersArray);
            }


            console.log("stateemails", NotifyRequestorEMails)
            // setNotifyRequestorEMail(transactionItem.NotifyRequestor?.EMail || "")
            SetnotifytorequestorID(transactionItem.NotifyRequestorId || [])

            // Parse FormData and set initial state
            const parsedFormData = JSON.parse(transactionItem.FormData);
            const updatedFormData = parsedFormData.map((item: any) => {
                if (item.name.startsWith("file_upload")) {
                    return {
                        ...item,
                        value: item.value ? item.value : null,
                    };
                }
                return item;
            });

            console.log("Updated Form Data:", updatedFormData);

            setFormData(updatedFormData);
            setParentFormAnswerCopy(updatedFormData);


            // Handle button visibility and edit mode based on status
            if (transactionItem.Status === "Draft") {
                setbuttonView(true);
                // setFormEditMode(false);
            }
            if (transactionItem.Status === "Returned" && transactionItem.DestinationQueue === "0") {
                setResubmitButton(true);
                // setFormEditMode(false);
            }
            const loginuser = await sp.web.currentUser();
            const createduser = transactionItem.AuthorId;
            console.log(loginuser.Id, createduser);
            console.log(ResubmitButton, buttonView);
            // console.log(transactionItem.Status === "Returned" || transactionItem.Status === "In-Progress"   && loginuser.Id === createduser);
            // if (transactionItem.Status === "Returned" || transactionItem.Status === "In-Progress" && loginuser.Id === createduser) { setCancelrequest(true) }


            setAppName(transactionItem.AppName);
            setRecentStatus(transactionItem.Status);
            setRecentTitle(transactionItem.Title);
            setLevel(transactionItem.Level);

            // setAppSeqNo(transactionItem.SeqNo);

            setFileFolder(transactionItem.SeqNo);

            // debugger;
            setRequesterName("Requester : " + transactionItem.Author.Title.split("|")[0]);
            setRequesterNamee(transactionItem.Author.Title.split("|")[0])
            // console.log('Author', transactionItem.Author.Title);

            // FETCHING DATA FROM AD

            const authorDataString = transactionItem.AuthorADdata; // Adjust the field name to 'AuthorADdata'
            const parsedAuthorData: AuthorADdata = JSON.parse(authorDataString);

            // Set the parsed data into the state
            setAuthorADdata(parsedAuthorData);
            console.log("AthorADdataaaaaaaaaaaaaaaaaaaa----->", authorADdata)



            const mappingMasterData = await sp.web.lists
                .getByTitle("MappingMaster")
                .items.select("Level", "Role", "Users_x002f_Groups/Title", "Users_x002f_Groups/Id")
                .expand("Users_x002f_Groups")
                .filter(`AppCode eq '${formMasterItems[0].IsParentForm ? transactionItem.CurrentAppCode : appCodeValue}'`)
                .get();

            const levelsWithRoles = mappingMasterData.map(item => ({
                Level: item.Level,
                Role: item.Role,
                User: item.Users_x002f_Groups ? item.Users_x002f_Groups.Title : "N/A",
            })).sort((a, b) => a.Level - b.Level);

            setMappingLevels(levelsWithRoles);
            let highestLevel
            if (mappingMasterData.length > 0) {
                highestLevel = Math.max(...mappingMasterData.map(item => item.Level));
            }
            const conditionMasterData = await sp.web.lists
                .getByTitle("ConditionsList")
                .items.select("Level", "Role", "PersonOrGroup/Title", "PersonOrGroup/Id")
                .expand("PersonOrGroup")
                .filter(`AppCode eq '${formMasterItems[0].IsParentForm ? transactionItem.CurrentAppCode : appCodeValue}'`)
                .get()

            const curApproverId = transactionItem.CurApproverId;

            let userRole = ""
            for (let i = 0; i < mappingMasterData.length; i++) {
                if (mappingMasterData[i]?.Users_x002f_Groups?.Id !== undefined && mappingMasterData[i]?.Users_x002f_Groups?.Id !== null) {
                    if (mappingMasterData[i].Level === transactionItem.Level && mappingMasterData[i].Users_x002f_Groups.Id === curApproverId && mappingMasterData[i].Role === "Creator") {

                        userRole = "Creator"
                    }
                    if (mappingMasterData[i].Level === transactionItem.Level && mappingMasterData[i].Users_x002f_Groups.Id === curApproverId && mappingMasterData[i].Role === "Reviewer") {

                        userRole = "Reviewer"
                    }
                    if (mappingMasterData[i].Level === transactionItem.Level && mappingMasterData[i].Users_x002f_Groups.Id === curApproverId && mappingMasterData[i].Role === "Task Performer") {
                        userRole = "Task Performer"
                    }
                    if (mappingMasterData[i].Level === transactionItem.Level && mappingMasterData[i].Users_x002f_Groups.Id === curApproverId && mappingMasterData[i].Role === "Approver") {
                        userRole = "Approver"
                    }
                }
            }
            for (let i = 0; i < conditionMasterData.length; i++) {

                if (conditionMasterData[i]?.PersonOrGroup?.Id !== undefined && conditionMasterData[i]?.PersonOrGroup?.Id !== null) {
                    if (conditionMasterData[i].Level === transactionItem.Level && conditionMasterData[i].PersonOrGroup.Id === curApproverId && conditionMasterData[i].Role === "Creator") {

                        userRole = "Creator"
                    }
                    if (conditionMasterData[i].Level === transactionItem.Level && conditionMasterData[i].PersonOrGroup.Id === curApproverId && conditionMasterData[i].Role === "Reviewer") {

                        userRole = "Reviewer"
                    }
                    if (conditionMasterData[i].Level === transactionItem.Level && conditionMasterData[i].PersonOrGroup.Id === curApproverId && conditionMasterData[i].Role === "Task Performer") {
                        userRole = "Task Performer"
                    }
                    if (conditionMasterData[i].Level === transactionItem.Level && conditionMasterData[i].PersonOrGroup.Id === curApproverId && conditionMasterData[i].Role === "Approver") {
                        userRole = "Approver"
                    }
                }
            }
            console.log("childappselected in fetchFormdata", CurrentChild);
            const thetransactionItem = (transactionItem.CurrentAppCode) ? (CurrentChild) ? CurrentChild : transactionItem : transactionItem;
            let endflow: boolean = false;
            if ((thetransactionItem.Status === "In-Progress" && Number(thetransactionItem.CurrentQueue) === highestLevel && userRole === "Creator") ||
                (thetransactionItem.Status === "In-Progress" && Number(thetransactionItem.CurrentQueue) === highestLevel && userRole === "Task Performer") ||
                (thetransactionItem.Status === "In-Progress" && Number(thetransactionItem.DestinationQueue) === highestLevel && userRole === "Task Performer")) {
                endflow = true;
                // setEndFlow(endflow)
            }
            // setLoading(true)
            // let ChildClicked = ""
            const flowjson = await FlowData(thetransactionItem, levelsWithRoles, endflow)
            console.log("flowjson------------------------------------------------------------------------", flowjson)
            setFlowJSON(flowjson)
            // setLoading(false)
            // Fetch and process FormMaster data
            if (formMasterItems[0].IsParentForm) {
                let RAWChildforms = await sp.web.lists.getByTitle("FormMaster").items.filter(`ParentAppCode eq '${appCodeValue}'`).get();

                const Childforms = RAWChildforms?.map((item:any)=>{
                    return { ...item, FormJSON: dataserviceobj.decryptjson(item.FormJSON)};
                })

                const currentAppCode = transactionItem?.CurrentAppCode;

                if (!currentAppCode) {
                    console.log("No CurrentAppCode found");
                }
                // Step 1: Check if CurrentAppCode exists in ChildtransactionItems
                const currentAppExists = ChildtransactionItems.some(
                    (item) => item?.AppCode === currentAppCode
                );

                // Step 2: If NOT found in ChildtransactionItems
                if (!currentAppExists) {
                    // Find that form in Childforms
                    const matchingForm = Childforms.find(
                        (form) => form?.AppCode === currentAppCode
                    );

                    if (matchingForm) {
                        const targetOrder = matchingForm.ChildOrder;

                        // Step 3: Get all forms with ChildOrder <= that target order
                        const filteredForms = Childforms.filter(
                            (form) => form.ChildOrder <= targetOrder
                        );

                        console.log("Filtered ChildForms:", filteredForms);

                        // Step 4: Update state
                        SetChildFormDetails(filteredForms);
                    } else {
                        console.log(
                            "CurrentAppCode not found in Childforms:",
                            currentAppCode
                        );
                        // SetChildFormDetails([]); // fallback
                    }
                } else if (currentAppExists) {
                    const childAppCodes = ChildtransactionItems.map((item) => item?.AppCode).filter(Boolean);

                    // Filter Childforms that match any of those AppCodes
                    const matchedForms = Childforms.filter((form) =>
                        childAppCodes.includes(form.AppCode)
                    );

                    console.log("Filtered ChildForms (exists case):", matchedForms);
                    SetChildFormDetails(matchedForms);
                }
                else {
                    console.log("CurrentAppCode already exists in ChildtransactionItems");
                    SetChildFormDetails(Childforms);
                }
            }
            const appSeqNo = thetransactionItem.SeqNo;
            setAppSeqNo(appSeqNo);


            // Fetch audit log data
            const workFlowLog = await sp.web.lists
                .getByTitle("WorkFlowAuditLog")
                .items.filter(`(AppCode eq '${transactionItem.AppCode}') and (SeqNo eq '${transactionItem.SeqNo}')`)
                .get();
            setAuditLog(workFlowLog);

            // Fetch files from NoteAttach library
            const seqNoFolder = transactionItem.SeqNo;
            const files = await sp.web.lists
                .getByTitle("NoteAttach")
                .items.select("FileLeafRef", "FileRef", "FileDirRef")
                .filter(`FileDirRef eq '${fileFetchUrl}/NoteAttach/${seqNoFolder}'`)
                .get();

            const fileNames = files.map((file, index) => {
                const fileName = file.FileLeafRef || file.Title || "Unnamed File";

                const fileUrl = file.FileRef;

                return { name: `${index + 1}. ${fileName}`, url: fileUrl };
            });
            setAttachmentFiles(fileNames);

            if (formMasterItems.length > 0) {
                const formMasterItem = formMasterItems[0];
                setreopenbuttonview(formMasterItem.VisibilityFlag && formMasterItem.Published !== false)
                // Get Master JSON (Form Structure)
                const parsedFormJSON = JSON.parse(formMasterItem.FormJSON);
                console.log("parsedFormJSON-------------", parsedFormJSON);
                // SetFormMasterJson(parsedFormJSON);
                try {
                    console.log("parsedformmmmmjson", parsedFormJSON)
                    const formdatefields: any[] = parsedFormJSON.filter((ele: any) => { return ele.field_name?.startsWith("date_picker") });
                    console.log("Formdatefields", formdatefields)
                    setdatefieldids(formdatefields)
                } catch (error) {
                    console.log("formdatefields error", error)

                }
                console.log("parsedFormData---------------------------------------", parsedFormData)
                const parsedFormDataArray = parsedFormData
                    .filter((dataItem: any) => dataItem.name.startsWith("file_upload"))
                    .map((dataItem: any) => ({
                        FileName: dataItem.name,
                        FileValue: dataItem.value,
                        FileId: dataItem.id
                    }));

                setMappedfile(parsedFormDataArray);
                console.log("parsedFormDataArray---------------------", parsedFormDataArray);

                // Create modified JSON structure
                const modifiedFormJSON = parsedFormJSON.map((item: any) => {
                    if (item.element === "FileUpload") {

                        const filename = fileNames.map((rawname) => {
                            const fname = rawname.name.replace(/^\d+\.\s*/, '').split('name')[0];
                            const controlname = rawname.name.split('name')[1];
                            if (controlname === item.field_name)
                                return fname;
                            //   return rawname.name.replace(/^\d+\.\s*/, '').split('name')[0];
                            return "";
                        }).filter(str => str.trim() !== "").join(', ')

                        return {
                            ...item,
                            text: "See Attachment Section",
                            label: filename ? `${item.label} : ${filename}` : `${item.label} : No file(s) available`,
                            className: item.className ? `${item.className} custom-red-label` : "custom-red-label",
                            style: { color: "red" }
                        };
                    }
                    return item;
                });

                console.log("EditTransaction---------modifiedFormJSON--------", modifiedFormJSON);
                setFormJSON(modifiedFormJSON);
                setParentFormJSONCopy(modifiedFormJSON);
            }

            // Fetch and process comments

            try {
                // Fetch data from the CommentsLog list
                const CommentItems = await sp.web.lists
                    .getByTitle("CommentsLog")
                    .items.filter(`(AppCode eq '${childappselected ? childappselected : appCode}') and (SeqNo eq '${appSeqNo}')`)
                    .expand("AttachmentFiles") // Include attachments
                    .select(
                        "Id",
                        "Comments",
                        "CommentedBy",
                        "CommentedByEmail",
                        "AppCode",
                        "SeqNo",
                        "AttachmentFiles/FileName",
                        "AttachmentFiles/ServerRelativeUrl",
                        "Created",
                        "Contents",
                        "Modified"
                    ) // Select required fields
                    .get();


                console.log("CommentItems------------------------", CommentItems)
                // const stripHTML = (html: string) => {
                //     const tempDiv = document.createElement("div");
                //     tempDiv.innerHTML = html;
                //     return tempDiv.textContent || tempDiv.innerText || "";
                // };

                // Process the data to include attachments
                const processedComments = CommentItems.map((comment) => ({
                    Id: comment.Id,
                    Comments: comment.Comments, // Remove HTML from Comments
                    CommentedBy: comment.CommentedBy,
                    CommentedByEmail: comment.CommentedByEmail,
                    AppCode: comment.AppCode,
                    SeqNo: comment.SeqNo,
                    Created: comment.Created,
                    Modified: comment.Modified,
                    Attachments: comment.AttachmentFiles.map((file: any) => ({
                        FileName: file.FileName,
                        FileUrl: file.ServerRelativeUrl,
                    })), // Process attachments
                }));
                processedComments.sort((a, b) => new Date(b.Modified).getTime() - new Date(a.Modified).getTime());
                console.log("Processed comments with all columns------", processedComments);
                setAllComments(processedComments);
            } catch (err) {
                console.error("Error fetching comments with all columns:", err);
            }


        } catch (error) {
            console.error("Error fetching form data:", error);
        } finally {
            setLoading(false);
        }
    };
    useEffect(() => {


        void fetchFormData();
        void fetchUserInfo();
        const applyClassToSelects = () => {
            document.querySelectorAll("select").forEach((el) => {
                el.classList.add("form-select");
            });
        };

        // Apply class initially
        applyClassToSelects();

        // Set up MutationObserver to watch for changes in the form container
        const observer = new MutationObserver(() => {
            applyClassToSelects();
        });

        // Observe the form element or the entire document
        observer.observe(document.body, { childList: true, subtree: true });

        // Cleanup observer on unmount
        return () => observer.disconnect();
    }, [id]);
    useEffect(() => {

        let TransformedTableJsons = [...tables]

        TransformedTableJsons?.forEach((field: any) => {
            const labelNodes = Array.from(document.querySelectorAll("label.static.form-label"));

            const targetLabel = labelNodes.find(label =>
                label.innerHTML.includes(field.content)
            );

            if (targetLabel && targetLabel.parentElement) {
                // const container = document.createElement("div");
                const parent = targetLabel.parentElement;
                // Replace the label within its parent
                //  targetLabel.parentElement.replaceChild(container, targetLabel);
                targetLabel.remove()
                ReactDOM.render(
                    <TableComponent tableConfig={field} onTableUpdate={handleTableUpdate} readonly={FormEditMode} seqno={appSeqNo} onRestrictionBreachUpdate={handleRestrictionBreachUpdate}
                    />,
                    parent
                );
            }
        });
    })
    const handletimesheetTableUpdate = (updatedTable: TimeSheetData) => {
        setTimesheettables(prevTables => {
            const updatedTables = prevTables.map(table =>
                table.id === updatedTable.id ? updatedTable : table
            );
            console.log('All timetables updated:', updatedTables);
            return updatedTables;
        });
        console.log("timetables", timesheettables);
    };
    const handletimesheetRestrictionBreachUpdate = (hasBreaches: boolean) => {
        setHasBreaches(hasBreaches);
        console.log("Has Restriction Breaches for this table:", hasBreaches);
    };



    useEffect(() => {

        let TransformedtimesheetTableJsons = [...timesheettables]

        TransformedtimesheetTableJsons?.forEach((field: any) => {
            const labelNodes = Array.from(document.querySelectorAll(".form-group a"));

            const targetLabel = labelNodes.find(label =>
                label.innerHTML.includes(field.content)
            );

            if (targetLabel && targetLabel.parentElement) {
                // const container = document.createElement("div");
                const parent = targetLabel.parentElement;
                // Replace the label within its parent
                //  targetLabel.parentElement.replaceChild(container, targetLabel);
                targetLabel.remove()
                ReactDOM.render(
                    <TimesheetComponent tableConfig={field} onTableUpdate={handletimesheetTableUpdate} readonly={FormEditMode} seqno={appSeqNo} onRestrictionBreachUpdate={handletimesheetRestrictionBreachUpdate}
                    />,
                    parent
                );
            }
        });
    })

    const getManagerFromAD = async (): Promise<number | null> => {
        try {
            // Get the current logged-in user
            const currentUser = await sp.web.currentUser();
            const accountName = currentUser.LoginName; // Use the LoginName field
            console.log("Current User LoginName: ", accountName);

            // Construct the PeopleManager endpoint
            const endpoint = `${mySiteUrl}/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='${encodeURIComponent(accountName)}'`;
            console.log("Constructed Endpoint: ", endpoint);

            // Fetch manager data
            const response = await fetch(endpoint, {
                method: "GET",
                headers: {
                    Accept: "application/json;odata=verbose",
                    "Content-Type": "application/json;odata=verbose",
                    "odata-version": "",
                },
            });

            // Check response status
            if (response.ok) {
                const data = await response.json();
                console.log("Fetched Data: ", data);
                const extendedManagers = data.d?.ExtendedManagers?.results || [];
                console.log("Extended Managers: ", extendedManagers);

                if (extendedManagers.length > 0) {
                    // Get the last manager in the array
                    const lastManager = extendedManagers[extendedManagers.length - 1];
                    // const email = lastManager.split("|")[2];
                    console.log("Manager Email: ", lastManager);

                    // Use PnP JS to get user by email
                    // debugger;
                    // const user = await sp.web.siteUsers.getByEmail("rakesh.chitte@cloudangles.com")();
                    const user = await sp.web.ensureUser(lastManager);
                    console.log("Manager User Details---------000000000: ", user);

                    // Return the manager's user ID
                    return user.data.Id;
                } else {
                    console.log("No managers found.");
                    return null;
                }
            } else {
                console.error("Error fetching user properties: ", response.status, response.statusText);
                return null;
            }
        } catch (error) {
            console.error("Error during fetch: ", error);
            return null;
        }
    };
    const handlepasteData = (data: any) => {
        setComment(data)
    }
    const deleteExistingFilesAndFolders = async () => {
        try {
            // Ensure fileFetchUrl is server-relative and normalized
            const normalizedFileFetchUrl = fileFetchUrl.startsWith('/')
                ? fileFetchUrl
                : `/${fileFetchUrl}`.replace(/\/+/g, '/'); // Normalize slashes


            const sequenceFolderUrl = `${normalizedFileFetchUrl}/TableAttachments/${fileFolder}`.replace(/\/+/g, '/');
            // Check if the main sequence folder exists
            let sequenceFolderExists = false;
            try {
                await sp.web.getFolderByServerRelativeUrl(sequenceFolderUrl).get();
                sequenceFolderExists = true;
            } catch (err) {
                console.warn(`Sequence folder not found: ${sequenceFolderUrl}`, err);
            }

            if (!sequenceFolderExists) {
                console.log(`No files to delete; sequence folder does not exist: ${sequenceFolderUrl}`);
                return;
            }

            // Iterate through tables
            for (const table of tables) {
                // console.log("tableffirst", table)
                const attachColumnName = table.columns.find((column: any) => column.columntype === "Attachment")?.headername || '';
                console.log("attachColumnName", attachColumnName);
                if (attachColumnName?.length === 0) {
                    console.log(`No attachment column found for table ${table.id}`);
                    continue;
                }

                const tableFolderUrl = `${sequenceFolderUrl}/${table.id}`.replace(/\/+/g, '/');

                // Check if table folder exists
                // let tableFolderExists = false;
                try {
                    await sp.web.getFolderByServerRelativeUrl(tableFolderUrl).get();
                    // tableFolderExists = true;
                } catch (err) {
                    console.warn(`Table folder not found: ${tableFolderUrl}`, err);
                    continue;
                }


                //     // Delete files in the table folder
                for (const row of table.rows || []) {
                    // console.log(" row[attachColumnName]",  row[attachColumnName]);
                    const cleanFileName = row[attachColumnName].name?.replace(/^\d+\.\s*/, "");
                    console.log("cleanFileName", cleanFileName);
                    if (!cleanFileName) {
                        console.log(`No file name found for row in table ${table.id}`);
                        continue;
                    }

                    const encodedFileName = encodeURIComponent(cleanFileName);
                    console.log("encodedFileName", encodedFileName);
                    const fileUrl = `${tableFolderUrl}/${encodedFileName}`.replace(/\/+/g, '/');

                    try {
                        await sp.web.getFileByServerRelativeUrl(fileUrl).delete();
                        console.log(`Successfully deleted file: ${fileUrl}`);
                    } catch (err) {
                        console.warn(`Failed to delete file: ${fileUrl}`, err);
                    }
                }



                // Delete the table folder
                try {
                    await sp.web.getFolderByServerRelativeUrl(tableFolderUrl).delete();
                    console.log(`Successfully deleted table folder: ${tableFolderUrl}`);
                } catch (err) {
                    console.warn(`Failed to delete table folder: ${tableFolderUrl}`, err);
                }
            }

            // Delete the main sequence folder
            try {
                await sp.web.getFolderByServerRelativeUrl(sequenceFolderUrl).delete();
                console.log(`Successfully deleted sequence folder: ${sequenceFolderUrl}`);
            } catch (err) {
                console.warn(`Failed to delete sequence folder: ${sequenceFolderUrl}`, err);
            }
        } catch (error) {
            console.error("Error in deleteExistingFilesAndFolders:", error);
            throw error; // Re-throw to handle in the caller
        }
    };
    const handleError = (error: unknown) => {
        if (error instanceof Error) {
            console.error("Error:", error.message);

            toast.current?.show({ severity: 'error', summary: 'Error', detail: error.message, life: 2000 });
        } else {
            console.error("Unexpected error", error);

            toast.current?.show({ severity: 'error', summary: '', detail: 'An unexpected error occurred.', life: 2000 });
        }
    };

    const handleSave = async (data: any) => {
        let toasterrors: string[] = []
        if (datefieldids.length > 0) {
            datefieldids.forEach(dateElement => {
                let datefieldonchange = document.querySelector(`input[name="${dateElement.field_name}"]`) as HTMLInputElement;
                if (datefieldonchange) {
                    datefieldonchange.addEventListener("change", (e: any) => {
                        let dateJson = data.find((element: any) => { return element.name === dateElement.field_name })
                        dateJson.value = e.target.value
                        console.log("sssssssssssss", e.target.value);
                        // data[5].value=e.target.value
                    })
                    datefieldonchange.dispatchEvent(new Event('change'))
                }

            });
        }
        setLoading(true)
        const files: { name: string; file: File }[] = [];

        let hasFileUploadField = false;
        let hasUploadedFiles = false;

        console.log(hasFileUploadField, hasUploadedFiles)
        // Process file upload fields and gather file details
        data.forEach((item: any) => {
            console.log("item", item)
            console.log("itemname", item.name)
            if (item.name && item.name.includes("file_upload")) {
                console.log("insideresult1fd", fileData)
                const uplfiles = fileData[item.name]
                console.log("uplfiles", uplfiles)
                if (uplfiles) {
                    uplfiles.map((file: File) => {
                        if (file && file.name) {
                            files.push({ name: file.name + "name" + item.name, file });
                            console.log("frominsideif", files);
                            hasUploadedFiles = true;
                        }
                    });
                }
            }

        })

        tables.forEach((table: any) => {
            const attachcolumnName = table.columns.find((column: any) => column.columntype === "Attachment")?.headername;

            table.rows.forEach((row: any) => { row[attachcolumnName]?.fileObject?.size > 1024 * 1024 ? toasterrors.push(`${table.content}: File size exceeded , max-allowed size is 1mb`) : null })
        }
        )

        if (toasterrors.length > 0) {
            toasterrors.forEach((error) => {
                setLoading(false)

                toast.current?.show({ severity: "error", summary: "", detail: error, life: 3000 });
            });

            setIsButtonDisabled(true);
            setTimeout(() => {
                setIsButtonDisabled(false);
            }, 5000);

            return;
        }

        const updatedFormData = data.map((item: any) => {
            if (item.name && item.name.includes("file_upload") && item.value && item.value.name) {
                return {
                    ...item,
                    value: item.value.name, // Replace file object with file name
                };
            }
            return item;
        });

        try {



            await sp.web.lists.getByTitle("WorkFlowProcessData").items.getById(Number(id)).update({
                FormData: JSON.stringify(updatedFormData), // Use updatedFormData here
                TableJSON: JSON.stringify(tables),
                TimeSheetJSON: JSON.stringify(timesheettables),
                MinDateArray: JSON.stringify(timesheettables.map((rec) => rec.fromdate)),
                MaxDateArray: JSON.stringify(timesheettables.map((rec) => rec.todate))
            });

            // Upload files as attachments if any
            for (const file of files) {
                await sp.web.lists
                    .getByTitle("WorkFlowProcessData")
                    .items.getById(Number(id))
                    .attachmentFiles.add(file.name, file.file);
            }

            const folderName = `NoteAttach/${appSeqNo}`;


            // Upload files to the subfolder if any
            for (const file of files) {
                await sp.web.getFolderByServerRelativeUrl(folderName).files.add(file.name, file.file, true);
            }
            // ------------------------------------------Saving table attachments to library-----------------------------------------------

            tables.forEach(async (table: any) => {
                const attachcolumnName = table.columns.find((column: any) => column.columntype === "Attachment")?.headername
                const boolvalue: boolean = table.rows.find((row: any) => row[attachcolumnName].name.length > 0)
                if (attachcolumnName && boolvalue) {
                    const tablefoldername = `TableAttachments/${appSeqNo}`;
                    // await sp.web.folders.add(tablefoldername);
                    const subfolderName = `${tablefoldername}/${table.id}`;
                    // await sp.web.folders.add(subfolderName);

                    for (const file of table.rows) {
                        console.log("----file=====", file)
                        const fileData = file[attachcolumnName];

                        if (fileData && fileData.name && fileData.fileObject) {
                            const uploadResult = await sp.web.getFolderByServerRelativeUrl(subfolderName).files.add(file[attachcolumnName].name, file[attachcolumnName].fileObject, true);
                            await uploadResult.file.listItemAllFields.select('Id').get()
                                .then(async (item) => {
                                    await sp.web.lists.getByTitle("TableAttachments")
                                        .items.getById(item.Id)
                                        .update({
                                            RowID: file.id
                                        });
                                });
                        }
                    }
                }


            })
            // ------------------------------------------Saving table attachments to library-----------------------------------------------
            setFormData([]);
            setParentFormAnswerCopy([]);
            setLoading(false)
            toast.current?.show({ severity: 'info', summary: '', detail: 'Your Request successfully saved as draft!', life: 2000 });
            // Reset form data after saving
            setTimeout(() => {
                navigate(-1)
            }, 1000);
        } catch (error: unknown) {
            handleError(error);
        }
    };
    const getFirstOverlappingContent = (data: any[]) => {
        // Validate input
        if (!Array.isArray(data)) {
            console.error("Invalid input: expected array, got", typeof data);
            return null;
        }

        // Helper function to parse dates safely
        const parseDate = (dateStr: string) => {
            if (!dateStr) return null;
            const date = new Date(dateStr);
            return isNaN(date.getTime()) ? null : date;
        };

        for (let i = 0; i < data.length; i++) {
            const entryA = data[i];
            if (!entryA) continue;

            const aFrom = parseDate(entryA.fromdate);
            const aTo = parseDate(entryA.todate);

            // Skip invalid entries
            if (!aFrom || !aTo) {
                console.warn(`Invalid dates in entry ${i}:`, entryA);
                continue;
            }

            for (let j = i + 1; j < data.length; j++) {
                const entryB = data[j];
                if (!entryB) continue;

                const bFrom = parseDate(entryB.fromdate);
                const bTo = parseDate(entryB.todate);

                // Skip invalid entries
                if (!bFrom || !bTo) {
                    console.warn(`Invalid dates in entry ${j}:`, entryB);
                    continue;
                }

                console.log(`Comparing entry ${i} (${aFrom} - ${aTo}) with entry ${j} (${bFrom} - ${bTo})`);

                // Check for overlap (simplified logic)
                if (aFrom <= bTo && aTo >= bFrom) {
                    console.log("Found overlap between:", entryA.content, "and", entryB.content);
                    return [
                        entryA.content || `Entry ${i}`,
                        entryB.content || `Entry ${j}`
                    ];
                }
            }
        }

        console.log("No overlapping entries found");
        return null;
    };




    // const handleUpdateForm = async (data: any) => {
    //     if (datefieldids.length > 0) {
    //         datefieldids.forEach(dateElement => {
    //             // debugger;
    //             let datefieldonchange = document.querySelector(`input[name="${dateElement.field_name}"]`) as HTMLInputElement;
    //             if (datefieldonchange) {
    //                 datefieldonchange.addEventListener("change", (e: any) => {
    //                     let dateJson = data.find((element: any) => { return element.name === dateElement.field_name })
    //                     dateJson.value = e.target.value
    //                     console.log("sssssssssssss", e.target.value);
    //                     // data[5].value=e.target.value
    //                 })
    //                 datefieldonchange.dispatchEvent(new Event('change'))
    //             }

    //         });

    //         console.log("Data set to formdata", data)
    //     }

    //     // UPDATE WITH A RESENT FILE NAME 
    //     const updatedFormData = data.map((item: any) => {
    //         if (item.name && item.name.includes("file_upload") && item.value && item.value.name) {

    //             return {
    //                 ...item,
    //                 value: fileData[item.name] || null,
    //             };
    //         }
    //         return item;
    //     });
    //     setChildFormData(updatedFormData)
    //     // setFormData(updatedFormData)
    //     // setParentFormAnswerCopy(updatedFormData)


    //     setLoading(true)
    //     const files: { name: string; file: File }[] = [];

    //     let hasFileUploadField = false;
    //     let hasUploadedFiles = false;

    //     console.log(hasFileUploadField, hasUploadedFiles)




    //     try {

    //         const toasterrors: string[] = [];
    //         if (hasBreaches) {
    //             toasterrors.push("Invalid Table Data")
    //         }
    //         const overlappingContent = getFirstOverlappingContent(timesheettables);
    //         console.log("Overlapping content result:", overlappingContent);

    //         if (overlappingContent) {
    //             toasterrors.push(`Overlapping timesheet entries found for: ${overlappingContent.join(", ")}`);
    //         }
    //         const result1 = formData.filter(
    //             (dataitem: any) => !hiddendata.some(subItem => subItem.id === dataitem.id)
    //         );

    //         result1.forEach((item: any) => {

    //             if (item.name && item.name.includes("file_upload")) {
    //                 console.log("insideresult1fd", fileData)
    //                 const uplfiles = fileData[item.name]
    //                 item.value = fileData[item.name];
    //                 console.log("uplfiles", uplfiles)
    //                 if (uplfiles) {
    //                     uplfiles.map((file: File) => {
    //                         if (file && file.name) {
    //                             files.push({ name: file.name + "name" + item.name, file });
    //                             console.log("frominsideif", files);
    //                         }
    //                     });
    //                 }
    //                 // else if (uplfiles) {
    //                 //        // Handle single file fallback
    //                 //        files.push({ name: uplfiles.name, file: item.value });
    //                 //        hasUploadedFiles = true; // At least one file is uploaded
    //                 //     }
    //             }



    //             // Checking from mappedfile if there WAS a value

    //             if (Array.isArray(FormMasterJson) && Array.isArray(mappedfile)) {
    //                 let formjsonitem = FormMasterJson.find((formjson: any) => {
    //                     let mappedFileItem = mappedfile.find((file: any) => file.FileId === item.id);

    //                     // Check conditions to update item.value
    //                     if (
    //                         mappedFileItem &&
    //                         formjson.required === true &&
    //                         mappedFileItem.FileValue &&
    //                         mappedFileItem.FileValue !== "" && mappedFileItem.FileValue !== null && item.value === null
    //                     ) {
    //                         item.value = mappedFileItem.FileValue;
    //                         console.log("maappeedd", item.value)
    //                         return true; // Ensures function returns true when condition matches
    //                     }

    //                     return formjson.id === item.id; // Always return this condition
    //                 });

    //                 attachmentFiles.forEach((filename) => {
    //                     (item.name === filename.name.replace(/^\d+\.\s*/, '').split("name")[1]) ? item.value = filename.name : null;
    //                     console.log("previous", item)
    //                 })

    //                 console.log("current item has --text--", formjsonitem, item.value, mappedfile);


    //                 switch (formjsonitem.text) {
    //                     case "File Upload":
    //                         if (formjsonitem.required && (!item.value || item.value === null)) {
    //                             toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : Please upload a file.`);
    //                         }
    //                         break;
    //                     case "Text Input":
    //                         if (formjsonitem.required && !item.value.trim()) {
    //                             toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : This field is required.`);
    //                         }
    //                         if (item.value.length > Number(formjsonitem.maxLength)) {
    //                             toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : Maximum ${formjsonitem.maxLength} characters are allowed`);
    //                         }
    //                         break;
    //                     case "Multi-line Input":
    //                         if (formjsonitem.required && !item.value.trim()) {
    //                             toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : This field is required.`);
    //                         }
    //                         if (item.value.length > Number(formjsonitem.maxLengthforMultiline)) {
    //                             toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : Maximum ${formjsonitem.maxLengthforMultiline} characters are allowed`);
    //                         }
    //                         break;
    //                     case "Email":
    //                         if (formjsonitem.required && !item.value.trim()) {
    //                             toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : This field is required.`);
    //                         }
    //                         if (item.value.trim() && !/^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/.test(item.value.trim())) {
    //                             toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : Invalid Email Address`);
    //                         }
    //                         break;
    //                     case "Currency Input":
    //                         if (formjsonitem.required && !item.value.trim()) {
    //                             toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : This field is required.`);
    //                         }
    //                         if (item.value.trim()) {
    //                             if (!/^\d*\.?\d*$/.test(item.value)) {
    //                                 toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : Only numeric values are allowed.`);
    //                             } else if (item.value.split('.')[0].length > 10) {
    //                                 toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : Invalid currency value. Please enter a number between 0 and 1,000,000,000.`);
    //                             } else if (item.value.includes('.') && item.value.split('.')[1].length > 2) {
    //                                 toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : Only up to 2 digits are allowed after the decimal point.`);
    //                             }
    //                         }
    //                         break;
    //                     case "Number Input":
    //                         if (formjsonitem.required && !item.value.trim()) {
    //                             toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : This field is required.`);
    //                         }
    //                         if (item.value.trim()) {

    //                             if (!/^\d*\.?\d*$/.test(item.value)) {
    //                                 toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : Only numeric values are allowed.`);
    //                             } else if (item.value.split('.')[0].length > 10) {
    //                                 toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : Only up to 10 digits are allowed.`);
    //                             } else if (item.value.includes('.') && item.value.split('.')[1].length > 2) {
    //                                 toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : Only up to 2 digits are allowed after the decimal point.`);
    //                             }
    //                         }
    //                         break;
    //                     case "Phone Number":
    //                         if (formjsonitem.required && !item.value.trim()) {
    //                             toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : This field is required.`);
    //                         }
    //                         if (item.value.trim()) {

    //                             if (!/^\d*$/.test(item.value.trim())) {
    //                                 toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : Only numbers are allowed`);
    //                             } else if (item.value.trim().length > 10) {
    //                                 toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : Phone number cannot be more than 10 digits`);
    //                             } else if (item.value.trim().length < 10 && item.value.trim().length > 0) {
    //                                 toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : Phone number must be exactly 10 digits`);
    //                             }
    //                         }
    //                         break;
    //                     case "Checkboxes":
    //                         if (formjsonitem.required && !item.value.length) {
    //                             toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : This field is required.`);
    //                         }
    //                         if ((item.value) && (formjsonitem.maxAllowed) && (item.value.length > formjsonitem.maxAllowedValue)) {
    //                             toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : maximum ${formjsonitem.maxAllowedValue} options are allowed `)
    //                         }

    //                         break;
    //                     case "Date":
    //                         if (formjsonitem.required && !item.value) {
    //                             toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : This field is required.`);
    //                         }

    //                         if (item.value && !isNaN(new Date(item.value).getTime()) && !(item.showTimeSelect && item.showTimeSelectOnly)) {

    //                             const itemDate = new Date(item.value);
    //                             let today = new Date();
    //                             today.setHours(0, 0, 0, 0)
    //                             let futuretoday = new Date()
    //                             futuretoday.setHours(23, 59, 59, 999)
    //                             console.log("today and itemdate", today, itemDate)
    //                             if (formjsonitem.restrictPastdays && itemDate < today) {
    //                                 toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : Past dates are not allowed.`);
    //                                 return;
    //                             }
    //                             if (formjsonitem.restrictFuturedays && itemDate > futuretoday) {
    //                                 toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : Future dates are not allowed.`);
    //                                 return;
    //                             }

    //                             if (formjsonitem.dateRange) {
    //                                 const minDate = formjsonitem.minDate
    //                                     ? new Date(convertDateFormat(formjsonitem.minDate))
    //                                     : null; // Min date
    //                                 let maxDate = formjsonitem.maxDate
    //                                     ? new Date(convertDateFormat(formjsonitem.maxDate))
    //                                     : null; // Max date
    //                                 if (maxDate) {
    //                                     maxDate.setHours(23, 59, 59, 999); // Sets the time to 23:59:59.999
    //                                 }
    //                                 // Validate only if date is provided
    //                                 // if (itemDate && !isNaN(itemDate.getTime())) { // Checks if date is valid    

    //                                 if ((minDate && maxDate) && (itemDate < minDate || itemDate > maxDate)) {
    //                                     console.log("mindate", minDate.toISOString().split("T")[0])
    //                                     console.log("mindate", minDate)
    //                                     console.log("maxdate", maxDate.toISOString().split("T")[0])
    //                                     console.log("onlyitemDate", itemDate)
    //                                     toasterrors.push(
    //                                         `${stripHtmlTags(formjsonitem.label)} : Please select a date between ${convertDateFormat(formjsonitem.minDate)} and ${convertDateFormat(formjsonitem.maxDate)}`

    //                                     );
    //                                     return;
    //                                 } else if (minDate && itemDate < minDate) {
    //                                     toasterrors.push(
    //                                         `${stripHtmlTags(formjsonitem.label)} : Date should be greater than or equal to ${convertDateFormat(formjsonitem.minDate)}`
    //                                     );
    //                                     return;
    //                                 } else if (maxDate && itemDate > maxDate) {
    //                                     toasterrors.push(
    //                                         `${stripHtmlTags(formjsonitem.label)} : Date should be less than or equal to ${convertDateFormat(formjsonitem.maxDate)}`
    //                                     );
    //                                     return;
    //                                 }
    //                             }
    //                             if (formjsonitem.allowOnlyCurrentYearDates && itemDate.getFullYear() !== today.getFullYear()) {
    //                                 toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : Please select a date within the current year (${today.getFullYear()}).`);

    //                             }
    //                             if (formjsonitem.weekdaySelectionOnly) {
    //                                 const dayOfWeek = itemDate.getDay(); // Get the day of the week (0 = Sunday, 6 = Saturday)
    //                                 if (dayOfWeek === 0 || dayOfWeek === 6) {
    //                                     toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : Please select a weekday (Monday to Friday).`);
    //                                 }
    //                             }


    //                         }
    //                         // else if (isNaN(new Date(item.value).getTime())) {
    //                         //   // Invalid date format
    //                         //   toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : Invalid date format`);
    //                         //   return;
    //                         // }
    //                         break;
    //                     default:
    //                         if (formjsonitem.required && !item.value?.length) {
    //                             toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : This field is required.`);
    //                         }
    //                         break;
    //                 }
    //             }
    //         });




    //         if (!Categorytype) {
    //             toasterrors.push("Request For Field is required");
    //         }
    //         else if (Categorytype.length > 80) {
    //             toasterrors.push("Request For : Only 80 characters are allowed ");
    //         }
    //         console.log("taabledata", tables)
    //         tables.forEach((table: any) => {
    //             const attachcolumnName = table.columns.find((column: any) => column.columntype === "Attachment")?.headername;
    //             table.rows.forEach((row: any) => { row[attachcolumnName]?.fileObject?.size > 1024 * 1024 ? toasterrors.push(`${table.content}: File size exceeded , max-allowed size is 1mb`) : null })
    //         }
    //         )


    //         if (toasterrors.length > 0) {
    //             toasterrors.forEach((error) => {
    //                 setLoading(false)
    //                 toast.current?.show({ severity: "error", summary: "", detail: error, life: 3000 });
    //             });

    //             setIsButtonDisabled(true);
    //             setTimeout(() => {
    //                 setIsButtonDisabled(false);
    //             }, 3000);

    //             return;
    //         }


    //         const returned = await sp.web.lists.getByTitle("WorkFlowProcessData").items.getById(Number(id)).get();
    //         const returnedStatus = returned.Status;
    //         const ZeroDestinationQueue = returned.DestinationQueue;
    //         const buttonAction = "Submit"

    //         setLoading(true);
    //         const routingRulesItems = await sp.web.lists.getByTitle("RoutingRules").items.filter(`(AppCode eq '${IsParentForm ? record.CurrentAppCode : appCode}') and (Actions eq '${buttonAction}')`).get();

    //         // debugger;

    //         const { CurrentQueue, DestinationQueue, Status } = routingRulesItems[0];

    //         if (routingRulesItems.length === 0) {
    //             setLoading(false)
    //             alert("Reviewers or Approvers are not available for further actions");
    //             return;
    //         }

    //         const mappingMasterItem = await sp.web.lists.getByTitle("MappingMaster").items
    //             .select("*", "Users_x002f_Groups/Id", "Users_x002f_Groups/Title")
    //             .expand("Users_x002f_Groups")
    //             .filter(`(Level eq '${DestinationQueue}') and (AppCode eq '${IsParentForm ? record.CurrentAppCode : appCode}')`)
    //             .get();


    //         let conditionlistItem: any[] = [];
    //         let result: boolean = false;

    //         try {
    //             // Fetching condition list items
    //             conditionlistItem = await sp.web.lists
    //                 .getByTitle("ConditionsList")
    //                 .items.select("*", "PersonOrGroup/Id", "PersonOrGroup/Title")
    //                 .expand("PersonOrGroup")
    //                 .filter(`(Level eq '${mappingMasterItem[0].Level}') and (AppCode eq '${IsParentForm ? record.CurrentAppCode : appCode}')`)
    //                 .get();

    //             console.log("-------conditionlistItem---------", conditionlistItem);

    //             // Ensure there is data to process
    //             if (conditionlistItem.length === 0) {
    //                 throw new Error("No data found in ConditionsList for the provided filter.");
    //             }

    //             // Parse the ConditionJson and process conditions
    //             const conditionlistItemJson: any = JSON.parse(conditionlistItem[0].ConditionJson);

    //             const mainoperator: string = conditionlistItemJson.operator;
    //             result = mainoperator === "And" ? true : false;

    //             conditionlistItemJson.conditions.forEach((element: any) => {
    //                 try {
    //                     const fieldId = FormMasterJson.find((val: any) => stripHtmlTags(val.label) === element.field)?.id;

    //                     if (!fieldId) {
    //                         throw new Error(`Field ID not found for field: ${element.field}`);
    //                     }

    //                     const fieldvalue = formData.find((d: any) => d.id === fieldId)?.value;

    //                     if (fieldvalue === undefined) {
    //                         throw new Error(`Field value not found for field ID: ${fieldId}`);
    //                     }

    //                     const boolvalue = ConditionValidator(fieldvalue, element.operator, element.value, element.elementType);

    //                     result = mainoperator === "And" ? boolvalue && result : boolvalue || result;

    //                     console.log("-------result inside--", result);
    //                 } catch (innerError: any) {
    //                     console.error("Error processing condition element:", innerError.message);
    //                 }
    //             });

    //             console.log("Final result:", result);
    //         } catch (error: any) {
    //             console.error("Error occurred in processing conditions:", error.message);
    //         }

    //         // Determine the userGroupId based on conditions
    //         let userGroupId: number | null = null;

    //         if (result && conditionlistItem.length > 0) {
    //             userGroupId = conditionlistItem[0]?.PersonOrGroup?.Id || null;
    //         } else {
    //             userGroupId = mappingMasterItem?.[0]?.Users_x002f_Groups?.Id || null;
    //         }

    //         console.log("User Group ID--------------------: ", userGroupId);



    //         if (userGroupId) {
    //             console.log("User Group ID: ", userGroupId);

    //         } else {
    //             console.log("User Group ID is null, fetching manager from AD...");

    //             try {
    //                 const managerId = await getManagerFromAD();

    //                 console.log("mmmmmmmmmmmmmmmmmmmmmmmmmId", managerId);
    //                 // debugger;
    //                 if (managerId) {
    //                     console.log("Immediate Manager: ", managerId);


    //                     userGroupId = managerId; // Retrieve the User ID
    //                     console.log("Author/User ID-----------------: ", userGroupId);
    //                     // Proceed with logic using userGroupId

    //                     console.error("No user found for the given email.");

    //                 } else {
    //                     console.log("No manager found or an error occurred.");
    //                 }
    //             } catch (error) {
    //                 console.error("Error fetching manager or user details: ", error);
    //             }
    //         }



    //         const MMLevel = mappingMasterItem[0].Level



    //         //-------------------------------------------------------------------------------------------------------
    //         if ((returnedStatus === "Returned") && (ZeroDestinationQueue === "0")) {

    //             let commentValidationValue = "false"

    //             if (ResubmitButton && comment.length > 1) {
    //                 commentValidationValue = "true"
    //                 try {
    //                     const rr = await sp.web.lists.getByTitle("CommentsLog").items.add({
    //                         AppCode: AppCodevalue,
    //                         Comments: comment,
    //                         CommentedBy: userName,
    //                         CommentedByEmail: userEmail,
    //                         SeqNo: appSeqNo
    //                     });

    //                     // Access the item ID from the result
    //                     const itemId = rr.data.Id;

    //                     // If a file is selected, add it to the list item
    //                     if (selectedFile.length > 0) {
    //                         for (const file of selectedFile) {
    //                             await sp.web.lists
    //                                 .getByTitle("CommentsLog")
    //                                 .items.getById(itemId)
    //                                 .attachmentFiles.add(file.name, file);
    //                         }
    //                         console.log("Files successfully uploaded!");
    //                     } else {
    //                         console.log("No file selected, skipping file upload.");
    //                     }

    //                 }
    //                 catch (err) {
    //                     console.log("error", err)
    //                 }
    //             }





    //             if (commentValidationValue === "true") {
    //                 const updatedFormData = data.map((item: any) => {
    //                     if (item.name && item.name.includes("file_upload") && item.value && item.value.name) {
    //                         return {
    //                             ...item,
    //                             value: item.value.name,
    //                         };
    //                     }
    //                     return item;
    //                 });

    //                 // Function to assign missing file names from attachmentFiles
    //                 const assignMissingFiles = (updatedFormData: any[], attachmentFiles: any[]) => {
    //                     let attachmentIndex = 0;

    //                     return updatedFormData.map((item: any) => {
    //                         if (
    //                             item.name.includes("file_upload") &&
    //                             (item.value === null || item.value === null) &&
    //                             attachmentIndex < attachmentFiles.length
    //                         ) {
    //                             return {
    //                                 ...item,
    //                                 value: attachmentFiles[attachmentIndex++].name.split('. ')[1], // Extract actual filename
    //                             };
    //                         }
    //                         return item;
    //                     });
    //                 };

    //                 // Assign missing file names dynamically
    //                 const finalUpdatedFormData = assignMissingFiles(updatedFormData, attachmentFiles);

    //                 try {
    //                     await sp.web.lists.getByTitle("WorkFlowProcessData").items.getById(Number(id)).update({

    //                         FormData: JSON.stringify(finalUpdatedFormData),
    //                         CurrentQueue: CurrentQueue,
    //                         DestinationQueue: DestinationQueue,
    //                         Status: Status,
    //                         CurApproverId: userGroupId,
    //                         Level: MMLevel,
    //                         CategoryType: Categorytype,
    //                         NotifyRequestorId: {
    //                             results: notifytorequestorID
    //                         },
    //                         TableJSON: JSON.stringify(tables),
    //                         TimeSheetJSON: JSON.stringify(timesheettables),
    //                         MinDateArray: JSON.stringify(timesheettables.map((rec) => rec.fromdate)),
    //                         MaxDateArray: JSON.stringify(timesheettables.map((rec) => rec.todate))

    //                     });
    //                     await sp.web.lists.getByTitle("WorkFlowLog").items.add({
    //                         AppCode: appCode,
    //                         Status: Status,
    //                         Domain: domainfield,
    //                         Level: MMLevel,
    //                         SeqNo: appSeqNo,
    //                         Role: routingRulesItems[0].Role,
    //                         Action: "Re-Submit",
    //                     })
    //                 } catch (error) {
    //                     console.error("error in updating filled data to list", error)
    //                 }


    //                 await deleteExistingFilesAndFolders();
    //                 tables.forEach(async (table: any) => {
    //                     const attachcolumnName = table.columns.find((column: any) => column.columntype === "Attachment")?.headername
    //                     const boolvalue: boolean = table.rows.find((row: any) => row[attachcolumnName].name.length > 0)
    //                     if (attachcolumnName && boolvalue) {
    //                         const tablefoldername = `TableAttachments/${appSeqNo}`;
    //                         await sp.web.folders.add(tablefoldername);
    //                         const subfolderName = `${tablefoldername}/${table.id}`;
    //                         await sp.web.folders.add(subfolderName);

    //                         for (const file of table.rows) {
    //                             console.log("----file=====", file)
    //                             const fileData = file[attachcolumnName];

    //                             if (fileData && fileData.name && fileData.fileObject) {
    //                                 const uploadResult = await sp.web.getFolderByServerRelativeUrl(subfolderName).files.add(file[attachcolumnName].name, file[attachcolumnName].fileObject, true);
    //                                 await uploadResult.file.listItemAllFields.select('Id').get()
    //                                     .then(async (item) => {
    //                                         await sp.web.lists.getByTitle("TableAttachments")
    //                                             .items.getById(item.Id)
    //                                             .update({
    //                                                 RowID: file.id
    //                                             });
    //                                     });
    //                             }
    //                         }
    //                     }


    //                 })
    //                 // ------------------------------------------Saving table attachments to library-----------------------------------------------

    //                 const folderName = `NoteAttach/${appSeqNo}`;


    //                 try {

    //                     // Check if folder exists
    //                     await sp.web.getFolderByServerRelativeUrl(folderName).get();
    //                 } catch (error) {
    //                     // If folder does not exist, create it
    //                     await sp.web.folders.add(folderName);
    //                 }

    //                 // Upload files to the subfolder
    //                 for (const file of files) {
    //                     await sp.web.getFolderByServerRelativeUrl(folderName).files.add(file.name, file.file, true);
    //                 }

    //                 toast.current?.show({ severity: 'success', summary: '', detail: 'Details re-submitted successfully ', life: 2000 });

    //                 setTimeout(() => {
    //                     navigate(-1);
    //                 }, 1000)
    //             }

    //             else {
    //                 setLoading(false)
    //                 toast.current?.show({ severity: 'error', summary: '', detail: 'Add a comment before submitting !', life: 2000 });
    //                 setHighlight(true)
    //             }

    //         }



    //         else {


    //             // const allItems = await sp.web.lists
    //             //     .getByTitle("WorkFlowProcessData")
    //             //     .items.orderBy("Created", false) // Ensure newest records come first
    //             //     .filter(`AppName eq '${appName}' `)
    //             //     .getAll()
    //             //     .catch((error) => {
    //             //         console.error("Error fetching items:", error);
    //             //         return [];
    //             //     });

    //             const formMasterItem2 = await sp.web.lists.getByTitle("FormMaster")
    //                 .items.select("FormPrefix,isWorkflowRequired")
    //                 .filter(`AppName eq '${appName}' `).get();

    //             let isNoWorkflowRequiredFlag = formMasterItem2[0]?.isWorkflowRequired === "No" ? true : false;

    //             let NoApprovalStatus = "";
    //             debugger;
    //             if (isNoWorkflowRequiredFlag) {
    //                 NoApprovalStatus = "Submitted";
    //             }
    //             else {
    //                 NoApprovalStatus = Status
    //             }

    //             // let latestNumber = 0;
    //             // console.log("allItems------------", allItems)
    //             // for (const item of allItems) {
    //             //     if (item.Status !== "Draft") {
    //             //         const titleParts = item.Title.split("/");

    //             //         console.log("titleParts----------------- !!!!!!!!!!!!!!", titleParts, titleParts.length);

    //             //         if (titleParts.length >= 3) {
    //             //             const numberPart = parseInt(titleParts[2], 10);
    //             //             console.log("numberPart", numberPart);

    //             //             if (!isNaN(numberPart) && numberPart > latestNumber) {
    //             //                 latestNumber = numberPart; // Store the highest number found
    //             //             }
    //             //         }
    //             //     }
    //             // }
    //             // const seqNo = await getCurrentUserSeqNo();
    //             // const date = new Date();
    //             // const month = date.toLocaleString('default', { month: 'short' });
    //             // const year = date.getFullYear().toString().slice(-2);
    //             // const monthYear = `${month}_${year}`;

    //             // const newTitle = `${formMasterItem2[0].FormPrefix}/${monthYear}/${latestNumber + 1}`;


    //             console.log("ddaattaa", data)

    //             try {

    //                 const aduserdetails = await fetchToken();
    //                 // debugger;
    //                 // console.log("aduserdetails--------------> ",aduserdetails)
    //                 // console.log(aduserdetails.employeeId)
    //                 let employeeId = "---";
    //                 let jobTitle = "---";
    //                 let department = "---";
    //                 let mobilePhone = "---"
    //                 // debugger;
    //                 if (aduserdetails) {
    //                     console.log("aduserdetails--------------> ", aduserdetails);
    //                     employeeId = aduserdetails.employeeId || "---";
    //                     jobTitle = aduserdetails.jobTitle || "---";
    //                     department = aduserdetails.department || "---";
    //                     mobilePhone = aduserdetails.mobilePhone || "---";
    //                 } else {
    //                     // If fetchToken fails, log the error and continue with NA values
    //                     console.error("Failed to fetch Auther user details from AD");
    //                 }

    //                 // Create the object to store in the ADuserData column
    //                 const adUserData = {
    //                     RequesterName: RequesterTitle,
    //                     Email: RequesterEmail,
    //                     EmployeeID: employeeId,
    //                     Designation: jobTitle,
    //                     Department: department,
    //                     ContactNo: mobilePhone
    //                 };

    //                 // Convert the adUserData object to a JSON string
    //                 const adUserDataJSON = JSON.stringify(adUserData);
    //                 await sp.web.lists.getByTitle("WorkFlowProcessData").items.getById(Number(id)).update({
    //                     // Title: newTitle,
    //                     // FormData: JSON.stringify(finalUpdatedFormData),
    //                     // FormData: JSON.stringify(data),
    //                     CurrentQueue: CurrentQueue,
    //                     DestinationQueue: DestinationQueue,
    //                     // Status: NoApprovalStatus,
    //                     CurApproverId: userGroupId,
    //                     Level: MMLevel,
    //                     // ManagersId: userGroupId,
    //                     CategoryType: Categorytype,
    //                     NotifyRequestorId: {
    //                         results: notifytorequestorID
    //                     },
    //                     // TableJSON: JSON.stringify(tables),
    //                     // TimeSheetJSON: JSON.stringify(timesheettables),
    //                     // MinDateArray: JSON.stringify(timesheettables.map((rec) => rec.fromdate)),
    //                     // MaxDateArray: JSON.stringify(timesheettables.map((rec) => rec.todate))
    //                 });

    //                 await sp.web.lists.getByTitle("WorkFlowProcessData").items.add({
    //                     Title: record.Title,
    //                     AuthorADdata: adUserDataJSON,
    //                     AppCode: record.CurrentAppCode,
    //                     ParentAppCode: appCode,
    //                     FormData: JSON.stringify(data),
    //                     ParentAppId: id,
    //                     Status: NoApprovalStatus,
    //                     SeqNo: appSeqNo,
    //                     Domain: domainfield,
    //                     Level: MMLevel,

    //                 })

    //                 await sp.web.lists.getByTitle("WorkFlowLog").items.add({
    //                     AppCode: record.CurrentAppCode,
    //                     Status: NoApprovalStatus,
    //                     Domain: domainfield,
    //                     Level: MMLevel,
    //                     SeqNo: appSeqNo,
    //                     Role: routingRulesItems[0].Role,
    //                     Action: buttonAction,
    //                 })

    //             } catch (error) {
    //                 console.error("Submit case patching to list failed due to : ", error)
    //             }

    //             await deleteExistingFilesAndFolders();
    //             // ------------------------------------------Saving table attachments to library-----------------------------------------------
    //             tables.forEach(async (table: any) => {
    //                 const attachcolumnName = table.columns.find((column: any) => column.columntype === "Attachment")?.headername
    //                 const boolvalue: boolean = table.rows.find((row: any) => row[attachcolumnName].name.length > 0)
    //                 if (attachcolumnName && boolvalue) {
    //                     const tablefoldername = `TableAttachments/${appSeqNo}`;
    //                     await sp.web.folders.add(tablefoldername);
    //                     const subfolderName = `${tablefoldername}/${table.id}`;
    //                     await sp.web.folders.add(subfolderName);

    //                     for (const file of table.rows) {
    //                         console.log("----file=====", file)
    //                         const fileData = file[attachcolumnName];

    //                         if (fileData && fileData.name && fileData.fileObject) {
    //                             const uploadResult = await sp.web.getFolderByServerRelativeUrl(subfolderName).files.add(file[attachcolumnName].name, file[attachcolumnName].fileObject, true);
    //                             await uploadResult.file.listItemAllFields.select('Id').get()
    //                                 .then(async (item) => {
    //                                     await sp.web.lists.getByTitle("TableAttachments")
    //                                         .items.getById(item.Id)
    //                                         .update({
    //                                             RowID: file.id
    //                                         });
    //                                 });
    //                         }
    //                     }
    //                 }


    //             })
    //             // ------------------------------------------Saving table attachments to library-----------------------------------------------

    //             // const itemId = newItem.data.Id;
    //             const folderName = `NoteAttach/${appSeqNo}`;


    //             try {
    //                 // Check if folder exists
    //                 await sp.web.getFolderByServerRelativeUrl(folderName).get();
    //             } catch (error) {
    //                 // If folder does not exist, create it
    //                 await sp.web.folders.add(folderName);
    //             }

    //             // Upload files to the subfolder
    //             for (const file of files) {
    //                 await sp.web.getFolderByServerRelativeUrl(folderName).files.add(file.name, file.file, true);
    //             }

    //             toast.current?.show({ severity: 'success', summary: '', detail: 'Data submitted successfully!', life: 2000 });
    //             navigate(-1);
    //         }

    //     } catch (error) {
    //         console.error("Error updating form data submit case:", error);
    //     } finally {
    //         setLoading(false);
    //     }


    // };


    // console.log("hare krishna")


    const handleUpdateForm = async (data: any) => {
        debugger;
        if (data.length === 0) {
            toast.current?.show({ severity: 'error', summary: '', detail: 'Please fill out the form before submitting.', life: 2000 });
            return
        }
        if (datefieldids.length > 0) {
            datefieldids.forEach(dateElement => {
                // debugger;
                let datefieldonchange = document.querySelector(`input[name="${dateElement.field_name}"]`) as HTMLInputElement;
                if (datefieldonchange) {
                    datefieldonchange.addEventListener("change", (e: any) => {
                        let dateJson = data.find((element: any) => { return element.name === dateElement.field_name })
                        dateJson.value = e.target.value
                        console.log("sssssssssssss", e.target.value);
                        // data[5].value=e.target.value
                    })
                    datefieldonchange.dispatchEvent(new Event('change'))
                }

            });

            console.log("Data set to formdata", data)
        }

        // UPDATE WITH A RESENT FILE NAME 
        const updatedFormData = data.map((item: any) => {
            if (item.name && item.name.includes("file_upload") && item.value && item.value.name) {

                return {
                    ...item,
                    value: fileData[item.name] || null,
                };
            }
            return item;
        });
        setChildFormData(updatedFormData)
        // setParentFormAnswerCopy(updatedFormData)


        setLoading(true)
        const files: { name: string; file: File }[] = [];

        let hasFileUploadField = false;
        let hasUploadedFiles = false;

        console.log(hasFileUploadField, hasUploadedFiles)




        try {

            const toasterrors: string[] = [];
            if (hasBreaches) {
                // toasterrors.push("Invalid Table Data")
                console.log("hasBreacheshasBreacheshasBreaches", hasBreaches)
            }
            const overlappingContent = getFirstOverlappingContent(timesheettables);
            console.log("Overlapping content result:", overlappingContent);

            if (overlappingContent) {
                toasterrors.push(`Overlapping timesheet entries found for: ${overlappingContent.join(", ")}`);
            }
            const result1 = data.filter(
                (dataitem: any) => !hiddendata.some(subItem => subItem.id === dataitem.id)
            );
            console.log("inupdateform formdata and result1", childformData, result1);

            result1.forEach((item: any) => {

                if (item.name && item.name.includes("file_upload")) {
                    console.log("insideresult1fd", fileData)
                    const uplfiles = fileData[item.name]
                    item.value = fileData[item.name];
                    console.log("uplfiles", uplfiles)
                    if (uplfiles) {
                        uplfiles.map((file: File) => {
                            if (file && file.name) {
                                files.push({ name: file.name + "name" + item.name, file });
                                console.log("frominsideif", files);
                            }
                        });
                    }
                    // else if (uplfiles) {
                    //        // Handle single file fallback
                    //        files.push({ name: uplfiles.name, file: item.value });
                    //        hasUploadedFiles = true; // At least one file is uploaded
                    //     }
                }
                if (Array.isArray(ChildFormJsons) && Array.isArray(mappedfile)) {
                    debugger
                    // console.log("insidefileupload", FormMasterJson)
                    let formjsonitem = ChildFormJsons.find((formjson: any) => {
                        let mappedFileItem = mappedfile.find((file: any) => file.FileId === item.id);

                        // Check conditions to update item.value
                        if (
                            mappedFileItem &&
                            formjson.required === true &&
                            mappedFileItem.FileValue &&
                            mappedFileItem.FileValue !== "" && mappedFileItem.FileValue !== null && item.value === null
                        ) {
                            item.value = mappedFileItem.FileValue;
                            console.log("maappeedd", item.value)
                            return true; // Ensures function returns true when condition matches
                        }

                        return formjson.id === item.id; // Always return this condition
                    });

                    attachmentFiles.forEach((filename) => {
                        (item.name === filename.name.replace(/^\d+\.\s*/, '').split("name")[1]) ? item.value = filename.name : null;
                        console.log("previous", item)
                    })

                    console.log("current item has --text--", formjsonitem, item.value, mappedfile);


                    switch (formjsonitem.text) {
                        case "File Upload":
                            if (formjsonitem.required && (!item.value || item.value === null)) {
                                toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : Please upload a file.`);
                            }
                            break;
                        case "Text Input":
                            if (formjsonitem.required && !item.value.trim()) {
                                toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : This field is required.`);
                            }
                            if (item.value.length > Number(formjsonitem.maxLength)) {
                                toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : Maximum ${formjsonitem.maxLength} characters are allowed`);
                            }
                            break;
                        case "Multi-line Input":
                            if (formjsonitem.required && !item.value.trim()) {
                                toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : This field is required.`);
                            }
                            if (item.value.length > Number(formjsonitem.maxLengthforMultiline)) {
                                toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : Maximum ${formjsonitem.maxLengthforMultiline} characters are allowed`);
                            }
                            break;
                        case "Email":
                            if (formjsonitem.required && !item.value.trim()) {
                                toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : This field is required.`);
                            }
                            if (item.value.trim() && !/^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/.test(item.value.trim())) {
                                toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : Invalid Email Address`);
                            }
                            break;
                        case "Currency Input":
                            if (formjsonitem.required && !item.value.trim()) {
                                toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : This field is required.`);
                            }
                            if (item.value.trim()) {
                                if (!/^\d*\.?\d*$/.test(item.value)) {
                                    toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : Only numeric values are allowed.`);
                                } else if (item.value.split('.')[0].length > 10) {
                                    toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : Invalid currency value. Please enter a number between 0 and 1,000,000,000.`);
                                } else if (item.value.includes('.') && item.value.split('.')[1].length > 2) {
                                    toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : Only up to 2 digits are allowed after the decimal point.`);
                                }
                            }
                            break;
                        case "Number Input":
                            if (formjsonitem.required && !item.value.trim()) {
                                toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : This field is required.`);
                            }
                            if (item.value.trim()) {

                                if (!/^\d*\.?\d*$/.test(item.value)) {
                                    toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : Only numeric values are allowed.`);
                                } else if (item.value.split('.')[0].length > 10) {
                                    toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : Only up to 10 digits are allowed.`);
                                } else if (item.value.includes('.') && item.value.split('.')[1].length > 2) {
                                    toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : Only up to 2 digits are allowed after the decimal point.`);
                                }
                            }
                            break;
                        case "Phone Number":
                            if (formjsonitem.required && !item.value.trim()) {
                                toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : This field is required.`);
                            }
                            if (item.value.trim()) {

                                if (!/^\d*$/.test(item.value.trim())) {
                                    toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : Only numbers are allowed`);
                                } else if (item.value.trim().length > 10) {
                                    toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : Phone number cannot be more than 10 digits`);
                                } else if (item.value.trim().length < 10 && item.value.trim().length > 0) {
                                    toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : Phone number must be exactly 10 digits`);
                                }
                            }
                            break;
                        case "Checkboxes":
                            if (formjsonitem.required && !item.value.length) {
                                toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : This field is required.`);
                            }
                            if ((item.value) && (formjsonitem.maxAllowed) && (item.value.length > formjsonitem.maxAllowedValue)) {
                                toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : maximum ${formjsonitem.maxAllowedValue} options are allowed `)
                            }

                            break;
                        case "Date":
                            if (formjsonitem.required && !item.value) {
                                toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : This field is required.`);
                            }

                            if (item.value && !isNaN(new Date(item.value).getTime()) && !(item.showTimeSelect && item.showTimeSelectOnly)) {

                                const itemDate = new Date(item.value);
                                let today = new Date();
                                today.setHours(0, 0, 0, 0)
                                let futuretoday = new Date()
                                futuretoday.setHours(23, 59, 59, 999)
                                console.log("today and itemdate", today, itemDate)
                                if (formjsonitem.restrictPastdays && itemDate < today) {
                                    toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : Past dates are not allowed.`);
                                    return;
                                }
                                if (formjsonitem.restrictFuturedays && itemDate > futuretoday) {
                                    toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : Future dates are not allowed.`);
                                    return;
                                }

                                if (formjsonitem.dateRange) {
                                    const minDate = formjsonitem.minDate
                                        ? new Date(convertDateFormat(formjsonitem.minDate))
                                        : null; // Min date
                                    let maxDate = formjsonitem.maxDate
                                        ? new Date(convertDateFormat(formjsonitem.maxDate))
                                        : null; // Max date
                                    if (maxDate) {
                                        maxDate.setHours(23, 59, 59, 999); // Sets the time to 23:59:59.999
                                    }
                                    // Validate only if date is provided
                                    // if (itemDate && !isNaN(itemDate.getTime())) { // Checks if date is valid    

                                    if ((minDate && maxDate) && (itemDate < minDate || itemDate > maxDate)) {
                                        console.log("mindate", minDate.toISOString().split("T")[0])
                                        console.log("mindate", minDate)
                                        console.log("maxdate", maxDate.toISOString().split("T")[0])
                                        console.log("onlyitemDate", itemDate)
                                        toasterrors.push(
                                            `${stripHtmlTags(formjsonitem.label)} : Please select a date between ${convertDateFormat(formjsonitem.minDate)} and ${convertDateFormat(formjsonitem.maxDate)}`

                                        );
                                        return;
                                    } else if (minDate && itemDate < minDate) {
                                        toasterrors.push(
                                            `${stripHtmlTags(formjsonitem.label)} : Date should be greater than or equal to ${convertDateFormat(formjsonitem.minDate)}`
                                        );
                                        return;
                                    } else if (maxDate && itemDate > maxDate) {
                                        toasterrors.push(
                                            `${stripHtmlTags(formjsonitem.label)} : Date should be less than or equal to ${convertDateFormat(formjsonitem.maxDate)}`
                                        );
                                        return;
                                    }
                                }
                                if (formjsonitem.allowOnlyCurrentYearDates && itemDate.getFullYear() !== today.getFullYear()) {
                                    toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : Please select a date within the current year (${today.getFullYear()}).`);

                                }
                                if (formjsonitem.weekdaySelectionOnly) {
                                    const dayOfWeek = itemDate.getDay(); // Get the day of the week (0 = Sunday, 6 = Saturday)
                                    if (dayOfWeek === 0 || dayOfWeek === 6) {
                                        toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : Please select a weekday (Monday to Friday).`);
                                    }
                                }


                            }
                            // else if (isNaN(new Date(item.value).getTime())) {
                            //   // Invalid date format
                            //   toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : Invalid date format`);
                            //   return;
                            // }
                            break;
                        default:
                            if (formjsonitem.required && !item.value?.length) {
                                toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : This field is required.`);
                            }
                            break;
                    }
                }
            });




            // if (!Categorytype) {
            //     toasterrors.push("Request For Field is required");
            // }
            // else if (Categorytype.length > 80) {
            //     toasterrors.push("Request For : Only 80 characters are allowed ");
            // }

            console.log("taabledata", tables)
            tables.forEach((table: any) => {
                const attachcolumnName = table.columns.find((column: any) => column.columntype === "Attachment")?.headername;
                table.rows.forEach((row: any) => { row[attachcolumnName]?.fileObject?.size > 1024 * 1024 ? toasterrors.push(`${table.content}: File size exceeded , max-allowed size is 1mb`) : null })
            }
            )


            if (toasterrors.length > 0) {
                toasterrors.forEach((error) => {
                    setLoading(false)
                    toast.current?.show({ severity: "error", summary: "", detail: error, life: 3000 });
                });

                setIsButtonDisabled(true);
                setTimeout(() => {
                    setIsButtonDisabled(false);
                }, 3000);

                return;
            }


            const returned = await sp.web.lists.getByTitle("WorkFlowProcessData").items.getById(Number(id)).get();
            const returnedStatus = returned.Status;
            const ZeroDestinationQueue = returned.DestinationQueue;
            const buttonAction = "Submit";
            const CorrectAppCode = IsParentForm ? record.CurrentAppCode : appCode

            setLoading(true);
            console.log("Routing ru;es app code: ", childselected);
            const routingRulesItems = await sp.web.lists.getByTitle("RoutingRules").items.filter(`(AppCode eq '${CorrectAppCode}') and (Actions eq '${buttonAction}')`).get();
            console.log("RoutingRules Items in updateform : ", routingRulesItems);
            // debugger;

            const { CurrentQueue, DestinationQueue, Status } = routingRulesItems[0];

            if (routingRulesItems.length === 0) {
                setLoading(false)
                alert("Reviewers or Approvers are not available for further actions");
                return;
            }

            const mappingMasterItem = await sp.web.lists.getByTitle("MappingMaster").items
                .select("*", "Users_x002f_Groups/Id", "Users_x002f_Groups/Title")
                .expand("Users_x002f_Groups")
                .filter(`(Level eq '${DestinationQueue}') and (AppCode eq '${CorrectAppCode}')`)
                .get();


            let conditionlistItem: any[] = [];
            let result: boolean = false;

            try {
                // Fetching condition list items
                conditionlistItem = await sp.web.lists
                    .getByTitle("ConditionsList")
                    .items.select("*", "PersonOrGroup/Id", "PersonOrGroup/Title")
                    .expand("PersonOrGroup")
                    .filter(`(Level eq '${mappingMasterItem[0].Level}') and (AppCode eq '${CorrectAppCode}')`)
                    .get();

                console.log("-------conditionlistItem---------", conditionlistItem);

                // Ensure there is data to process
                if (conditionlistItem.length === 0) {
                    throw new Error("No data found in ConditionsList for the provided filter.");
                }

                // Parse the ConditionJson and process conditions
                const conditionlistItemJson: any = JSON.parse(conditionlistItem[0].ConditionJson);

                const mainoperator: string = conditionlistItemJson.operator;
                result = mainoperator === "And" ? true : false;

                conditionlistItemJson.conditions.forEach((element: any) => {
                    try {
                        const fieldId = ChildFormJsons.find((val: any) => stripHtmlTags(val.label) === element.field)?.id;

                        if (!fieldId) {
                            throw new Error(`Field ID not found for field: ${element.field}`);
                        }

                        const fieldvalue = childformData.find((d: any) => d.id === fieldId)?.value;

                        if (fieldvalue === undefined) {
                            throw new Error(`Field value not found for field ID: ${fieldId}`);
                        }

                        const boolvalue = ConditionValidator(fieldvalue, element.operator, element.value, element.elementType);

                        result = mainoperator === "And" ? boolvalue && result : boolvalue || result;

                        console.log("-------result inside--", result);
                    } catch (innerError: any) {
                        console.error("Error processing condition element:", innerError.message);
                    }
                });

                console.log("Final result:", result);
            } catch (error: any) {
                console.error("Error occurred in processing conditions:", error.message);
            }

            // Determine the userGroupId based on conditions
            let userGroupId: number | null = null;

            if (result && conditionlistItem.length > 0) {
                userGroupId = conditionlistItem[0]?.PersonOrGroup?.Id || null;
            } else {
                userGroupId = mappingMasterItem?.[0]?.Users_x002f_Groups?.Id || null;
            }

            console.log("User Group ID--------------------: ", userGroupId);



            if (userGroupId) {
                console.log("User Group ID: ", userGroupId);

            } else {
                console.log("User Group ID is null, fetching manager from AD...");

                try {
                    const managerId = await getManagerFromAD();

                    console.log("mmmmmmmmmmmmmmmmmmmmmmmmmId", managerId);
                    // debugger;
                    if (managerId) {
                        console.log("Immediate Manager: ", managerId);


                        userGroupId = managerId; // Retrieve the User ID
                        console.log("Author/User ID-----------------: ", userGroupId);
                        // Proceed with logic using userGroupId

                        console.error("No user found for the given email.");

                    } else {
                        console.log("No manager found or an error occurred.");
                    }
                } catch (error) {
                    console.error("Error fetching manager or user details: ", error);
                }
            }



            const MMLevel = mappingMasterItem[0].Level;
            // let [ParentAnswers, ChildAnswers] = AnswerJsonSeparator(ParentFormJSONCopy, updatedFormData);
            // console.log(" from commentvalidation value ParentAnswers , ChildAnswers ", ParentAnswers, ChildAnswers);


            //-------------------------------------------------------------------------------------------------------
            if ((returnedStatus === "Returned") && (ZeroDestinationQueue === "0")) {

                let commentValidationValue = "false"

                if (ResubmitButton && comment.length > 1) {
                    commentValidationValue = "true"
                    try {
                        const rr = await sp.web.lists.getByTitle("CommentsLog").items.add({
                            AppCode: IsParentForm ? record.CurrentAppCode : appCode,
                            Comments: comment,
                            CommentedBy: userName,
                            CommentedByEmail: userEmail,
                            SeqNo: appSeqNo,
                        });

                        // Access the item ID from the result
                        const itemId = rr.data.Id;

                        // If a file is selected, add it to the list item
                        if (selectedFile.length > 0) {
                            for (const file of selectedFile) {
                                await sp.web.lists
                                    .getByTitle("CommentsLog")
                                    .items.getById(itemId)
                                    .attachmentFiles.add(file.name, file);
                            }
                            console.log("Files successfully uploaded!");
                        } else {
                            console.log("No file selected, skipping file upload.");
                        }

                    }
                    catch (err) {
                        console.log("error", err)
                    }
                }


                if (commentValidationValue === "true") {
                    const updatedFormData = data.map((item: any) => {
                        if (item.name && item.name.includes("file_upload") && item.value && item.value.name) {
                            return {
                                ...item,
                                value: item.value.name,
                            };
                        }
                        return item;
                    });

                    // Function to assign missing file names from attachmentFiles
                    const assignMissingFiles = (updatedFormData: any[], attachmentFiles: any[]) => {
                        let attachmentIndex = 0;

                        return updatedFormData.map((item: any) => {
                            if (
                                item.name.includes("file_upload") &&
                                (item.value === null || item.value === null) &&
                                attachmentIndex < attachmentFiles.length
                            ) {
                                return {
                                    ...item,
                                    value: attachmentFiles[attachmentIndex++].name.split('. ')[1], // Extract actual filename
                                };
                            }
                            return item;
                        });
                    };

                    // Assign missing file names dynamically
                    const finalUpdatedFormData = assignMissingFiles(updatedFormData, attachmentFiles);
                    console.log("Final Updated Form Data with assigned files:", finalUpdatedFormData);
                    debugger;
                    try {
                        await sp.web.lists.getByTitle("WorkFlowProcessData").items.getById(Number(id)).update({
                            // FormData: JSON.stringify(ParentAnswers),
                            CurrentQueue: CurrentQueue,
                            DestinationQueue: DestinationQueue,
                            Status: Status,
                            CurApproverId: userGroupId,
                            Level: MMLevel,
                            // CategoryType: Categorytype,
                            // NotifyRequestorId: {
                            //     results: notifytorequestorID
                            // },
                            // TableJSON: JSON.stringify(tables),
                            // TimeSheetJSON: JSON.stringify(timesheettables),
                            // MinDateArray: JSON.stringify(timesheettables.map((rec) => rec.fromdate)),
                            // MaxDateArray: JSON.stringify(timesheettables.map((rec) => rec.todate))

                        });

                        debugger;
                        if (IsParentForm) {
                            const childseqNo = (appSeqNo) ? appSeqNo : await getCurrentUserSeqNo();

                            const childid = ChildtransactionItems.filter((item: any) => item.AppCode === record.CurrentAppCode)[0].Id
                            console.log("Child ids :", childid);
                            // setChildids(childid);
                            await sp.web.lists.getByTitle("WorkFlowProcessData").items.getById(Number(childid)).update({
                                FormData: JSON.stringify(finalUpdatedFormData),
                                CurrentQueue: CurrentQueue,
                                DestinationQueue: DestinationQueue,
                                Status: Status,
                                CurApproverId: userGroupId,
                                Level: MMLevel,
                                // CategoryType: Categorytype,
                                // NotifyRequestorId: {
                                //     results: notifytorequestorID
                                // },
                                TableJSON: JSON.stringify(tables),
                                TimeSheetJSON: JSON.stringify(timesheettables),
                                MinDateArray: JSON.stringify(timesheettables.map((rec) => rec.fromdate)),
                                MaxDateArray: JSON.stringify(timesheettables.map((rec) => rec.todate))
                            });
                            const folderName = `NoteAttach/${childseqNo}`;


                            try {

                                // Check if folder exists
                                await sp.web.getFolderByServerRelativeUrl(folderName).get();
                            } catch (error) {
                                // If folder does not exist, create it
                                await sp.web.folders.add(folderName);
                            }

                            // Upload files to the subfolder
                            for (const file of files) {
                                await sp.web.getFolderByServerRelativeUrl(folderName).files.add(file.name, file.file, true);
                            }


                        }
                        await sp.web.lists.getByTitle("WorkFlowLog").items.add({
                            AppCode: IsParentForm ? record.CurrentAppCode : appCode,
                            Status: Status,
                            Domain: domainfield,
                            Level: MMLevel,
                            SeqNo: appSeqNo,
                            Role: routingRulesItems[0].Role,
                            Action: "Re-Submit",
                        })
                    }
                    catch (error) {
                        console.error("error in updating filled data to list", error)
                    }


                    await deleteExistingFilesAndFolders();
                    tables.forEach(async (table: any) => {
                        const attachcolumnName = table.columns.find((column: any) => column.columntype === "Attachment")?.headername
                        const boolvalue: boolean = table.rows.find((row: any) => row[attachcolumnName].name.length > 0)
                        if (attachcolumnName && boolvalue) {
                            const tablefoldername = `TableAttachments/${appSeqNo}`;
                            await sp.web.folders.add(tablefoldername);
                            const subfolderName = `${tablefoldername}/${table.id}`;
                            await sp.web.folders.add(subfolderName);

                            for (const file of table.rows) {
                                console.log("----file=====", file)
                                const fileData = file[attachcolumnName];

                                if (fileData && fileData.name && fileData.fileObject) {
                                    const uploadResult = await sp.web.getFolderByServerRelativeUrl(subfolderName).files.add(file[attachcolumnName].name, file[attachcolumnName].fileObject, true);
                                    await uploadResult.file.listItemAllFields.select('Id').get()
                                        .then(async (item) => {
                                            await sp.web.lists.getByTitle("TableAttachments")
                                                .items.getById(item.Id)
                                                .update({
                                                    RowID: file.id
                                                });
                                        });
                                }
                            }
                        }


                    })
                    // ------------------------------------------Saving table attachments to library-----------------------------------------------

                    const folderName = `NoteAttach/${appSeqNo}`;


                    try {

                        // Check if folder exists
                        await sp.web.getFolderByServerRelativeUrl(folderName).get();
                    } catch (error) {
                        // If folder does not exist, create it
                        await sp.web.folders.add(folderName);
                    }

                    // Upload files to the subfolder
                    for (const file of files) {
                        await sp.web.getFolderByServerRelativeUrl(folderName).files.add(file.name, file.file, true);
                    }

                    toast.current?.show({ severity: 'success', summary: '', detail: 'Details re-submitted successfully ', life: 2000 });

                    setTimeout(() => {
                        navigate(-1);
                    }, 1000)
                }

                else {
                    setLoading(false)
                    toast.current?.show({ severity: 'error', summary: '', detail: 'Add a comment before submitting !', life: 2000 });
                    setHighlight(true)
                }

            }



            else {
                // const allItems = await sp.web.lists
                //     .getByTitle("WorkFlowProcessData")
                //     .items.orderBy("Created", false) // Ensure newest records come first
                //     .filter(`AppCode eq '${appCode}' and Domain eq '${domainfield}' `)
                //     .getAll()
                //     .catch((error) => {
                //         console.error("Error fetching items:", error);
                //         return [];
                //     });

                // const formMasterItem2 = await sp.web.lists.getByTitle("FormMaster")
                //     .items.select("FormPrefix,isWorkflowRequired")
                //     .filter(`AppName eq '${appName}' `).get();

                // let isNoWorkflowRequiredFlag = formMasterItem2[0]?.isWorkflowRequired === "No" ? true : false;

                // let NoApprovalStatus = Status;
                // debugger;
                // if (isNoWorkflowRequiredFlag) {
                //     NoApprovalStatus = "Submitted";
                // }
                // else {
                //     NoApprovalStatus = Status
                // }

                // let latestNumber = 0;
                //Rrmoved this because we are not creating new Title for the Other child forms. 
                // console.log("allItems------------", allItems)
                // for (const item of allItems) {
                //     if (item.Status !== "Draft") {
                //         const titleParts = item.Title.split("/");

                //         console.log("titleParts----------------- !!!!!!!!!!!!!!", titleParts, titleParts.length);

                //         if (titleParts.length >= 3) {
                //             const numberPart = parseInt(titleParts[2], 10);
                //             console.log("numberPart", numberPart);

                //             if (!isNaN(numberPart) && numberPart > latestNumber) {
                //                 latestNumber = numberPart; // Store the highest number found
                //             }
                //         }
                //     }
                // }
                // const seqNo = await getCurrentUserSeqNo();
                // const date = new Date();
                // const month = date.toLocaleString('default', { month: 'short' });
                // const year = date.getFullYear().toString().slice(-2);
                // const monthYear = `${month}_${year}`;
                const updatedApprovers = [...AllApprovers, userEmail]

                const newTitle = record.Title;
                // (ChildtransactionItems.length===1)?`${formMasterItem2[0].FormPrefix}/${monthYear}/${latestNumber + 1}`:record.Title;
                debugger;

                console.log("ddaattaa", newTitle)

                debugger;
                try {
                    await sp.web.lists.getByTitle("WorkFlowProcessData").items.getById(Number(id)).update({
                        // Title: newTitle,
                        // FormData: JSON.stringify(ParentAnswers),
                        CurrentQueue: CurrentQueue,
                        DestinationQueue: DestinationQueue,
                        Status: Status,
                        CurApproverId: userGroupId,
                        Level: MMLevel,
                        ManagersId: userGroupId,
                        AllApprovers: JSON.stringify(updatedApprovers),
                        // CategoryType: Categorytype,
                        // CurrentAppCode: IsParentForm ? record.CurrentAppCode : "",
                        // NotifyRequestorId: {
                        //     results: notifytorequestorID
                        // },
                        // TableJSON: JSON.stringify(tables),
                        // TimeSheetJSON: JSON.stringify(timesheettables),
                        // MinDateArray: JSON.stringify(timesheettables.map((rec) => rec.fromdate)),
                        // MaxDateArray: JSON.stringify(timesheettables.map((rec) => rec.todate))
                    });
                    const aduserdetails = await fetchToken();
                    // debugger;
                    // console.log("aduserdetails--------------> ",aduserdetails)
                    // console.log(aduserdetails.employeeId)
                    let employeeId = "---";
                    let jobTitle = "---";
                    let department = "---";
                    let mobilePhone = "---"
                    // debugger;
                    if (aduserdetails) {
                        console.log("aduserdetails--------------> ", aduserdetails);
                        employeeId = aduserdetails.employeeId || "---";
                        jobTitle = aduserdetails.jobTitle || "---";
                        department = aduserdetails.department || "---";
                        mobilePhone = aduserdetails.mobilePhone || "---";
                    } else {
                        // If fetchToken fails, log the error and continue with NA values
                        console.error("Failed to fetch Auther user details from AD");
                    }

                    // Create the object to store in the ADuserData column
                    const adUserData = {
                        RequesterName: RequesterTitle,
                        Email: RequesterEmail,
                        EmployeeID: employeeId,
                        Designation: jobTitle,
                        Department: department,
                        ContactNo: mobilePhone
                    };


                    const adUserDataJSON = JSON.stringify(adUserData);
                    const childseqNo = await getCurrentUserSeqNo();
                    const ChildAppCode = record.CurrentAppCode;
                    const ChildAppName = ChildFormDetails.find(ele => ele.AppCode === ChildAppCode).AppName;
                    const childupdatedChildApprovers = [userEmail];


                    // console.log("formmasterItem2", formMasterItem2);
                    await sp.web.lists.getByTitle("WorkFlowProcessData").items.add({
                        Title: newTitle,
                        FormData: JSON.stringify(updatedFormData),
                        AppName: ChildAppName,
                        AppCode: ChildAppCode,
                        ParentAppCode: appCode,
                        Domain: domainfield,
                        Status: Status,
                        SeqNo: childseqNo,
                        ParentAppId: Number(id),
                        CurApproverId: userGroupId,
                        AuthorADdata: adUserDataJSON,
                        CurrentQueue: CurrentQueue,
                        DestinationQueue: DestinationQueue,
                        Level: mappingMasterItem[0].Level,
                        ManagersId: userGroupId,
                        CategoryType: Categorytype,
                        AllApprovers: JSON.stringify(childupdatedChildApprovers),
                        NotifyRequestorId: {
                            results: notifytorequestorID
                        },
                        TableJSON: JSON.stringify(tables),
                        TimeSheetJSON: JSON.stringify(timesheettables),
                        MinDateArray: JSON.stringify(timesheettables.map((rec) => rec.fromdate)),
                        MaxDateArray: JSON.stringify(timesheettables.map((rec) => rec.todate))
                    });


                    let AppCodeee = appCode;
                    let ParentAppCodeee = "";
                    let Sequence = appSeqNo;
                    if (IsParentForm) {
                        AppCodeee = record.CurrentAppCode;
                        ParentAppCodeee = appCode;
                        Sequence = childseqNo;
                    }

                    await sp.web.lists.getByTitle("WorkFlowLog").items.add({
                        AppCode: AppCodeee,
                        Status: Status,
                        Domain: domainfield,
                        Level: MMLevel,
                        SeqNo: Sequence,
                        ParentAppCode: ParentAppCodeee,
                        Role: routingRulesItems[0].Role,
                        Action: buttonAction,
                    })

                    const folderName = `NoteAttach/${childseqNo}`;


                    try {

                        // Check if folder exists
                        await sp.web.getFolderByServerRelativeUrl(folderName).get();
                    } catch (error) {
                        // If folder does not exist, create it
                        await sp.web.folders.add(folderName);
                    }
                    // Upload files to the subfolder
                    for (const file of files) {
                        await sp.web.getFolderByServerRelativeUrl(folderName).files.add(file.name, file.file, true);
                    }



                } catch (error) {
                    console.error("Submit case patching to list failed due to : ", error)
                }

                await deleteExistingFilesAndFolders();
                // ------------------------------------------Saving table attachments to library-----------------------------------------------
                tables.forEach(async (table: any) => {
                    const attachcolumnName = table.columns.find((column: any) => column.columntype === "Attachment")?.headername
                    const boolvalue: boolean = table.rows.find((row: any) => row[attachcolumnName].name.length > 0)
                    if (attachcolumnName && boolvalue) {
                        const tablefoldername = `TableAttachments/${appSeqNo}`;
                        await sp.web.folders.add(tablefoldername);
                        const subfolderName = `${tablefoldername}/${table.id}`;
                        await sp.web.folders.add(subfolderName);

                        for (const file of table.rows) {
                            console.log("----file=====", file)
                            const fileData = file[attachcolumnName];

                            if (fileData && fileData.name && fileData.fileObject) {
                                const uploadResult = await sp.web.getFolderByServerRelativeUrl(subfolderName).files.add(file[attachcolumnName].name, file[attachcolumnName].fileObject, true);
                                await uploadResult.file.listItemAllFields.select('Id').get()
                                    .then(async (item) => {
                                        await sp.web.lists.getByTitle("TableAttachments")
                                            .items.getById(item.Id)
                                            .update({
                                                RowID: file.id
                                            });
                                    });
                            }
                        }
                    }


                })
                // ------------------------------------------Saving table attachments to library-----------------------------------------------

                // const itemId = newItem.data.Id;
                // const folderName = `NoteAttach/${appSeqNo}`;


                // try {
                //     // Check if folder exists
                //     await sp.web.getFolderByServerRelativeUrl(folderName).get();
                // } catch (error) {
                //     // If folder does not exist, create it
                //     await sp.web.folders.add(folderName);
                // }

                // // Upload files to the subfolder
                // for (const file of files) {
                //     await sp.web.getFolderByServerRelativeUrl(folderName).files.add(file.name, file.file, true);
                // }

                toast.current?.show({ severity: 'success', summary: '', detail: 'Data submitted successfully!', life: 2000 });
                navigate(-1);
            }

        } catch (error) {
            console.error("Error updating form data submit case:", error);
        } finally {
            setLoading(false);
        }


    };



    const handleDocumentDelete = async (fileName: string, fileId: string) => {
        // debugger
        console.log("fileId--------------- nin delete -------- ", fileId, mappedfile);

        // Clean file name to remove leading numbers
        const cleanFileName = fileName.replace(/^\d+\.\s*/, "");

        // Encode filename to handle special characters
        const encodedFileName = encodeURIComponent(cleanFileName);

        // Ensure correct relative path (must start with `/sites/...`)
        let fileRelativePath = `/${fileFetchUrl}/NoteAttach/${fileFolder}/${encodedFileName}`;


        try {
            // Attempt to delete the file from SharePoint
            await sp.web.getFileByServerRelativeUrl(fileRelativePath).delete();

            // Fetch the list item from SharePoint
            const transactionItem = await sp.web.lists
                .getByTitle("WorkFlowProcessData")
                .items.getById(Number(id))
                .select("*", "Title, Created, Author/Title", "FormData")
                .expand("Author")
                .get();



            // Parse and update FormData
            const parsedFormData = JSON.parse(transactionItem.FormData);
            const updatedFormData = parsedFormData.map((item: any) => {
                if (item.name.startsWith("file_upload")) {
                    return {
                        ...item,
                        value: item.value === cleanFileName ? null : item.value, // Remove deleted file
                    };
                }
                return item;
            });


            // Update the list item with modified FormData
            await sp.web.lists.getByTitle("WorkFlowProcessData").items.getById(Number(id)).update({
                FormData: JSON.stringify(updatedFormData),
            });

            // Remove the file from UI state
            setAttachmentFiles(prevFiles => prevFiles.filter(file => file.name !== fileName));

            toast.current?.show({ severity: 'success', summary: '', detail: 'File deleted successfully!', life: 2000 });

            // Refresh form data
            void fetchFormData();
            setSelectTab("AllDetails");

        } catch (deleteError) {
            console.error("Error deleting file:", deleteError);
            toast.current?.show({ severity: 'error', summary: '', detail: `Error deleting file`, life: 2000 });
        }
    };

    const handleCancel = () => {
        navigate(-1);
    };
    const handleBlur = (data: any) => {
        if (datefieldids.length > 0) {
            datefieldids.forEach(dateElement => {
                let datefieldonchange = document.querySelector(`input[name="${dateElement.field_name}"]`) as HTMLInputElement;
                if (datefieldonchange) {
                    datefieldonchange.addEventListener("change", (e: any) => {
                        let dateJson = data.find((element: any) => { return element.name === dateElement.field_name })
                        dateJson.value = e.target.value
                        console.log("sssssssssssss", e.target.value);
                        // data[5].value=e.target.value
                    })
                    datefieldonchange.dispatchEvent(new Event('change'))
                }

            });

            const updatedData = data.map((item: any) => {
                if (item.name.startsWith("file_upload")) {
                    return {
                        ...item,
                        value: fileData[item.name] || null, // Keep the file if it exists in state

                    };
                }
                return item;
            });

            setChildFormData(updatedData);
            // setParentFormAnswerCopy(updatedData);
        }
    }

    //     const handleFileChange = (event: Event, fieldName: string, fileId: string) => {
    //         const target = event.target as HTMLInputElement | null;
    //         console.log("targetfiles",target)
    //         if (target && target.files && target.files.length > 0) {
    //             // const file = target.files[0];
    //             const files = Array.from(target.files);
    // console.log("fromhandlefilechangefiles",files)


    //             // Update file state immediately
    //             setFileData((prev) => ({
    //                 ...prev,
    //                 [fieldName]: files,
    //             }));
    // console.log("inhandlefilechangefiledata",fileData)
    //             // Update form data immediately
    //             setFormData((prevData) =>
    //                 prevData.map((el) =>
    //                     el.name === fieldName ? { ...el, value: files } : el
    //                 )
    //             );
    //         }
    //     };

    // useEffect(() => {
    //     // Attach file input change events once the component mounts
    //     formData.forEach((item) => {
    //         if (item.name.startsWith("file_upload")) {
    //             const fileInputElement = document.querySelector(
    //                 `input[name="${item.name}"]`
    //             ) as HTMLInputElement | null;

    //             if (fileInputElement) {
    //                 console.log("uploadelements",fileInputElement)
    //                 fileInputElement.onchange = null;
    //                 // fileInputElement.onchange = (event) => {
    //                 //     handleFileChange(event, item.name, item.id);

    //                 // };
    //                 fileInputElement.addEventListener("change",(event) => handleFileChange(event, item.name, item.id))

    //             }
    //         }
    //     });
    // }, [formData]);



    const handleChange = (data: any) => {
        // debugger

        if (!data || !Array.isArray(data)) {
            console.error("handleChange received invalid data:", data);
            return;
        }

        funforVisibility(data)
        console.log("Event change in form -----------", data);
        // funforVisibility(data)
        // Preserve existing file values
        const updatedData = data.map((item: any) => {
            if (item.name.startsWith("file_upload")) {


                return {
                    ...item,
                    value: fileData[item.name] !== undefined ? fileData[item.name] : item.value,
                };
            }
            return item;
        });
        console.log("updatedData in FormComponent ------------", updatedData)
        // setFormData(updatedData);
        console.log("Updated dataaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa", updatedData)
        setChildFormData(updatedData);

        // setChildFormJsons(updatedData)
        // setParentFormAnswerCopy(updatedData);

        // Attach file change event listeners for file inputs
        // data.forEach((item: any) => {
        //     if (item.name.startsWith("file_upload")) {
        //         const fileInputElement = document.querySelector(
        //             `input[name="${item.name}"]`
        //         ) as HTMLInputElement | null;

        //         if (fileInputElement) {


        //             fileInputElement.onchange = (event) => {
        //                 const target = event.target as HTMLInputElement | null;
        //                 if (target && target.files && target.files.length > 0) {
        //                     // const file = target.files[0];
        //                     const files = Array.from(target.files);
        //                     console.log("File uploaded: ", files);

        //                     // Update file state
        //                     setFileData((prev) => ({
        //                         ...prev,
        //                         [item.name]: files,
        //                     }));

        //                     // Update formData
        //                     setFormData((prevData: any) =>
        //                         prevData.map((el: any) =>
        //                             el.name === item.name ? { ...el, value: files } : el
        //                         )
        //                     );
        //                 }
        //             };
        //         }
        //     }
        // });
    };


    const handleCommnetSaveOnly = async () => {
        let commentValidationValue = "false"
        if (comment.length < 1) {
            // alert("plese add comment")
        }
        else {
            try {


                const rr = await sp.web.lists.getByTitle("CommentsLog").items.add({
                    AppCode: AppCodevalue,
                    Comments: comment,
                    CommentedBy: userName,
                    CommentedByEmail: userEmail,
                    SeqNo: appSeqNo
                });

                const itemId = rr.data.Id;

                if (selectedFile.length > 0) {
                    for (const file of selectedFile) {
                        await sp.web.lists
                            .getByTitle("CommentsLog")
                            .items.getById(itemId)
                            .attachmentFiles.add(file.name, file);
                    }
                    console.log("Files successfully uploaded!");
                } else {
                    console.log("No file selected, skipping file upload.");
                }


                commentValidationValue = "true"
                setComment("")
                setSelectedFile([])
            }
            catch (err) {
                console.log("error", err)
            }
        }
        if (commentValidationValue === "true") {
            toast.current?.show({ severity: 'success', summary: '', detail: 'Comment added successfully', life: 2000 });
        } else {
            toast.current?.show({
                severity: 'warn',
                summary: '',
                detail: 'Please add your comment before saving',
                life: 2000
            });
            setHighlight(true);
            setSelectTab('AllDetails')
        }

    }

    const handleChatResponse = (severity: any, message: any) => {
        toast.current?.show({ severity, detail: message, life: 1000 });
    };

    // const itemId = rr.data.Id;
    const handlereqCancel = async () => {

        let messages = comment.trim().length == 0 ? 'comment is required' : `"comment can't be single character`;
        if (comment.trim().length < 2) { console.log("invalid Comment"); toast.current?.show({ severity: 'error', detail: messages, life: 1000 }); }
        else {
            try {
                await sp.web.lists
                    .getByTitle("WorkFlowProcessData")
                    .items.getById(Number(id)).update(
                        { Status: 'Cancelled', CurApproverId: null, DestinationQueue: '' });

                const workflowlog = await sp.web.lists.getByTitle("WorkFlowLog").items.add({
                    AppCode: appCode,
                    Status: 'Cancelled',
                    Domain: domainfield,
                    Level: -1,
                    SeqNo: appSeqNo,
                    Role: "Creator",
                    Action: "Cancel",
                })
                console.log("------------------workflowlog", workflowlog)

                toast.current?.show({ severity: 'success', detail: `Request cancelled successfully`, life: 1000 });
                console.log("request canceled ('-_-')");
                setOpenpop(false);
                await handleCommnetSaveOnly();

                setTimeout(() => {
                    navigate(-1);
                }, 1000);

            }
            catch (err) {
                console.log(err, "Error, couldnot Cancel the request (!_!)");
                toast.current?.show({ severity: 'info', detail: "Couldn't cancel the request", life: 1000 });
            }
        }
    }
    const handleDeleteRequest = async () => {
        try {
            console.log('Filename', appSeqNo)
            await sp.web.lists.getByTitle("WorkFlowProcessData").items.getById(Number(id)).delete();
            toast.current?.show({
                severity: 'success',
                summary: '',
                detail: `Deleted Successfully`,
                life: 3000,
            });
            let fileRelativePath = `/${fileFetchUrl}/NoteAttach/${appSeqNo}`;
            try { await sp.web.getFolderByServerRelativeUrl(fileRelativePath).delete(); }
            catch { console.log(`${appSeqNo} Folder not found ('O_Q')`) }
            setOpenpop(false);
            console.log(`${id} Record, ${appSeqNo} Folder with attachments is Deleted (x_x)`);

            setTimeout(() => {
                navigate(-1);
            }, 1000);
        }

        catch {
            console.log("Error, couldnot Delete the request (!_!)")
            toast.current?.show({ severity: 'info', detail: 'Request cnnot be  deleted...', life: 3000 });
        }
    }


    const handlecancelchange = (e: string) => {
        console.log("clicked button is ----------", e)
        if (e === "cancel") { setOpencanpop(true) }
        setOpenpop(true)
    }
    const handleReopen = async () => {
        try {
            if (comment.trim().length > 0) {
                await handleUpdateForm(formData)
                await sp.web.lists.getByTitle("CommentsLog").items.add({
                    AppCode: AppCodevalue,
                    Comments: comment,
                    CommentedBy: userName,
                    CommentedByEmail: userEmail,
                    SeqNo: appSeqNo
                });

                setisreopen(false)
                toast?.current?.clear();
                toast?.current?.show({
                    severity: "success",
                    detail: "Request re-opened successfully.",
                    life: 2000
                })
            } else {
                toast?.current?.clear();
                toast?.current?.show({
                    severity: 'warn',
                    detail: 'Please enter the reason to Reopen',
                    life: 2000
                })
            }
        } catch (err) {
            toast?.current?.clear();
            toast?.current?.show({
                severity: "error",
                detail: "Request re-open failed.",
                life: 2000
            })
            console.error("error in reopening request", err)
        }
    }
    const handleUserDomain = (UserDomain: string) => {
        setIsSameDomain(UserDomain)

    };

    //  -------------------------------------------- DropDown for Extenal user ---------------------------------

    // const handleInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    //     const value = e.target.value;
    //     setSearchText(value);

    //     if (value.trim() === '') {
    //         setFilteredUsers(myOrganizationUsers.slice(0, 10));
    //     } else {
    //         const filtered = myOrganizationUsers
    //             .filter((user) =>
    //                 user.displayName.toLowerCase().includes(value.toLowerCase())
    //             )
    //             .slice(0, 10);
    //         setFilteredUsers(filtered);
    //     }

    //     setShowDropdown(true);
    // };

    // const handleUserSelect = async (user: User) => {
    //     // Avoid duplicates and enforce max 5 users
    //     const alreadySelected = selectedUsers.find(u => u.mail === user.mail);
    //     if (alreadySelected || selectedUsers.length >= 5) return;

    //     try {
    //         await sp.web.ensureUser(user.mail);

    //         const result = await sp.web.siteUsers
    //             .filter(`Email eq '${user.mail}'`)
    //             .select('Id,Email')
    //             .get();

    //         const userId = result[0]?.Id;
    //         if (userId) {
    //             setSelectedUsers(prev => [...prev, user]);
    //             SetnotifytorequestorID(prev => [...prev, userId]);
    //         }
    //     } catch (error) {
    //         console.error("Error selecting user:", error);
    //     }

    //     setSearchText('');
    //     setFilteredUsers([]);
    //     setShowDropdown(false);
    // };

    // const removeSelectedUser = (userMail: string) => {
    //     const updatedUsers = selectedUsers.filter(u => u.mail !== userMail);
    //     setSelectedUsers(updatedUsers);

    //     SetnotifytorequestorID(prev =>
    //         prev.filter((_, idx) => selectedUsers[idx].mail !== userMail)
    //     );
    // };
    // Function to update FlowJSON, Requester Details, and Comments when child tab changes
    const updateChildTabData = async (childFormAppCode: string, childOrder: number | string, ChildFormJson: any, matchedChildData?: any) => {
        try {
            // console.log(`📊 Updating child tab data for: ${childFormAppCode}, ChildOrder: ${childOrder}`);
            const childTransaction = ChildtransactionItems[Number(childOrder) - 1];
            if (!childTransaction || matchedChildData?.Status === "Returned") {
                debugger;
                // const parsedFormJSON = JSON.parse(ChildFormJson);
                const formdatefields: any[] = ChildFormJson.filter((ele: any) => { return ele.field_name?.startsWith("date_picker") });
                console.log("selected child Formdatefields", formdatefields)
                setdatefieldids(formdatefields);
                const selectedFormJSON = ChildFormJson
                const dependentFields = selectedFormJSON.filter((ele: any) => { //All child fields having conditions
                    return ele?.dependentVisibility
                })
                setSelectedFormJSON(selectedFormJSON)
                SetChildFieldJsons(dependentFields)
                const dependentShowFields = selectedFormJSON.filter((ele: any) => { //Child fields which should be hidden initially
                    return ele?.dependentVisibility && ele?.controlVisibility
                })
                console.log("dependentFields", dependentFields)
                console.log("dependentShowFields", dependentShowFields)
                const SourceFieldNames = [...new Set(dependentFields.map((ele: any) => ele?.VisibilityCondition[0]?.SourceField))]
                console.log("SourceFieldNames", SourceFieldNames)

                const SourceFields = selectedFormJSON.filter((ele: any) => {
                    return ele?.label && SourceFieldNames.includes(stripHtmlTags(ele?.label))
                })
                console.log("SourceFields", SourceFields)

                SetSourceFieldJsons(SourceFields)
            }
            if (matchedChildData) {
                debugger;
                const parsedFormDataArray = childformData
                    .filter((dataItem: any) => dataItem.name.startsWith("file_upload"))
                    .map((dataItem: any) => ({
                        FileName: dataItem.name,
                        FileValue: dataItem.value,
                        FileId: dataItem.id
                    }));

                debugger
                console.log("parsedFormDataArrayyyyyyyyyyyyyyyyyyyyyyyyy", parsedFormDataArray)
                setMappedfile(parsedFormDataArray);
            } else { setMappedfile([]) }

            //    const childTransaction = await ChildtransactionItems.filter(t =>{console.log(t.AppCode); t.AppCode === childFormAppCode})[0];
            debugger;
            if (!childTransaction) {
                console.log("❌ No transaction found for ChildOrder:", childOrder);

                setAuthorADdata(null);
                setAllComments([]);
                setFlowJSON([]);
                setAttachmentFiles([]);

            }
            else {


                console.log(`✓ Transaction found - ID: ${childTransaction.Id}, AppCode: ${childTransaction.AppCode}, SeqNo: ${childTransaction.SeqNo}`);
                // console.log(`  Status: ${childTransaction.Status}`);
                // console.log(`  Has AuthorADdata: ${!!childTransaction.AuthorADdata}`);

                // === Fetch FlowJSON for the child transaction ===
                console.log(`🔄 Fetching FlowJSON for AppCode: ${childTransaction.AppCode}`);
                let mappingMasterData = await sp.web.lists
                    .getByTitle("MappingMaster")
                    .items.select("Level", "Role", "Users_x002f_Groups/Title")
                    .expand("Users_x002f_Groups")
                    .filter(`AppCode eq '${childTransaction.AppCode}'`)
                    .get();

                // Fallback to FormMaster AppCode if needed
                if (mappingMasterData.length === 0) {
                    console.log(`  Trying fallback with FormMaster AppCode: ${childFormAppCode}`);
                    mappingMasterData = await sp.web.lists
                        .getByTitle("MappingMaster")
                        .items.select("Level", "Role", "Users_x002f_Groups/Title")
                        .expand("Users_x002f_Groups")
                        .filter(`AppCode eq '${childFormAppCode}'`)
                        .get();
                    console.log(`⚠️ No MappingMaster data found — setting empty FlowJSON`);
                    setFlowJSON([]);
                } else {
                    const levelsWithRoles = mappingMasterData
                        .map(item => ({
                            Level: item.Level,
                            Role: item.Role,
                            User: item.Users_x002f_Groups ? item.Users_x002f_Groups.Title : "N/A",
                        }))
                        .sort((a, b) => a.Level - b.Level);

                    const highestLevel = Math.max(...mappingMasterData.map(item => item.Level));
                    const endflow =
                        (childTransaction.Status === "In-Progress" &&
                            Number(childTransaction.CurrentQueue) === highestLevel) ||
                        (childTransaction.Status === "In-Progress" &&
                            Number(childTransaction.DestinationQueue) === highestLevel);

                    console.log(`  Generating FlowJSON with endflow: ${endflow}`);
                    const flowjson = await FlowData(childTransaction, levelsWithRoles, endflow);
                    console.log(`  Generated FlowJSON with ${flowjson.length} nodes`);
                    setFlowJSON(flowjson);
                }

                // Update Requester Details (AuthorADdata)
                if (childTransaction.AuthorADdata && childTransaction.AuthorADdata.trim() !== "") {
                    console.log(`✓ Updating requester details for child form`);
                    const parsedAuthorData = JSON.parse(childTransaction.AuthorADdata);
                    setAuthorADdata(parsedAuthorData);
                } else {
                    console.log(`⚠️ No requester details available for this child form`);
                    setAuthorADdata(null);
                }
                const seqNoFolder = childTransaction.SeqNo;
                const files = await sp.web.lists
                    .getByTitle("NoteAttach")
                    .items.select("FileLeafRef", "FileRef", "FileDirRef")
                    .filter(`FileDirRef eq '${fileFetchUrl}/NoteAttach/${seqNoFolder}'`)
                    .get();

                const fileNamesss = files.map((file, index) => {
                    const fileName = file.FileLeafRef || file.Title || "Unnamed File";

                    const fileUrl = file.FileRef;

                    return { name: `${index + 1}. ${fileName}`, url: fileUrl };
                });
                setAttachmentFiles(fileNamesss);
                const childmodifiedFormJSON = ChildFormJson.map((item: any) => {
                    if (item.element === "FileUpload") {

                        const filename = fileNamesss.map((rawname) => {
                            const fname = rawname.name.replace(/^\d+\.\s*/, '').split('name')[0];
                            const controlname = rawname.name.split('name')[1];
                            if (controlname === item.field_name)
                                return fname;
                            //   return rawname.name.replace(/^\d+\.\s*/, '').split('name')[0];
                            return "";
                        }).filter(str => str.trim() !== "").join(', ')

                        return {
                            ...item,
                            text: "See Attachment Section",
                            label: filename ? `${item.label} : ${filename}` : `${item.label} : No file(s) available`,
                            className: item.className ? `${item.className} custom-red-label` : "custom-red-label",
                            style: { color: "red" }
                        };
                    }
                    return item;
                });

                // === Fetch comments for this child form ===
                // console.log(`📝 Fetching comments for AppCode: ${childTransaction.AppCode}, SeqNo: ${childTransaction.SeqNo}`);

                const CommentItems = await sp.web.lists
                    .getByTitle("CommentsLog")
                    .items.filter(`(AppCode eq '${childTransaction.AppCode}') and (SeqNo eq '${childTransaction.SeqNo}')`)
                    .expand("AttachmentFiles")
                    .select(
                        "Id",
                        "Comments",
                        "CommentedBy",
                        "CommentedByEmail",
                        "AppCode",
                        "SeqNo",
                        "AttachmentFiles/FileName",
                        "AttachmentFiles/ServerRelativeUrl",
                        "Created",
                        "Modified"
                    )
                    .get();

                console.log(`✓ Found ${CommentItems.length} comments`);
                const processedComments = CommentItems.map((comment) => ({
                    Id: comment.Id,
                    Comments: comment.Comments,
                    CommentedBy: comment.CommentedBy,
                    CommentedByEmail: comment.CommentedByEmail,
                    AppCode: comment.AppCode,
                    SeqNo: comment.SeqNo,
                    Created: comment.Created,
                    Modified: comment.Modified,
                    Attachments: comment.AttachmentFiles.map((file: any) => ({
                        FileName: file.FileName,
                        FileUrl: file.ServerRelativeUrl,
                    })),
                }));
                processedComments.sort((a, b) => new Date(b.Modified).getTime() - new Date(a.Modified).getTime());
                setAllComments(processedComments);


                console.log(`✅ Child tab data updated successfully for ${childFormAppCode} (Transaction AppCode: ${childTransaction.AppCode})`);
                return childmodifiedFormJSON;
            }
        } catch (error) {
            console.error("Error updating child tab data:", error);
        }
    };

    const downloadWithName = async (url: string, newFileName: string) => {
        const response = await fetch(url);
        const blob = await response.blob();
        const blobUrl = URL.createObjectURL(blob);

        const link = document.createElement("a");
        link.href = blobUrl;
        link.download = newFileName;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(blobUrl);
    };

    // function childTabChange() {
    //     setSecChildFormRaised(true);
    //     console.log(SecChildFormRaised)
    //     throw new Error('Function not implemented.');
    // }

    // ECLIPSE POP UP ------------------------------------------------------------------

    return (
        <>
            {(loading) ? <LoadingSpinner></LoadingSpinner> : (record.Domain === domainfield && (record.AuthorId === userId || isAdmin) ?
                <div style={{ display: 'flex', flexDirection: 'column' }}>
                    <Toast ref={toast} />

                    < TopNavBar onUserDomainRetrieved={handleUserDomain} />
                    <div style={{ display: 'flex' }}>
                        <SideBar activeMenu="" />

                        <div className='FC_mainDiv' style={{ width: '100%' }}>
                            <div className="FC_header-container">
                                <div className="FC_title-section">
                                    <div className="FC_back-icon">
                                        <img onClick={() =>

                                            navigate(-1 as unknown as To, preservedfilters)} src={require('../../../assets/Images/previous.png')} alt="backicon" title='Back' />
                                    </div>
                                    <div>
                                        <p className="FC_FormName">{appName}</p>
                                    </div>
                                </div>

                                <div className="FC_requester-name">
                                    <p style={{ margin: 0 }}>{RequesterName}</p>
                                </div>
                            </div>

                            {IsParentForm && (
                                <ParentWithChildren
                                    parentName="Main Parent Form"
                                    children={ChildFormDetails.map(ele => ({
                                        id: ele.Id,
                                        name: ele.AppName,
                                        ChildFormJson: JSON.parse(ele.FormJSON),
                                        AppCode: ele.AppCode,
                                        ChildOrder: ele.ChildOrder
                                    }))}
                                    onChildClick={async (child) => {
                                        console.log("Clicked:", child);
                                        setChildselected(child.AppCode);
                                        setchildappselected(child.AppCode);
                                        console.log(childappselected);
                                        const matchedChild = ChildtransactionItems.find(
                                            ele => ele.AppCode === child.AppCode
                                        );
                                        setreopenbuttonview(false)
                                        setMultiFormSeqNo(matchedChild?.SeqNo);
                                        setIsAccordionVisible(!matchedChild);

                                        // console.log("Tab Clickedddddddddddddddddddddd", isAccordionVisible)

                                        setIsChildReturned(matchedChild?.Status === "Returned")

                                        setFormEditMode(!!matchedChild && matchedChild.Status !== "Returned");

                                        // if (!matchedChild || matchedChild.Status?.toLowerCase() === "returned") {
                                        //     setResubmitButton(true)
                                        //     setFormEditMode(false);
                                        // }
                                        // else{setResubmitButton(false);
                                        //     setFormEditMode(true);
                                        //     }

                                        const childOrder = child.ChildOrder || 1;
                                        const currentchildanswerdata = matchedChild
                                            ? JSON.parse(matchedChild.FormData)
                                            : [];

                                        setChildFormJsons(child.ChildFormJson);
                                        setChildFormData(currentchildanswerdata);
                                        const childformjson = await updateChildTabData(String(child.id), childOrder, child.ChildFormJson, matchedChild);
                                        const completeformjson = [...ParentFormJSONCopy, ...(Array.isArray(childformjson) ? childformjson : [])];

                                        const completeformanswers = [
                                            ...ParentFormAnswerCopy,
                                            ...currentchildanswerdata
                                        ];
                                        setFormJSON(completeformjson);
                                        setFormData(completeformanswers);


                                        setRecentStatus(matchedChild?.Status);
                                        // const flowjson = await FlowData(trnxItem, mappingLevels, EndFlow)
                                        // setFlowJSON(flowjson)
                                        console.log("Form Data Updated ✅", {
                                            completeformanswers,
                                            completeformjson
                                        });

                                    }}
                                />
                            )}


                            {!isAccordionVisible && recentStatus !== "Draft" && (<Accordion activeIndex={recentStatus !== "Submitted" ? 0 : 1} className='wholeAccordion'>

                                {recentStatus !== "Submitted" && <AccordionTab style={{ marginBottom: "3px" }} header={
                                    <div className="header-container">
                                        <span className="title">{recentTitle}</span>
                                        <div style={{ display: "flex", alignItems: "center" }} onClick={(e) => { e.preventDefault(); e.stopPropagation(); }}>
                                            {/* Download icon with PDF text to the right */}
                                            <div style={{
                                                display: "flex",
                                                alignItems: "center",
                                                marginRight: "10px",
                                                cursor: "pointer",
                                                marginBottom: "1px"
                                            }}
                                            >
                                                <i
                                                    className="fas fa-download"
                                                    style={{
                                                        fontSize: "1rem",
                                                        color: 'white',
                                                        marginRight: "4px" // Space between icon and text
                                                    }}
                                                    onClick={handleWatermarkedDownload}
                                                />
                                                <span style={{
                                                    fontSize: "0.7rem",
                                                    color: "white",
                                                    textTransform: "uppercase",
                                                    fontWeight: "bold" // Makes PDF text slightly more prominent
                                                }}
                                                    onClick={handleWatermarkedDownload}>
                                                    PDF
                                                </span>
                                            </div>
                                            <span
                                                className={`status ${recentStatus === 'In-Progress' ? 'status-inprogress' :
                                                    recentStatus === 'Completed' ? 'status-approved' :
                                                        recentStatus === 'Rejected' ? 'status-rejected' : ''
                                                    }`}
                                            >
                                                {recentStatus}
                                            </span>
                                        </div>
                                    </div>
                                }
                                >
                                    <div className={`circle-sequence ${mappingLevels.length > 5 ? '' : 'circleflex'}`}>
                                        {

                                            FlowJSON.map((levelData: NodeItem, index: number) => {
                                                const { role, name, modified, color, status } = levelData.data;

                                                return (
                                                    <div className="approval-step" key={levelData.id}>
                                                        <div className="role-label">{role}</div>
                                                        <div className="circle" style={{ backgroundColor: color }} title={status}>
                                                            <div className="name" title={name} >
                                                                {name}
                                                            </div>
                                                        </div>
                                                        <div className="date-label">{modified}</div>

                                                        {index !== FlowJSON.length - 1 && <div className="arrow-line" />}
                                                        {index == FlowJSON.length - 1 && (recentStatus === "Completed" || recentStatus === "Rejected") && (
                                                            <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                                                                <div className="arrow-line"></div>

                                                            </div>
                                                        )}
                                                    </div>
                                                )
                                            })
                                        }
                                        {(recentStatus === "Completed" || recentStatus === "Rejected") &&
                                            <div style={{ paddingLeft: '5px', paddingTop: '30px' }}>
                                                <span style={{ display: "inline-block", fontSize: "24px", color: "#000000", background: "#42d40c", borderRadius: "25px", padding: "0px 10px" }}>✓</span>
                                            </div>}
                                    </div>
                                </AccordionTab>}

                                <AccordionTab style={{ marginBottom: "3px" }} header={
                                    <div className="header-container">
                                        <span className="title">Requester Details</span>

                                        {(recentStatus === "Submitted") && <div className="header-container">
                                            <div style={{ display: "flex", alignItems: "center" }} onClick={(e) => { e.preventDefault(); e.stopPropagation(); }}>
                                                {/* Download icon with PDF text to the right */}
                                                <div style={{
                                                    display: "flex",
                                                    alignItems: "center",
                                                    marginRight: "10px",
                                                    cursor: "pointer",
                                                    marginBottom: "1px"
                                                }}
                                                >
                                                    <i
                                                        className="fas fa-download"
                                                        style={{
                                                            fontSize: "1rem",
                                                            color: 'white',
                                                            marginRight: "4px",

                                                        }}
                                                        onClick={handleWatermarkedDownload}
                                                    />
                                                    <span style={{
                                                        fontSize: "0.7rem",
                                                        color: "white",
                                                        textTransform: "uppercase",
                                                        fontWeight: "bold"
                                                    }}
                                                        onClick={handleWatermarkedDownload}>
                                                        PDF
                                                    </span>
                                                </div>
                                                <span
                                                    className={`status ${'status-approved'}`}
                                                >
                                                    {recentStatus}
                                                </span>
                                            </div>
                                        </div>}
                                    </div>
                                }>
                                    <div className="container">
                                        {authorADdata ? (
                                            <>
                                                <div className="row">
                                                    <div className="col-sm-4 col-12">
                                                        <span style={{ fontWeight: "500", fontSize: "12px" }}>
                                                            Requester Name : {authorADdata?.RequesterName}
                                                        </span>
                                                    </div>
                                                    <div className="col-sm-4 col-12">
                                                        <span style={{ fontWeight: "500", fontSize: "12px" }}>
                                                            Mail : {authorADdata?.Email}
                                                        </span>
                                                    </div>
                                                    <div className="col-sm-4 col-12">
                                                        <span style={{ fontWeight: "500", fontSize: "12px" }}>
                                                            Employee ID : {authorADdata?.EmployeeID}
                                                        </span>
                                                    </div>
                                                </div>
                                                <div className="row">
                                                    <div className="col-sm-4 col-12">
                                                        <span style={{ fontWeight: "500", fontSize: "12px" }}>
                                                            Designation : {authorADdata?.Designation}
                                                        </span>
                                                    </div>
                                                    <div className="col-sm-4 col-12">
                                                        <span style={{ fontWeight: "500", fontSize: "12px" }}>
                                                            Department : {authorADdata?.Department}
                                                        </span>
                                                    </div>
                                                    <div className="col-sm-4 col-12">
                                                        <span style={{ fontWeight: "500", fontSize: "12px" }}>
                                                            Contact No : {authorADdata?.ContactNo}
                                                        </span>
                                                    </div>
                                                </div>
                                            </>
                                        ) : (
                                            <div style={{ fontWeight: "500", fontSize: "12px", color: "gray" }}>No data found</div>
                                        )}
                                    </div>
                                </AccordionTab>

                                {recentStatus !== "Submitted" && <AccordionTab style={{ marginBottom: "3px" }} header={
                                    <div className="header-container">
                                        <span className="title">Comments Log</span>
                                    </div>}>

                                    <div className="container accordian_comment_container">

                                        <DataTable
                                            className="hover-table"
                                            value={allComments}
                                            responsiveLayout="scroll"
                                            currentPageReportTemplate="Showing {first} to {last} of {totalRecords}"
                                            dataKey="ID"
                                            emptyMessage="No Comments yet, Add comment"
                                        >


                                            <Column
                                                header="Comments"
                                                sortable
                                                body={(rowData) => {
                                                    const comment = rowData.Comments || '';

                                                    const parser = new DOMParser();
                                                    const parsedHtml = parser.parseFromString(comment, 'text/html');
                                                    const hasTable = parsedHtml.querySelector('table') !== null;
                                                    const hasImage = parsedHtml.querySelector('img') !== null;

                                                    const getStrippedText = (html: string) => {
                                                        const tempDiv = document.createElement('div');
                                                        tempDiv.innerHTML = html;
                                                        return tempDiv.textContent || tempDiv.innerText || '';
                                                    };

                                                    const cleanText = getStrippedText(comment);
                                                    let labelToShow = '';
                                                    let isClickable = false;

                                                    if (!hasTable && !hasImage) {
                                                        // Only text
                                                        if (cleanText.length > 75) {
                                                            labelToShow = `${cleanText.slice(0, 75)}... See more`;
                                                            isClickable = true;
                                                        } else {
                                                            labelToShow = cleanText;
                                                        }
                                                    } else if ((hasImage && !cleanText) || (hasTable && !cleanText) || (hasTable || hasImage) || (hasImage && hasTable)) {
                                                        // Only image or only table
                                                        labelToShow = 'Check comment...';
                                                        isClickable = true;
                                                    } else {
                                                        // Text with image or table
                                                        if (cleanText.length > 75) {
                                                            labelToShow = `${cleanText.slice(0, 75)}... See more`;
                                                        } else {
                                                            labelToShow = cleanText;
                                                        }
                                                        isClickable = true;
                                                    }

                                                    const handleClick = () => {
                                                        setSelectedComment(rowData);
                                                        setIsDialogVisible(true);
                                                    };

                                                    return (
                                                        <div
                                                            style={{
                                                                whiteSpace: 'pre-wrap',
                                                                wordBreak: 'break-word',
                                                                display: '-webkit-box',
                                                                WebkitLineClamp: 2,
                                                                WebkitBoxOrient: 'vertical',
                                                                overflow: 'hidden',
                                                                textOverflow: 'ellipsis',
                                                                maxWidth: '100%',
                                                                color: isClickable ? '#007bff' : 'inherit',
                                                                cursor: isClickable ? 'pointer' : 'default',
                                                            }}
                                                            onClick={isClickable ? handleClick : undefined}
                                                        >
                                                            {labelToShow}
                                                        </div>
                                                    );
                                                }}
                                                bodyStyle={{
                                                    whiteSpace: 'normal',
                                                    wordBreak: 'break-word',
                                                    width: '36%',
                                                    maxWidth: '36%',
                                                    padding: '1px 3px',
                                                    paddingLeft: '12px',
                                                }}
                                                headerStyle={{
                                                    width: '36%',
                                                    maxWidth: '36%',
                                                    padding: '3px',
                                                    fontSize: '14px',
                                                }}
                                            />


                                            <Column
                                                field="CommentedBy"
                                                header="Commented By"
                                                sortable
                                                bodyStyle={{ whiteSpace: 'pre-wrap', wordBreak: 'break-word', width: '16%', padding: '1px 3px', paddingLeft: "14px" }}
                                                headerStyle={{ width: '16%', padding: '0px', fontSize: "14px" }}
                                            />
                                            <Column
                                                field="Created"
                                                header="Created Date"
                                                sortable
                                                body={(rowData) => {
                                                    const date = new Date(rowData.Created);
                                                    return date.toLocaleString('en-GB', {
                                                        day: '2-digit',
                                                        month: 'short',
                                                        year: 'numeric',
                                                        hour: '2-digit',
                                                        minute: '2-digit',
                                                        second: '2-digit',
                                                        hour12: true,
                                                    });
                                                }}
                                                bodyStyle={{ whiteSpace: 'pre-wrap', wordBreak: 'break-word', width: '22%', padding: '1px 3px' }}
                                                headerStyle={{ width: '22%', padding: '0px', fontSize: "14px" }}
                                            />


                                            <Column
                                                header="Attachments"
                                                body={(rowData) => {
                                                    const count = rowData.Attachments?.length || 0;

                                                    return count > 0 ? (
                                                        <div
                                                            onClick={() => {
                                                                setSelectedComment(rowData);
                                                                setIsDialogVisible(true);
                                                            }}
                                                            style={{
                                                                display: 'flex',
                                                                alignItems: 'center',
                                                                cursor: 'pointer',
                                                                color: '#007bff',
                                                            }}
                                                        >
                                                            <svg
                                                                xmlns="http://www.w3.org/2000/svg"
                                                                width="20"
                                                                height="20"
                                                                viewBox="0 0 24 24"
                                                                fill="none"
                                                                stroke="currentColor"
                                                                strokeWidth="2"
                                                                strokeLinecap="round"
                                                                strokeLinejoin="round"
                                                                className="feather feather-file-text"
                                                                style={{ marginRight: '6px', marginLeft: "-16px" }}
                                                            >
                                                                <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                                                                <polyline points="14 2 14 8 20 8"></polyline>
                                                                <line x1="16" y1="13" x2="8" y2="13"></line>
                                                                <line x1="16" y1="17" x2="8" y2="17"></line>
                                                                <line x1="10" y1="9" x2="8" y2="9"></line>
                                                            </svg>
                                                            {count} Attachment{count > 1 ? 's' : ''}
                                                        </div>
                                                    ) : (
                                                        <span style={{ marginRight: '6px', marginLeft: "-16px" }}>No Attachments</span>
                                                    );
                                                }}
                                                bodyStyle={{ whiteSpace: 'pre-wrap', wordBreak: 'break-word', width: '24%', padding: '9px 18px' }}
                                                headerStyle={{ width: '24%', padding: '0px', fontSize: "14px", }}
                                            />



                                        </DataTable>
                                    </div>

                                </AccordionTab>}
                            </Accordion>)
                            }



                            <Dialog
                                header="Comment Details"
                                visible={isDialogVisible}
                                onHide={() => setIsDialogVisible(false)}
                                style={{ width: '50vw', background: "grey", backgroundColor: 'red' }}
                                modal
                                className="Comment_dialog_Header"
                            >
                                {selectedComment && (
                                    <table className='Comment_Table'  >
                                        <tbody>
                                            <tr>
                                                <td className='Comment_td_header' >Comment:</td>
                                                <td className='Comment_td_detail' >
                                                    {(() => {
                                                        const comment = selectedComment.Comments || '';
                                                        const containsTable = comment.includes('<table');
                                                        const containsImage = comment.includes('<img');

                                                        if (containsTable || containsImage) {
                                                            return (
                                                                <div
                                                                    dangerouslySetInnerHTML={{
                                                                        __html: comment,
                                                                    }}
                                                                />
                                                            );
                                                        } else {
                                                            // Strip HTML tags from plain comments
                                                            const getStrippedText = (html: string) => {
                                                                const tempDiv = document.createElement("div");
                                                                tempDiv.innerHTML = html;
                                                                return tempDiv.textContent || tempDiv.innerText || "";
                                                            };

                                                            return getStrippedText(comment);
                                                        }
                                                    })()}
                                                </td>
                                            </tr>
                                            <tr>
                                                <td className='Comment_td_header' >Commented By:</td>
                                                <td className='Comment_td_detail'>{selectedComment.CommentedBy}</td>
                                            </tr>
                                            <tr>
                                                <td className='Comment_td_header' >Created Date:</td>
                                                <td className='Comment_td_detail'>
                                                    {new Date(selectedComment.Created).toLocaleString()}
                                                </td>
                                            </tr>
                                            {(selectedComment.Attachments?.length ?? 0) > 0 && (
                                                <tr>
                                                    <td className='Comment_td_header' >Attachments:</td>
                                                    <td className='Comment_td_detail'>
                                                        <ul style={{ paddingLeft: 0, margin: 0, listStyleType: 'none' }}>
                                                            {selectedComment.Attachments?.map((file, index) => (
                                                                <li key={index}>
                                                                    <a
                                                                        href={file.FileUrl}
                                                                        target="_blank"
                                                                        rel="noopener noreferrer"
                                                                    >
                                                                        {file.FileName}
                                                                    </a>
                                                                </li>
                                                            ))}
                                                        </ul>
                                                    </td>
                                                </tr>
                                            )}
                                        </tbody>
                                    </table>
                                )}
                            </Dialog>

                            <div style={{ marginTop: "12px" }}>
                                <ul >
                                    <div className='Select_Btn'>
                                        <li
                                            style={{ borderTopLeftRadius: '5px' }}
                                            className={selectTab === 'AllDetails' ? 'onSelected' : 'onUnSelected'}
                                            onClick={() => setSelectTab('AllDetails')}
                                        >
                                            All Details
                                        </li>
                                        <li
                                            className={selectTab === 'Attachment' ? 'onSelected' : 'onUnSelected'}
                                            onClick={() => setSelectTab('Attachment')}
                                        >
                                            Attachment
                                        </li>

                                    </div>

                                    {/* <li

                                className={selectTab === 'Comment' ? 'onSelected' : 'onUnSelected'}
                                onClick={() => setSelectTab('Comment')}
                            >
                                Comments Log
                            </li> */}
                                    {/* <li
                                className={selectTab === 'WorkflowLog' ? 'onSelected' : 'onUnSelected'}
                                onClick={() => setSelectTab('WorkflowLog')}
                            >
                                Workflow Log
                            </li> */}
                                </ul>
                            </div>





                            {loading ? (
                                <LoadingSpinner />
                                // <></>
                            ) : (
                                <div className='fo-mid1'>
                                    {selectTab === 'AllDetails' && (
                                        <>
                                            <div className='p-tabview-panel' style={{ paddingTop: 6 }}>

                                                <div className='d-flex align-items-center custom_field_main' style={{ gap: 0, alignItems: 'center' }}>
                                                    {buttonView || ResubmitButton ?
                                                        // <div className='form-group col-md-4' style={{ marginLeft: -4 }} >
                                                        //     <label htmlFor="requestfor" style={{ marginTop: 5 }} className='form-label'>Request For <span style={{ color: 'red' }}> *</span></label>
                                                        //     <input type="text" id='requestfor' className='  form-controlll form-control' readOnly={!(buttonView || ResubmitButton)} value={Categorytype} onChange={(e) => setCategorytype(e.target.value)} />
                                                        // </div>
                                                        <div className='form-group col-md-4' style={{ marginLeft: -4 }} >
                                                            <label htmlFor="requestfor" style={{ marginTop: 5 }} className='form-label'>Request For <span style={{ color: 'red' }}> *</span></label>
                                                            <input type="text" id='requestfor' className='  form-controlll form-control' readOnly value={Categorytype} onChange={(e) => setCategorytype(e.target.value)} />
                                                        </div>
                                                        :
                                                        <div className='col-md-4' style={{ marginLeft: -4, marginTop: '14px' }} >
                                                            <label htmlFor="requestfor" style={{ marginTop: 5, paddingBottom: '6px' }} className='form-label'>Request For <span style={{ color: 'red' }}> *</span></label> <br />
                                                            <input type="text" id='requestfor' className='view_form-controlll' value={Categorytype} readOnly />
                                                        </div>
                                                    }


                                                    {
                                                        buttonView || ResubmitButton ? (
                                                            <div className='form-group right-section' style={{ marginLeft: -4 }} >
                                                                <label htmlFor="requestfor" style={{ marginTop: 5 }} className='form-label FC_StaticFields'>Notify To </label>

                                                                <PeoplePicker
                                                                    context={getPeoplePickerContext()}
                                                                    personSelectionLimit={5}
                                                                    //Added disabled for readonly
                                                                    disabled={true}
                                                                    required={false}
                                                                    onChange={onChange}
                                                                    showHiddenInUI={false}
                                                                    principalTypes={[PrincipalType.User]}
                                                                    resolveDelay={1000}
                                                                    defaultSelectedUsers={NotifyRequestorEMails}
                                                                    resultFilter={(results: any[]) =>
                                                                        results.filter(persona => {
                                                                            if (isSameDomain && String(isSameDomain) === "NO") {
                                                                                let email = persona?.loginName || '';
                                                                                console.log("persona in if ", persona)
                                                                                return email.includes(`${domainfield}`);
                                                                            }
                                                                            else if (isSameDomain) {
                                                                                console.log("persona in else ", persona)
                                                                                let email = persona?.loginName || '';
                                                                                return !email.includes('#ext#');
                                                                            }
                                                                        })}
                                                                />

                                                            </div>

                                                        ) : (
                                                            <div className='col-md-8 people_notify_to'>
                                                                <PeoplePicker
                                                                    context={getPeoplePickerContext()}
                                                                    titleText="Notify To"
                                                                    personSelectionLimit={5}
                                                                    required={false}
                                                                    disabled={true}
                                                                    showHiddenInUI={false}
                                                                    principalTypes={[PrincipalType.User]}
                                                                    resolveDelay={1000}
                                                                    defaultSelectedUsers={NotifyRequestorEMails}
                                                                />
                                                            </div>
                                                        )
                                                    }

                                                </div>
                                                {(isAccordionVisible || isChildReturned) ?
                                                    <div className="form-component-mode">

                                                        <ReactFormGenerator
                                                            key={JSON.stringify(ParentFormAnswerCopy)}
                                                            // onBlur={(e) => handleBlur(e)}
                                                            data={ParentFormJSONCopy}
                                                            read_only={true}
                                                            form_action=""
                                                            form_method=""
                                                            answer_data={ParentFormAnswerCopy}
                                                        // onChange={(e) => handleChange(e)}
                                                        // onChange={(updatedData) => setFormData(updatedData)}
                                                        />

                                                        <ReactFormGenerator
                                                            key={JSON.stringify(ChildFormJsons)}
                                                            onBlur={(e) => handleBlur(e)}
                                                            data={ChildFormJsons}
                                                            read_only={false}
                                                            form_action=""
                                                            form_method=""
                                                            answer_data={childformData}
                                                            onChange={(e) => handleChange(e)}
                                                        // onChange={(updatedData) => setFormData(updatedData)}
                                                        />
                                                    </div>
                                                    :
                                                    <div className="form-component-mode">
                                                        <ReactFormGenerator
                                                            key={JSON.stringify(formData)}
                                                            // onBlur={(e) => handleBlur(e)}
                                                            data={formJSON}
                                                            read_only={true}
                                                            form_action=""
                                                            form_method=""
                                                            answer_data={formData}
                                                            onChange={(e) => handleChange(e)}
                                                        // onChange={(updatedData) => setFormData(updatedData)}
                                                        />
                                                    </div>
                                                }

                                            </div>

                                            <div>
                                                { }
                                            </div>

                                            {recentStatus === "Returned" ?
                                                <div style={{ marginLeft: '1.5rem' }}>
                                                    <h6 style={{ marginTop: '10px', color: "white", marginLeft: '7px' }}>Add your comment here</h6>
                                                    <div className="comment-box">
                                                        <div style={{ flexShrink: 0, flexBasis: "50%", color: "#B1B1B1" }}>
                                                            <PasteArea imagetabledata={handlepasteData} />
                                                        </div>


                                                        <div style={{ flexShrink: 0, flexBasis: "20%" }}>

                                                            <div className="comment_doc_upload_container">
                                                                {/* Wrapping both the image and label inside the label tag */}
                                                                <label
                                                                    htmlFor="fileInput"
                                                                    style={{
                                                                        cursor: "pointer",
                                                                        //   border: "2px dashed #63489d",
                                                                        backgroundColor: "rgb(7, 13, 25)",
                                                                        display: "inline-flex",
                                                                        alignItems: "center",
                                                                        justifyContent: "center",
                                                                        flexDirection: "column",
                                                                        borderRadius: "5px",
                                                                        width: "100%",
                                                                        height: "40px",
                                                                        fontSize: "15px",
                                                                        fontWeight: 400,
                                                                        marginTop: "10px",
                                                                        gap: "10px",
                                                                        color: "white"
                                                                    }}
                                                                >
                                                                    <img src={require("../../../assets/Images/UploadIcon.png")} style={{ width: "12%" }} />
                                                                    Upload Attachment
                                                                </label>

                                                                <input
                                                                    type="file"
                                                                    multiple
                                                                    onChange={onFileChange}
                                                                    style={{ display: 'none' }}
                                                                    id="fileInput"
                                                                />
                                                            </div>

                                                            <button onClick={handleCommnetSaveOnly} style={{ backgroundColor: "green", color: "white", padding: "6px 17px", fontWeight: "bold", margin: "14px 20px", alignItems: "center" }}>Save</button>
                                                        </div>

                                                        <div style={{ flexShrink: 0, flexBasis: "25%" }}>

                                                            {selectedFile.length > 0 && (
                                                                <ul className='Comment_file_upload_ul' >
                                                                    {selectedFile.map((file, index) => (
                                                                        <li key={index} className='Comment_file_upload_li'>
                                                                            <span>{file.name}</span>
                                                                            <svg
                                                                                xmlns="http://www.w3.org/2000/svg"
                                                                                viewBox="0 0 24 24"
                                                                                onClick={() => handleRemoveFile(index)}
                                                                                style={{
                                                                                    width: '15px',
                                                                                    height: "15px",
                                                                                    fill: "white", // default color
                                                                                    cursor: "pointer",
                                                                                    marginLeft: "8px"
                                                                                }}
                                                                            >
                                                                                <g>
                                                                                    <path fill="none" d="M0 0h24v24H0z" />
                                                                                    <path d="M4 8h16v13a1 1 0 0 1-1 1H5a1 1 0 0 1-1-1V8zm2 2v10h12V10H6zm3 2h2v6H9v-6zm4 0h2v6h-2v-6zM7 5V3a1 1 0 0 1 1-1h8a1 1 0 0 1 1 1v2h5v2H2V5h5zm2-1v1h6V4H9z" />
                                                                                </g>
                                                                            </svg>
                                                                        </li>
                                                                    ))}
                                                                </ul>
                                                            )}


                                                        </div>

                                                    </div>



                                                </div>

                                                : " "}

                                            {/* {recentStatus !== "Draft" ? <ChatBox context={context} appSeqNo={appSeqNo} userName={userName} onChatResponse={handleChatResponse} appcode={AppCodevalue} /> : ""} */}


                                        </>)}


                                    {selectTab === 'Attachment' && (
                                        <div className='p-tabview-panel'>
                                            {attachmentFiles.length === 0 ? (
                                                <li>No attachments available</li>
                                            ) : (
                                                attachmentFiles.map((file, index) => (
                                                    <div key={index}>

                                                        <section style={{ display: "flex", justifyContent: "space-between" }}>
                                                            <p className='Attachment_fileName'>{file.name.split("name")[0]}</p>
                                                            {isAccordionVisible || buttonView ? <button className="doc_delete" onClick={() => handleDocumentDelete(file.name, file.id)}>Delete</button> : ""}

                                                        </section>
                                                        {/* <p className='Attachment_fileName'>{file.name}</p> */}
                                                        {file.name.endsWith('.pdf') ? (
                                                            <embed src={file.url} width="100%" height="600px" type="application/pdf" />
                                                        ) : file.name.endsWith('.jpg') || file.name.endsWith('.jpeg') || file.name.endsWith('.png') ? (
                                                            <img src={file.url} alt={file.name} style={{ width: '100%', height: 'auto' }} />
                                                        ) : file.name.endsWith('.docx') || file.name.endsWith('.pptx') ? (
                                                            <div>
                                                                <p>File preview not available. You can download and view it:</p>
                                                                {/* <a href={file.url} target="_blank" rel="noopener noreferrer">{file.name.split("name")[0]}</a> */}
                                                                <a href='#'
                                                                    onClick={async (e) => {
                                                                        e.preventDefault();
                                                                        await downloadWithName(file.url, file.name.replace(/^\d+\.\s*/, '').split("name")[0]);
                                                                    }}
                                                                    target="_blank"
                                                                    rel="noopener noreferrer">{file.name.split("name")[0]}
                                                                </a>
                                                            </div>
                                                        ) : (
                                                            <div>
                                                                <p>File type not supported for preview. You can download and view it:</p>
                                                                {/* <a href={file.url} target="_blank" rel="noopener noreferrer">{file.name.split("name")[0]}</a> */}
                                                                <a href='#'
                                                                    onClick={async (e) => {
                                                                        e.preventDefault();
                                                                        await downloadWithName(file.url, file.name.replace(/^\d+\.\s*/, '').split("name")[0]);
                                                                    }}
                                                                    target="_blank"
                                                                    rel="noopener noreferrer">{file.name.split("name")[0]}
                                                                </a>
                                                            </div>
                                                            // ------ ending for multi file attachment code
                                                        )}
                                                    </div>
                                                ))
                                            )}
                                        </div>
                                    )}
                                    {selectTab === 'Comment' && (
                                        <div className='p-tabview-panel'>

                                            <DataTable
                                                className="hover-table"
                                                value={allComments}
                                                responsiveLayout="scroll"
                                                currentPageReportTemplate="Showing {first} to {last} of {totalRecords}"
                                                dataKey="ID"
                                                emptyMessage="No Comments yet, Add comment"
                                            >
                                                <Column
                                                    field="Comments"
                                                    header="Comments"
                                                    sortable
                                                    bodyStyle={{ whiteSpace: 'pre-wrap', wordBreak: 'break-word', width: '28%' }}
                                                    headerStyle={{ width: '26%' }}
                                                />
                                                <Column
                                                    field="CommentedBy"
                                                    header="Commented By"
                                                    sortable
                                                    bodyStyle={{ whiteSpace: 'pre-wrap', wordBreak: 'break-word', width: '22%' }}
                                                    headerStyle={{ width: '20%' }}
                                                />
                                                <Column
                                                    field="Created"
                                                    header="Created Date"
                                                    sortable
                                                    body={(rowData) => {
                                                        const date = new Date(rowData.Created);
                                                        return date.toLocaleString('en-GB', {
                                                            day: '2-digit',
                                                            month: 'short',
                                                            year: 'numeric',
                                                            hour: '2-digit',
                                                            minute: '2-digit',
                                                            second: '2-digit',
                                                            hour12: true,
                                                        });
                                                    }}
                                                    bodyStyle={{ whiteSpace: 'pre-wrap', wordBreak: 'break-word', width: '22%' }}
                                                    headerStyle={{ width: '20%' }}
                                                />
                                                <Column
                                                    header="Attachments"
                                                    body={(rowData) => {
                                                        return rowData.Attachments && rowData.Attachments.length > 0 ? (
                                                            <ul style={{ paddingLeft: '0', margin: 0, listStyleType: 'none' }}>
                                                                {rowData.Attachments.map((file: any, index: any) => (
                                                                    <li key={index} style={{ marginBottom: '5px', display: 'flex', alignItems: 'center' }}>
                                                                        {/* SVG Icon */}
                                                                        <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" className="feather feather-file-text" style={{ marginLeft: '0px', marginRight: '5px' }}>
                                                                            <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"></path>
                                                                            <polyline points="14 2 14 8 20 8"></polyline>
                                                                            <line x1="16" y1="13" x2="8" y2="13"></line>
                                                                            <line x1="16" y1="17" x2="8" y2="17"></line>
                                                                            <line x1="10" y1="9" x2="8" y2="9"></line>
                                                                        </svg>
                                                                        {/* File Name Link */}
                                                                        <button
                                                                            // href={file.FileUrl}
                                                                            // target="_blank"
                                                                            onClick={() => {
                                                                                window.open(file.FileUrl, '_blank');
                                                                            }}
                                                                            rel="noopener noreferrer"
                                                                            style={{
                                                                                all: 'unset',           // Reset all default button styles
                                                                                textDecoration: 'none', // Remove underline if needed
                                                                                color: 'black',         // Set the text color
                                                                                background: 'transparent', // Remove background color
                                                                                border: 'none',         // Remove border
                                                                                cursor: 'pointer',      // Make it look clickable
                                                                                padding: 0,             // Remove padding
                                                                                font: 'inherit',        // Inherit font style from parent (same as normal text)
                                                                                display: 'inline',      // Make the button inline with other elements
                                                                            }}
                                                                        >
                                                                            {file.FileName}
                                                                        </button>
                                                                    </li>
                                                                ))}
                                                            </ul>
                                                        ) : (
                                                            'No Attachments'
                                                        );
                                                    }}
                                                    bodyStyle={{ whiteSpace: 'pre-wrap', wordBreak: 'break-word', width: '28%' }}
                                                    headerStyle={{ width: '26%', textAlign: 'center' }}
                                                />


                                            </DataTable>



                                        </div>
                                    )}
                                    {selectTab === 'WorkflowLog' && (
                                        <div className='p-tabview-panel'>

                                            <DataTable value={AuditLog}>
                                                <Column field="Title" header="Workflow Log" />
                                            </DataTable>
                                        </div>
                                    )}
                                    {isChildReturned ?
                                        <div className='form-btn'>
                                            {/* {Cancelrequest && (<button className='all-btn' onClick={(e) => { handlecancelchange("cancel") }}>Cancel Request</button>
                                            )} */}
                                            <button disabled={isButtonDisabled} className='Sbmt-btn all-btn' onClick={async () => { if (childformData) { await handleUpdateForm(childformData) } }}>Re-Submit</button>

                                            <button className="all-btn" onClick={handleCancel}>Close</button>
                                        </div> :
                                        isAccordionVisible ?
                                            <div className='form-btn'>
                                                <button
                                                    onClick={async () => {
                                                        if (childformData.length > 0 || tables.length > 0 || timesheettables.length > 0) {
                                                            await handleSave(childformData);
                                                        } else {
                                                            // setError("Please fill out the form before saving.");
                                                            toast.current?.show({ severity: 'error', summary: '', detail: 'Please fill out the form before saving.', life: 2000 });
                                                        }
                                                    }}
                                                    className='all-btn' style={{ display: "none" }}
                                                >
                                                    Save
                                                </button>

                                                <button disabled={isButtonDisabled} className='Sbmt-btn all-btn' onClick={async () => { if (childformData) { await handleUpdateForm(childformData) } }}>Submit</button>
                                                <button className='all-btn' style={{ display: "none" }} onClick={(e) => { handlecancelchange("delete") }}>Delete Request</button>
                                                <button className="all-btn" onClick={handleCancel}>Close</button>

                                            </div> : <div className='form-btn'>
                                                {(reopenbuttonview && (recentStatus === "Completeddddd" || recentStatus === "Rejecteddddd")) && <button className='Sbmt-btn all-btn' style={{ display: "none" }} onClick={() => setisreopen(true)}>Re open</button>}
                                                {/* {Cancelrequest && (<button className='all-btn' onClick={(e) => { handlecancelchange("cancel") }}>Cancel Request</button>)} */}
                                                <button className="all-btn" onClick={handleCancel}>Close</button>

                                            </div>

                                    }
                                </div>
                            )}
                            <div>
                                {Openpop && (selectTab === 'AllDetails') && (
                                    <div className="custom-dark-modal-overlay" style={{ height: '100%' }} >
                                        <div className=" custom-dark-modal" style={Opencanpop ? ResubmitButton ? {
                                            minWidth: '285px',
                                            maxWidth: '285px'
                                        } : {} : {
                                            minWidth: '285px',
                                            maxWidth: '285px'
                                        }}>
                                            {Opencanpop ? <div> <h5 style={{ margin: "4.5% 0 0 5%", fontSize: 14 }} >Do you want to cancel the request ?</h5>

                                                <div className="comment-box form-group" style={{ display: 'flex', justifyContent: "space-between", width: "90.5%", margin: "20px" }}>

                                                    {!ResubmitButton && <textarea
                                                        onChange={(e) => setComment(e.target.value)}
                                                        placeholder="Comment..."
                                                        value={comment}
                                                        required={true}
                                                        // className={highlight ? 'highlight' : ''}
                                                        style={{
                                                            border: '1px solid #6A6E79',
                                                            backgroundColor: '#161D2F',
                                                            color: 'white',
                                                            outline: 'none',
                                                            width: "95%",
                                                            height: "50px"

                                                        }}
                                                    ></textarea>}
                                                </div>
                                            </div>

                                                : <h5 style={{ marginTop: "25px", textAlign: "center", fontSize: 15 }}>Do you want to Delete the request ?</h5>}

                                            {Opencanpop ? <div className="modal-buttons" style={{ padding: '0px 20px 20px 20px' }}>

                                                <button className="newlogocolorbtn" onClick={handlereqCancel}>Ok</button>
                                                <button className="newblackcolorbtn" onClick={() => setOpenpop(false)}>Cancel</button>
                                            </div> : <div className="modal-buttons" style={{ padding: '23px 0px 25px 22px' }}>
                                                <button className="newlogocolorbtn" onClick={handleDeleteRequest}>Ok</button>
                                                <button className="newblackcolorbtn" onClick={() => setOpenpop(false)}>Cancel</button>
                                            </div>}
                                        </div>
                                    </div>
                                )}
                            </div>
                            {isreopen && <div className="custom-dark-modal-overlay">
                                <div className="custom-dark-modal" style={{ minHeight: 180 }}>
                                    <h5 className='addsteptitle' style={{ fontSize: '14px', backgroundColor: 'transparent', marginLeft: 5 }}>{`Reason for Reopen:`}</h5>
                                    <div className="form-group" style={{ padding: '0px 20px 10px 20px' }}>
                                        <div><PasteArea imagetabledata={handlepasteData} containerClassName="short-height"></PasteArea></div>
                                        <p style={{ marginTop: 10, fontSize: 13 }}>Note: Reopening this ticket will restart the approval flow.</p>
                                        <div className="modal-buttons"
                                        // style={{ gap: '5%', marginTop: '20px' }}
                                        >
                                            {/* <div> */}
                                            <button
                                                className="newlogocolorbtn"
                                                // style={{ marginRight: '20px' }}
                                                onClick={() => handleReopen()}
                                            >
                                                Yes
                                            </button>
                                            <button
                                                className="newblackcolorbtn"
                                                onClick={() => setisreopen(false)}
                                            >
                                                No
                                            </button>
                                            {/* </div> */}
                                        </div>
                                    </div>

                                </div>
                            </div>}
                            {(recentStatus !== "Draft" && !isAccordionVisible) ? <ChatBox context={context} appSeqNo={IsParentForm ? multiFormSeqNo : appSeqNo} userName={userName} onChatResponse={handleChatResponse} appcode={AppCodevalue} isSameDomain={isSameDomain} /> : ""}
                        </div>
                    </div>

                </div > : <NoAccess></NoAccess >
            )}
        </>
    );
}

export default ChildForm;