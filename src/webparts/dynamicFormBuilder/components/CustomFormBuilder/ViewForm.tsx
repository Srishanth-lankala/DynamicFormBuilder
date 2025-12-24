import * as React from 'react';
import { useEffect, useState } from 'react';
import { useParams, useNavigate, useLocation, To } from 'react-router-dom';
import { sp } from "@pnp/sp/presets/all";
import { ReactFormGenerator } from 'react-form-builder2';
import TopNavBar from './TopBar';
import LoadingSpinner from './Loading';
import './FormGenerator.css';
import { DataTable } from 'primereact/datatable';
import { Column } from 'primereact/column';
import { mySiteUrl } from './ConfigURL/All_URLs';
import { Toast } from 'primereact/toast';
import { Accordion, AccordionTab } from 'primereact/accordion';
import { IDynamicFormBuilderProps } from '../IDynamicFormBuilderProps';
import ChatBox from './UtilityComponents/ChatBox';
import { PeoplePicker, PrincipalType, IPeoplePickerContext } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { fileFetchUrl } from './ConfigURL/All_URLs';
import { useGlobalState } from './GlobalVariable/GlobalStateContext';
import PasteArea from './PasteArea';
import TableComponent, { TableData } from './TableStructure/TableComponent';
import TimesheetComponent, { TimeSheetData } from './Timesheetstructure/TimesheetComponent';
import ReactDOM from 'react-dom';
import { CommentData } from './EditTransactionForm';
import { Dialog } from 'primereact/dialog';
import { fetchTenantUser } from './FetchTenantUser/fetchTenantUser';
import logoimage from '../../assets/Images/logosmartoffctrp.png';
import {
    pdf
} from '@react-pdf/renderer';
import { degrees, PDFDocument, rgb } from 'pdf-lib';
import MyDocument from './MyDocument';
import SideBar from './Sidebar/SideBar';
import { FlowData, NodeItem } from './AccordionWorkflow/FlowJSON';
import NoAccess from './NoAccessScreen';
import ParentWithChildren from './ParentWithChildren/ParentWithChildren';
// import ChildForm from './ChildForm/ChildForm';
import { dataservice } from './encryptionutil';
import { SharePointService, type TemplateRecord } from './services/SharePointService';
import { PdfGeneratorService } from './services/PdfGeneratorService';
import { FileText, Download, Loader } from 'lucide-react';


interface AuthorADdata {
    RequesterName: string;
    Email: string;
    EmployeeID: string;
    Designation: string;
    Department: string;
    ContactNo: string
}
const ViewForm: React.FC<IDynamicFormBuilderProps> = ({ context }) => {
    type Approver = {
        email: string;
        id: number
    }
    const toast = React.useRef<Toast>(null);
    const { id } = useParams<{ id: string }>(); // Retrieve the ID from the route
    const [formData, setFormData] = useState<any[]>([]);
    const [formUpdatedFlag, setformUpdatedFlag] = useState<boolean>(false);
    const [formJSON, setFormJSON] = useState<any>(null);
    const [loading, setLoading] = useState<boolean>(true);
    const [isCurApprover, setIsCurApprover] = useState<boolean>(false); // State to track if the user is CurApprover
    const [, setCurApproverEmail] = useState<string | null>(null);
    const [AppCodevalue, setAppCodevalue] = useState<string>('');
    const [ProcessDataDestiQueue, setProcessDataDestiQueue] = useState<any[]>([]);
    const [AppName, setAppName] = useState<string | null>(null);
    const [selectTab, setSelectTab] = useState("AllDetails");
    const [AuditLog, setAuditLog] = useState<any[]>([]);
    const [attachmentFiles, setAttachmentFiles] = useState<any[]>([]);
    const [childAttachmentFiles, setChildAttachmentFiles] = useState<any[]>([]);
    const [Categorytype, setCategorytype] = useState<string>("");
    const [NotifyRequestorEMails, setNotifyRequestorEMails] = useState<string[]>([]);
    const [SourceFieldJsons, SetSourceFieldJsons] = useState<any[]>([])
    const [ChildFieldJsons, SetChildFieldJsons] = useState<any[]>([])
    const [SrcFieldAnswer, SetSrcFieldAnswer] = useState<any[]>([])
    const [approvedShow, setApprovedShow] = useState("")
    const [comment, setComment] = useState("")
    const [allComments, setAllComments] = useState<any[]>([])
    const [appSeqNo, setAppSeqNo] = useState("")
    const [userName, setUserName] = React.useState("");
    const [userEmail, setUserEmail] = React.useState("");
    const [AllApprovers, setAllApprovers] = useState<Approver[]>([]);
    const [recentStatus, setRecentStatus] = useState("")
    const [WholeRecord, setWholeRecord] = useState<any>()
    const [recentTitle, setRecentTitle] = useState("")
    const [level, setLevel] = useState('0')
    const [, setHighlight] = useState(false);
    const navigate = useNavigate()
    const [mappingLevels, setMappingLevels] = useState<{ Level: string; Role: string; User: string }[]>([]);
    const [currManager, setCurrMnaager] = useState(0)
    const [adminEditAcces, setAdminEditAccess] = useState(false)
    const [RejectButton, setRejectButton] = useState(false)
    const [userRole, setUserRole] = useState("")
    const [RequesterName, setRequesterName] = useState("")
    const [, setRequesterNamee] = useState("")
    const [trnxItem, setTransactionItem] = useState<any>();
    const [selectedFile, setSelectedFile] = useState<File[]>([]);
    const [authorADdata, setAuthorADdata] = useState<AuthorADdata | null>(null);
    const [selectedComment, setSelectedComment] = useState<CommentData | null>(null);
    const [isDialogVisible, setIsDialogVisible] = useState(false);
    const [FinalPopup, setFinalPopup] = useState<boolean>(false);
    const [FinalCost, setFinalCost] = useState<number>(0)
    const [Finaldescription, setFinaldescription] = useState<string>('')
    const [domainfield, setDomainField] = useState<string>("");
    const [preservedfilters, setPreservedFilters] = useState<any>({})
    const [DynamicStatuses, setDynamicStatuses] = useState<string[]>([])
    const [selectedDynamicStatus, setselectedDynamicStatus] = useState<string>("")
    const [maxLevel, setMaxLevel] = useState<number>(0)
    const [pdfanswerjson, setPdfAnswerJson] = useState<any>({});
    const [FlowJSON, setFlowJSON] = useState<NodeItem[]>([])
    const [childappselected, setchildappselected] = useState<string | null>(null);
    const [IsParentForm, setIsParentForm] = useState<boolean>(false);
    const [ChildFormDetails, SetChildFormDetails] = useState<any[]>([]);
    const [ChildtransactionItems, setChildtransactionItems] = useState<any[]>([]);
    const [ParentFormJSONCopy, setParentFormJSONCopy] = useState<any>(null);
    const [ParentFormAnswerCopy, setParentFormAnswerCopy] = useState<any[]>([]);
    const [CurrentChildId, setCurrentChildId] = useState<number>(0);

    // we are checking the nextChild for appCode
    const [nextAppCode, setNextAppCode] = useState<string | null>(null);

    const [filteredChildFormDetails, setFilteredChildFormDetails] = useState<any[]>([]);

    const [isSameDomain, setIsSameDomain] = useState("")
    const [isAdmin, setIsAdmin] = useState(false);
    const [userId, setUserId] = useState<number>();
    const [selectedChildStatus, setSelectedChildStatus] = useState(false);
    // this multiFormSeqNo state for now using for filter the chatbox i.e. communication channel
    const [multiFormSeqNo, setMultiFormSeqNo] = useState("");
    const dataserviceobj = new dataservice();
    const location = useLocation();

    const { setGlobalVariable } = useGlobalState();
    sp.setup({
        sp: {
            baseUrl: mySiteUrl,
        },
    });
    const [tables, setTables] = React.useState<TableData[]>([]);
    const [timesheettables, setTimesheettables] = useState<TimeSheetData[]>([]);
    const [templates, setTemplates] = useState<TemplateRecord[]>([]);
    const [selectedTemplate, setSelectedTemplate] = useState<string>('');
    const [isGeneratingPDF, setIsGeneratingPDF] = useState(false);


    const handleTableUpdate = (updatedTable: TableData) => {
        setTables(prevTables => {
            const updatedTables = prevTables.map(table =>
                table.id === updatedTable.id ? updatedTable : table
            );
            return updatedTables;
        });
        //console.log("tables", tables);
    };
    const handletimesheetTableUpdate = (updatedTable: TimeSheetData) => {
        setTimesheettables(prevTables => {
            const updatedTables = prevTables.map(table =>
                table.id === updatedTable.id ? updatedTable : table
            );
            //console.log('All timetables updated:', updatedTables);
            return updatedTables;
        });
        //console.log("timetables", timesheettables);
    };

    function getPeoplePickerContext(): IPeoplePickerContext {
        return {
            absoluteUrl: context.pageContext.web.absoluteUrl,
            msGraphClientFactory: context.msGraphClientFactory,
            spHttpClient: context.spHttpClient,
        };
    }
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
        return base64;
    };

    const loadAppTemplates = async () => {
        if (!AppCodevalue) {
            console.log("PdfGenerator: No AppCodevalue found, skipping template load.");
            return;
        }
        try {
            console.log(`PdfGenerator: Loading templates for AppCode: "${AppCodevalue}"`);
            const allTemplates = await SharePointService.getAllTemplates();
            const filtered = allTemplates.filter(t =>
                t.AppCode?.toString().trim().toLowerCase() === AppCodevalue.toString().trim().toLowerCase()
            );
            console.log(`PdfGenerator: Found ${filtered.length} matching templates.`);
            setTemplates(filtered);
            if (filtered.length > 0) {
                setSelectedTemplate(filtered[0].id);
            }
        } catch (error) {
            console.error("Error loading templates:", error);
        }
    };

    useEffect(() => {
        if (AppCodevalue) {
            void loadAppTemplates();
        }
    }, [AppCodevalue]);


    const handleGeneratePDF = async () => {
        setIsGeneratingPDF(true);
        try {
            // Re-fetch templates to ensure we have the absolute latest mappings from SharePoint
            await loadAppTemplates();

            if (!selectedTemplate) {
                alert('Please select a template');
                setIsGeneratingPDF(false);
                return;
            }

            const templateToUse = templates.find(t => t.id === selectedTemplate);
            if (!templateToUse) {
                alert('Template not found');
                setIsGeneratingPDF(false);
                return;
            }

            // Prepare Record Data
            let recordData: Record<string, any> = {
                ...formData.reduce((acc, curr) => {
                    const updated = { ...acc };
                    if (curr.id) updated[curr.id] = curr.value;
                    if (curr.name) updated[curr.name] = curr.value;
                    if (curr.field_name) updated[curr.field_name] = curr.value;
                    if (curr.custom_name) updated[curr.custom_name] = curr.value;
                    return updated;
                }, {} as Record<string, any>),
                ...tables.reduce((acc, curr) => {
                    const updated = { ...acc };
                    if (curr.id) updated[curr.id] = curr;
                    return updated;
                }, {} as Record<string, any>),
                ...timesheettables.reduce((acc, curr) => {
                    const updated = { ...acc };
                    if (curr.id) updated[curr.id] = curr;
                    return updated;
                }, {} as Record<string, any>)
            };

            let mergedSchema = [...(formJSON || [])];

            // If it's a parent form, merge data from all child transactions
            if (IsParentForm) {
                console.log(`[PdfGenerator] Parent form detected. Merging data from ${ChildtransactionItems.length} child transactions.`);

                // Merge Schemas first so we can resolve references for child fields
                if (ChildFormDetails && ChildFormDetails.length > 0) {
                    ChildFormDetails.forEach(childForm => {
                        try {
                            const schema = typeof childForm.FormJSON === 'string' ? JSON.parse(childForm.FormJSON) : childForm.FormJSON;
                            if (Array.isArray(schema)) {
                                mergedSchema = [...mergedSchema, ...schema];
                            }
                        } catch (e) {
                            console.error("Error merging child schema:", e);
                        }
                    });
                }

                // Merge Answers
                if (ChildtransactionItems && ChildtransactionItems.length > 0) {
                    ChildtransactionItems.forEach(child => {
                        try {
                            const childFormData = JSON.parse(child.FormData || '[]');
                            const childTableData = JSON.parse(child.TableJSON || '[]');
                            const childTimesheetData = JSON.parse(child.TimeSheetJSON || '[]');

                            childFormData.forEach((curr: any) => {
                                if (curr.id) recordData[curr.id] = curr.value;
                                if (curr.name) recordData[curr.name] = curr.value;
                                if (curr.field_name) recordData[curr.field_name] = curr.value;
                                if (curr.custom_name) recordData[curr.custom_name] = curr.value;
                            });

                            childTableData.forEach((curr: any) => {
                                if (curr.id) recordData[curr.id] = curr;
                            });

                            childTimesheetData.forEach((curr: any) => {
                                if (curr.id) recordData[curr.id] = curr;
                            });
                        } catch (e) {
                            console.error('Error merging child transaction data:', e);
                        }
                    });
                }
            }

            await PdfGeneratorService.generatePDF(
                templateToUse,
                recordData,
                mergedSchema
            );

            alert('PDF Document Generated Successfully!');
        } catch (error) {
            console.error('PDF Generation Error:', error);
            alert('Failed to generate PDF document');
        } finally {
            setIsGeneratingPDF(false);
        }
    };
    useEffect(() => {
        if (Array.isArray(formJSON)) {
            const pdfjson: any = {};
            formJSON.forEach((element: any) => {
                const match = formData.find((e: any) => e.id === element.id);
                if (!match) return;
                const elementType = element.element;
                let readableValue = "";
                if (elementType === "FileUpload") {
                    return;
                }
                if (Array.isArray(match.value)) {
                    readableValue = match.value.map((val: any) => {
                        const option = element.options?.find((opt: any) => opt.key === val);
                        return option ? option.text : val;
                    }).join(",");
                }
                else if (elementType === "Dropdown") {
                    const option = element.options?.find((opt: any) => opt.value === match.value);
                    readableValue = option ? option.text : match.value;
                } else {
                    readableValue = match.value;
                }

                pdfjson[stripHtmlTags(element.label)] = readableValue;

            })
            setPdfAnswerJson(pdfjson);
        }

    }, [formData, formJSON])


    const handleWatermarkedDownload = async () => {
        // Step 1: Create PDF from react-pdf component
        const blob = await pdf(<MyDocument data={pdfanswerjson} loghistory={AuditLog} tables={[...tables, ...timesheettables]} logo={getBase64Logo} txitem={trnxItem} />).toBlob();

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
        a.download = `SmartofficeNxt/${AppName}/${trnxItem.Title}.pdf`;
        a.click();
        URL.revokeObjectURL(url);
    };
    const AnswerJsonSeparator = (ParentCopy: any[], Answers: any[]) => {
        let ParentAnswers: any[] = []
        let ChildAnswers: any[] = []
        Answers.forEach((ele) => {
            ParentCopy.find(elem => elem.id === ele.id) ? ParentAnswers.push(ele) : ChildAnswers.push(ele)
        })

        return [ParentAnswers, ChildAnswers]
    }
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
                //console.log("operation failed ");
                break;
        }
    }
    function stripHtmlTags(label: any) { return label?.replace(/<\/?[^>]+(>|$)/g, "")?.trim() || ""; }
    const funforVisibility = (data: any) => {
        try {
            let filtereddata = data.filter((item: any) => {
                return SourceFieldJsons.find(ele => ele.id === item.id)
            })
            let idboolobjArray = compareValues(filtereddata, SrcFieldAnswer)
            idboolobjArray.forEach((element: any) => {
                if (!element.isSame) {
                    let field = SourceFieldJsons.find(ele => ele.id === element.id)
                    ChildFieldJsons.forEach((childjson) => {
                        if (childjson?.VisibilityCondition[0]?.SourceField === stripHtmlTags(field.label)) {
                            let leftvalue
                            if (childjson?.VisibilityCondition[0]?.element === "RadioButtons") {
                                let optionjson = field.options.find((lfv: any) => {
                                    return lfv.key === element.value[0]
                                });
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
                            let newbool = ConditionValidator(
                                leftvalue,
                                childjson?.VisibilityCondition[0]?.Operator,
                                childjson?.VisibilityCondition[0]?.TargetValue,
                                field.element)

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


                            if ((childjson?.controlVisibility && newbool) || (!childjson?.controlVisibility && !newbool)) {
                                (visiblechangeelement as HTMLElement).style.display = "block"
                            }
                            else if ((!childjson?.controlVisibility && newbool) || (childjson?.controlVisibility && !newbool)) {
                                (visiblechangeelement as HTMLElement).style.display = "none"
                            }
                        }
                    })
                }
            });
            SetSrcFieldAnswer(filtereddata)

        } catch (err) {
            console.log("error in funforvisibility", err)
        }
    }
    const initialfunforvisibility = async () => {
        const transactionItem = await sp.web.lists
            .getByTitle("WorkFlowProcessData")
            .items.getById(Number(id))
            .select("*", "Title, Created, Author/Title,CategoryType,NotifyRequestor/EMail ,Managers/Title")
            .expand("Author", "NotifyRequestor", "Managers")
            .get();
        setTransactionItem(transactionItem)
        const RAWformmasteritem = await sp.web.lists.getByTitle("FormMaster").items.filter(`AppCode eq '${transactionItem.AppCode}'`).top(1).get();
        const formmasteritem = RAWformmasteritem?.map((item: any) => {
            return { ...item, FormJSON: dataserviceobj.decryptjson(item.FormJSON) };
        })
        setDynamicStatuses(JSON.parse(formmasteritem[0].DynamicStatuses) || [])
        const selectedFormJSON = JSON.parse(formmasteritem[0].FormJSON)
        const dependentFields = selectedFormJSON.filter((ele: any) => { //All child fields having conditions
            return ele?.dependentVisibility
        })
        SetChildFieldJsons(dependentFields)
        // const dependentShowFields = selectedFormJSON.filter((ele: any) => { //Child fields which should be hidden initially
        //     return ele?.dependentVisibility && ele?.controlVisibility
        // })
        // console.log("dependentShowFields", dependentShowFields)
        const SourceFieldNames = [...new Set(dependentFields.map((ele: any) => ele?.VisibilityCondition[0]?.SourceField))]
        const SourceFields = selectedFormJSON.filter((ele: any) => {
            return ele?.label && SourceFieldNames.includes(stripHtmlTags(ele?.label))
        })
        SetSourceFieldJsons(SourceFields)
        const transformedArray = SourceFields.map((item: any) => ({
            id: item.id,
            name: item.field_name,
            custom_name: item.field_name,
            value: item.element === "RadioButtons" ? [] : ""
        }));
        SetSrcFieldAnswer(transformedArray)
    }
    const checkDomReady = () => {
        const formContainer = document.querySelector('.react-form-builder-form'); // Adjust selector to match your form container
        if (formContainer) {
            funforVisibility(JSON.parse(trnxItem.FormData || []));
        } else {
            requestAnimationFrame(checkDomReady); // Keep checking until DOM is ready
        }
    };
    useEffect(() => {
        if (ChildFieldJsons.length > 0) {
            requestAnimationFrame(checkDomReady);
        }
    }, [ChildFieldJsons, SourceFieldJsons, SrcFieldAnswer, trnxItem])

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

    const onChange = (items: any[]) => {
        // console.log("Selected People:", items);
    };

    useEffect(() => {
        const applyClassToSelects = () => {
            document.querySelectorAll("select").forEach((el) => {
                el.classList.add("form-select");
            });
        };
        setPreservedFilters(location.state || {})
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
    }, []);



    const fetchUserInfo = async () => {

        try {
            const domain = await fetchTenantUser();
            setDomainField(domain.TenantUsers)
            setUserId(domain.CurrentUserId)
            const currentUser = await sp.web.currentUser();
            let x = currentUser.Title.split("|")[0]
            setUserName(x);
            setUserEmail(currentUser.Email)
        } catch (error) {
            console.error('Error fetching user data', error);
        }
    };

    const checkGroupMembership = async (groupId: number, userId: number) => {
        try {
            const groupMembers = await sp.web.siteGroups.getById(groupId).users.get();
            return groupMembers.some(member => member.Id === userId);
        } catch (error) {
            console.error("Error checking group membership:", error);
            return false;
        }
    };
    const handleChangeDynamicStatus = (e: any) => {
        setselectedDynamicStatus(e.target.value)
    }
    const handleSaveDynamicStatus = async () => {
        if (selectedDynamicStatus.trim().length === 0 || !DynamicStatuses.includes(selectedDynamicStatus)) {
            toast?.current?.clear()
            toast?.current?.show({
                severity: 'error',
                // summary: 'No value',
                detail: 'Selet a status to Save',
                life: 3000
            })
        } else {

            try {
                let Id = CurrentChildId && CurrentChildId ? CurrentChildId : id
                await sp.web.lists.getByTitle("WorkFlowProcessData").items.getById(Number(Id)).update({
                    Status: selectedDynamicStatus
                })
                setRecentStatus(selectedDynamicStatus)
                // await sp.web.lists.getByTitle('WorkFlowLog').items.add()
                toast?.current?.clear()
                toast?.current?.show({
                    severity: 'success',
                    // summary: 'No value',
                    detail: 'Status Updated',
                    life: 3000
                })
            }
            catch (err) {
                //console.log("error in status update", err)
                toast?.current?.clear()
                toast?.current?.show({
                    severity: 'error',
                    // summary: 'No value',
                    detail: 'Error in updating status.Please try again',
                    life: 3000
                })
            }
        }
    }

    const isAdminUser = async () => {
        const currentUser = await sp.web.currentUser.get();
        const authUsers = await sp.web.lists.getByTitle('AuthList').items
            .filter(`AdminUser eq 'Yes' and AuthName/Id eq ${currentUser.Id}`)
            .select('AuthName/Id')
            .expand('AuthName')
            .getAll();
        const isAdminResult = authUsers.length > 0;
        //console.log("📋 AuthList query result:", authUsers);
        //console.log("✅ Is Admin:", isAdminResult);

        setIsAdmin(isAdminResult);
    };


    const fetchFormData = async () => {
        setLoading(true);
        // setAdminEditAccess(false)
        try {

            const transactionItem = await sp.web.lists.getByTitle("WorkFlowProcessData").items.getById(Number(id)).select("*", "Title, Created,Domain,AuthorId,Author/Title,NotifyRequestor/EMail,TableJSON,TimeSheetJSON,Author/Id,CurApprover/Title,CurrentAppCode")
                .expand("Author", "NotifyRequestor", "CurApprover").get();


            const ChildtransactionItems = await sp.web.lists.getByTitle("WorkFlowProcessData")
                .items.filter(`ParentAppId eq '${transactionItem.Id}'`)
                .select("*", "Title, Created,Domain,CurrentAppCode,AuthorId, Author/Title,CategoryType,NotifyRequestor/EMail,TableJSON,TimeSheetJSON,CurApprover/Title")
                .expand("Author", "NotifyRequestor", "CurApprover")
                .orderBy("Id", true)
                .get();
            //console.log("🔍 Child Transaction Items:", ChildtransactionItems.length);
            ChildtransactionItems.forEach((item, index) => {
                //console.log(`  Child ${index + 1}: ID=${item.ID}, AppCode=${item.AppCode}, Status=${item.Status}, Title=${item.Title}`);
                //console.log(`    AuthorADdata: ${item.AuthorADdata ? 'EXISTS' : 'NULL/EMPTY'}`);
            });
            setChildtransactionItems(ChildtransactionItems)

            const CurrentChild = ChildtransactionItems.find(
                (child) => child.AppCode === transactionItem.CurrentAppCode
            );
            // setAppSeqNo(transactionItem.SeqNo)

            if (CurrentChild) {
                setCurrentChildId(CurrentChild.ID);
                // const appSeqNo = 
                // setAppSeqNo(CurrentChild.SeqNo);
                //console.log("✅ Matching child ID:", CurrentChild.ID);
            } else {
                //console.log("⚠️ No matching child found for CurrentAppCode:", transactionItem.CurrentAppCode);
            }


            setTables(JSON.parse(transactionItem.TableJSON) || []) // Assuming TableJSON is a JSON string
            setTimesheettables(JSON.parse(transactionItem.TimeSheetJSON) || [])
            if (transactionItem.NotifyRequestor && transactionItem.NotifyRequestor.length > 0) {
                const userEmails = transactionItem.NotifyRequestor.map((user: { EMail: any; }) => user?.EMail);
                setNotifyRequestorEMails(userEmails); // Save emails in state
            }
            //console.log("transactionItem from view form", transactionItem)
            await isAdminUser()
            //debugger;
            setApprovedShow(transactionItem.Status)
            setCategorytype(transactionItem.CategoryType)
            setGlobalVariable(JSON.stringify({ appcode: transactionItem.AppCode, appname: transactionItem.AppName }));
            const parsedFormData = JSON.parse(transactionItem.FormData);
            const updatedFormData = parsedFormData.map((item: any) => {
                if (item.name.startsWith("file_upload")) {
                    return {
                        ...item,
                        value: item.value ? item.value : "No file",
                    };
                }
                return item;
            });
            setFormData(updatedFormData);
            setParentFormAnswerCopy(updatedFormData);
            setWholeRecord(transactionItem)
            setAppName(transactionItem.AppName)


            const appCode = transactionItem.AppCode;
            // setAppCodevalue(appCode);

            //Checking if CurrentAppCode exists in the WorkFlowProcessData then use it else use AppCode column data
            const DynamicAppCode = transactionItem?.CurrentAppCode ?? appCode;
            setAppCodevalue(DynamicAppCode);

            const curApproverId = transactionItem.CurApproverId;
            const ProcessDataDestiQueue = transactionItem.DestinationQueue;
            setCurApproverEmail(curApproverId);

            setProcessDataDestiQueue(ProcessDataDestiQueue)
            // console.log("CurAppover email from list ", curApproverEmail)

            setCurrMnaager(transactionItem.ManagersId)

            setRecentStatus(transactionItem.Status)
            setselectedDynamicStatus(transactionItem.Status)
            setRecentTitle(transactionItem.Title)
            setLevel(transactionItem.Level)
            setRequesterName("Requester : " + transactionItem.Author.Title.split("|")[0]);
            setRequesterNamee(transactionItem.Author.Title.split("|")[0])
            const approvers = transactionItem.AllApprovers
                ? JSON.parse(transactionItem.AllApprovers)
                : [];

            // Ensure the parsed data is an array
            if (Array.isArray(approvers)) {
                setAllApprovers(approvers as Approver[]);
            } else {
                setAllApprovers([]);
            }
            const authorDataString = transactionItem.AuthorADdata; // Adjust the field name to 'AuthorADdata'
            const parsedAuthorData: AuthorADdata = JSON.parse(authorDataString);

            // Set the parsed data into the state
            setAuthorADdata(parsedAuthorData);

            // Fetch levels from MappingMaster list for the specific AppCode
            const mappingMasterData = await sp.web.lists
                .getByTitle("MappingMaster")
                .items.select("Level", "Role", "Users_x002f_Groups/Title", "Users_x002f_Groups/Id", "CanEdit")
                .expand("Users_x002f_Groups")
                .filter(`AppCode eq '${DynamicAppCode}'`)
                .get();
            let highestLevel
            if (mappingMasterData.length > 0) {
                highestLevel = Math.max(...mappingMasterData.map(item => item.Level));
                setMaxLevel(highestLevel)
            } else {
                //console.log("Highest level did not set for skip button visibility case")
            }
            const conditionMasterData = await sp.web.lists
                .getByTitle("ConditionsList")
                .items.select("Level", "Role", "PersonOrGroup/Title", "PersonOrGroup/Id")
                .expand("PersonOrGroup")
                .filter(`AppCode eq '${DynamicAppCode}'`)
                .get()
            const levelsWithRoles = mappingMasterData.map(item => ({
                Level: item.Level,
                Role: item.Role,
                User: item.Users_x002f_Groups ? item.Users_x002f_Groups.Title : "N/A",

            })).sort((a, b) => a.Level - b.Level);

            // setAdminEditAccess(mappingMasterData.CanEdit)
            let userRolevar
            for (let i = 0; i < mappingMasterData.length; i++) {


                if (
                    mappingMasterData[i]?.CanEdit && // Ensure CanEdit exists
                    (mappingMasterData[i].CanEdit === "true" || mappingMasterData[i].CanEdit === "True") &&
                    mappingMasterData[i]?.Users_x002f_Groups?.Id && // Ensure Users_x002f_Groups and Id exist
                    mappingMasterData[i].Users_x002f_Groups.Id === curApproverId
                ) {
                    setAdminEditAccess(true);
                    break;
                } else {
                    setAdminEditAccess(false);
                }


                if (mappingMasterData[i]?.Users_x002f_Groups?.Id !== undefined && mappingMasterData[i]?.Users_x002f_Groups?.Id !== null) {
                    if (mappingMasterData[i].Level === transactionItem.Level && mappingMasterData[i].Users_x002f_Groups.Id === curApproverId && mappingMasterData[i].Role === "Creator") {
                        setRejectButton(true);
                        userRolevar = "Creator"
                    }
                    if (mappingMasterData[i].Level === transactionItem.Level && mappingMasterData[i].Users_x002f_Groups.Id === curApproverId && mappingMasterData[i].Role === "Reviewer") {
                        setRejectButton(true);
                        userRolevar = "Reviewer"
                    }
                    if (mappingMasterData[i].Level === transactionItem.Level && mappingMasterData[i].Users_x002f_Groups.Id === curApproverId && mappingMasterData[i].Role === "Task Performer") {
                        setRejectButton(true);
                        userRolevar = "Task Performer"
                    }
                    if (mappingMasterData[i].Level === transactionItem.Level && mappingMasterData[i].Users_x002f_Groups.Id === curApproverId && mappingMasterData[i].Role === "Approver") {
                        userRolevar = "Approver"
                    }
                }
            }
            for (let i = 0; i < conditionMasterData.length; i++) {

                if (conditionMasterData[i]?.PersonOrGroup?.Id !== undefined && conditionMasterData[i]?.PersonOrGroup?.Id !== null) {
                    if (conditionMasterData[i].Level === transactionItem.Level && conditionMasterData[i].PersonOrGroup.Id === curApproverId && conditionMasterData[i].Role === "Creator") {
                        setRejectButton(true);
                        userRolevar = "Creator"
                    }
                    if (conditionMasterData[i].Level === transactionItem.Level && conditionMasterData[i].PersonOrGroup.Id === curApproverId && conditionMasterData[i].Role === "Reviewer") {
                        setRejectButton(true);
                        userRolevar = "Reviewer"
                    }
                    if (conditionMasterData[i].Level === transactionItem.Level && conditionMasterData[i].PersonOrGroup.Id === curApproverId && conditionMasterData[i].Role === "Task Performer") {
                        setRejectButton(true);
                        userRolevar = "Task Performer"
                    }
                    if (conditionMasterData[i].Level === transactionItem.Level && conditionMasterData[i].PersonOrGroup.Id === curApproverId && conditionMasterData[i].Role === "Approver") {
                        userRolevar = "Approver"
                    }
                }
            }
            setUserRole(userRolevar || "")
            setMappingLevels(levelsWithRoles);
            //debugger;
            //console.log("childappselected in fetchFormdata", CurrentChildId);
            const thetransactionItem = (transactionItem.CurrentAppCode) ? (CurrentChild) ? CurrentChild : transactionItem : transactionItem;

            let endflow: boolean = false;
            if ((thetransactionItem.Status === "In-Progress" && Number(thetransactionItem.CurrentQueue) === highestLevel && userRolevar === "Creator") ||
                (thetransactionItem.Status === "In-Progress" && Number(thetransactionItem.CurrentQueue) === highestLevel && userRolevar === "Task Performer") ||
                (thetransactionItem.Status === "In-Progress" && Number(thetransactionItem.DestinationQueue) === highestLevel && userRolevar === "Task Performer")) {
                endflow = true;
            }

            const flowjson = await FlowData(thetransactionItem, levelsWithRoles, endflow);
            //console.log("flowjson------------------------------------------------------------------------", flowjson);
            setFlowJSON(flowjson);
            const appSeqNo = thetransactionItem.SeqNo;
            setAppSeqNo(appSeqNo);



            // Fetch the corresponding FormJSON from the FormMaster list based on AppName
            const RAWformMasterItems = await sp.web.lists.getByTitle("FormMaster").items.filter(`AppCode eq '${appCode}'`).get();
            const formMasterItems = RAWformMasterItems?.map((item: any) => {
                return { ...item, FormJSON: dataserviceobj.decryptjson(item.FormJSON) };
            })
            const Isparentform = formMasterItems[0].IsParentForm;
            setIsParentForm(Isparentform);

            if (formMasterItems[0].IsParentForm === true) {
                let RAWChildforms = await sp.web.lists.getByTitle("FormMaster").items.filter(`ParentAppCode eq '${appCode}'`).get();
                const Childforms = RAWChildforms?.map((item: any) => {
                    return { ...item, FormJSON: dataserviceobj.decryptjson(item.FormJSON) };
                })
                SetChildFormDetails(Childforms);

                // Only run if Childforms exist
                if (Childforms && Childforms.length > 0) {

                    // Current child AppCode from transaction
                    const currentAppCode = transactionItem.CurrentAppCode;


                    // Find current child in Childforms
                    const currentForm = Childforms.find(
                        (child) => child.AppCode === currentAppCode
                    );

                    if (currentForm) {
                        const currentOrder = currentForm.ChildOrder;

                        // Find the next child form (ChildOrder = currentOrder + 1)
                        const nextForm = Childforms.find(
                            (child) => child.ChildOrder === currentOrder + 1
                        );

                        if (nextForm) {
                            setNextAppCode(nextForm.AppCode);
                            //console.log("Next child form:", nextForm.AppCode);
                        } else {
                            setNextAppCode('');
                            //console.log(" This is the last child form.");
                        }
                    } else {
                        //console.log("Current child not found in Childforms array.");
                    }
                }

            }
            if (formMasterItems.length > 0) {
                //debugger
                const formMasterItem = formMasterItems[0];
                // const IsParentForm  = formMasterItem.IsParentForm

                // console.log("IsParentFormmmmmmmmmmmmmmmmmmmmm",IsParentForm);

                // Get Master JSON (Form Structure)
                const parsedFormJSON = JSON.parse(formMasterItem.FormJSON);
                setFormJSON(parsedFormJSON);
                setParentFormJSONCopy(parsedFormJSON);
                // Map parsed form data for quick lookup
                // const parsedFormDataMap = new Map(
                //     parsedFormData
                //         .filter((dataItem: any) => dataItem.name.startsWith("file_upload"))
                //         .map((dataItem: any) => [dataItem.name, dataItem.value])
                // );
                // console.log("parsedFormDataMap---------------------", parsedFormDataMap)
                // -------- starting code for multi file attachment
                // Fetch files from the NoteAttach document library
                const seqNoFolder = transactionItem.SeqNo; // Assuming SeqNo is available from the transaction item
                const files = await sp.web.lists.getByTitle("NoteAttach").items.select("FileLeafRef", "FileRef", "FileDirRef").filter(`FileDirRef eq '${fileFetchUrl}/NoteAttach/${seqNoFolder}'`).get();

                const fileNames = files.map((file, index) => {
                    const fileName = file.FileLeafRef || file.Title || "Unnamed File"; // Use FileLeafRef or Title if available
                    const fileUrl = file.FileRef;
                    return { name: `${index + 1}. ${fileName}`, url: fileUrl }; // Return an object with name and URL
                });
                setAttachmentFiles(fileNames);

                // Create modified JSON structure
                const modifiedFormJSON = parsedFormJSON.map((item: any) => {
                    if (item.element === "FileUpload") {
                        // const filename = fileNames.map((item) => item.name.replace(/^\d+\.\s*/, '')).join(', ')
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
                setFormJSON(modifiedFormJSON);
                setParentFormJSONCopy(modifiedFormJSON);
            }

            // filtering who is approver-----------------------------------
            const currentUser = await sp.web.currentUser.get();
            const userId = currentUser.Id;
            // if (userId === curApproverId) {
            //     setIsCurApprover(true);
            // }
            if (!Isparentform) {
                for (const mappingItem of mappingMasterData) {
                    const groupOrUserId = mappingItem.Users_x002f_Groups?.Id;
                    if (groupOrUserId) {
                        // Check if it's a SharePoint group
                        const isGroup = await sp.web.siteGroups.getById(groupOrUserId).get().then(() => true).catch(() => false);
                        if (isGroup && groupOrUserId === curApproverId) {
                            // Check if user is a member of the group
                            const isMember = await checkGroupMembership(groupOrUserId, userId);
                            if (isMember) {
                                setIsCurApprover(true);
                                break;
                            }
                        } else if (curApproverId === userId) {
                            // Direct user ID match
                            setIsCurApprover(true);
                            break;
                        }
                    }
                }
            }
            const WorkFlowLog = await sp.web.lists.getByTitle("WorkFlowAuditLog").items.filter(`(AppCode eq '${transactionItem.AppCode}') and (SeqNo eq '${transactionItem.SeqNo}')`).get();
            setAuditLog(WorkFlowLog);
            // Fetch files from the NoteAttach document library
            const seqNoFolder = transactionItem.SeqNo; // Assuming SeqNo is available from the transaction item
            const files = await sp.web.lists.getByTitle("NoteAttach").items.select("FileLeafRef", "FileRef", "FileDirRef").filter(`FileDirRef eq '${fileFetchUrl}/NoteAttach/${seqNoFolder}'`).get();
            const fileNames = files.map((file, index) => {
                const fileName = file.FileLeafRef || file.Title || "Unnamed File"; // Use FileLeafRef or Title if available
                // Construct the full URL without duplicating the site path
                // const fileUrl = `${mySiteUrl}${file.FileRef.replace(/^\/sites\/Smartofficetst\/SmartOfficeREDEV/i, '')}`;    //tst
                // const fileUrl = `${mySiteUrl}${file.FileRef.replace(/^\/sites\/Smartoffice\/SmartOfficenxt/i, '')}`;            //Prod
                // const fileUrl = `${mySiteUrl}${file.FileRef.replace(/^\/sites\/SmartOfficeNxtDemo/i, '')}`;                     // Demo
                const fileUrl = file.FileRef;
                return { name: `${index + 1}. ${fileName}`, url: fileUrl }; // Return an object with name and URL
            });
            setAttachmentFiles(fileNames);
            // comment -------------------------------------------
            try {
                // Fetch data from the CommentsLog list
                const ComApp = IsParentForm ? childappselected : appCode;
                const CommentItems = await sp.web.lists
                    .getByTitle("CommentsLog")
                    .items.filter(`(AppCode eq '${ComApp}') and (SeqNo eq '${appSeqNo}')`)
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
                        "Modified",
                    ) // Select required fields
                    .get();
                // Helper function to remove HTML from communication
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
                setAllComments(processedComments);
            } catch (err) {
                console.error("Error fetching communication with all columns:", err);
            }

        } catch (error) {
            console.error("Error fetching form data:", error);
        } finally {
            setLoading(false);
        }
    };

    const validateFileSize = (file: File) => {
        const maxSizeInMB = 3;
        const fileSizeInMB = file.size / (1024 * 1024);
        return fileSizeInMB <= maxSizeInMB;
    };
    const onFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
        if (event.target.files) {
            const newFiles = Array.from(event.target.files);
            const validFiles: File[] = [];
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
    // const fetchCommunicationLog = async () => {
    //     try {
    //         const items = await sp.web.lists.getByTitle("CommunicationLog")
    //             .items.select("Title", "Chat", "Author/Title", "Created", "MessageTo/Title", "MessageTo/ID", "MessageTo/EMail")
    //             .expand("Author", "MessageTo")
    //             .filter(`Title eq '${appSeqNo}'`)
    //             .get();
    //         console.log("Filtered Comments:", items);
    //     } catch (error) {
    //         console.error("Error fetching communication:", error);
    //     }
    // };
    // Function to check if user is part of a child form's workflow
    // const isUserInWorkflow = async (childAppCode: string, userId: number): Promise<boolean> => {
    //     try {
    //         // Get all workflow mappings for this child form
    //         const mappings = await sp.web.lists.getByTitle("MappingMaster")
    //             .items.select("Users_x002f_Groups/Id")
    //             .expand("Users_x002f_Groups")
    //             .filter(`AppCode eq '${childAppCode}'`)
    //             .get();

    //         for (const mapping of mappings) {
    //             const groupOrUserId = mapping.Users_x002f_Groups?.Id;
    //             if (!groupOrUserId) continue;

    //             // Check if it's a direct user match
    //             if (groupOrUserId === userId) {
    //                 return true;
    //             }

    //             // Check if it's a group and user is member
    //             try {
    //                 const isGroup = await sp.web.siteGroups.getById(groupOrUserId).get().then(() => true).catch(() => false);
    //                 if (isGroup) {
    //                     const groupUsers = await sp.web.siteGroups.getById(groupOrUserId).users.get();
    //                     if (groupUsers.some((user: any) => user.Id === userId)) {
    //                         return true;
    //                     }
    //                 }
    //             } catch (error) {
    //                 // Not a group, skip
    //                 continue;
    //             }
    //         }
    //         return false;
    //     } catch (error) {
    //         console.error("Error checking user in workflow:", error);
    //         return false;
    //     }
    // };

    // Function to check if all previous child form workflows are completed
    // const arePreviousWorkflowsCompleted = (childFormDetails: any[], childTransactions: any[], targetChildOrder: number): boolean => {
    //     // Get all child forms with lower order than target
    //     const previousChildren = childFormDetails.filter(child => child.ChildOrder < targetChildOrder);

    //     console.log(`🔍 Checking if previous workflows completed for ChildOrder ${targetChildOrder}`);
    //     console.log(`Previous children to check:`, previousChildren.map(c => `${c.AppName} (Order: ${c.ChildOrder})`));

    //     // Check if all previous child forms have completed workflows
    //     for (const prevChild of previousChildren) {
    //         // Try both CurrentAppCode and AppCode for matching
    //         const transaction = childTransactions.find(t =>
    //             t.CurrentAppCode === prevChild.AppCode || t.AppCode === prevChild.AppCode
    //         );

    //         if (!transaction) {
    //             // No transaction exists for this child form - workflow not started/completed
    //             console.log(`❌ No transaction found for ${prevChild.AppName} - previous workflow not complete`);
    //             return false;
    //         }

    //         // Check if workflow is completed (Status should be final approval status)
    //         const status = (transaction.Status || transaction.RecentStatus || '').toLowerCase();
    //         const isCompleted = status === "approved" || status === "completed" || status === "closed";

    //         console.log(`📋 ${prevChild.AppName} status: "${status}" - Completed: ${isCompleted}`);

    //         if (!isCompleted) {
    //             console.log(`❌ ${prevChild.AppName} workflow not completed yet`);
    //             return false;
    //         }
    //     }

    //     console.log(`✅ All previous workflows are completed`);
    //     return true;
    // };

    const filterVisibleChildren = async (
        childFormDetails: any[],
        childTransactions: any[],
        isAdminUser: boolean,
        parentTransaction: any
    ): Promise<any[]> => {
        debugger;
        //console.log("=== FILTERING CHILD FORMS WITH USER WORKFLOW CHECK ===");
        //console.log("Total child forms:", childFormDetails.length);
        //console.log("Total child transactions:", childTransactions.length);
        //console.log("Is Admin User:", isAdmin);

        // If admin, show all child forms
        const initial: string[] = [];
        ChildtransactionItems.forEach((item) => {
            initial.push(item.AppCode);
        });
        // if (!initial.includes(WholeRecord.CurrentAppCode)) { initial.push(WholeRecord.CurrentAppCode) }
        const adminapps = ChildFormDetails.filter((child) => initial.includes(child.AppCode));
        if (isAdminUser) {
            //console.log("✅ Admin user - showing all child forms");
            return adminapps;
        }

        const currentUser = await sp.web.currentUser.get();
        const userId = currentUser.Id;

        const visibleChildren: any[] = [];

        // Sort child forms by ChildOrder
        const sortedChildren = [...childFormDetails].sort((a, b) => a.ChildOrder - b.ChildOrder);

        for (const child of sortedChildren) {
            //console.log(`\n🔎 Checking child form: ${child.AppName} (Order: ${child.ChildOrder}, AppCode: ${child.AppCode})`);
            const childNotifyToId = child.NotifyToId;
            // Find corresponding transaction for this child (if any)
            const correspondingChildTx = childTransactions.find(t => t.AppCode === child.AppCode || t.CurrentAppCode === child.AppCode);

            if (!correspondingChildTx) {
                //console.log(`❌ HIDDEN: No transaction exists yet for ${child.AppName}`);
                // const mappingMasterData1 = await sp.web.lists
                //                 .getByTitle("MappingMaster")
                //                 .items.select("Level", "Role", "Users_x002f_Groups/Title", "Users_x002f_Groups/Id", "CanEdit")
                //                 .expand("Users_x002f_Groups")
                //                 .filter(`AppCode eq '${child.AppCode}' and Level eq 0`)
                //                 .get();
                //             for (const mappingItem of mappingMasterData1) {
                //                     const groupOrUserId = mappingItem.Users_x002f_Groups?.Id;
                //                     if (groupOrUserId) {
                //                         // Check if it's a SharePoint group
                //                         const isGroup = await sp.web.siteGroups.getById(groupOrUserId).get().then(() => true).catch(() => false);
                //                         const isnotifygroup = await sp.web.siteGroups.getById(child.NotifyToId).get().then(() => true).catch(() => false);
                //                         if (isGroup  || isnotifygroup) {

                //                             // Check if user is a member of the group
                //                             const isMember = await checkGroupMembership(groupOrUserId, Number(userId)) 
                //                             const isnotifytoid = (isnotifygroup) ? await checkGroupMembership(child.NotifyToId, Number(userId)) : false;
                //                             if (isMember || isnotifytoid) {
                //                                 visibleChildren.push(child);
                //                             }
                //                             //if the id is not group then check if current approver the creator of next child

                //                         }    
                //                         else{
                //                             if(groupOrUserId === trnxItem?.CurApproverId || child.NotifyToId === userId){
                //                                 visibleChildren.push(child);
                //                         }                                   
                //                     }
                // } 
                continue;
                // }
            }
            else {
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
                const notifyrequestorId = trnxItem.NotifyRequestorId;
                let isMember = false;
                let isfinalmember = false;
                const isgroup = await sp.web.siteGroups.getById(notifyrequestorId).get().then(() => true).catch(() => false);
                const isnotifygroup = await sp.web.siteGroups.getById(childNotifyToId).get().then(() => true).catch(() => false);
                if (isgroup) {
                    isMember = await checkGroupMembership(notifyrequestorId, Number(userId))
                }
                else {
                    isMember = Number(notifyrequestorId) === Number(userId);
                }
                if (isnotifygroup) {
                    isfinalmember = await checkGroupMembership(childNotifyToId, Number(userId))
                }

                else {
                    isfinalmember = Number(childNotifyToId) === Number(userId);
                }
                const isTransactionParticipant = (txAuthorId && Number(txAuthorId) === Number(userId)) || (txCurApproverId && Number(txCurApproverId) === Number(userId)) || (txAllApproversEmails && txAllApproversEmails.includes(currentUserEmail) || isMember || isfinalmember);

                // Fallback to mapping membership if needed
                const isUserInvolved = isTransactionParticipant
                //|| await isUserInWorkflow(child.AppCode, userId);

                if (!isUserInvolved) {
                    //console.log(`❌ HIDDEN: User not a participant in transaction or mapping for ${child.AppName}`);

                    continue;
                }

                //console.log(`✓ User is participant for ${child.AppName}`);

                // Check if all previous workflows are completed
                // const previousCompleted = arePreviousWorkflowsCompleted(childFormDetails, childTransactions, child.ChildOrder);

                // if (!previousCompleted) {
                //     console.log(`❌ HIDDEN: Previous workflows not completed for ${child.AppName}`);
                //     continue;
                // }

                // Special-case for Form 4: require parent workflow and this child transaction to be completed
                // if (child.AppName && child.AppName.toString().toLowerCase().includes('form 4')) {
                //     const parentStatus = (parentTransaction?.Status || '').toString().toLowerCase();
                //     const parentCompleted = parentStatus === 'completed' || parentStatus === 'approved' || parentStatus === 'closed';
                //     if (!parentCompleted) {
                //         console.log(`❌ HIDDEN: Parent workflow not completed for ${child.AppName}`);
                //         continue;
                //     }

                //     const correspondingChildTx = childTransactions.find(t => t.AppCode === child.AppCode || t.CurrentAppCode === child.AppCode);
                //     const childStatus = (correspondingChildTx?.Status || '').toString().toLowerCase();
                //     const childCompleted = childStatus === 'completed' || childStatus === 'approved' || childStatus === 'closed';
                //     if (!childCompleted) {
                //         console.log(`❌ HIDDEN: ${child.AppName} transaction is not completed (status: ${childStatus})`);
                //         continue;
                //     }
                //     console.log(`✅ Special-case passed: Parent and ${child.AppName} workflows completed`);
                // }

                //console.log(`✅ VISIBLE: Showing ${child.AppName} to user`);
                visibleChildren.push(child);
            }
        }

        return visibleChildren;
    };


    // Function to update FlowJSON, Requester Details, and Comments when child tab changes
    const updateChildTabData = async (childFormAppCode: string, childOrder: number | string, ChildFormJson: any) => {
        try {
            // console.log(`📊 Updating child tab data for: ${childFormAppCode}, ChildOrder: ${childOrder}`);

            // Find the correct transaction by ChildOrder
            //console.log("childformapcode: ", childFormAppCode);
            const childTransaction = ChildtransactionItems[Number(childOrder) - 1];
            //    const childTransaction = await ChildtransactionItems.filter(t =>{console.log(t.AppCode); t.AppCode === childFormAppCode})[0];
            //debugger;
            if (!childTransaction) {
                //console.log("❌ No transaction found for ChildOrder:", childOrder);

                setAuthorADdata(null);
                setAllComments([]);
                setFlowJSON([]);
                setChildAttachmentFiles([]);


            }
            else {


                //console.log(`✓ Transaction found - ID: ${childTransaction.Id}, AppCode: ${childTransaction.AppCode}, SeqNo: ${childTransaction.SeqNo}`);
                //console.log(`  Status: ${childTransaction.Status}`);
                //console.log(`  Has AuthorADdata: ${!!childTransaction.AuthorADdata}`);

                // Update Requester Details (AuthorADdata)
                if (childTransaction.AuthorADdata && childTransaction.AuthorADdata.trim() !== "") {
                    //console.log(`✓ Updating requester details for child form`);
                    const parsedAuthorData = JSON.parse(childTransaction.AuthorADdata);
                    setAuthorADdata(parsedAuthorData);
                } else {
                    //console.log(`⚠️ No requester details available for this child form`);
                    setAuthorADdata(null);
                }
                const seqNoFolder = childTransaction.SeqNo; // Assuming SeqNo is available from the transaction item
                const files = await sp.web.lists.getByTitle("NoteAttach").items.select("FileLeafRef", "FileRef", "FileDirRef").filter(`FileDirRef eq '${fileFetchUrl}/NoteAttach/${seqNoFolder}'`).get();

                const fileNamesss = files.map((file, index) => {
                    const fileName = file.FileLeafRef || file.Title || "Unnamed File"; // Use FileLeafRef or Title if available
                    const fileUrl = file.FileRef;
                    return { name: `${index + 1}. ${fileName}`, url: fileUrl }; // Return an object with name and URL
                });
                setChildAttachmentFiles(fileNamesss);

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
                //console.log(`📝 Fetching comments for AppCode: ${childTransaction.AppCode}, SeqNo: ${childTransaction.SeqNo}`);
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

                //console.log(`✓ Found ${CommentItems.length} comments`);
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

                // === Fetch FlowJSON for the child transaction ===
                //console.log(`🔄 Fetching FlowJSON for AppCode: ${childTransaction.AppCode}`);
                let mappingMasterData = await sp.web.lists
                    .getByTitle("MappingMaster")
                    .items.select("Level", "Role", "Users_x002f_Groups/Title")
                    .expand("Users_x002f_Groups")
                    .filter(`AppCode eq '${childTransaction.AppCode}'`)
                    .get();

                // Fallback to FormMaster AppCode if needed
                if (mappingMasterData.length === 0) {
                    //console.log(`  Trying fallback with FormMaster AppCode: ${childFormAppCode}`);
                    mappingMasterData = await sp.web.lists
                        .getByTitle("MappingMaster")
                        .items.select("Level", "Role", "Users_x002f_Groups/Title")
                        .expand("Users_x002f_Groups")
                        .filter(`AppCode eq '${childFormAppCode}'`)
                        .get();
                    //console.log(`⚠️ No MappingMaster data found — setting empty FlowJSON`);
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

                    //console.log(`  Generating FlowJSON with endflow: ${endflow}`);
                    const flowjson = await FlowData(childTransaction, levelsWithRoles, endflow);
                    //console.log(`  Generated FlowJSON with ${flowjson.length} nodes`);
                    setFlowJSON(flowjson);
                }

                //console.log(`✅ Child tab data updated successfully for ${childFormAppCode} (Transaction AppCode: ${childTransaction.AppCode})`);
                return childmodifiedFormJSON;
            }
        } catch (error) {

            console.error("Error updating child tab data:", error);
        }
    };


    useEffect(() => {
        void fetchUserInfo();
        void fetchFormData();
        // void fetchCommunicationLog();
        // void fetchFiles();
    }, [id]);

    // Filter child forms when ChildFormDetails, ChildtransactionItems, or isAdmin change
    useEffect(() => {
        const applyChildFiltering = async () => {
            //console.log("🔄 Child filtering useEffect triggered");
            //console.log("ChildFormDetails.length:", ChildFormDetails.length);
            //console.log("ChildtransactionItems.length:", ChildtransactionItems.length);
            //console.log("isAdmin state in useeffect:", isAdmin);

            if (ChildFormDetails.length > 0 && ChildtransactionItems.length > 0) {
                const visibleChildren = await filterVisibleChildren(
                    ChildFormDetails,
                    ChildtransactionItems,
                    isAdmin,
                    // false,
                    WholeRecord
                );
                setFilteredChildFormDetails(visibleChildren);
            }
        };

        void applyChildFiltering();
    }, [ChildFormDetails, ChildtransactionItems, isAdmin]);


    useEffect(() => {
        let TransformedTableJsons = [...tables]
        TransformedTableJsons.forEach((field: any) => {
            const labelNodes = Array.from(document.querySelectorAll("label.static.form-label"));
            const targetLabel = labelNodes.find(label =>
                label.innerHTML.includes(field.content)
            );
            if (targetLabel && targetLabel.parentElement) {
                const parent = targetLabel.parentElement;
                targetLabel.remove()
                ReactDOM.render(
                    <TableComponent tableConfig={field} onTableUpdate={handleTableUpdate} readonly={true} seqno={appSeqNo} />,
                    parent
                );
            }
        });
    })

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
                    <TimesheetComponent tableConfig={field} onTableUpdate={handletimesheetTableUpdate} readonly={true} seqno={appSeqNo}
                    />,
                    parent
                );
            }
        });
    })

    const lastfunction = async () => {
        setLoading(true);
        const trimmedComment = comment?.trim() || "";
        const isFilePresent = selectedFile && selectedFile.length > 0;
        if (isFilePresent && trimmedComment.length === 0) {
            setLoading(false);
            toast.current?.show({
                severity: 'warn',
                summary: '',
                detail: 'Please enter a comment when uploading a file.',
                life: 2000
            });
            setSelectTab('AllDetails')
            setHighlight(true);
            return;
        }
        if (trimmedComment.length > 0) {
            debugger;
            const Approvers = [...AllApprovers, userEmail]

            if (IsParentForm === true) {

                try {
                    let childAllApprovers = ChildtransactionItems.find((item) => item.AppCode === WholeRecord.CurrentAppCode)?.AllApprovers;
                    childAllApprovers = JSON.parse(childAllApprovers);
                    childAllApprovers = [...childAllApprovers, userEmail];

                    //debugger;

                    let varStatusss: string;
                    let varDestQueuee: string;
                    let varCurApproverId: number | null;
                    let CurrentAppCode: string | null;
                    let nextAppCreator: number | null = null;
                    //debugger
                    if (nextAppCode) {

                        const nextAppCreatorData = await sp.web.lists.getByTitle("MappingMaster").items
                            .select("*", "Users_x002f_Groups/Id", "Users_x002f_Groups/Title")
                            .expand("Users_x002f_Groups")
                            .filter(`(Level eq 0) and (AppCode eq '${nextAppCode}')`)
                            .get();

                        if (nextAppCreatorData.length > 0) {
                            nextAppCreator = nextAppCreatorData[0].Users_x002f_GroupsId || null;
                        }

                        varStatusss = "In-Progress";
                        varDestQueuee = "0";
                        varCurApproverId = nextAppCreator;
                        CurrentAppCode = nextAppCode;


                        await sp.web.lists.getByTitle("WorkFlowProcessData").items.getById(Number(id)).update({
                            CurrentQueue: WholeRecord.DestinationQueue,
                            DestinationQueue: varDestQueuee,
                            Status: varStatusss,
                            CurApproverId: varCurApproverId,
                            Level: 0,
                            FinalCost: FinalCost || 0,
                            FinalDescription: Finaldescription || '',
                            CurrentAppCode: CurrentAppCode,
                            AllApprovers: JSON.stringify(Approvers)
                        });
                        await sp.web.lists.getByTitle("WorkFlowProcessData").items.getById(Number(CurrentChildId)).update({
                            Status: "Completed", AllApprovers: JSON.stringify(childAllApprovers)
                        });
                    }
                    else {
                        varStatusss = "Completed";
                        varCurApproverId = null;
                        varDestQueuee = "";

                        await sp.web.lists.getByTitle("WorkFlowProcessData").items.getById(Number(id)).update({
                            CurrentQueue: WholeRecord.DestinationQueue,
                            DestinationQueue: varDestQueuee,
                            Status: varStatusss,
                            CurApproverId: varCurApproverId,
                            Level: 0,
                            FinalCost: FinalCost || 0,
                            FinalDescription: Finaldescription || '',
                            AllApprovers: JSON.stringify(Approvers),
                        });

                        await sp.web.lists.getByTitle("WorkFlowProcessData").items.getById(Number(CurrentChildId)).update({
                            Status: "Completed",
                            AllApprovers: JSON.stringify(childAllApprovers),
                        });
                    }
                    await sp.web.lists.getByTitle("WorkFlowLog").items.add({
                        AppCode: AppCodevalue,
                        Status: "Completed",
                        Domain: domainfield,
                        Level: 0,
                        SeqNo: appSeqNo,
                        Role: userRole,
                        Action: "Close",
                    })

                } catch (err) {
                    console.log("final close button patching error in multiform")
                }
            }
            else {
                debugger;
                try {

                    await sp.web.lists.getByTitle("WorkFlowProcessData").items.getById(Number(id)).update({
                        CurrentQueue: WholeRecord.DestinationQueue,
                        DestinationQueue: "",
                        Status: "Completed",
                        CurApproverId: null,
                        Level: 0,
                        FinalCost: FinalCost || 0,
                        FinalDescription: Finaldescription || '',
                        AllApprovers: JSON.stringify(Approvers),
                    });
                    await sp.web.lists.getByTitle("WorkFlowLog").items.add({
                        AppCode: AppCodevalue,
                        Status: "Completed",
                        Domain: domainfield,
                        Level: 0,
                        SeqNo: appSeqNo,
                        Role: userRole,
                        Action: "Close",
                    })
                } catch (err) {
                    console.log("final close button patching error in single form")
                }
            }

            const rr = await sp.web.lists.getByTitle("CommentsLog").items.add({
                AppCode: AppCodevalue,
                Comments: comment,
                CommentedBy: userName,
                CommentedByEmail: userEmail,
                SeqNo: appSeqNo
            });

            // Access the item ID from the result
            const itemId = rr.data.Id;

            // If a file is selected, add it to the list item
            if (selectedFile?.length > 0) {
                for (const file of selectedFile) {
                    await sp.web.lists
                        .getByTitle("CommentsLog")
                        .items.getById(itemId)
                        .attachmentFiles.add(file.name, file);
                }
            } else {
                console.log("No file selected, skipping file upload.");
            }
            setLoading(false);
            toast.current?.show({ severity: 'success', summary: '', detail: "Request completed successfully", life: 2000 });
            setLoading(true);
            setTimeout(() => {

                navigate('/InProcess');
            }, 1000)
        }
        else {
            setLoading(false);
            toast.current?.show({
                severity: 'warn',
                summary: '',
                detail: 'Please add your comment before closing the request',
                life: 2000
            });
            setSelectTab('AllDetails')
            setHighlight(true);
        }
    }
    const handleApprove = async () => {
        debugger;
        setLoading(true);
        let commentValidationValue = "false"
        const trimmedComment = comment?.trim() || "";
        const isFilePresent = selectedFile && selectedFile.length > 0;
        const folderPath = `${fileFetchUrl}/CommentAttachments/${appSeqNo}`;
        const files = await sp.web.lists.getByTitle("CommentAttachments")
            .items
            .filter(`FileDirRef eq '${folderPath}'`)
            .select("FileLeafRef", "FileRef", "Author/Title", "FileUploadedbyEmail")
            .expand("Author")
            .get();
        console.log("all files Inside approveeeeeeeeeeee---------------", files)
        if (userRole === "Approver") {
            commentValidationValue = "true"
        }
        if (isFilePresent && trimmedComment.length === 0) {
            setLoading(false);
            toast.current?.show({
                severity: 'warn',
                summary: '',
                detail: 'Please enter a comment when uploading a file.',
                life: 2000
            });
            setSelectTab('AllDetails')
            setHighlight(true);
            return;
        }
        if (comment.trim().length > 0) {
            commentValidationValue = "true"
        }

        if (commentValidationValue === "true") {
            //debugger;
            const buttonAction = "Approve"
            const routingRulesItems = await sp.web.lists.getByTitle("RoutingRules").items.filter(`(CurrentQueue eq '${ProcessDataDestiQueue}') and (AppCode eq '${IsParentForm ? WholeRecord.CurrentAppCode : AppCodevalue}') and (Actions eq '${buttonAction}')`).get();
            if (routingRulesItems.length > 0) {

                const { CurrentQueue, DestinationQueue, Status } = routingRulesItems[0];
                const mappingMasterItem = await sp.web.lists.getByTitle("MappingMaster").items
                    .select("*", "Users_x002f_Groups/Id", "Users_x002f_Groups/Title")
                    .expand("Users_x002f_Groups")
                    .filter(`(Level eq '${DestinationQueue}') and (AppCode eq '${AppCodevalue}')`)
                    .get();
                let conditionlistItem: any[] = [];
                let conditionlistItem1: any[] = [];
                let result: boolean = false
                let result1: boolean = false
                try {
                    conditionlistItem = await sp.web.lists
                        .getByTitle("ConditionsList")
                        .items.select("*", "PersonOrGroup/Id", "PersonOrGroup/Title")
                        .expand("PersonOrGroup")
                        .filter(`(Level eq '${mappingMasterItem[0].Level}') and (AppCode eq '${IsParentForm ? WholeRecord.CurrentAppCode : AppCodevalue}')`)
                        .get();
                    conditionlistItem1 = await sp.web.lists
                        .getByTitle("ConditionsList")
                        .items.select("*", "PersonOrGroup/Id", "PersonOrGroup/Title", "FinalApprove")
                        .expand("PersonOrGroup")
                        .filter(`(Level eq '${mappingMasterItem[0].Level - 1}') and (AppCode eq '${IsParentForm ? WholeRecord.CurrentAppCode : AppCodevalue}')`)
                        .get();
                    if (conditionlistItem.length > 0) {
                        const conditionlistItemJson: any = JSON.parse(conditionlistItem[0].ConditionJson);
                        const mainoperator: string = conditionlistItemJson.operator;
                        result = mainoperator === "And";

                        conditionlistItemJson.conditions.forEach((element: any) => {
                            try {
                                const fieldId = formJSON.find((val: any) => {
                                    return stripHtmlTags(val.label) === element.field;
                                })?.id;

                                if (!fieldId) {
                                    throw new Error(`Field ID not found for field: ${element.field}`);
                                }

                                const fieldvalue = formData.find((d: any) => d.id === fieldId)?.value;

                                if (fieldvalue === undefined) {
                                    throw new Error(`Field value not found for field ID: ${fieldId}`);
                                }

                                const boolvalue = ConditionValidator(fieldvalue, element.operator, element.value, element.elementType);

                                result = mainoperator === "And" ? boolvalue && result : boolvalue || result;

                            } catch (innerError) {
                                console.error("Error processing condition element:", innerError);
                            }
                        });
                    }
                    if (conditionlistItem1.length > 0) {
                        const conditionlistItemJson: any = JSON.parse(conditionlistItem1[0].ConditionJson);
                        const mainoperator: string = conditionlistItemJson.operator;
                        result1 = mainoperator === "And";

                        conditionlistItemJson.conditions.forEach((element: any) => {
                            try {
                                const fieldId = formJSON.find((val: any) => {
                                    return stripHtmlTags(val.label) === element.field;
                                })?.id;

                                if (!fieldId) {
                                    throw new Error(`Field ID not found for field: ${element.field}`);
                                }

                                const fieldvalue = formData.find((d: any) => d.id === fieldId)?.value;

                                if (fieldvalue === undefined) {
                                    throw new Error(`Field value not found for field ID: ${fieldId}`);
                                }

                                const boolvalue = ConditionValidator(fieldvalue, element.operator, element.value, element.elementType);

                                result1 = mainoperator === "And" ? boolvalue && result1 : boolvalue || result1;

                            } catch (innerError) {
                                console.error("Error processing condition element:", innerError);
                            }
                        });
                    }

                } catch (error) {
                    console.error("Error occurred in processing conditions:", error);
                }
                //console.log("Maapingmastrt items in handle approver :", mappingMasterItem);
                const MMLevel = mappingMasterItem[0].Level;
                // let lastlevel = mappingMasterItem.length - 1
                let userGroupId
                // if (MMLevel === lastlevel) {
                //     userGroupId = null;
                // } else

                if (result && conditionlistItem.length > 0) {
                    userGroupId = conditionlistItem[0].PersonOrGroup?.Id || null; // Safely access PersonOrGroup.Id
                } else {
                    userGroupId = mappingMasterItem[0]?.Users_x002f_Groups?.Id || null; // Fallback logic
                }
                if (mappingMasterItem.length === 0) {
                    // alert("Reviewers or Approvers are not available for further actions");
                    setLoading(false)
                    toast.current?.show({ severity: 'error', summary: '', detail: 'Reviewers or Approvers are not available for further actions', life: 2000 });
                    return;
                }
                try {
                    if (trimmedComment.length > 0) {
                        const rr = await sp.web.lists.getByTitle("CommentsLog").items.add({
                            AppCode: (IsParentForm) ? WholeRecord.CurrentAppCode : AppCodevalue,
                            Comments: comment,
                            CommentedBy: userName,
                            CommentedByEmail: userEmail,
                            SeqNo: appSeqNo
                        });
                        // Access the item ID from the result
                        const itemId = rr.data.Id;
                        // If a file is selected, add it to the list item
                        if (selectedFile?.length > 0) {
                            for (const file of selectedFile) {
                                await sp.web.lists
                                    .getByTitle("CommentsLog")
                                    .items.getById(itemId)
                                    .attachmentFiles.add(file.name, file);
                            }
                        } else {
                            console.log("No file selected, skipping file upload.");
                        }
                    }
                }
                catch (err) {
                    console.log("error", err)
                }
                // if (DestinationQueue === "0" && Status === "In-Progress") {
                //     userGroupId = WholeRecord.AuthorId
                // }
                //debugger;
                // const updatedFormData = JSON.stringify(formData);
                const [parentFormData1, childFormData1] = AnswerJsonSeparator(formJSON, formData);
                const parentFormData = JSON.stringify(parentFormData1);
                const childFormData = JSON.stringify(childFormData1);

                let childupdatedChildApprovers: any;
                if (IsParentForm) {
                    const Childtranallapprover = ChildtransactionItems.find(child => child.Id === CurrentChildId);
                    let childAllApprovers = (Childtranallapprover.AllApprovers !== null) ? JSON.parse(Childtranallapprover.AllApprovers) : [];
                    childupdatedChildApprovers = (!childAllApprovers.includes(userEmail)) ? [...childAllApprovers, userEmail] : [...childAllApprovers];
                }

                if (formUpdatedFlag === true) {

                    try {

                        const updatedApprovers = [...AllApprovers, userEmail]
                        await sp.web.lists.getByTitle("WorkFlowProcessData").items.getById(Number(id)).update({
                            FormData: parentFormData, //updating form without seperating the child form details
                            CurrentQueue: CurrentQueue,
                            DestinationQueue: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? "" : DestinationQueue,
                            Status: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? "Completed" : Status,
                            CurApproverId: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? null : userGroupId,
                            Level: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? 0 : MMLevel,
                            // AllApprovers: updatedApprovers
                            // CurrentAppCode : nextAppCode,
                            FinalCost: FinalCost || 0,
                            FinalDescription: Finaldescription || '',
                            AllApprovers: JSON.stringify(updatedApprovers),
                        });
                        if (IsParentForm) {
                            await sp.web.lists.getByTitle("WorkFlowProcessData").items.getById(Number(CurrentChildId)).update({
                                FormData: childFormData, //updating form without seperating the child form details
                                CurrentQueue: CurrentQueue,
                                DestinationQueue: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? "" : DestinationQueue,
                                Status: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? "Completed" : Status,
                                CurApproverId: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? null : userGroupId,
                                Level: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? 0 : MMLevel,
                                // AllApprovers: updatedApprovers
                                FinalCost: FinalCost || 0,
                                FinalDescription: Finaldescription || '',
                                AllApprovers: JSON.stringify(childupdatedChildApprovers)
                            });
                        }
                        setFinalPopup(false)
                        const workflowlog = await sp.web.lists.getByTitle("WorkFlowLog").items.add({
                            AppCode: IsParentForm ? WholeRecord.CurrentAppCode : AppCodevalue,
                            Status: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? "Completed" : Status,
                            Domain: domainfield,
                            Level: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? 0 : MMLevel,
                            SeqNo: appSeqNo,
                            Role: routingRulesItems[0].Role,
                            Action: buttonAction,
                            ParentAppCode: AppCodevalue
                        })
                        console.log("------------------workflowlog", workflowlog)

                    }
                    catch (err) {
                        console.log("error", err)
                    }
                }
                else {
                    const updatedApprovers = [...AllApprovers, userEmail]
                    //console.log(userGroupId, "usergroupppppppppppppppIDDDDDDDDDDDDDDDD")
                    try {
                        if (IsParentForm) {
                            //debugger;
                            if (nextAppCode && Status === "Completed") {

                                //debugger;
                                const nextAppCreatorData = await sp.web.lists.getByTitle("MappingMaster").items
                                    .select("*", "Users_x002f_Groups/Id", "Users_x002f_Groups/Title")
                                    .expand("Users_x002f_Groups")
                                    .filter(`(Level eq 0) and (AppCode eq '${nextAppCode}')`)
                                    .get();

                                let nextAppCreatorsss: any;

                                if (nextAppCreatorData.length > 0) {
                                    nextAppCreatorsss = nextAppCreatorData[0].Users_x002f_GroupsId || null;
                                }

                                await sp.web.lists.getByTitle("WorkFlowProcessData").items.getById(Number(id)).update({
                                    CurrentQueue: WholeRecord.DestinationQueue,
                                    DestinationQueue: "0",
                                    Status: "In-Progress",
                                    CurApproverId: nextAppCreatorsss,
                                    Level: 0,
                                    AllApprovers: JSON.stringify(updatedApprovers),
                                    FinalCost: FinalCost || 0,
                                    FinalDescription: Finaldescription || '',
                                    CurrentAppCode: nextAppCode,
                                });
                                await sp.web.lists.getByTitle("WorkFlowProcessData").items.getById(CurrentChildId).update({
                                    CurrentQueue: CurrentQueue,
                                    DestinationQueue: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? "" : DestinationQueue,
                                    Status: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? "Completed" : Status,
                                    CurApproverId: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? null : userGroupId,
                                    Level: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? 0 : MMLevel,
                                    AllApprovers: JSON.stringify(childupdatedChildApprovers),
                                    FinalCost: FinalCost || 0,
                                    FinalDescription: Finaldescription || ''
                                })
                            }
                            else if (!nextAppCode && Status === "Completed") {
                                await sp.web.lists.getByTitle("WorkFlowProcessData").items.getById(Number(id)).update({
                                    CurrentQueue: CurrentQueue,
                                    DestinationQueue: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? "" : DestinationQueue,
                                    Status: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? "Completed" : Status,
                                    CurApproverId: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? null : null,
                                    Level: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? 0 : MMLevel,
                                    AllApprovers: JSON.stringify(updatedApprovers),
                                    FinalCost: FinalCost || 0,
                                    FinalDescription: Finaldescription || '',
                                    // AllApprovers: updatedApprovers
                                });
                                await sp.web.lists.getByTitle("WorkFlowProcessData").items.getById(Number(CurrentChildId)).update({
                                    CurrentQueue: CurrentQueue,
                                    DestinationQueue: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? "" : DestinationQueue,
                                    Status: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? "Completed" : Status,
                                    CurApproverId: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? null : null,
                                    Level: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? 0 : MMLevel,
                                    AllApprovers: JSON.stringify(childupdatedChildApprovers),
                                    FinalCost: FinalCost || 0,
                                    FinalDescription: Finaldescription || ''
                                })
                            }
                            else {
                                await sp.web.lists.getByTitle("WorkFlowProcessData").items.getById(Number(id)).update({
                                    CurrentQueue: CurrentQueue,
                                    DestinationQueue: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? "" : DestinationQueue,
                                    Status: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? "Completed" : Status,
                                    CurApproverId: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? null : userGroupId,
                                    Level: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? 0 : MMLevel,
                                    AllApprovers: JSON.stringify(updatedApprovers),
                                    FinalCost: FinalCost || 0,
                                    FinalDescription: Finaldescription || '',
                                    // AllApprovers: updatedApprovers
                                });
                                await sp.web.lists.getByTitle("WorkFlowProcessData").items.getById(Number(CurrentChildId)).update({
                                    CurrentQueue: CurrentQueue,
                                    DestinationQueue: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? "" : DestinationQueue,
                                    Status: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? "Completed" : Status,
                                    CurApproverId: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? null : userGroupId,
                                    Level: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? 0 : MMLevel,
                                    AllApprovers: JSON.stringify(childupdatedChildApprovers),
                                    FinalCost: FinalCost || 0,
                                    FinalDescription: Finaldescription || ''
                                })
                            }

                        }
                        else {
                            //debugger
                            await sp.web.lists.getByTitle("WorkFlowProcessData").items.getById(Number(id)).update({
                                CurrentQueue: CurrentQueue,
                                DestinationQueue: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? "" : DestinationQueue,
                                Status: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? "Completed" : Status,
                                CurApproverId: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? null : Status === "Completed" ? null : userGroupId,
                                Level: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? 0 : MMLevel,
                                AllApprovers: JSON.stringify(updatedApprovers),
                                FinalCost: FinalCost || 0,
                                FinalDescription: Finaldescription || ''
                                // AllApprovers: updatedApprovers
                            });
                        }
                        setFinalPopup(false)
                        const workflowlog = await sp.web.lists.getByTitle("WorkFlowLog").items.add({
                            AppCode: IsParentForm ? WholeRecord.CurrentAppCode : AppCodevalue,
                            Status: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? "Completed" : Status,
                            Domain: domainfield,
                            Level: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? 0 : MMLevel,
                            SeqNo: appSeqNo,
                            Role: routingRulesItems[0].Role,
                            Action: buttonAction,
                        })
                        console.log("------------------workflowlog", workflowlog)
                    }
                    catch (err) {
                        toast.current?.show({ severity: 'error', summary: '', detail: 'Something went wrong , please contact your admin', life: 2000 });
                        return;
                    }
                }
                // alert("You have approved the request");
                let text = "Request Approved Successfully"
                if (userRole === "Reviewer") {
                    text = "Review Submitted Successfully"
                }
                if (userRole === "Task Performer") {
                    text = "Submitted Successfully"
                }
                toast.current?.show({ severity: 'success', summary: '', detail: text, life: 2000 });
                setLoading(true);
                setTimeout(() => {

                    navigate('/InProcess');
                }, 1000)
            }
            else {
                // alert("Routing Rules are not available for further actions");
                setLoading(false);
                toast.current?.show({ severity: 'error', summary: '', detail: 'Routing Rules are not available for further actions', life: 2000 });
                return;
            }
        }
        else {
            // alert("Cannot approve");
            setLoading(false);
            toast.current?.show({
                severity: 'warn',
                summary: '',
                detail: 'Please add your comment before approving',
                life: 2000
            });
            setSelectTab('AllDetails')
            setHighlight(true);
            // setSelectTab('Comment')
        }


    };

    /*Skip functionality*/
    const handleSkip = async () => {
        //debugger;
        setLoading(true);
        let commentValidationValue = "false"
        const folderPath = `${fileFetchUrl}/CommentAttachments/${appSeqNo}`;

        const files = await sp.web.lists.getByTitle("CommentAttachments")
            .items
            .filter(`FileDirRef eq '${folderPath}'`)
            .select("FileLeafRef", "FileRef", "Author/Title", "FileUploadedbyEmail")
            .expand("Author")
            .get();
        console.log("all files inside skip ---------------", files);

        const iscommentthere = comment.trim().length > 0
        const isFilePresent = selectedFile && selectedFile.length > 0;

        if ((iscommentthere) || (!iscommentthere && !isFilePresent)) {
            commentValidationValue = "true"
        }

        if (commentValidationValue === "true") {
            const buttonAction = "Skip";
            const routingRulesItems = await sp.web.lists.getByTitle("RoutingRules").items
                .filter(`(CurrentQueue eq '${ProcessDataDestiQueue}') and (AppCode eq '${AppCodevalue}') and (Actions eq '${buttonAction}')`)
                .get();
            if (routingRulesItems.length > 0) {
                const { CurrentQueue, DestinationQueue, Status } = routingRulesItems[0];
                const mappingMasterItem = await sp.web.lists.getByTitle("MappingMaster").items
                    .select("*", "Users_x002f_Groups/Id", "Users_x002f_Groups/Title")
                    .expand("Users_x002f_Groups")
                    .filter(`(Level eq '${DestinationQueue}') and (AppCode eq '${AppCodevalue}')`)
                    .get();
                let conditionlistItem: any[] = [];
                let conditionlistItem1: any[] = [];
                let result: boolean = false;
                let result1: boolean = false;
                try {
                    conditionlistItem = await sp.web.lists.getByTitle("ConditionsList").items
                        .select("*", "PersonOrGroup/Id", "PersonOrGroup/Title")
                        .expand("PersonOrGroup")
                        .filter(`(Level eq '${mappingMasterItem[0]?.Level}') and (AppCode eq '${AppCodevalue}')`)
                        .get();
                    conditionlistItem1 = await sp.web.lists.getByTitle("ConditionsList").items
                        .select("*", "PersonOrGroup/Id", "PersonOrGroup/Title", "FinalApprove")
                        .expand("PersonOrGroup")
                        .filter(`(Level eq '${mappingMasterItem[0]?.Level - 1}') and (AppCode eq '${AppCodevalue}')`)
                        .get();
                    if (conditionlistItem.length > 0) {
                        const conditionlistItemJson: any = JSON.parse(conditionlistItem[0].ConditionJson);
                        const mainoperator: string = conditionlistItemJson.operator;
                        result = mainoperator === "And";
                        conditionlistItemJson.conditions.forEach((element: any) => {
                            try {
                                const fieldId = formJSON.find((val: any) => stripHtmlTags(val.label) === element.field)?.id;
                                const fieldvalue = formData.find((d: any) => d.id === fieldId)?.value;
                                const boolvalue = ConditionValidator(fieldvalue, element.operator, element.value, element.elementType);
                                result = mainoperator === "And" ? boolvalue && result : boolvalue || result;
                            } catch (innerError) {
                                console.error("Error processing condition element:", innerError);
                            }
                        });
                    }
                    if (conditionlistItem1.length > 0) {
                        const conditionlistItemJson: any = JSON.parse(conditionlistItem1[0].ConditionJson);
                        const mainoperator: string = conditionlistItemJson.operator;
                        result1 = mainoperator === "And";
                        conditionlistItemJson.conditions.forEach((element: any) => {
                            try {
                                const fieldId = formJSON.find((val: any) => stripHtmlTags(val.label) === element.field)?.id;
                                const fieldvalue = formData.find((d: any) => d.id === fieldId)?.value;
                                const boolvalue = ConditionValidator(fieldvalue, element.operator, element.value, element.elementType);
                                result1 = mainoperator === "And" ? boolvalue && result1 : boolvalue || result1;
                            } catch (innerError) {
                                console.error("Error processing condition element:", innerError);
                            }
                        });
                    }
                } catch (error) {
                    console.error("Error occurred in processing conditions:", error);
                }
                const MMLevel = mappingMasterItem[0]?.Level ?? 0;
                const lastlevel = mappingMasterItem.length - 1;

                let userGroupId = null;
                if (MMLevel !== lastlevel) {
                    if (result && conditionlistItem.length > 0) {
                        userGroupId = conditionlistItem[0].PersonOrGroup?.Id || null;
                    } else {
                        userGroupId = mappingMasterItem[0]?.Users_x002f_Groups?.Id || null;
                    }
                }
                try {
                    const trimmedComment = comment?.trim() || "";
                    if (trimmedComment.length > 0) {
                        const rr = await sp.web.lists.getByTitle("CommentsLog").items.add({
                            AppCode: IsParentForm ? WholeRecord.CurrentAppCode : AppCodevalue,
                            Comments: trimmedComment,
                            CommentedBy: userName,
                            CommentedByEmail: userEmail,
                            SeqNo: appSeqNo
                        });
                        const itemId = rr.data.Id;
                        // Upload files if available
                        if (isFilePresent) {
                            for (const file of selectedFile) {
                                await sp.web.lists
                                    .getByTitle("CommentsLog")
                                    .items.getById(itemId)
                                    .attachmentFiles.add(file.name, file);
                            }
                            //console.log("Files uploaded!");
                        }
                        toast.current?.show({
                            severity: 'success',
                            summary: 'Success',
                            detail: 'Comment added successfully.',
                            life: 3000
                        });
                    }
                } catch (err) {
                    console.error("Error:", err);
                    toast.current?.show({
                        severity: 'error',
                        summary: 'Submission Failed',
                        detail: 'An error occurred while submitting the comment.',
                        life: 3000
                    });
                }

                let childupdatedChildApprovers: any;
                if (IsParentForm) {
                    const Childtranallapprover = ChildtransactionItems.find(child => child.Id === CurrentChildId);
                    let childAllApprovers = (Childtranallapprover.AllApprovers !== null) ? JSON.parse(Childtranallapprover.AllApprovers) : [];
                    childupdatedChildApprovers = (!childAllApprovers.includes(userEmail)) ? [...childAllApprovers, userEmail] : [...childAllApprovers];
                }

                const updatedApprovers = [...AllApprovers, userEmail];
                if (formUpdatedFlag === true) {
                    const [parentFormData1, childFormData1] = AnswerJsonSeparator(formJSON, formData);
                    const parentFormData = JSON.stringify(parentFormData1);
                    const childFormData = JSON.stringify(childFormData1);

                    try {
                        //debugger;
                        await sp.web.lists.getByTitle("WorkFlowProcessData").items.getById(Number(id)).update({
                            FormData: parentFormData,
                            CurrentQueue: CurrentQueue,
                            DestinationQueue: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? "" : DestinationQueue,
                            Status: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? "Completed" : Status,
                            CurApproverId: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? null : userGroupId,
                            Level: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? 0 : MMLevel,
                            AllApprovers: JSON.stringify(updatedApprovers)
                        });
                        if (IsParentForm) {
                            await sp.web.lists.getByTitle("WorkFlowProcessData").items.getById(Number(CurrentChildId)).update({
                                FormData: childFormData,
                                CurrentQueue: CurrentQueue,
                                DestinationQueue: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? "" : DestinationQueue,
                                Status: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? "Completed" : Status,
                                CurApproverId: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? null : userGroupId,
                                Level: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? 0 : MMLevel,
                                AllApprovers: JSON.stringify(childupdatedChildApprovers)
                            });
                        }

                        setFinalPopup(false)
                        const workflowlog = await sp.web.lists.getByTitle("WorkFlowLog").items.add({
                            AppCode: IsParentForm ? WholeRecord.CurrentAppCode : AppCodevalue,
                            Status: "Skipped",
                            Domain: domainfield,
                            Level: MMLevel,
                            Role: routingRulesItems[0].Role,
                            SeqNo: appSeqNo,
                            Action: buttonAction,
                            ParentAppCode: IsParentForm ? WholeRecord.AppCode : AppCodevalue
                        })
                        console.log("------------------workflowlog", workflowlog)
                    } catch (err) {
                        console.log("error", err);
                    }
                } else {
                    try {
                        await sp.web.lists.getByTitle("WorkFlowProcessData").items.getById(Number(id)).update({
                            CurrentQueue: CurrentQueue,
                            DestinationQueue: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? "" : DestinationQueue,
                            Status: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? "Completed" : Status,
                            CurApproverId: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? null : userGroupId,
                            Level: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? 0 : MMLevel,
                            AllApprovers: JSON.stringify(updatedApprovers)
                        });
                        if (IsParentForm) {
                            await sp.web.lists.getByTitle("WorkFlowProcessData").items.getById(Number(CurrentChildId)).update({
                                // FormData: childFormData,
                                CurrentQueue: CurrentQueue,
                                DestinationQueue: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? "" : DestinationQueue,
                                Status: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? "Completed" : Status,
                                CurApproverId: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? null : userGroupId,
                                Level: result1 && conditionlistItem1[0]?.FinalApprove === "Yes" ? 0 : MMLevel,
                                AllApprovers: JSON.stringify(childupdatedChildApprovers)
                            });
                        }
                        setFinalPopup(false)
                        const workflowlog = await sp.web.lists.getByTitle("WorkFlowLog").items.add({
                            AppCode: IsParentForm ? WholeRecord.CurrentAppCode : AppCodevalue,
                            Status: "Skipped",
                            Domain: domainfield,
                            Level: MMLevel,
                            Role: routingRulesItems[0].Role,
                            SeqNo: appSeqNo,
                            Action: buttonAction,
                            ParentAppCode: IsParentForm ? WholeRecord.AppCode : AppCodevalue
                        })
                        console.log("------------------workflowlog", workflowlog)
                    } catch (err) {
                        console.log("error", err);
                    }
                }
                toast.current?.show({ severity: 'success', summary: '', detail: 'You have skipped the request', life: 2000 });
                setLoading(true);
                setTimeout(() => {
                    navigate('/InProcess');
                }, 1000);
            } else {
                setLoading(false);
                toast.current?.show({ severity: 'error', summary: '', detail: 'Routing Rules are not available for further actions', life: 2000 });
                return;
            }
        }
        else {
            // alert("Cannot approve");
            setLoading(false);
            toast.current?.show({
                severity: 'warn',
                summary: '',
                detail: 'Please enter a comment when uploading a file.',
                life: 2000
            });
            setSelectTab('AllDetails')
            setHighlight(true);
            // setSelectTab('Comment')
        }
    };

    const handleonclickchild = async (appCode: string, status: string) => {
        debugger;
        if (status === "In-Progress" || DynamicStatuses.includes(status)) {
            // const mappingMasterData1 = await sp.web.lists
            //         .getByTitle("MappingMaster")
            //         .items.select("Level", "Role", "Users_x002f_Groups/Title", "Users_x002f_Groups/Id", "CanEdit")
            //         .expand("Users_x002f_Groups")
            //         .filter(`AppCode eq '${appCode}'`)
            //         .get();
            // for (const mappingItem of mappingMasterData1) {

            // for (const mappingItem of mappingMasterData) {
            const groupOrUserId = WholeRecord.CurApproverId;
            if (groupOrUserId) {
                // Check if it's a SharePoint group
                const isGroup = await sp.web.siteGroups.getById(groupOrUserId).get().then(() => true).catch(() => false);
                if (isGroup) {

                    // Check if user is a member of the group
                    const isMember = await checkGroupMembership(groupOrUserId, Number(userId));
                    if (isMember) {
                        setIsCurApprover(true);
                        return;
                    }
                } else if (WholeRecord?.CurApproverId === userId) {
                    // Direct user ID match
                    setIsCurApprover(true);
                    return;
                }
            }


            else { setIsCurApprover(false) }
        }
    }

    const handleReject = async () => {
        //debugger;
        setLoading(true);
        let commentValidationValue = "false"
        if (comment.trim().length < 1) {
            // alert("plese add comment")
        }
        else {
            try {
                const rr = await sp.web.lists.getByTitle("CommentsLog").items.add({
                    AppCode: IsParentForm ? WholeRecord.CurrentAppCode : AppCodevalue,
                    Comments: comment,
                    CommentedBy: userName,
                    CommentedByEmail: userEmail,
                    SeqNo: appSeqNo
                });

                const itemId = rr.data.Id;

                // If a file is selected, add it to the list item
                if (selectedFile.length > 0) {
                    for (const file of selectedFile) {
                        await sp.web.lists
                            .getByTitle("CommentsLog")
                            .items.getById(itemId)
                            .attachmentFiles.add(file.name, file);
                    }
                    //console.log("Files successfully uploaded!");
                } else {
                    console.log("No file selected, skipping file upload.");
                }

                commentValidationValue = "true"
            }
            catch (err) {
                console.log("error", err)
            }
        }


        if (commentValidationValue === "true") {
            //console.log("commentValidationValue----------", commentValidationValue)
            const buttonAction = "Reject"
            const routingRulesItems = await sp.web.lists.getByTitle("RoutingRules").items.filter(`(CurrentQueue eq '${ProcessDataDestiQueue}') and (AppCode eq '${IsParentForm ? WholeRecord.CurrentAppCode : AppCodevalue}') and (Actions eq '${buttonAction}')`).get();
            if (routingRulesItems.length > 0) {

                const { CurrentQueue, DestinationQueue, Status } = routingRulesItems[0];
                const mappingMasterItem = await sp.web.lists.getByTitle("MappingMaster").items
                    .select("*", "Users_x002f_Groups/Id", "Users_x002f_Groups/Title")
                    .expand("Users_x002f_Groups")
                    .filter(`(Level eq '${DestinationQueue}') and (AppCode eq '${AppCodevalue}')`)
                    .get();
                const MMLevel = mappingMasterItem[0].Level;
                const userGroupId = null
                // let MMLevel
                // let userGroupId = null
                // if (DestinationQueue === null || undefined) {
                //     userGroupId = null
                //     MMLevel = CurrentQueue
                // }
                // else {
                //     userGroupId = mappingMasterItem[0].Users_x002f_Groups.Id;
                //     MMLevel = mappingMasterItem[0].Level;
                // }

                if (mappingMasterItem.length === 0) {
                    setLoading(false)
                    alert("Reviewers or Approvers are not available for further actions");
                    return;
                }
                // const newApprover = {
                //     email: userEmail
                // }
                let childupdatedChildApprovers: any;
                if (IsParentForm) {
                    const Childtranallapprover = ChildtransactionItems.find(child => child.Id === CurrentChildId);
                    let childAllApprovers = (Childtranallapprover.AllApprovers !== null) ? JSON.parse(Childtranallapprover.AllApprovers) : [];
                    childupdatedChildApprovers = (!childAllApprovers.includes(userEmail)) ? [...childAllApprovers, userEmail] : [...childAllApprovers];
                }
                const updatedApprovers = [...AllApprovers, userEmail]
                await sp.web.lists.getByTitle("WorkFlowProcessData").items.getById(Number(id)).update({
                    CurrentQueue: CurrentQueue,
                    DestinationQueue: DestinationQueue,
                    Status: Status,
                    CurApproverId: userGroupId,
                    Level: MMLevel,
                    AllApprovers: JSON.stringify(updatedApprovers)
                });
                // alert("You have Rejected the request");
                if (IsParentForm) {
                    await sp.web.lists.getByTitle("WorkFlowProcessData").items.getById(Number(CurrentChildId)).update({
                        CurrentQueue: CurrentQueue,
                        DestinationQueue: DestinationQueue,
                        Status: Status,
                        CurApproverId: userGroupId,
                        Level: MMLevel,
                        AllApprovers: JSON.stringify(childupdatedChildApprovers)
                    });
                }
                toast.current?.show({ severity: 'success', summary: '', detail: 'You have Rejected the request', life: 2000 });
                const workflowlog = await sp.web.lists.getByTitle("WorkFlowLog").items.add({
                    AppCode: IsParentForm ? WholeRecord.CurrentAppCode : AppCodevalue,
                    Status: Status,
                    Domain: domainfield,
                    Level: MMLevel,
                    SeqNo: appSeqNo,
                    Role: routingRulesItems[0].Role,
                    Action: buttonAction,
                    ParentAppCode: IsParentForm ? WholeRecord.AppCode : ""
                })
                console.log("------------------workflowlog", workflowlog)
                setLoading(true);

                setTimeout(() => {

                    navigate('/InProcess');
                }, 3000)

            }
            else {
                // alert("Routing Rules are not available for further actions");
                setLoading(false)
                toast.current?.show({ severity: 'warn', summary: '', detail: 'Routing Rules are not available for further actions', life: 2000 });
                return;
            }
        }
        else {
            setLoading(false)
            // alert("cannot reject without comment")
            toast.current?.show({
                severity: 'warn',
                summary: '',
                detail: 'Please add your comment before reject',
                life: 2000
            });
            setHighlight(true);
            setSelectTab('AllDetails')
            // setSelectTab('Comment')
        }



    };
    const handlepasteData = (data: any) => {
        //console.log("===------------------------------------------===================paste data", data);
        setComment(data)
    }
    const handleReturn = async () => {
        setLoading(true)
        //console.log("return manager -----------------", currManager)
        let commentValidationValue = "false"
        if (comment.trim().length < 1) {
            // alert("plese add comment")
        }
        else {
            try {
                const rr = await sp.web.lists.getByTitle("CommentsLog").items.add({
                    AppCode: IsParentForm ? WholeRecord.CurrentAppCode : AppCodevalue,
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
                    //console.log("Files successfully uploaded!");
                } else {
                    console.log("No file selected, skipping file upload.");
                }


                commentValidationValue = "true"
            }
            catch (err) {
                console.log("error", err)
            }
        }


        if (commentValidationValue === "true") {
            const buttonAction = "Return"
            const routingRulesItems = await sp.web.lists.getByTitle("RoutingRules").items.filter(`(CurrentQueue eq '${ProcessDataDestiQueue}') and (AppCode eq '${IsParentForm ? WholeRecord.CurrentAppCode : AppCodevalue}') and (Actions eq '${buttonAction}')`).get();
            if (routingRulesItems.length > 0) {

                const { CurrentQueue, DestinationQueue, Status } = routingRulesItems[0];
                const mappingMasterItem = await sp.web.lists.getByTitle("MappingMaster").items
                    .select("*", "Users_x002f_Groups/Id", "Users_x002f_Groups/Title")
                    .expand("Users_x002f_Groups")
                    .filter(`(Level eq '${DestinationQueue}') and (AppCode eq '${IsParentForm ? WholeRecord.CurrentAppCode : AppCodevalue}')`)
                    .get();

                const MMLevel = mappingMasterItem[0].Level;
                //console.log("mappingMasterItem[0].Users_x002f_Groups.Id--------------", mappingMasterItem[0].Users_x002f_Groups, MMLevel)
                let userGroupId;
                if (MMLevel === 1 && mappingMasterItem[0].Users_x002f_Groups === undefined) {

                    userGroupId = currManager
                }
                else {
                    userGroupId = mappingMasterItem[0].Users_x002f_Groups.Id;
                }
                console.log(userGroupId);

                if (mappingMasterItem.length === 0) {
                    // alert("Reviewers or Approvers are not available for further actions");
                    toast.current?.show({ severity: 'warn', summary: '', detail: 'Routing Rules are not available for further actions', life: 2000 });
                    return;
                }
                const creator = ChildtransactionItems.find((item) => item.AppCode === WholeRecord.CurrentAppCode);
                const authorId = creator ? creator.AuthorId : null;
                let childupdatedChildApprovers: any;
                if (IsParentForm) {
                    const Childtranallapprover = ChildtransactionItems.find(child => child.Id === CurrentChildId);
                    let childAllApprovers = (Childtranallapprover.AllApprovers !== null) ? JSON.parse(Childtranallapprover.AllApprovers) : [];
                    childupdatedChildApprovers = (!childAllApprovers.includes(userEmail)) ? [...childAllApprovers, userEmail] : [...childAllApprovers];
                }

                const updatedApprovers = [...(AllApprovers || ""), userEmail]
                await sp.web.lists.getByTitle("WorkFlowProcessData").items.getById(Number(id)).update({
                    CurrentQueue: CurrentQueue,
                    DestinationQueue: DestinationQueue,
                    Status: Status,
                    CurApproverId: (IsParentForm) ? authorId : null,
                    Level: MMLevel,
                    AllApprovers: JSON.stringify(updatedApprovers)

                });
                if (IsParentForm) {
                    await sp.web.lists.getByTitle("WorkFlowProcessData").items.getById(Number(CurrentChildId)).update({
                        CurrentQueue: CurrentQueue,
                        DestinationQueue: DestinationQueue,
                        Status: Status,
                        CurApproverId: (IsParentForm) ? authorId : null,
                        Level: MMLevel,
                        AllApprovers: JSON.stringify(childupdatedChildApprovers)

                    });
                }

                // alert("You have returned the request");
                toast.current?.show({ severity: 'info', summary: '', detail: 'You have returned the request', life: 2000 });
                const workflowlog = await sp.web.lists.getByTitle("WorkFlowLog").items.add({
                    AppCode: IsParentForm ? WholeRecord.CurrentAppCode : AppCodevalue,
                    Status: Status,
                    Domain: domainfield,
                    Level: MMLevel,
                    SeqNo: appSeqNo,
                    Role: routingRulesItems[0].Role,
                    Action: buttonAction,
                    ParentAppCode: IsParentForm ? WholeRecord.AppCode : ""
                })

                console.log("------------------workflowlog", workflowlog)
                setLoading(true);

                setTimeout(() => {

                    navigate('/InProcess');
                }, 3000)
            }
            else {
                // alert("Routing Rules are not available for further actions");
                setLoading(false)
                toast.current?.show({ severity: 'warn', summary: '', detail: 'Workflow is not designed.Please contact your admin', life: 2000 });
                return;
            }
        } else {
            setLoading(false)
            // alert("Cannot return the note");
            toast.current?.show({
                severity: 'warn',
                summary: '',
                detail: 'Please add your comment before returning',
                life: 2000
            });
            setHighlight(true);
            setSelectTab('AllDetails')
            // setSelectTab('Comment');
        }

    };

    const openFinalCostPopup = () => {
        if (userRole === "Task Performer") {

            const trimmedComment = comment?.trim() || "";
            const isFilePresent = selectedFile && selectedFile.length > 0;
            if (isFilePresent && trimmedComment.length === 0) {
                setLoading(false);
                toast.current?.show({
                    severity: 'warn',
                    summary: '',
                    detail: 'Please enter a comment when uploading a file.',
                    life: 2000
                });
                setSelectTab('AllDetails')
                setHighlight(true);
                return;
            }
            else if (trimmedComment.length === 0) {
                toast.current?.show({
                    severity: 'warn',
                    summary: '',
                    detail: 'Please add your comment before closing the request',
                    life: 2000
                });
                setSelectTab('AllDetails')
                setHighlight(true);
            }
            else {
                setFinalPopup(true)
            }
        } else {
            void lastfunction()
        }
    }
    const closeFinalPopup = () => {
        setFinalPopup(false);
        setFinalCost(0);
        setFinaldescription('')

    }
    const handleCancel = () => {

        navigate('/InProcess');

    };


    const handleChatResponse = (severity: any, message: any) => {
        toast.current?.show({ severity, detail: message, life: 1000 });
    };

    const handleCommnetSaveOnly = async () => {
        let commentValidationValue = "false"
        if (comment.trim().length < 1) {
            // alert("plese add comment")
        }
        else {
            try {


                const rr = await sp.web.lists.getByTitle("CommentsLog").items.add({
                    AppCode: IsParentForm ? WholeRecord.CurrentAppCode : AppCodevalue,
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
                } else {
                    console.log("No file selected, skipping file upload.");
                }


                commentValidationValue = "true"
                setComment("")
                setSelectedFile([])
                void fetchFormData()
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
                detail: 'Please add your comment to save',
                life: 2000
            });
            setHighlight(true);
            setSelectTab('AllDetails')
        }

    }
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

    const handleUserDomain = (UserDomain: string) => {
        setIsSameDomain(UserDomain)

    };

    return (
        <>
            {(loading) ? <LoadingSpinner></LoadingSpinner> : (WholeRecord.Domain === domainfield ?
                <div style={{ display: 'flex', flexDirection: 'column' }}>
                    <Toast ref={toast} />
                    <TopNavBar onUserDomainRetrieved={handleUserDomain} />
                    <div style={{ display: 'flex' }}>
                        <SideBar activeMenu="" />
                        <div className='FC_mainDiv' style={{ width: '100%' }}>
                            <div className="FC_header-container">
                                <div className="FC_title-section">
                                    <div className="FC_back-icon">
                                        <img onClick={() =>

                                            navigate(-1 as unknown as To, preservedfilters)} src={require('../../assets/Images/previous.png')} alt="backicon" title='Back' />
                                    </div>
                                    <div>
                                        <p className="FC_FormName">{AppName}</p>
                                    </div>
                                    <div className="header-template-generator" style={{
                                        display: 'flex',
                                        alignItems: 'center',
                                        gap: '8px',
                                        background: 'rgba(255, 255, 255, 0.1)',
                                        padding: '4px 8px',
                                        borderRadius: '6px',
                                        border: '1px solid rgba(255, 255, 255, 0.2)',
                                        marginLeft: '20px'
                                    }}>
                                        {templates.length > 0 ? (
                                            <>
                                                <FileText size={16} style={{ color: 'white' }} />
                                                <select
                                                    value={selectedTemplate}
                                                    onChange={(e) => setSelectedTemplate(e.target.value)}
                                                    style={{
                                                        border: 'none',
                                                        background: 'transparent',
                                                        color: 'white',
                                                        fontSize: '12px',
                                                        fontWeight: '600',
                                                        outline: 'none',
                                                        cursor: 'pointer',
                                                        minWidth: '120px'
                                                    }}
                                                >
                                                    {templates.map(t => (
                                                        <option key={t.id} value={t.id} style={{ color: 'black' }}>{t.name}</option>
                                                    ))}
                                                </select>
                                                <button
                                                    onClick={handleGeneratePDF}
                                                    disabled={isGeneratingPDF || !selectedTemplate}
                                                    style={{
                                                        background: '#4ade80',
                                                        color: '#14532d',
                                                        border: 'none',
                                                        padding: '4px 10px',
                                                        borderRadius: '4px',
                                                        fontSize: '11px',
                                                        fontWeight: 'bold',
                                                        display: 'flex',
                                                        alignItems: 'center',
                                                        gap: '4px',
                                                        cursor: 'pointer'
                                                    }}
                                                >
                                                    {isGeneratingPDF ? <Loader size={12} className="animate-spin" /> : <Download size={12} />}
                                                    {isGeneratingPDF ? '...' : 'PDF'}
                                                </button>
                                            </>
                                        ) : (
                                            <span style={{ fontSize: '10px', color: 'rgba(255,255,255,0.6)' }}>
                                                {AppCodevalue ? `No PDF Templates` : `Searching...`}
                                            </span>
                                        )}
                                    </div>
                                </div>

                                <div className="FC_requester-name">
                                    <p style={{ margin: 0 }}>{RequesterName}</p>
                                </div>
                            </div>

                            {IsParentForm && <ParentWithChildren
                                parentName="Form Panel"
                                children={filteredChildFormDetails.map((ele) => ({
                                    id: ele.AppCode,
                                    name: ele.AppName,
                                    ChildFormJson: JSON.parse(ele.FormJSON),
                                    AppCode: ele.AppCode,
                                    ChildOrder: ele.ChildOrder
                                }))}
                                onChildClick={async (child: any) => {
                                    debugger
                                    //console.log("🖱️ Child tab clicked:", child.name, "AppCode:", child.id, "ChildOrder:", child.ChildOrder);
                                    const childOrder = child.ChildOrder || 1;
                                    let currentchildtransactionitem = ChildtransactionItems[childOrder - 1] || ChildtransactionItems[0];
                                    setTables(JSON.parse(currentchildtransactionitem?.TableJSON) || []);
                                    setTimesheettables(JSON.parse(currentchildtransactionitem?.TimeSheetJSON) || []);
                                    setSelectedChildStatus(currentchildtransactionitem?.Status === "In-Progress" || DynamicStatuses.includes(currentchildtransactionitem?.Status) ? true : false);
                                    let childformjson = await updateChildTabData(String(child.id), childOrder, child.ChildFormJson);
                                    setchildappselected(child.AppCode);
                                    setMultiFormSeqNo(currentchildtransactionitem?.SeqNo);
                                    //console.log(childappselected);
                                    let completeformjson = [...ParentFormJSONCopy, ...childformjson];
                                    let currentchildanswerdata = currentchildtransactionitem ? JSON.parse(currentchildtransactionitem.FormData) : [];
                                    //console.log("currentchildanswerdata", currentchildanswerdata);
                                    //console.log("ParentFormAnswerCopy", ParentFormAnswerCopy);
                                    let completeformanswers = [...ParentFormAnswerCopy, ...currentchildanswerdata];
                                    //console.log("completeformanswers", completeformanswers);
                                    setFormData(completeformanswers);
                                    // console.log("pm parent with child clicked status", currentchildtransactionitem);
                                    //console.log("completeformjson", completeformjson);
                                    setFormJSON(completeformjson);
                                    setRecentStatus(currentchildtransactionitem?.Status);
                                    setselectedDynamicStatus(currentchildtransactionitem?.Status)
                                    setApprovedShow(currentchildtransactionitem?.Status);
                                    void handleonclickchild(child.AppCode, currentchildtransactionitem.Status);
                                }}
                            />}
                            {recentStatus !== "Draft" && <Accordion activeIndex={recentStatus !== "Submitted" ? 0 : 1} className='wholeAccordion'>

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
                                                className={`status ${recentStatus === 'In-Progress' ? 'status-inprogress' :
                                                    recentStatus === 'Completed' ? 'status-approved' :
                                                        recentStatus === 'Rejected' ? 'status-rejected' : ''
                                                    }`}
                                            >
                                                {recentStatus}
                                            </span>
                                        </div>
                                    </div>
                                }>
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

                                    <div className="container">

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



                            </Accordion>}

                            <Dialog
                                header="Comment Details"
                                visible={isDialogVisible}
                                onHide={() => setIsDialogVisible(false)}
                                style={{ width: '50vw', background: "grey" }}
                                modal
                                className="Comment_dialog_Header"
                            >
                                {selectedComment && (
                                    <table className='Comment_Table' >
                                        <tbody>
                                            <tr>
                                                <td className='Comment_td_header'>Comment:</td>
                                                <td className='Comment_td_detail'>
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
                                                <td className='Comment_td_header'>Commented By:</td>
                                                <td className='Comment_td_detail'>{selectedComment.CommentedBy}</td>
                                            </tr>
                                            <tr>
                                                <td className='Comment_td_header'>Created Date:</td>
                                                <td className='Comment_td_detail'>
                                                    {new Date(selectedComment.Created).toLocaleString()}
                                                </td>
                                            </tr>
                                            {(selectedComment.Attachments?.length ?? 0) > 0 && (
                                                <tr>
                                                    <td className='Comment_td_header'>Attachments:</td>
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

                                </ul>
                            </div>

                            {loading ? (
                                <LoadingSpinner />
                            ) : (
                                <><div className='fo-mid1'>
                                    {selectTab === 'AllDetails' && (
                                        <>
                                            <div className='p-tabview-panel' style={{ paddingTop: 6 }}>
                                                <div className='custom_field_main form-component-mode'>


                                                    <div className='form-view-mode left-section' >
                                                        <label className='FC_StaticFields_label'>Request For <span style={{ color: 'red' }}> *</span></label>
                                                        <input type="text" className='FC_StaticFields_input' value={Categorytype} readOnly />
                                                    </div>

                                                    <div className='form-view-mode right-section'>
                                                        <label className='FC_StaticFields_label'>Notify To </label>
                                                        <PeoplePicker
                                                            context={getPeoplePickerContext()}
                                                            personSelectionLimit={5}
                                                            required={false}
                                                            onChange={onChange}
                                                            showHiddenInUI={false}
                                                            principalTypes={[PrincipalType.User]}
                                                            resolveDelay={1000}
                                                            defaultSelectedUsers={NotifyRequestorEMails}
                                                            disabled={true}
                                                        />
                                                    </div>

                                                </div>




                                                {adminEditAcces && isCurApprover ?

                                                    <div className="form-view-mode">
                                                        <ReactFormGenerator
                                                            // key={JSON.stringify(formData)}
                                                            data={formJSON}
                                                            read_only={false}
                                                            form_action=""
                                                            form_method=""
                                                            answer_data={formData}
                                                            onChange={(updatedData) => {
                                                                setFormData(updatedData);
                                                                setParentFormAnswerCopy(updatedData); setformUpdatedFlag(true)
                                                            }
                                                            }
                                                        />
                                                    </div>

                                                    :

                                                    <div className="form-view-mode">
                                                        <ReactFormGenerator
                                                            key={JSON.stringify(formData)}
                                                            data={formJSON}
                                                            read_only={true}
                                                            form_action=""
                                                            form_method=""
                                                            answer_data={formData}
                                                        />
                                                    </div>

                                                }



                                            </div>

                                            {isCurApprover ?
                                                <>
                                                    {isCurApprover && (approvedShow === "In-Progress" || DynamicStatuses.includes(approvedShow)) && (
                                                        <div style={{ marginLeft: '1.5rem' }}>
                                                            {(DynamicStatuses.length > 0 && userRole === "Task Performer" && !(Number(WholeRecord.CurrentQueue) !== maxLevel && Number(WholeRecord.DestinationQueue) !== maxLevel)) &&
                                                                <>
                                                                    <h6 style={{ marginTop: '10px', color: "white", marginLeft: '7px', fontSize: 14 }}>Update your status if required</h6>
                                                                    <div className='d-flex gap-3'>
                                                                        <select style={{ fontSize: 13 }} value={selectedDynamicStatus} onChange={handleChangeDynamicStatus} className='form-select col-md-4'>
                                                                            <option value="">Select a status</option>
                                                                            {DynamicStatuses.map((status) => {
                                                                                return <option value={status}>{status}</option>
                                                                            })}
                                                                        </select>
                                                                        <button className='newlogocolorbtn' onClick={handleSaveDynamicStatus}>Update Status</button>
                                                                    </div>
                                                                </>
                                                            }</div>
                                                    )}
                                                    {recentStatus !== "Submitted" && <div style={{ marginLeft: '1.5rem' }}>
                                                        <h6 style={{ marginTop: '10px', color: "white", marginLeft: '7px', fontSize: 14 }}>Add your comment here</h6>

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
                                                                        <img src={require("../../assets/Images/UploadIcon.png")} style={{ width: "12%" }} />
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



                                                                <button onClick={handleCommnetSaveOnly} className="all-btn">Save</button>
                                                            </div>

                                                            <div style={{ flexShrink: 0, flexBasis: "25%" }}>
                                                                {selectedFile.length > 0 && (
                                                                    <ul className='Comment_file_upload_ul'>
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
                                                                                        fill: "#3b82f6", // default color
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

                                                    </div>}


                                                </>
                                                : " "}





                                            {/* <ChatBox context={context} appSeqNo={appSeqNo} userName={userName} onChatResponse={handleChatResponse} appcode={AppCodevalue} /> */}
                                        </>


                                    )}
                                    {/* selectedFile */}
                                    {(selectTab === 'Attachment' && !IsParentForm) && (
                                        <div className='p-tabview-panel'>
                                            {attachmentFiles.length === 0 ? (
                                                <li>No attachments available</li>
                                            ) : (
                                                attachmentFiles.map((file, index) => (
                                                    <div key={index}>
                                                        <p className='Attachment_fileName'>{file.name.split("name")[0]}</p>
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
                                                            // ------- ending code for multi file attachment
                                                        )}
                                                    </div>
                                                ))
                                            )}
                                        </div>
                                    )}
                                    {(selectTab === 'Attachment' && IsParentForm) && (
                                        <div className='p-tabview-panel'>
                                            {/* Parent Attachment Section */}
                                            <div>
                                                <span className="attachment-header" style={{ backgroundColor: '#212B46', padding: '0.5rem', borderRadius: '5px', display: 'block' }}>Parent Attachments</span>
                                                {attachmentFiles.length === 0 ? (
                                                    <li>No attachments available</li>
                                                ) : (
                                                    attachmentFiles.map((file, index) => (
                                                        <div key={index}>
                                                            <p className='Attachment_fileName Attachment_fileName_BorderTop' style={{ background: "#212b4608", borderBottom: "none" }}>{file.name.split("name")[0]}</p>
                                                            {file.name.endsWith('.pdf') ? (
                                                                <embed src={file.url} width="100%" height="600px" type="application/pdf" />
                                                            ) : file.name.endsWith('.jpg') || file.name.endsWith('.jpeg') || file.name.endsWith('.png') ? (
                                                                <img src={file.url} alt={file.name} style={{ width: '100%', height: 'auto' }} />
                                                            ) : file.name.endsWith('.docx') || file.name.endsWith('.pptx') ? (
                                                                <div>
                                                                    <p>File preview not available. You can download and view it:</p>
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
                                                                    <a href='#'
                                                                        onClick={async (e) => {
                                                                            e.preventDefault();
                                                                            await downloadWithName(file.url, file.name.replace(/^\d+\.\s*/, '').split("name")[0]);
                                                                        }}
                                                                        target="_blank"
                                                                        rel="noopener noreferrer">{file.name.split("name")[0]}
                                                                    </a>
                                                                </div>
                                                            )}
                                                        </div>
                                                    ))
                                                )}
                                            </div>
                                            {/* Child Attachment Section */}
                                            <div>
                                                <span className="attachment-header" style={{ backgroundColor: '#212B46', padding: '0.5rem', borderRadius: '5px', display: 'block', marginTop: '10px' }}>Child Attachments</span>
                                                {childAttachmentFiles.length === 0 ? (
                                                    <li>No attachments available</li>
                                                ) : (
                                                    childAttachmentFiles.map((file, index) => (
                                                        <div key={index}>
                                                            <p className='Attachment_fileName Attachment_fileName_BorderTop' style={{ background: "#212b4608", borderBottom: "none" }}>{file.name.split("name")[0]}</p>
                                                            {file.name.endsWith('.pdf') ? (
                                                                <embed src={file.url} width="100%" height="600px" type="application/pdf" />
                                                            ) : file.name.endsWith('.jpg') || file.name.endsWith('.jpeg') || file.name.endsWith('.png') ? (
                                                                <img src={file.url} alt={file.name} style={{ width: '100%', height: 'auto' }} />
                                                            ) : file.name.endsWith('.docx') || file.name.endsWith('.pptx') ? (
                                                                <div>
                                                                    <p>File preview not available. You can download and view it:</p>
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
                                                                    <a href='#'
                                                                        onClick={async (e) => {
                                                                            e.preventDefault();
                                                                            await downloadWithName(file.url, file.name.replace(/^\d+\.\s*/, '').split("name")[0]);
                                                                        }}
                                                                        target="_blank"
                                                                        rel="noopener noreferrer">{file.name.split("name")[0]}
                                                                    </a>
                                                                </div>
                                                            )}
                                                        </div>
                                                    ))
                                                )}
                                            </div>
                                        </div>
                                    )}
                                    {/* comment AND files ----------------------------------------------------------------------- */}
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
                                                    headerStyle={{ width: '26%', textAlign: 'center' }}
                                                />
                                                <Column
                                                    field="CommentedBy"
                                                    header="Commented By"
                                                    sortable
                                                    bodyStyle={{ whiteSpace: 'pre-wrap', wordBreak: 'break-word', width: '22%' }}
                                                    headerStyle={{ width: '20%', textAlign: 'center' }}
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
                                                    headerStyle={{ width: '20%', textAlign: 'center' }}
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
                                                            "No Attachments"
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

                                            <DataTable
                                                value={AuditLog}
                                                responsiveLayout="scroll"
                                                paginator
                                                paginatorTemplate="CurrentPageReport FirstPageLink PrevPageLink PageLinks NextPageLink LastPageLink RowsPerPageDropdown"
                                                currentPageReportTemplate="Showing {first} to {last} of {totalRecords}"
                                                rows={10}
                                                dataKey="ID"
                                                emptyMessage="No transactions found."
                                                loading={loading} // Show loading indicator
                                            >
                                                <Column field="Title" header="Workflow Log" sortable />

                                            </DataTable>
                                        </div>
                                    )}

                                    {isCurApprover && (approvedShow === "In-Progress" || DynamicStatuses.includes(approvedShow)) && (
                                        <div className='d-flex justify-content-between align-items-center w-100'>


                                            <div className='form-btn'>
                                                {
                                                    ((Number(WholeRecord.CurrentQueue) === maxLevel && userRole === "Creator") ||
                                                        (Number(WholeRecord.CurrentQueue) === maxLevel && userRole === "Task Performer") ||
                                                        (Number(WholeRecord.DestinationQueue) === maxLevel && userRole === "Task Performer")) &&
                                                    <button className='all-btn' onClick={openFinalCostPopup}>Mark as Done</button>
                                                }
                                                {userRole === "Task Performer" && Number(WholeRecord.CurrentQueue) !== maxLevel && Number(WholeRecord.DestinationQueue) !== maxLevel &&
                                                    <button className='all-btn' onClick={handleApprove}>Submit</button>
                                                }
                                                {userRole !== "Task Performer" && userRole !== "Creator" &&
                                                    <button className='all-btn' onClick={handleApprove}>{userRole === "Reviewer" ? "Submit Review" : "Approve"}</button>
                                                }

                                                {((Number(maxLevel) !== Number(level)) && (userRole === "Reviewer" || userRole === "Approver")) && (
                                                    <button className='all-btn' onClick={handleSkip}>Skip</button>
                                                )}

                                                {!RejectButton && (
                                                    <button className='all-btn' onClick={handleReject}>Reject</button>
                                                )}
                                                {userRole !== "Creator" &&
                                                    <button className='all-btn' onClick={handleReturn}>Return</button>
                                                }
                                                {/* <button className=" all-btn" onClick={handleCancel}>Cancel</button> */}
                                            </div>

                                        </div>

                                    )}

                                    {IsParentForm && selectedChildStatus && isCurApprover && (
                                        <div className='d-flex justify-content-between align-items-center w-100'>

                                            <div className='form-btn'>
                                                {
                                                    ((Number(WholeRecord.CurrentQueue) === maxLevel && userRole === "Creator") ||
                                                        (Number(WholeRecord.CurrentQueue) === maxLevel && userRole === "Task Performer") ||
                                                        (Number(WholeRecord.DestinationQueue) === maxLevel && userRole === "Task Performer")) &&
                                                    <button className='all-btn' onClick={openFinalCostPopup}>Mark as Done</button>
                                                }
                                                {userRole === "Task Performer" && Number(WholeRecord.CurrentQueue) !== maxLevel && Number(WholeRecord.DestinationQueue) !== maxLevel &&
                                                    <button className='all-btn' onClick={handleApprove}>Submit</button>
                                                }
                                                {userRole !== "Task Performer" && userRole !== "Creator" &&
                                                    <button className='all-btn' onClick={handleApprove}>{userRole === "Reviewer" ? "Submit Review" : "Approve"}</button>
                                                }

                                                {((Number(maxLevel) !== Number(level)) && (userRole === "Reviewer" || userRole === "Approver")) && (
                                                    <button className='all-btn' onClick={handleSkip}>Skip</button>
                                                )}

                                                {!RejectButton && (
                                                    <button className='all-btn' onClick={handleReject}>Reject</button>
                                                )}
                                                {userRole !== "Creator" &&
                                                    <button className='all-btn' onClick={handleReturn}>Return</button>
                                                }
                                                {/* <button className=" all-btn" onClick={handleCancel}>Cancel</button>
                                                 */}
                                            </div>

                                        </div>
                                    )}

                                    {/* {IsParentForm &&!selectedChildStatus && (!isCurApprover || recentStatus === "Submitted") && ( */}
                                    <div className='form-btn d-flex align-items-center' style={{ gap: '10px' }}>
                                        <button className="all-btn" onClick={handleCancel}>Close</button>
                                    </div>
                                </div>
                                </>
                            )}
                            <div>
                                {userRole === "Task Performer" && FinalPopup && (
                                    <div className="custom-dark-modal-overlay">
                                        <div className="custom-dark-modal finalcost">
                                            {/* <h2 style={{ fontSize: '1.5rem' }}>Add Final Cost</h2> */}
                                            <div className="form-group" style={{ padding: '0px 20px 10px 20px' }}>
                                                <div className="form-control-group">
                                                    <label>Total Cost <span style={{ color: 'red' }}>*</span></label>
                                                    <input
                                                        type="number"
                                                        className="form-control formcontrol-extra"
                                                        value={FinalCost === 0 ? '' : FinalCost}
                                                        onChange={(e) => {
                                                            const val = e.target.value;
                                                            setFinalCost(val === '' ? 0 : Number(val));
                                                        }}
                                                        required
                                                        min="0"
                                                        step="1"
                                                    />
                                                </div>
                                                <div className="form-control-group">
                                                    <label>Description</label>
                                                    <textarea
                                                        className="form-control"
                                                        value={Finaldescription}
                                                        onChange={(e: any) => setFinaldescription(e.target.value)}
                                                        rows={3}
                                                    />
                                                </div>

                                                <div className="modal-buttons"
                                                // style={{ gap: '5%', marginTop: '30px' }}
                                                >
                                                    {/* <div> */}
                                                    <button
                                                        className="newlogocolorbtn"
                                                        // style={{ marginRight: '20px' }}
                                                        onClick={lastfunction}
                                                        disabled={FinalCost === null || isNaN(FinalCost) || FinalCost < 0}
                                                    >
                                                        Done
                                                    </button>
                                                    <button
                                                        className="newblackcolorbtn"
                                                        onClick={closeFinalPopup}
                                                    >
                                                        Cancel
                                                    </button>
                                                    {/* </div> */}
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                )}
                            </div>
                        </div>
                        <ChatBox context={context} appSeqNo={IsParentForm ? multiFormSeqNo : appSeqNo} userName={userName} onChatResponse={handleChatResponse} appcode={AppCodevalue} isSameDomain={isSameDomain} />
                    </div>
                </div > : <NoAccess></NoAccess >
            )}
        </>
    );

};

export default ViewForm;