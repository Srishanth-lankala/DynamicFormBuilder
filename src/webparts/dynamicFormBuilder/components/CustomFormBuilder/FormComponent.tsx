import * as React from 'react';
import { useState, useRef, useEffect } from 'react';
import { sp } from "@pnp/sp/presets/all";
import { ReactFormGenerator } from 'react-form-builder2';
import './FormGenerator.css';
import { Toast } from 'primereact/toast';
import { mySiteUrl } from './ConfigURL/All_URLs';
import { PeoplePicker, PrincipalType, IPeoplePickerContext } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { fetchToken } from './ADService/ADService';
import LoadingSpinner from './Loading';
import { BaseWebPartContext } from '@microsoft/sp-webpart-base';
import { fetchTenantUser } from './FetchTenantUser/fetchTenantUser';
// import { fetchAllUsers } from './ADService/ADService';
import TableComponent, { TableData } from './TableStructure/TableComponent';
import TimesheetComponent, { TimeSheetData } from './Timesheetstructure/TimesheetComponent';
import ReactDOM from 'react-dom';
import { useLocation, useNavigate } from 'react-router-dom';
import TopNavBar from './TopBar';
import SideBar from './Sidebar/SideBar';
// import ParentChildTabs from './ParentChildTabs';
import ParentWithChildren from "./ParentWithChildren/ParentWithChildren";

import { dataservice } from './encryptionutil';
import { fetchWorkflowLimit } from './CustomHooks/TransactionCount';

interface IFormComponentProps {
  context: BaseWebPartContext;
}
// interface User {
//   userPrincipalName: any;
//   displayName: string;
//   mail: string;
// }

const FormComponent: React.FC<IFormComponentProps> = ({ context }) => {
  const location = useLocation();
  const [isSameDomain, setIsSameDomain] = useState("");
  const [selectedFormJSON, setSelectedFormJSON] = useState<any[]>([]);
  const [selectedFormTitle, setSelectedFormTitle] = useState<string>("");
  const [selectedFormAppCode, setSelectedFormAppCode] = useState<string>("");
  const [formData, setFormData] = useState<any[]>([]);
  const [isButtonDisabled, setIsButtonDisabled] = useState(false);
  const [RequesterName, setRequesterName] = useState("")
  const [RequesterTitle, setRequesterTitle] = useState("")
  const [RequesterEmail, setRequesterEmail] = useState("");
  const [datefieldids, setdatefieldids] = useState<any[]>([]);
  const [domainfield, setDomainField] = useState<string>("");
  // const [error, setError] = useState<string | null>(null);
  // ------- starting code for multi file attachment
  const [fileData, setFileData] = useState<Record<string, File[] | null>>({});
  const [Categorytype, setCategorytype] = useState<string>("");
  const [loading, setLoading] = useState<boolean>(false);
  const [SourceFieldJsons, SetSourceFieldJsons] = useState<any[]>([])
  const [ChildFieldJsons, SetChildFieldJsons] = useState<any[]>([])
  const [SrcFieldAnswer, SetSrcFieldAnswer] = useState<any[]>([])
  // const [notifytorequestor,Setnotifytorequestor]=useState<string>("");
  const [hiddendata, SetHiddenData] = useState<any[]>([])
  const [childselected, setChildselected] = useState<string>("");

  // const [selectedUsers, setSelectedUsers] = useState<User[]>([]);
  const [notifytorequestorID, SetnotifytorequestorID] = useState<number[]>([]);

  // const [myOrganizationUsers, setMyOrganizationUsers] = useState<User[]>([]);
  // const [searchText, setSearchText] = useState('');
  // const [filteredUsers, setFilteredUsers] = useState<User[]>([]);
  // const [showDropdown, setShowDropdown] = useState(false);
  const [tables, setTables] = useState<TableData[]>([]);
  const [ChildFormDetails, setChildFormDetails] = useState<any[]>([]);
  const [timesheettables, setTimesheetTables] = useState<TimeSheetData[]>([]);
  // const [, setActiveChildFormJSON] = useState<any[]>([]);
  const [ParentFormCopy, setParentFormCopy] = useState<any[]>([]);
  // const [mulitpletimesheeterrs, setMulitpletimesheeterrs] = useState<any[]>([]);

  const [tabclicked, settabclicked] = useState<boolean>(true);

  const [hasBreaches, setHasBreaches] = useState<boolean>(false);
  const navigate = useNavigate(); // State to store breaches
  const IsParentForm = location.state.selectedformdetails.IsParentForm

  const dataserviceobj = new dataservice();

  useEffect(() => {
    //console.log(location.state, "-------------------loccccccc");
    setSelectedFormJSON(JSON.parse(location.state.selectedformdetails?.FormJSON) || []);
    setParentFormCopy(JSON.parse(location.state.selectedformdetails?.FormJSON) || []);

    setSelectedFormTitle(location.state.selectedformdetails?.AppName || "");
    setSelectedFormAppCode(location.state.selectedformdetails?.AppCode || "");

  }, [])

  const fetchChildForms = async () => {
    debugger;
    const RAWChildForms = await sp.web.lists.getByTitle("FormMaster")
      .items
      .filter(`ParentAppCode eq '${location.state.selectedformdetails.AppCode}'`)
      .getAll();
    const ChildForms = RAWChildForms?.map((item: any) => {
      return { ...item, FormJSON: dataserviceobj.decryptjson(item.FormJSON) };
    })
    //console.log("childforms", ChildForms);
    if (ChildForms.length > 0) {
      let firstchildform = []
      firstchildform.push(ChildForms[0])
      setChildFormDetails(firstchildform)
    }
  }
  const handleTableUpdate = (updatedTable: TableData) => {
    setTables(prevTables => {
      const updatedTables = prevTables.map(table =>
        table.id === updatedTable.id ? updatedTable : table
      );
      //console.log('All tables updated:', updatedTables);
      return updatedTables;
    });
  };

  const AnswerJsonSeparator = (ParentCopy: any[], Answers: any[]) => {
    let ParentAnswers: any[] = []
    let ChildAnswers: any[] = []
    Answers.forEach((ele) => {
      ParentCopy.find(elem => elem.id === ele.id) ? ParentAnswers.push(ele) : ChildAnswers.push(ele)
    })

    return [ParentAnswers, ChildAnswers]
  }
  const handleRestrictionBreachUpdate = (hasBreaches: boolean) => {
    setHasBreaches(hasBreaches);
    console.log("Has Restriction Breaches for this table:", hasBreaches);
  };
  useEffect(() => {
    //console.log("yttttttttttttttttttttttt", location.state.selectedformdetails)
    void fetchChildForms();
    const TableStructureJsons = selectedFormJSON.filter(ele => ele.element === "Label" && ele?.text === "Table")
    const TransformedTableJsons = TableStructureJsons?.map((ele) => {
      return ({
        id: ele.id,
        content: stripHtmlTags(ele.content),
        text: ele.text,
        numberOfColumns: ele.numberofcolumns,
        columns: ele.columns,
        rows: []
      })
    })
    setTables(TransformedTableJsons || []);
    TransformedTableJsons.forEach((field: any) => {
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
          <TableComponent tableConfig={field} onTableUpdate={handleTableUpdate} onRestrictionBreachUpdate={handleRestrictionBreachUpdate}
          />,
          parent
        );
      }
    });

  }, [selectedFormJSON])

  const handleTimeTableUpdate = (updatedTable: TimeSheetData) => {
    //console.log('Updated TimeTable:', updatedTable);
    setTimesheetTables(prevTables => {
      const updatedTables = prevTables.map(table =>
        table.id === updatedTable.id ? updatedTable : table
      );
      //console.log('All timetables updated:', updatedTables);
      return updatedTables;
    });
    // if(updatedTable.fromdate && updatedTable.todate){
    //   new Date(updatedTable.fromdate) > new Date(updatedTable.todate) ? (() =>{toast?.current?.clear()
    //   toast?.current?.show({
    //       severity: "error",
    //       detail: "from date should be less than to date",
    //       life: 2000
    //   })}):''
    // }
  };
  const handleTimeRestrictionBreachUpdate = (hasBreaches: boolean) => {
    setHasBreaches(hasBreaches);
    //console.log("Has Restriction Breaches for this table:", hasBreaches);
  };
  useEffect(() => {
    const TimeTableStructureJsons = selectedFormJSON.filter(ele => ele.element === "Download" && ele?.text === "Timesheet")
    const TransformedTimeTableJsons = TimeTableStructureJsons?.map((ele) => {
      return ({
        id: ele.id,
        content: stripHtmlTags(ele.content),
        text: ele.text,
        numberOfColumns: ele.numberofcolumns,
        columns: ele.columns,
        rows: [],
        fromdate: ele.fromdate?.value,
        todate: ele.todate?.value,
        weekends: ele.weekends,
        maxHoursPerCell: ele.maxHoursPerCell,
        minHoursPerCell: ele.minHoursPerCell,
      })
    })
    setTimesheetTables(TransformedTimeTableJsons || []);
    TransformedTimeTableJsons.forEach((field: any) => {
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
          <TimesheetComponent tableConfig={field} onTableUpdate={handleTimeTableUpdate} onRestrictionBreachUpdate={handleTimeRestrictionBreachUpdate}
          />,
          parent
        );
      }
    });

  }, [selectedFormJSON])

  // useEffect(() => {
  //   type uploadelement = {
  //     field_name: string;
  //     id: string;
  //     element: string;
  //     text: string;
  //     group_name: string;
  //     required: boolean;
  //     VisibilityCondition: any[];
  //     canHavePageBreakBefore: boolean;
  //     canHaveAlternateForm: boolean;
  //     canHaveDisplayHorizontal: boolean;
  //     canHaveOptionCorrect: boolean;
  //     canHaveOptionValue: boolean;
  //     canPopulateFromApi: boolean;
  //     label: string;
  //     dirty: boolean
  //   }
  //   const fileuploadelement: uploadelement[] = selectedFormJSON.filter(item => item.element === "FileUpload" && item.text === "File Upload");
  //   console.log("fileelement", fileuploadelement);

  //   if (!fileuploadelement) return;

  //   const fileuploadnode = document.querySelector(`[name="${fileuploadelement[0].field_name}"]`);

  //   console.log("fieldname", fileuploadnode);
  //   fileuploadnode?.setAttribute('multiple', '');


  //   if (!fileuploadnode) return;

  //   const handlefilchange = (e: Event) => {
  //     setTimeout(() => {
  //       const target = e.target as HTMLInputElement;
  //       const files = target.files ? Array.from(target.files) : [];
  //       const fileListContainer = document.querySelector('.file-upload-preview');
  //       console.log("domeleeeeeeeeeee", fileListContainer);
  //       // Clear previous list
  //       if (fileListContainer) {
  //         fileListContainer.innerHTML = '';

  //         // Display new list
  //         files.forEach((file) => {
  //           const fileItem = document.createElement('div');
  //           fileItem.textContent = `Name: ${file.name} — Size: ${Math.ceil(file.size / 1024)} KB`;
  //           fileListContainer.appendChild(fileItem);
  //         })
  //       }

  //     }, 1);

  //   }

  //   fileuploadnode.addEventListener('change', handlefilchange);

  //   return () => {
  //     fileuploadnode.removeEventListener('change', handlefilchange);
  //   }

  // }, [])
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
    //console.log("fileeleeemenen", fileElements)
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
                //console.log("clearbtn", clearBtn)
                //console.log("fromobserve", files);
                setFileData((prev) => ({
                  ...prev,
                  [fieldName]: files,
                }));
                //console.log("newfiledata", fieldName)
                if (fileListContainer) {
                  fileListContainer.innerHTML = "";
                  files.forEach((file) => {
                    const fileItem = document.createElement("div");
                    fileItem.textContent = `Name: ${file.name} — Size: ${Math.ceil(file.size / 1024)} KB`;
                    fileListContainer.appendChild(fileItem);
                  });
                }
                const handleClearClick = () => {
                  //console.log(`Cleared files for ${fieldName}`);

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

  useEffect(() => {
    //console.log("filedatachanged", fileData);

  }, [fileData])



  // const fetchingTenantUsers = async () => {
  //   try {
  //     const allUsers = await fetchAllUsers();
  //     // if (allUsers) {  // Ensure data is available before updating state
  //     //   console.log("Fetched Users:------------------------------------ ", allUsers);
  //     //   setMyOrganizationUsers(allUsers)
  //     //   // setFilteredUsers(allUsers.slice(0, 10));
  //     // } else {
  //     //   console.warn("No users found or API call failed.");
  //     // }
  //   } catch (error) {
  //     console.error("Error fetching users:", error);
  //   }
  // }




  // useEffect(() => {

  //   setFilteredUsers(myOrganizationUsers.slice(0, 10));
  // }, [myOrganizationUsers]);



  function getPeoplePickerContext(): IPeoplePickerContext {
    return {
      absoluteUrl: context.pageContext.web.absoluteUrl,
      msGraphClientFactory: context.msGraphClientFactory,
      spHttpClient: context.spHttpClient,
    };
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

  useEffect(() => {
    document.querySelectorAll("select").forEach((el) => {
      el.classList.add("form-select");          //for chevrondown visibility in dropdown 
    });

    try {
      const formdatefields: any[] = selectedFormJSON.filter((ele: any) => { return ele.field_name?.startsWith("date_picker") });
      setdatefieldids(formdatefields)
    } catch (error) {
      //console.log("formdatefields error", error)
    }
  }, []);


  // function stripHtmlTags(label: any) { return label.replace(/<\/?[^>]+(>|$)/g, "").trim(); }
  function stripHtmlTags(label: any) { return label?.replace(/<\/?[^>]+(>|$)/g, "")?.trim() || ""; }

  useEffect(() => {
    const dependentFields = selectedFormJSON.filter((ele: any) => { //All child fields having conditions
      return ele?.dependentVisibility
    })
    SetChildFieldJsons(dependentFields)
    const dependentShowFields = selectedFormJSON.filter((ele: any) => { //Child fields which should be hidden initially
      return ele?.dependentVisibility && ele?.controlVisibility
    })
    SetHiddenData(dependentShowFields)
    //console.log("dependentFields", dependentFields)
    const dependentShowElements = dependentShowFields.map((elem: any) => { //child elements which should be hidden initially
      if (elem?.parentId) {
        if (elem.element === "Checkboxes") {
          const idPart = elem?.options?.[0]?.key;

          const matchedElement = Array.from(document.querySelectorAll('input[id]')).find(
            el => el.id.includes(idPart)
          );

          return matchedElement?.closest('.SortableItem.rfb-item')?.parentElement as HTMLElement;

        }
        return document.querySelector(`[name="${elem.field_name}"]`)?.closest('.SortableItem.rfb-item')?.parentElement as HTMLElement
      } else {
        if (elem.element === "Checkboxes") {
          const idPart = elem?.options?.[0]?.key;

          const matchedElement = Array.from(document.querySelectorAll('input[id]')).find(
            el => el.id.includes(idPart)
          );

          return matchedElement?.closest('.SortableItem.rfb-item') as HTMLElement;

        }
        return document.querySelector(`[name="${elem.field_name}"]`)?.closest('.SortableItem.rfb-item') as HTMLElement
      }
    })

    //console.log("dependentElements", dependentShowElements)

    dependentShowElements.forEach((ele: any) => { //hiding these elements
      if (ele) {
        (ele as HTMLElement).style.display = "none";
      }
    })

    const SourceFieldNames = [...new Set(dependentFields.map((ele: any) => ele?.VisibilityCondition[0]?.SourceField))]

    //console.log("SourceFieldNames", SourceFieldNames)

    const SourceFields = selectedFormJSON.filter((ele: any) => {
      return ele?.label && SourceFieldNames.includes(stripHtmlTags(ele?.label))
    })
    //console.log("SourceFields", SourceFields)

    SetSourceFieldJsons(SourceFields)

    const transformedArray = SourceFields.map((item: any) => ({
      id: item.id,
      name: item.field_name,
      custom_name: item.field_name,
      value: item.element === "RadioButtons" ? [] : ""
    }));
    //console.log("transformedarray", transformedArray)
    SetSrcFieldAnswer(transformedArray)



  }, [selectedFormJSON])

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

  const funforVisibility = (data: any) => {
    //console.log("Data in funforvisibility", data)
    //console.log("sourcefieldjsons====", SourceFieldJsons)
    //console.log("srcfieldans", SrcFieldAnswer)
    let filtereddata = data.filter((item: any) => {
      return SourceFieldJsons.find(ele => ele.id === item.id)
    })
    //console.log("filtereddata", filtereddata)
    //console.log("ChildFieldJsons", ChildFieldJsons)
    let idboolobjArray = compareValues(filtereddata, SrcFieldAnswer)
    //console.log("idboolobjArray", idboolobjArray)
    let updatedFormData = [...data];
    idboolobjArray.forEach((element: any) => {
      if (!element.isSame || true) {
        let field = SourceFieldJsons.find(ele => ele.id === element.id)
        //console.log("field", field)
        ChildFieldJsons.forEach((childjson) => {
          //console.log("689745689=====", childjson?.VisibilityCondition[0]?.SourceField)
          //console.log("stripHtmlTags(field.label)", stripHtmlTags(field.label))
          if (childjson?.VisibilityCondition[0]?.SourceField === stripHtmlTags(field.label)) {
            let leftvalue
            if (childjson?.VisibilityCondition[0]?.element === "RadioButtons") {
              //console.log("element.options====", field.options)
              let optionjson = field.options.find((lfv: any) => {
                return lfv.key === element.value[0]
              });
              //console.log("lfvv====", optionjson)
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
            //console.log("parameters condn validator", leftvalue, childjson?.VisibilityCondition[0]?.Operator, childjson?.VisibilityCondition[0]?.TargetValue, field.element)

            let newbool = ConditionValidator(
              leftvalue,
              childjson?.VisibilityCondition[0]?.Operator,
              childjson?.VisibilityCondition[0]?.TargetValue,
              field.element)
            //console.log("newbool", newbool)
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

            //console.log("visiblechangeelement=====", visiblechangeelement)
            if ((childjson?.controlVisibility && newbool) || (!childjson?.controlVisibility && !newbool)) {
              if (visiblechangeelement) {
                (visiblechangeelement as HTMLElement).style.display = "block";
              }
              SetHiddenData(prev => prev.filter(ele => ele.id !== childjson?.id))
            }
            else if ((!childjson?.controlVisibility && newbool) || (childjson?.controlVisibility && !newbool)) {
              if (visiblechangeelement) {
                (visiblechangeelement as HTMLElement).style.display = "none";
              }
              SetHiddenData(prev => [...prev, childjson])
              // Get the field element
              const fieldElement = document.querySelector(`[name="${childjson.field_name}"]`);
              if (childjson.element === "Checkboxes") {
                //console.log("if entered")
                childjson.options.forEach((checkbox: any) => {
                  const idPart = checkbox?.key;
                  //console.log("foreach entered")
                  const matchedElement = Array.from(document.querySelectorAll('input[type="checkbox"][id]')).find(
                    el => el.id.includes(idPart)
                  ) as HTMLInputElement | undefined;

                  if (matchedElement) {
                    //console.log("false block entered")
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

  useEffect(()=>{
  //console.log("tabclicked")
  funforVisibility(formData);
},[tabclicked])
 

  const toast = useRef<Toast>(null);
  let toasterrors: string[] = []
  sp.setup({
    sp: {
      baseUrl: mySiteUrl,
    },
  });

  const getloggedInUser = async () => {

    const curUser = await sp.web.currentUser();
    const domain = await fetchTenantUser();
    setDomainField(domain.TenantUsers)
    const RequesterTitle = curUser.Title.split("|")[0];
    setRequesterTitle(RequesterTitle);
    setRequesterEmail(curUser.Email)
    setRequesterName("Requester: " + RequesterTitle);
  }

  const onChange = async (items: any[]) => {
    //console.log("Selected People:", items);

    try {
      const users = await Promise.all(items.map(async (useritem) => {
        const user = await sp.web.siteUsers.filter(`Email eq '${useritem.secondaryText}'`).select('Id,Email').get()
        return user[0]?.Id
      }))

      //console.log("AllUserIds", users);

      const validUserIds = users.filter(id => id !== undefined);
      //console.log("Selected valid User IDs:", validUserIds);
      SetnotifytorequestorID(validUserIds || [])

    }
    catch (err) {
      //console.log(err);
    }
  };

  useEffect(() => {
    void getloggedInUser();
    // void fetchingTenantUsers()
    // void PrefixFetch()
  }, []);



  const generateRandomString = (length: number) => {
    const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
    return Array.from({ length }, () => characters.charAt(Math.floor(Math.random() * characters.length))).join('');
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

  const getRoutingRules = async (appCode: string, buttonAction: string) => {
    return await sp.web.lists.getByTitle("RoutingRules").items.orderBy("Modified", false).filter(`(AppCode eq '${appCode}') and (Actions eq '${buttonAction}')`).get();
  };

  const getCurrentUserSeqNo = async () => {
    const currentUser = await sp.web.currentUser.get();
    const displayNamePrefix = currentUser.Title.substring(0, 6).toUpperCase();
    const randomString = generateRandomString(6);
    return `${displayNamePrefix}-${randomString}`;
  };
  function convertDateFormat(date: string) {
    let [year, month, day] = date.split('-')
    return `${month}/${day}/${year}`
  }


  const getManagerFromAD = async (): Promise<number | null> => {
    try {
      // Get the current logged-in user
      const currentUser = await sp.web.currentUser();
      const accountName = currentUser.LoginName; // Use the LoginName field
      //console.log("Current User LoginName: ", accountName);

      // Construct the PeopleManager endpoint
      const endpoint = `${mySiteUrl}/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='${encodeURIComponent(accountName)}'`;
      //console.log("Constructed Endpoint: ", endpoint);

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
        //console.log("Fetched Data: ", data);
        const extendedManagers = data.d?.ExtendedManagers?.results || [];
        //console.log("Extended Managers: ", extendedManagers);

        if (extendedManagers.length > 0) {
          // Get the last manager in the array
          const lastManager = extendedManagers[extendedManagers.length - 1];
          // const email = lastManager.split("|")[2];
          //console.log("Manager Email: ", lastManager);

          // Use PnP JS to get user by email

          // const user = await sp.web.siteUsers.getByEmail("rakesh.chitte@cloudangles.com")();
          const user = await sp.web.ensureUser(lastManager);
          //console.log("Manager User Details---------000000000: ", user);

          // Return the manager's user ID
          return user.data.Id;
        } else {
          //console.log("No managers found.");
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

  const handleBlur = (data: any) => {
    if (datefieldids.length > 0) {
      datefieldids.forEach(dateElement => {
        let datefieldonchange = document.querySelector(`input[name="${dateElement.field_name}"]`) as HTMLInputElement;
        if (datefieldonchange) {
          datefieldonchange.addEventListener("change", (e: any) => {
            let dateJson = data.find((element: any) => { return element.name === dateElement.field_name })
            dateJson.value = e.target.value;
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

      // setFormData(data);
      setFormData(updatedData);
      //console.log("fromhandleblur", updatedData);
      //console.log("justdata", data)
    }
  }
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

        //console.log(`Comparing entry ${i} (${aFrom} - ${aTo}) with entry ${j} (${bFrom} - ${bTo})`);

        // Check for overlap (simplified logic)
        if (aFrom <= bTo && aTo >= bFrom) {
          //console.log("Found overlap between:", entryA.content, "and", entryB.content);
          return [
            entryA.content || `Entry ${i}`,
            entryB.content || `Entry ${j}`
          ];
        }
      }
    }

    //console.log("No overlapping entries found");
    return null;
  };

  const handleSubmit = async (data: any) => {
    const { recordscount, trnxRecLimitReached } = await fetchWorkflowLimit(domainfield, selectedFormAppCode);
    if (trnxRecLimitReached) {
      toast.current?.clear();
      toast.current?.show({
        severity: 'warn',
        summary: 'Freemium subscription limit reached',
        detail: `${selectedFormTitle}: You reached ${recordscount} transactions. Please upgrade your license.`,
        life: 8000,
        style: { height: 100 }
      });
      return;
    }
    if (data.length === 0 || tables.length > 0 || timesheettables.length > 0) {
      toast.current?.show({ severity: 'error', summary: '', detail: 'Please fill out the form before submitting.', life: 2000 });
      return
    }
    if (datefieldids.length > 0) {
      datefieldids.forEach(dateElement => {
        let datefieldonchange = document.querySelector(`input[name="${dateElement.field_name}"]`) as HTMLInputElement;
        if (datefieldonchange) {
          datefieldonchange.addEventListener("change", (e: any) => {
            let dateJson = data.find((element: any) => { return element.name === dateElement.field_name })
            dateJson.value = e.target.value
          })
          datefieldonchange.dispatchEvent(new Event('change'))
        }
      });
    }
    setLoading(true)
    //console.log("Submitted data--------------", data);

    // function stripHtmlTags(label: any) { return label.replace(/<\/?[^>]+(>|$)/g, "").trim(); }

    const buttonAction = "Submit";
    const files: { name: string; file: File }[] = [];

    const parentFiles: { name: string; file: File }[] = [];
    const childFiles: { name: string; file: File }[] = [];

    // const tablefiles: { name: string; file: File }[] = [];

    let hasFileUploadField = false;
    let hasUploadedFiles = false;

    console.log(hasFileUploadField, hasUploadedFiles)
    const toasterrors: string[] = [];
    if (hasBreaches) {
      toasterrors.push("Invalid Table Data")
    }
    const result1 = data.filter(
      (dataitem: any) => !hiddendata.some(subItem => subItem.id === dataitem.id)
    );

    result1.forEach((item: any) => {
      if (item.name && item.name.includes("file_upload")) {
        //console.log("insideresult1fd", fileData)
        const uplfiles = fileData[item.name];
        item.value = fileData[item.name];
        //console.log("uplfiles", uplfiles)
        if (uplfiles) {
          uplfiles.map((file: File) => {
            if (file && file.name) {
              files.push({ name: file.name + "name" + item.name, file });
              //console.log("frominsideif", files);
            }
          });
        }
      }

      let formjsonitem = selectedFormJSON.find((formjson) => {
        return formjson.id === item.id;
      });
      switch (formjsonitem.text) {
        case "File Upload":
          if (formjsonitem.required && (!item.value)) {
            //console.log("insidecheck", item.value, !item.value)
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
            //console.log("today and itemdate", today, itemDate)
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
          break;
        default:
          if (formjsonitem.required && !item.value?.length) {
            toasterrors.push(`${stripHtmlTags(formjsonitem.label)} : This field is required.`);
          }
          break;
      }
    });
    if (!Categorytype) {
      toasterrors.push("Request For Field is required");
    }
    else if (Categorytype.length > 80) {
      toasterrors.push("Request For : Only 80 characters are allowed ");
    }
    //console.log("taabledata", tables)
    tables.forEach((table: any) => {
      const attachcolumnName = table.columns.find((column: any) => column.columntype === "Attachment")?.headername;

      table.rows.forEach((row: any) => { row[attachcolumnName]?.fileObject?.size > 1024 * 1024 ? toasterrors.push(`${table.content}: File size exceeded , max-allowed size is 1mb`) : null })
    }
    )

    const overlappingContent = getFirstOverlappingContent(timesheettables);
    if (overlappingContent) {
      toasterrors.push(`Overlapping timesheet entries found for: ${overlappingContent.join(", ")}`);
    }
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
    const encformjson = await dataserviceobj.encryptjson(updatedFormData);
    console.log("encryptedrecord", updatedFormData, encformjson);
    //console.log("From the handle submit function : Form Data: -->", updatedFormData);
    //console.log("Files:", files);

    debugger;
    const appName = selectedFormTitle || "";
    const appCode = selectedFormAppCode || "";

    try {
      debugger
      let routingRulesItems = []
      if (ChildFormDetails.length > 0) {
        routingRulesItems = await getRoutingRules(ChildFormDetails[0].AppCode, buttonAction);
      }
      else {
        routingRulesItems = await getRoutingRules(appCode, buttonAction);
      }

      if (routingRulesItems.length === 0) {
        setLoading(false)
        toast.current?.show({ severity: 'error', summary: '', detail: 'Reviewers or Approvers are not available for further actions', life: 2000 });
        return;
      }
      // pushing into the data into the sharepoint list.

      const { CurrentQueue, DestinationQueue, Status } = routingRulesItems[0];
      const mappingMasterItem = await sp.web.lists
        .getByTitle("MappingMaster")
        .items.orderBy("Modified", false).select("*", "Users_x002f_Groups/Id", "Users_x002f_Groups/Title")
        .expand("Users_x002f_Groups")
        .filter(`(Level eq '${DestinationQueue}') and (AppCode eq '${ChildFormDetails.length > 0 ? ChildFormDetails[0].AppCode : appCode}')`)
        .get();
      //console.log("mappingMasterItem", mappingMasterItem);


      let conditionlistItem: any[] = [];
      let result: boolean = false;

      try {
        // Fetch data from the "ConditionsList" list
        conditionlistItem = await sp.web.lists
          .getByTitle("ConditionsList")
          .items.orderBy("Modified", false).select("*", "PersonOrGroup/Id", "PersonOrGroup/Title")
          .expand("PersonOrGroup")
          .filter(`(Level eq '${mappingMasterItem[0].Level}') and (AppCode eq '${ChildFormDetails.length > 0 ? ChildFormDetails[0].AppCode : appCode}')`)
          .get();

        //console.log("-------conditionlistItem---------", conditionlistItem);

        // Check if the data exists
        if (conditionlistItem.length === 0) {
          throw new Error("No data found in ConditionsList for the provided filter.");

        }

        // Parse the ConditionJson
        const conditionlistItemJson: any = JSON.parse(conditionlistItem[0].ConditionJson);
        const mainoperator: string = conditionlistItemJson.operator;
        result = mainoperator === "And" ? true : false;

        // Process each condition in the JSON
        conditionlistItemJson.conditions.forEach((element: any) => {
          try {
            // Find fieldId from selectedFormJSON
            const fieldId = selectedFormJSON.find((val: any) => stripHtmlTags(val.label) === element.field)?.id;

            if (!fieldId) {
              throw new Error(`Field ID not found for field: ${element.field}`);
            }

            // Find the corresponding field value from data
            const fieldvalue = data.find((d: any) => d.id === fieldId)?.value;

            if (fieldvalue === undefined) {
              throw new Error(`Field value not found for field ID: ${fieldId}`);
            }

            // Validate the condition
            const boolvalue = ConditionValidator(fieldvalue, element.operator, element.value, element.elementType);

            // Update the result based on the operator
            result = mainoperator === "And" ? boolvalue && result : boolvalue || result;

            //console.log("-------result inside--", result);
          } catch (innerError: any) {
            console.error("Error processing condition element:", innerError.message);
          }
        });

        //console.log("Final result:", result);
      } catch (error: any) {
        console.error("Error occurred while processing ConditionsList:", error.message);

      }

      // Log the final result after the try-catch block
      //console.log("-------result--", result);
      // const userGroupId = result? conditionlistItem[0].PersonOrGroup.Id: mappingMasterItem[0].Users_x002f_Groups.Id;
      //  let userGroupId =  result? conditionlistItem[0].PersonOrGroup.Id:mappingMasterItem?.[0]?.Users_x002f_Groups?.Id || null;
      let userGroupId: number | null = null;

      if (result && conditionlistItem.length > 0) {
        // Safely access conditionlistItem[0].PersonOrGroup.Id if conditionlistItem is not empty
        userGroupId = conditionlistItem[0]?.PersonOrGroup?.Id || null

      } else {
        // Safely access mappingMasterItem[0].Users_x002f_Groups.Id
        userGroupId = mappingMasterItem?.[0]?.Users_x002f_Groups?.Id || null;
      }

      //console.log("User Group ID:", userGroupId);

      //console.log("User Group ID--------------------: ", userGroupId);

      if (userGroupId) {
        //console.log("User Group ID: ", userGroupId);
        // Proceed with logic using userGroupId
      } else {
        //console.log("User Group ID is null, fetching manager from AD...");

        try {
          const managerId = await getManagerFromAD();

          if (managerId) {
            //console.log("Immediate Manager: ", managerId);


            userGroupId = managerId;
          } else {
            //console.log("No manager found or an error occurred.");
          }
        } catch (error) {
          console.error("Error fetching manager or user details: ", error);

        }
      }

      const allItems = await sp.web.lists
        .getByTitle("WorkFlowProcessData")
        .items // Ensure newest records come first
        .filter(`AppCode eq '${appCode}' and Domain eq '${domainfield}' `).orderBy("Created", false)
        .getAll()
        .catch((error) => {
          console.error("Error fetching items:", error);
          return [];
        });

      const formMasterItem2 = await sp.web.lists.getByTitle("FormMaster")
        .items.select("FormPrefix", "isWorkflowRequired")
        .filter(`AppName eq '${appName}' and Domain eq '${domainfield}' `).get();

      let isNoWorkflowRequiredFlag = formMasterItem2[0]?.isWorkflowRequired === "No" ? true : false;

      let latestNumber = 0;
      debugger;
      //console.log("allItems------------", allItems)
      for (const item of allItems) {
        if (item.Status !== "Draft") {
          const titleParts = item.Title.split("/");

          //console.log("titleParts----------------- !!!!!!!!!!!!!!", titleParts, titleParts.length);

          if (titleParts.length >= 3) {
            const numberPart = parseInt(titleParts[2], 10);
            //console.log("numberPart", numberPart);

            if (!isNaN(numberPart) && numberPart > latestNumber) {
              latestNumber = numberPart; // Store the highest number found
            }
          }
        }
      }
      const seqNo = await getCurrentUserSeqNo();
      const date = new Date();
      const month = date.toLocaleString('default', { month: 'short' });
      const year = date.getFullYear().toString().slice(-2);
      const monthYear = `${month}_${year}`;

      const newTitle = `${formMasterItem2[0].FormPrefix}/${monthYear}/${latestNumber + 1}`;

      //console.log("New Title:", newTitle);

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
        //console.log("aduserdetails--------------> ", aduserdetails);
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

      // Convert the adUserData object to a JSON string
      const adUserDataJSON = JSON.stringify(adUserData);
      //here
      let [ParentAnswers, ChildAnswers] = AnswerJsonSeparator(ParentFormCopy, updatedFormData)

      if (IsParentForm) {
        const AllFilteredDataExceptHidden = data.filter(
          (dataitem: any) => !hiddendata.some(subItem => subItem.id === dataitem.id)
        );

        AllFilteredDataExceptHidden.forEach((item: any) => {
          if (item.name && item.name.includes("file_upload")) {
            const uplfiles = fileData[item.name];
            if (uplfiles) {
              uplfiles.forEach((file: File) => {
                if (file && file.name) {
                  const isParent = ParentFormCopy.find(elem => elem.id === item.id);
                  if (isParent) {
                    parentFiles.push({ name: file.name + "name" + item.name, file });


                  } else {
                    childFiles.push({ name: file.name + "name" + item.name, file });
                  }
                }
              });
            }
          }
        });
      }


      let NoApprovalStatus = "";
      debugger;
      if (isNoWorkflowRequiredFlag) {
        NoApprovalStatus = "Submitted";
      }
      else {
        NoApprovalStatus = Status
      }
      //console.log("mappingMastrerItmes", mappingMasterItem);
      debugger;
      const newItem = await sp.web.lists.getByTitle("WorkFlowProcessData").items.add({
        Title: newTitle,
        // FormData: JSON.stringify(updatedFormData),
        FormData: JSON.stringify(ParentAnswers),
        AppName: appName,
        AppCode: appCode,
        CurrentAppCode: ChildFormDetails[0]?.AppCode,
        Domain: domainfield,
        Status: NoApprovalStatus,
        CurApproverId: userGroupId,
        CurrentQueue: CurrentQueue,
        DestinationQueue: DestinationQueue,
        Level: mappingMasterItem[0].Level,
        SeqNo: seqNo,
        ManagersId: userGroupId,
        AuthorADdata: adUserDataJSON,
        CategoryType: Categorytype,
        NotifyRequestorId: {
          results: notifytorequestorID
        },
        TableJSON: JSON.stringify(IsParentForm ? [] : tables),
        TimeSheetJSON: JSON.stringify(IsParentForm ? [] : timesheettables),
        MinDateArray: JSON.stringify(IsParentForm ? [] : timesheettables.map((rec) => rec.fromdate)),
        MaxDateArray: JSON.stringify(IsParentForm ? [] : timesheettables.map((rec) => rec.todate))
        // MultipleUsersId:{
        //   results:notifytorequestorID
        // }
        // NotifyRequestorId:{
        //   "results":[
        //     {Id:23}
        //   ]
        // }
      });
      const seqNochild = await getCurrentUserSeqNo();
      if (ChildFormDetails.length > 0) {
        await sp.web.lists.getByTitle("WorkFlowProcessData").items.add({
          Title: newTitle,
          FormData: JSON.stringify(ChildAnswers),
          AppName: ChildFormDetails[0].AppName,
          AppCode: ChildFormDetails[0].AppCode,
          ParentAppCode: appCode,
          ParentAppId: newItem.data.Id,
          SeqNo: seqNochild,
          AuthorADdata: adUserDataJSON,
          Domain: domainfield,
          Status: Status,
          CurApproverId: userGroupId,
          CurrentQueue: CurrentQueue,
          DestinationQueue: DestinationQueue,
          Level: mappingMasterItem[0].Level,
          ManagersId: userGroupId,
          CategoryType: Categorytype,
          NotifyRequestorId: {
            results: notifytorequestorID
          },
          TableJSON: JSON.stringify(tables),
          TimeSheetJSON: JSON.stringify(timesheettables),
          MinDateArray: JSON.stringify(timesheettables.map((rec) => rec.fromdate)),
          MaxDateArray: JSON.stringify(timesheettables.map((rec) => rec.todate))
        });

        //NoteAttach of child Files
        const ChildfolderName = `NoteAttach/${seqNochild}`;
        await sp.web.folders.add(ChildfolderName);
        for (const file of childFiles) {
          await sp.web.getFolderByServerRelativeUrl(ChildfolderName).files.add(file.name, file.file, true);
        }
      }

      let AppCodeee = appCode;
      let ParentAppCodeee = "";
      let Sequence = seqNo;

      if (ChildFormDetails.length > 0) {
        AppCodeee = ChildFormDetails[0]?.AppCode;
        ParentAppCodeee = appCode;
        Sequence = seqNochild;
      }

      const workflowlog = await sp.web.lists.getByTitle("WorkFlowLog").items.add({
        AppCode: AppCodeee,
        ParentAppCode: ParentAppCodeee,
        Status: NoApprovalStatus,
        Domain: domainfield,
        Level: mappingMasterItem[0].Level,
        SeqNo: Sequence,
        Role: routingRulesItems[0].Role,
        Action: buttonAction
      })
      console.log("------------------workflowlog", workflowlog)
      // for incrementing the count in formmaster after a record creation in workflowprocessdata
      const formMasterItem = await sp.web.lists.getByTitle("FormMaster")
        .items.select("Id", "TransactionCount")
        .filter(`AppName eq '${appName}' and Domain eq '${domainfield}' `).top(1).get();

      const recordscount = await sp.web.lists.getByTitle("WorkFlowProcessData")
        .items.filter(`AppName eq '${appName}' and Domain eq '${domainfield}'`).getAll();
      const processdatacount = recordscount.length;
      if (formMasterItem.length > 0) {
        const itemId = formMasterItem[0].Id;

        await sp.web.lists.getByTitle("FormMaster")
          .items.getById(itemId)
          .update({ TransactionCount: processdatacount })
          .catch(error => {
            console.error("Conflict detected! Please retry.", error);
          });
      }
      // for incrementing the count in formmaster after a record creation in workflowprocessdata

      const itemId = newItem.data.Id;
      console.log(itemId)


      if (!IsParentForm) {
        const folderName = `NoteAttach/${seqNo}`;
        await sp.web.folders.add(folderName);

        // Upload files to the subfolder if any
        for (const file of files) {
          await sp.web.getFolderByServerRelativeUrl(folderName).files.add(file.name, file.file, true);
        }
      }
      else {
        const ParentForlderName = `NoteAttach/${seqNo}`;
        await sp.web.folders.add(ParentForlderName);

        // Upload files to the subfolder if any
        for (const file of parentFiles) {
          await sp.web.getFolderByServerRelativeUrl(ParentForlderName).files.add(file.name, file.file, true);
        }
      }

      // ------------------------------------------Saving table attachments to library-----------------------------------------------

      tables.forEach(async (table: any) => {
        const attachcolumnName = table.columns.find((column: any) => column.columntype === "Attachment")?.headername
        const boolvalue: boolean = table.rows.find((row: any) => row[attachcolumnName].name.length > 0)
        if (attachcolumnName && boolvalue) {
          const tablefoldername = `TableAttachments/${seqNo}`;
          await sp.web.folders.add(tablefoldername);
          const subfolderName = `${tablefoldername}/${table.id}`;
          await sp.web.folders.add(subfolderName);

          for (const file of table.rows) {
            //console.log("----file=====", file)
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

      // setLoading(false);
      setFormData([]);
      setLoading(false)
      toast.current?.show({ severity: 'success', summary: '', detail: 'Your Request has been submitted successfully!', life: 2000 });


      setTimeout(() => {
        navigate(-1)
      }, 1000);

    } catch (error: unknown) {
      handleError(error);
      //console.log("error0=-------------------", error)
      setLoading(false)
      toast.current?.show({ severity: 'error', summary: '', detail: `${error}`, life: 2000 });

    }

  };


  const handleSave = async (data: any) => {
    if (datefieldids.length > 0) {
      datefieldids.forEach(dateElement => {
        let datefieldonchange = document.querySelector(`input[name="${dateElement.field_name}"]`) as HTMLInputElement;
        if (datefieldonchange) {
          datefieldonchange.addEventListener("change", (e: any) => {
            let dateJson = data.find((element: any) => { return element.name === dateElement.field_name })
            dateJson.value = e.target.value
            //console.log("sssssssssssss", e.target.value);
            // data[5].value=e.target.value
          })
          datefieldonchange.dispatchEvent(new Event('change'))
        }

      });
    }
    setLoading(true)
    const files: { name: string; file: File }[] = [];

    const parentFiles: { name: string; file: File }[] = [];
    const childFiles: { name: string; file: File }[] = [];


    let hasFileUploadField = false;
    let hasUploadedFiles = false;

    console.log(hasFileUploadField, hasUploadedFiles)
    // Process file upload fields and gather file details
    if (!IsParentForm) {
      data.forEach((item: any) => {
        //console.log("item", item)
        //console.log("itemname", item.name)
        if (item.name && item.name.includes("file_upload")) {
          //console.log("insideresult1fd", fileData)
          const uplfiles = fileData[item.name]
          //console.log("uplfiles", uplfiles)
          if (uplfiles) {
            uplfiles.map((file: File) => {
              if (file && file.name) {
                files.push({ name: file.name + "name" + item.name, file });
                //console.log("frominsideif", files);
                hasUploadedFiles = true;
              }
            });
          }
        }

      })
    }


    //console.log("taabledata", tables)
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

    // const updatedFormData = data.map((item: any) => {
    //   if (item.name && item.name.includes("file_upload") && item.value && item.value.name) {
    //     return {
    //       ...item,
    //       value: item.value.name, // Replace file object with file name
    //     };
    //   }
    //   return item;
    // });

    // console.log("From the handle save function : Form Data: -->", updatedFormData);
    // console.log("Files:", files);

    const appName = selectedFormTitle || "";
    const appCode = selectedFormAppCode || "";
    const formPrefix = appName
      .split(" ")
      .map((word: any) => word.charAt(0).toUpperCase())
      .join("") + "/";

    try {
      const buttonAction = "Save";
      // const routingRulesItems = await getRoutingRules(appCode, buttonAction);
      let routingRulesItems = [];
      if (ChildFormDetails.length > 0) {
        routingRulesItems = await getRoutingRules(ChildFormDetails[0].AppCode, buttonAction);
      }
      else {
        routingRulesItems = await getRoutingRules(appCode, buttonAction);
      }
      if (routingRulesItems.length === 0) {
        // setError("No routing rules found for saving.");
        setLoading(false)
        toast.current?.show({ severity: 'error', summary: '', detail: 'No routing rules found for saving.', life: 2000 });
        return;
      }
      // toast.current?.show({ severity: 'info', summary: '', detail: 'Data successfully saved as draft!', life: 2000 });
      const { CurrentQueue, DestinationQueue, Status } = routingRulesItems[0];
      const seqNo = await getCurrentUserSeqNo();
      const newTitle = `${formPrefix}Draft`;


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
        //console.log("aduserdetails--------------> ", aduserdetails);
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

      // Convert the adUserData object to a JSON string
      const adUserDataJSON = JSON.stringify(adUserData);

      // const encryptedjson = await dataserviceobj.encryptjson(updatedFormData);

      let [ParentAnswers, varChildAnswers] = AnswerJsonSeparator(ParentFormCopy, data);

      let ChildAnswers = varChildAnswers.map((item: any) => {
        if (item.name && item.name.includes("file_upload") && item.value && item.value.name) {
          return {
            ...item,
            value: item.value.name, // Replace file object with file name
          };
        }
        return item;
      });
      if (IsParentForm) {
        data.forEach((item: any) => {
          if (item.name && item.name.includes("file_upload")) {
            const uplfiles = fileData[item.name];
            if (uplfiles) {
              uplfiles.forEach((file: File) => {
                if (file && file.name) {
                  const isParent = ParentFormCopy.find(elem => elem.id === item.id);
                  if (isParent) {
                    parentFiles.push({ name: file.name + "name" + item.name, file });
                  } else {
                    childFiles.push({ name: file.name + "name" + item.name, file });
                  }
                }
              });
            }
          }
        });
      }

      debugger;
      console.log("From the handle save function : Form Data: -->", ChildAnswers);
      //console.log("Files:", files);

      const newItem = await sp.web.lists.getByTitle("WorkFlowProcessData").items.add({
        Title: newTitle,
        FormData: JSON.stringify(ParentAnswers), // Use updatedFormData here
        AppName: appName,
        AppCode: appCode,
        Status: Status,
        Domain: domainfield,
        CurApproverId: null, // No approver for Save action
        CurrentQueue: CurrentQueue,
        DestinationQueue: DestinationQueue,
        Level: DestinationQueue,
        SeqNo: seqNo,
        AuthorADdata: adUserDataJSON,
        CategoryType: Categorytype,
        CurrentAppCode: childselected,
        NotifyRequestorId: {
          results: notifytorequestorID
        },

        TableJSON: JSON.stringify(IsParentForm ? [] : tables),
        TimeSheetJSON: JSON.stringify(IsParentForm ? [] : timesheettables),

      });

      const itemId = newItem.data.Id;
      console.log(itemId)

      if (ChildFormDetails.length > 0) {
        debugger;
        const seqNochild = await getCurrentUserSeqNo();
        //console.log("childformdetials in save ------->>>>>>.>>>.>.>>>>>", ChildFormDetails);
        const newchild = await sp.web.lists.getByTitle("WorkFlowProcessData").items.add({
          Title: newTitle,
          FormData: JSON.stringify(varChildAnswers), // Use updatedFormData here
          AppName: ChildFormDetails[0].AppName,
          AppCode: ChildFormDetails[0].AppCode,
          Status: Status,
          ParentAppCode: newItem.data.AppCode,
          ParentAppId: newItem.data.Id,
          Domain: domainfield,
          CurApproverId: null, // No approver for Save action
          CurrentQueue: CurrentQueue,
          DestinationQueue: DestinationQueue,
          Level: DestinationQueue,
          SeqNo: seqNochild,
          AuthorADdata: adUserDataJSON,
          CategoryType: Categorytype,
          NotifyRequestorId: {
            results: notifytorequestorID
          },
          TableJSON: JSON.stringify(tables),
          TimeSheetJSON: JSON.stringify(timesheettables)
        });

        console.log("child inserted: ", newchild.data);

        // for (const file of childFiles) {
        //   await sp.web.lists
        //     .getByTitle("WorkFlowProcessData")
        //     .items.getById(newchild.data.Id)
        //     .attachmentFiles.add(file.name, file.file);
        // }

        const ChildFolderName = `NoteAttach/${seqNochild}`;
        await sp.web.folders.add(ChildFolderName);

        // Upload files to the subfolder if any
        for (const file of childFiles) {
          await sp.web.getFolderByServerRelativeUrl(ChildFolderName).files.add(file.name, file.file, true);
        }

      }

      if (!IsParentForm) {
        //   for (const file of files) {
        //   await sp.web.lists
        //     .getByTitle("WorkFlowProcessData")
        //     .items.getById(itemId)
        //     .attachmentFiles.add(file.name, file.file);
        // }

        const folderName = `NoteAttach/${seqNo}`;
        await sp.web.folders.add(folderName);

        // Upload files to the subfolder if any
        for (const file of files) {
          await sp.web.getFolderByServerRelativeUrl(folderName).files.add(file.name, file.file, true);
        }
      }
      else {
        const ParentFolderName = `NoteAttach/${seqNo}`;
        await sp.web.folders.add(ParentFolderName);

        // Upload files to the subfolder if any
        for (const file of parentFiles) {
          await sp.web.getFolderByServerRelativeUrl(ParentFolderName).files.add(file.name, file.file, true);
        }
      }
      // ------------------------------------------Saving table attachments to library-----------------------------------------------

      tables.forEach(async (table: any) => {
        const attachcolumnName = table.columns.find((column: any) => column.columntype === "Attachment")?.headername
        const boolvalue: boolean = table.rows.find((row: any) => row[attachcolumnName].name.length > 0)
        if (attachcolumnName && boolvalue) {
          const tablefoldername = `TableAttachments/${seqNo}`;
          await sp.web.folders.add(tablefoldername);
          const subfolderName = `${tablefoldername}/${table.id}`;
          await sp.web.folders.add(subfolderName);

          for (const file of table.rows) {
            //console.log("----file=====", file)
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

  useEffect(() => {
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
  }, []);

  const handleChange = (data: any) => {
    if (!data || !Array.isArray(data)) {
      console.error("handleChange received invalid data:", data);
      return;
    }

    funforVisibility(data)

    //console.log("Event change in form -----------", data);

    // // Preserve existing file values
    // const updatedData = data.map((item: any) => {
    //   if (item.name.startsWith("file_upload")) {
    //     return {
    //       ...item,
    //       value: fileData[item.name] || null, // Keep the file if it exists in state
    //     };
    //   }
    //   return item;
    // });

    // console.log("updatedData in FormComponent ------------", updatedData);
    // setFormData(updatedData);
    setFormData(data);
  };


  // Use `useEffect` to update `formData` whenever `fileData` changes
  useEffect(() => {
    setFormData((prevData: any) =>
      prevData.map((el: any) =>
        el.name.startsWith("file_upload") && fileData[el.name]
          ? { ...el, value: fileData[el.name] }
          : el
      )
    );
  }, [fileData]); // This ensures that file uploads are correctly reflected in `formData`
  // ---------- ending code for multi file attachment

  //  // Separate file input change handler
  // const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>, itemName: string) => {
  //   const target = event.target;
  //   if (target && target.files && target.files.length > 0) {
  //     // const file = target.files[0];
  //     const files = Array.from(target.files);
  //     console.log("File uploaded: ", files);

  //     // Update file state first
  //     setFileData((prev) => ({
  //       ...prev,
  //       [itemName]: files,
  //     }));
  //   }
  // };


  // useEffect(() => {

  //   const fileuploadcontrol = document.querySelectorAll('input[type="file"]');

  //   fileuploadcontrol.forEach((fileInput) => {
  //     const itemName = fileInput.getAttribute("name") || "";
  //     fileInput.addEventListener("change", (event) => handleFileChange(event as any, itemName));
  //   });

  //   return () => {
  //     fileuploadcontrol.forEach((fileInput) => {
  //       const itemName = fileInput.getAttribute("name") || "";
  //       fileInput.removeEventListener("change", (event) => handleFileChange(event as any, itemName));
  //     });
  //   };
  // }, [formData]);

  // NOTIFY TO FOR INTERNAL AND EXTERNAL USERS ----------------------------------------------------




  // const handleDropdownChange = async (selectedEmail: string) => {


  //   if (!selectedEmail) {
  //     SetnotifytorequestorID([]);
  //     return;
  //   }

  //   try {
  //     // Optional: ensure user exists in SharePoint
  //     await sp.web.ensureUser(selectedEmail);

  //     const user = await sp.web.siteUsers
  //       .filter(`Email eq '${selectedEmail}'`)
  //       .select('Id,Email')
  //       .get();

  //     const userId = user[0]?.Id;

  //     if (userId) {
  //       SetnotifytorequestorID([userId]);
  //     } else {
  //       console.warn("User ID not found for:", selectedEmail);
  //     }
  //   } catch (error) {
  //     console.error("Error in dropdown user selection:", error);
  //   }
  // };

  // const handleInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
  //   const value = e.target.value;
  //   setSearchText(value);

  //   if (value.trim() === '') {
  //     setFilteredUsers(myOrganizationUsers.slice(0, 10));
  //   } else {
  //     const filtered = myOrganizationUsers
  //       .filter((user) =>
  //         user.displayName.toLowerCase().includes(value.toLowerCase())
  //       )
  //       .slice(0, 10);
  //     setFilteredUsers(filtered);
  //   }

  //   setShowDropdown(true);
  // };


  // const handleUserSelect = async (user: User) => {
  //   const alreadySelected = selectedUsers.find(u => u.mail === user.mail);
  //   if (alreadySelected || selectedUsers.length >= 5) return;

  //   try {
  //     await sp.web.ensureUser(user.mail);
  //     const result = await sp.web.siteUsers
  //       .filter(`Email eq '${user.mail}'`)
  //       .select('Id,Email')
  //       .get();

  //     const userId = result[0]?.Id;
  //     if (userId) {
  //       setSelectedUsers(prev => [...prev, user]);
  //       SetnotifytorequestorID(prev => [...prev, userId]);
  //     }
  //   } catch (error) {
  //     console.error("Error selecting user:", error);
  //   }

  //   // Avoid triggering filter logic on blank input
  //   setSearchText('');
  //   setShowDropdown(false);

  //   // Set default filtered list again, just in case
  //   setFilteredUsers(myOrganizationUsers.slice(0, 10));
  // };

  const handleUserDomain = (UserDomain: string) => {
    setIsSameDomain(UserDomain)

  };




  // const removeSelectedUser = (userMail: string) => {
  //   const updatedUsers = selectedUsers.filter(u => u.mail !== userMail);
  //   setSelectedUsers(updatedUsers);

  //   SetnotifytorequestorID(prev =>
  //     prev.filter((_, idx) => selectedUsers[idx].mail !== userMail)
  //   );
  // };


  return (
    <div style={{ width: '100%' }}>
      {loading ? (
        <LoadingSpinner />
      ) : (
        <div >
          <Toast ref={toast} />
          <TopNavBar onUserDomainRetrieved={handleUserDomain} />

          <div style={{ display: 'flex' }}>
            <SideBar activeMenu={''} />
            <div className='FC_mainDiv' style={{ width: '100%' }}>
              <div className="FC_header-container">
                <div className="FC_title-section">
                  <div className="FC_back-icon">
                    <img onClick={() => navigate(-1)} src={require('../../assets/Images/previous.png')} alt="backicon" title='Home' />
                  </div>
                  <div>
                    <p className="FC_FormName">{selectedFormTitle}</p>
                  </div>
                </div>

                <div className="FC_requester-name">
                  <p style={{ margin: 0 }}>{RequesterName}</p>
                </div>
              </div>

              {IsParentForm && <ParentWithChildren
                parentName="Form Panel"
                children={ChildFormDetails.map((ele) => ({ id: ele.AppCode, name: ele.AppName, ChildFormJson: JSON.parse(ele.FormJSON), AppCode: ele.AppCode, ChildOrder: ele.ChildOrder }))}
                onChildClick={(child) => {
                  //console.log("Clicked:", child);
                  // setActiveChildFormJSON(child.ChildFormJson)
                  setChildselected(child.AppCode);
                  let completeformjson = [...JSON.parse(location.state.selectedformdetails?.FormJSON), ...child.ChildFormJson]
                  settabclicked(prev=>!prev);
                  //console.log("completeformjson", completeformjson)
                  setSelectedFormJSON(completeformjson);
                  // 👉 load form JSON or navigate to child form
                }}
              />}
              <div className='fo-mid1'>
                <div className='p-tabview-panel' style={{ paddingTop: 6 }}>

                  <div className='custom_field_main form-component-mode'>
                    <div className='form-edit-mode left-section' >
                      <label className='FC_StaticFields_label'> Request For <span style={{ color: 'red' }}> *</span></label>
                      <input type="text" className='FC_StaticFields_input' onChange={(e) => setCategorytype(e.target.value)} />
                    </div>



                    <div className='form-edit-mode right-section'>
                      <label className='FC_StaticFields_label'>Notify To</label>
                      <PeoplePicker
                        context={getPeoplePickerContext()}
                        personSelectionLimit={5}
                        required={false}
                        onChange={onChange}
                        showHiddenInUI={false}
                        principalTypes={[PrincipalType.User]}
                        resolveDelay={1000}
                        resultFilter={(results: any[]) =>
                          results.filter(persona => {
                            if (isSameDomain && String(isSameDomain) === "NO") {
                              let email = persona?.loginName || '';
                              //console.log("persona in if ", persona)
                              return email.includes(`${domainfield}`);
                            }
                            else if (isSameDomain) {
                              //console.log("persona in else ", persona)
                              let email = persona?.loginName || '';
                              return !email.includes('#ext#');
                            }
                          })}
                      />
                    </div>

                  </div>


                  <div className="form-component-mode">
                    {/* <ReactFormGenerator
                      data={selectedFormJSON}
                      form_action="/submit"
                      form_method="POST"
                      hide_actions={true}
                      skip_validations={false}
                      onChange={(e) => handleChange(e)}
                      onBlur={(e) => handleBlur(e)}
                    /> */}
                    <ReactFormGenerator
                      data={selectedFormJSON}
                      form_action="/submit"
                      form_method="POST"
                      hide_actions={true}
                      skip_validations={false}
                      onChange={(e) => handleChange(e)}
                      onBlur={(e) => handleBlur(e)}
                    />
                    {/* {
                                          ChildFormDetails.length > 0 &&
                                          <ReactFormGenerator
                                            data={activeChildFormJSON}
                                            form_action="/submit"
                                            form_method="POST"
                                            hide_actions={true}
                                            skip_validations={false}
                                            onChange={(e) => handleChange(e)}
                                            onBlur={(e) => handleBlur(e)}
                                          />
                    
                                        } */}

                  </div>
                </div>
                <div>



                  <div className='form-btn'>
                    <button
                      disabled={isButtonDisabled}
                      onClick={async () => {
                        if (formData.length > 0 || tables.length > 0 || timesheettables.length > 0) {
                          await handleSubmit(formData);
                        } else {
                          // setError("Please fill out the form before submitting.");
                          toast.current?.show({ severity: 'error', summary: '', detail: 'Please fill out the form before submitting.', life: 2000 });

                        }
                      }}
                      className={toasterrors.length === 0 ? 'Sbmt-btn all-btn' : 'btn btn-default'}
                    // className={`Sbmt-btn all-btn ${toasterrors.length === 0 ? "" : "disabled"}`}
                    >
                      Submit
                    </button>

                    <button
                      onClick={async () => {
                        if (formData.length > 0 || tables.length > 0 || timesheettables.length > 0) {
                          await handleSave(formData);
                        } else {
                          // setError("Please fill out the form before saving.");
                          toast.current?.show({ severity: 'error', summary: '', detail: 'Please fill out the form before saving.', life: 2000 });
                        }
                      }}
                      className='all-btn'
                    >
                      Save
                    </button>
                    <button className="all-btn " onClick={() => { navigate(-1) }}>Cancel</button>
                  </div>
                </div>
              </div>
            </div>
          </div>
        </div>)}
    </div>
  );
};

export default FormComponent;