import React, { useEffect, useState } from "react";
import { sp } from "@pnp/sp/presets/all";
import { mySiteUrl } from "../ConfigURL/All_URLs";
import { PeoplePicker, PrincipalType, IPeoplePickerContext } from "@pnp/spfx-controls-react/lib/PeoplePicker";
// import { BaseWebPartContext } from '@microsoft/sp-webpart-base';
// import * as pdfjsLib from 'pdfjs-dist';
// import worker from 'pdfjs-dist/build/pdf.worker.entry.js';
import { fileFetchUrl } from '../ConfigURL/All_URLs';
import { fetchTenantUser } from "../FetchTenantUser/fetchTenantUser";
import { IDynamicFormBuilderProps } from '../../IDynamicFormBuilderProps';
import MyZoneCommentBox from "./MyZoneCommentBox";
import LoadingSpinner from '../Loading';
import { DataTable } from "primereact/datatable";
import Attachicon from '../../../assets/Images/attachicon.png';
import searchIcon from '../../../assets/Images/searchicon.svg';
import OrganiserIcon from '../../../assets/Images/OrganiserIcon.svg';
import DateIcon from '../../../assets/Images/DateIcon.svg';
import TimeIcon from '../../../assets/Images/TimeIcon.svg';
import LocationIcon from '../../../assets/Images/LocationIcon.svg';
import AttendeesIcon from '../../../assets/Images/AttendeesIcon.svg';
import commentsicon from '../../../assets/Images/CommentsIcons.png';
import { Button } from "primereact/button";
import { Calendar } from "primereact/calendar";
import Sidebar from '../Sidebar/SideBar'
import TopNavBar from "../TopBar";
import './myzone.css';
import { Column } from "primereact/column";


// pdfjsLib.GlobalWorkerOptions.workerSrc = worker;
// interface myzoneprops{
//     context:BaseWebPartContext;
// }
interface tabledata {
    Id: Number;
    TaskId: string;
    Description: string;
    status: string;
    createdon: string;
    createdby: string;
    AssignedTo: string;
    AssignedToId: Number;
    AssignedToEMail: string;
    DueDate: string;
}

const MyZone: React.FC<IDynamicFormBuilderProps> = ({ context }) => {

    const [Togglestr, setTogglestr] = useState<string>("To Do");
    const [filtereddata, setFiltereddata] = useState<tabledata[]>();
    const [originaldata, setoriginaldata] = useState<tabledata[]>();
    const [originalmeetingitems, setoriginalmeetingitems] = useState<any>();
    const [meetingitems, setmeetingitems] = useState<any>();
    const [statusitems, setstatusitems] = useState<string[]>();
    const [DomainName, setDomainName] = useState<string>("");
    const [showpopup, setshowpopup] = useState<boolean>(false);
    const [isediting, setisediting] = useState<boolean>();
    const [formdata, setformdata] = useState<any>();
    const [TaskId, setTaskId] = useState<string>();
    const [activerowId, setactiverowId] = useState<Number | null>(null);
    const [currentuser, setcurrentuser] = useState<string>("");
    const [currentusertype, setcurrentusertype] = useState<string>();
    const [popupTaskId, setpopupTaskId] = useState<string>("");
    const [popupCoords, setPopupCoords] = useState({ top: 0, left: 0 });
    const [selectedmeeting, setselectedmeeting] = useState<any>();
    const [meetingpopup, setmeetingpopup] = useState<boolean>(false);
    const [isloading, setisloading] = useState<boolean>(true);
    const [isSameDomain, setIsSameDomain] = useState<string>("");
    const [externalGroups, setExternalGroups] = useState<string[]>([]);
    // const [latestcomments,setlatestcomments] = useState<Record<string,string>>({});

    sp.setup({
        sp: {
            baseUrl: mySiteUrl,
        },
    });
    function getPeoplePickerContext(): IPeoplePickerContext {
        return {
            absoluteUrl: context.pageContext.web.absoluteUrl,
            msGraphClientFactory: context.msGraphClientFactory,
            spHttpClient: context.spHttpClient,
        };
    }

    const filtergroups = async (data: any) => {
        const currentUser = await sp.web.currentUser.get();
        const currentuserId = currentUser.Id;

        const sitegroups = await sp.web.siteGroups.get();
        const sitegroupIds = new Set(sitegroups.map((group) => group.Id));

        const usergroups = await sp.web.currentUser.groups();
        const usergroupIds = new Set(usergroups.map(usergrp => usergrp.Id))

        let filtereditems: any[] = [];

        for (const item of data) {

            if (item.AssignedTo) {
                if (item.AssignedTo?.Id === currentuserId) {
                    filtereditems.push(item);
                }

                if (sitegroupIds.has(item.AssignedTo?.Id) && usergroupIds.has(item.AssignedTo?.Id)) {
                    filtereditems.push(item);
                }

                if (item.AssignedTo?.Title === "All Employees") {
                    filtereditems.push(item);
                }
            }
        }

        return filtereditems;
    }
    const fetchprocessitems = async () => {
        try {
            setisloading(true);
            const Domain = await fetchTenantUser();
            setDomainName(Domain.TenantUsers);
            const allitems = await sp.web.lists.getByTitle("MyZoneProcessList").items.select("ID", "TaskID", "Description", "Status", "Author/Id", "AssignedTo/Id", "Author/Title", "AssignedTo/Title", "AssignedTo/EMail", "DueDate", "Created", "Unassigned").expand("Author", "AssignedTo").filter(`Domain eq '${Domain.TenantUsers}'`).getAll();
            //console.log("allitemsinprocess", allitems);
            const gentaskid = allitems.reduce((highest, item) => {
                if (item.Id > highest) {
                    return item.Id;
                }
                return highest
            }, 0);
            const today = new Date();
            const formattedDate = today.toLocaleDateString('en-US', {
                month: 'short',
                year: '2-digit',
            }).replace(' ', '');
            const resultid = `${formattedDate}-${gentaskid}`;
            // const unassigned= allitems.filter(item => item.Unassigned === true);
            // let fetcheditems
            //console.log("currentusertype", currentusertype)
            //     if(currentusertype != "OnlyUser"){
            //      fetcheditems= allitems;
            //      console.log("adminuser",fetcheditems)
            //     }
            //     else{
            //        fetcheditems = allitems.map((item:any)=>{
            //         if(item.AssignedTo){
            //             return item
            //         }
            //         else{
            //             return {
            //                 ...item,
            //                 AssignedTo: {Id:0,Title:'',EMail:''}
            //             }
            //         }

            //     })
            //     console.log("unassigned",fetcheditems)

            // }

            setTaskId(resultid)
            // setTaskId(String(gentaskid));
            //console.log("highid", typeof gentaskid, typeof TaskId)
            const currentUser = await sp.web.currentUser.get();
            //console.log("currrr", currentUser)
            // setcurrentuser(currentUser.Title);
            let filtered = allitems;
            if (Togglestr === "To Do") {
                // if(currentusertype != "OnlyUser"){
                //     filtered=fetcheditems;
                // }
                // else{
                //    filtered= fetcheditems?.filter((item ) => item.Author.Id === currentUser.Id);
                // }
                filtered = allitems?.filter((item) => item.AssignedTo && item.Author.Id === currentUser.Id);

            }
            if (Togglestr === "My Actions") {
                //   filtered= fetcheditems?.filter((item ) => item.AssignedTo?.Id === currentUser.Id);
                const filterassigned = await filtergroups(allitems);
                //console.log("filterassigned", filterassigned)
                filtered = filterassigned;
            }
            if (Togglestr === "UnAssigned") {
                filtered = allitems.filter((item) => !item.AssignedTo)
                //console.log("uuuuuuuunassigned", filtered)
            }
            const cleaneddata: tabledata[] = filtered.map((item) => ({
                Id: item.Id,
                TaskId: item.TaskID,
                Description: item.Description,
                status: item.Status,
                createdon: item.Created,
                createdby: item.Author?.Title ?? ' ',  // make each record into string: number | string pairs to make sure there are no string: {id: 2} pairs
                AssignedTo: item.AssignedTo?.Title ?? ' ',
                AssignedToId: item.AssignedTo?.Id,
                AssignedToEMail: item.AssignedTo?.EMail,
                DueDate: item.DueDate,
            }));
            // cleaneddata.sort((a,b)=>{
            //   return a.TaskId < b.TaskId ? 1 : -1 ;
            // })
            cleaneddata.sort((a, b) => new Date(b.createdon).getTime() - new Date(a.createdon).getTime());
            setoriginaldata(cleaneddata);
            setFiltereddata(cleaneddata);

        }
        catch (error) {
            setisloading(false);
            //console.log("error fetching process items", error)
        }
        finally {
            setisloading(false);
        }
    }

    const fetchmeetings = async () => {
        try {
            setisloading(true);
            const fetchedmeetingitems = await sp.web.lists.getByTitle("MyZoneMeeting").items.select("MeetingID", "Attendees", "Time", "MeetingDate", "Location", "MeetingTitle", "Organiser", "Created").getAll();
            //console.log("meetingitems", fetchedmeetingitems)
            let meetingitems;
            if (currentusertype === "BothUser Admin" || currentusertype === "OnlyAdmin") {
                meetingitems = fetchedmeetingitems;
            }
            else if (currentusertype === "OnlyUser") {
                const username = currentuser?.split(" | ")[0];
                meetingitems = fetchedmeetingitems.filter(item =>
                    item.Attendees?.toLowerCase().includes(username.toLowerCase()));
            }

            const MOMfiles = await sp.web.lists
                .getByTitle("MyMeetings")
                .items.select("FileLeafRef", "FileRef", "FileDirRef")
                .filter(`FileDirRef eq '${fileFetchUrl}/MyMeetings/MOM'`)
                .get();
            //console.log("momfiles", MOMfiles)
            const MOMfileMap = new Map();
            MOMfiles.forEach(file => {
                const meetingIdFromFile = file.FileLeafRef?.split('MOM')[1]?.split('.')[0]; //Get "123" from "123.pdf" meetingid is filename
                MOMfileMap.set(meetingIdFromFile, {
                    name: file?.FileLeafRef,
                    url: file?.FileRef,
                });
            });



            const Actionitemsfiles = await sp.web.lists
                .getByTitle("MyMeetings")
                .items.select("FileLeafRef", "FileRef", "FileDirRef")
                .filter(`FileDirRef eq '${fileFetchUrl}/MyMeetings/ActionItems'`)
                .get();
            const actionmap = new Map();

            Actionitemsfiles.forEach((file) => {
                const meetingid = file.FileLeafRef?.split('AI')[1]?.split('.')[0]
                actionmap.set(meetingid, {
                    name: file?.FileLeafRef,
                    url: file?.FileRef,
                })
            })
            //console.log("actionmap", actionmap)
            // Attach file info to each meeting item
            const updatedmeetingitems = meetingitems?.map(meeting => {
                const fileInfo = MOMfileMap.get(String(meeting.MeetingID));
                const actionInfo = actionmap.get(String(meeting.MeetingID));
                //console.log("actionInfo", actionInfo)
                return {
                    ...meeting,
                    MOMfileName: fileInfo?.name || null,
                    MOMfileUrl: fileInfo?.url || null,
                    ActionfileName: actionInfo?.name || null,
                    ActionfileUrl: actionInfo?.url || null,
                };
            });
            //console.log("meetingitems", updatedmeetingitems);
            updatedmeetingitems?.sort((a, b) => new Date(b.Created).getTime() - new Date(a.Created).getTime());
            setoriginalmeetingitems(updatedmeetingitems);
            setmeetingitems(updatedmeetingitems);

        }
        catch (err) {
            setisloading(false);
            console.log("error in fetchmeetings", err)
        }
        finally {
            setisloading(false)
        }
    }

    const fetchstatuses = async () => {
        try {
            const statuses = await sp.web.lists.getByTitle("MyZoneStatus").items.select("Status").getAll();
            const statusstrings = statuses?.map((item: any) => {
                return item.Status;
            });
            setstatusitems(statusstrings);
            //console.log("status", statusstrings)
        }
        catch (err) {
            console.log("error fetching status master", err)
        }
    }

    const fetchcurrentuser = async () => {

        const currentUser = await sp.web.currentUser.get();
        //console.log("currrr", currentUser)
        setcurrentuser(currentUser.Title);
    }

    useEffect(() => {
        if (Togglestr === "My Meetings") {
            fetchmeetings().catch((error) => {
                console.log("error in fetching meeting", error)
            })
        }
        else {
            fetchprocessitems().catch((error) => {
                console.log("error in fetching process", error)
            })
        }

        // handletabchange(Togglestr).catch(Error)
    }, [currentuser, currentusertype, Togglestr])

    useEffect(() => {
        fetchstatuses().catch((error) => {
            console.log("error in fetching statuses", error)
        })
        fetchcurrentuser().catch((error) => {
            console.log("error fetching currentuser", error);
        })
    }, [])

    useEffect(() => {
        //console.log("form data", formdata)
    }, [formdata])

    useEffect(() => {
        //console.log("filtered data", filtereddata)
    }, [filtereddata])

    // const latestcommentmap = (TaskId:string, latestcomment:string) => {
    //     setlatestcomments((prev) =>({
    //         ...prev,
    //         [TaskId]: latestcomment,
    //     }))

    // }

    const handlechange = (columnname: string, val: string | Number | null) => {
        setformdata((prev: any) => ({
            ...prev,
            [columnname]: val
        })
        )
    }
    const handlepeoplepickerchange = async (items: any[]) => {
        //console.log("items", items);
        try {
            const users = await Promise.all(items.map(async (useritem) => {
                // const shrptuser= await sp.web.siteUsers.filter(`Email eq '${useritem.secondaryText}'`).select("Id,Email").get();
                // return shrptuser[0]?.Id;
                return Number(useritem.id)
            }));
            // const selecteditem = items?[0];
            // const assignedtologinname = selecteditem?.loginname;

            //console.log("shrptuserids", users);
            const validUserIds = users.filter(id => id !== undefined);
            //console.log("Selected valid User IDs:", validUserIds);
            handlechange("AssignedTo", validUserIds[0]);
            handlechange("AssignedToId", validUserIds[0]);
            // handlechange("AssignedToLoginName",assignedtologinname);
        }
        catch (err) {
            console.log(err)
        }
    }
    const handlecreate = async () => {
        //console.log("insidesaveeeee")
        // formdata.forEach((item:any)=>{
        //     if(item.value){
        //         return
        //     }
        // })
        if (!formdata.Description || !formdata.AssignedTo || !formdata.DueDate) return;
        try {
            await sp.web.lists.getByTitle("MyZoneProcessList").items.add({
                Description: formdata.Description,
                // AssignedToId: {
                //     results : formdata.AssignedTo[0]
                // }
                AssignedToId: formdata.AssignedTo,
                DueDate: formdata.DueDate,
                TaskID: TaskId,
                Domain: DomainName,
                Status: "To Do",

            })
            //console.log("data saved", formdata)
            setshowpopup(false);

        }
        catch (err) {
            console.log("error saving", err);
        }
    }
    const handleupdate = async () => {
        const itemId = formdata.Id;
        try {
            await sp.web.lists.getByTitle("MyZoneProcessList").items.getById(Number(itemId)).update({
                Description: formdata.Description,
                AssignedToId: formdata.AssignedToId,
                DueDate: formdata.DueDate,
            })
            //console.log("data updated", formdata)
            setshowpopup(false);

        }
        catch (err) {
            console.log("error updating", err);
        }
    }
    const handlecommentstoggle = (rowId: any) => {
        //console.log("comment", rowId)
        //console.log("activerowidddd", activerowId)
        setactiverowId(prev => (prev === rowId ? null : rowId)
        )
    }

    useEffect(() => {
        const fetchExternalGroups = async () => {
            let extgrps: any[] = []
            //console.log("domainandissame", DomainName, isSameDomain)
            if (isSameDomain && isSameDomain === "YES") {
                //console.log("DomainNameinIFFF", DomainName)
                extgrps = await sp.web.lists
                    .getByTitle("ExternalGroupsList")
                    .items.select("Domain,GroupName/Title").expand("GroupName").filter(`Domain ne '${DomainName}'`).get()
            }
            else if (isSameDomain) {
                //console.log("INELSE", DomainName)
                extgrps = await sp.web.lists
                    .getByTitle("ExternalGroupsList")
                    .items.select("Domain,GroupName/Title").expand("GroupName").filter(`Domain eq '${DomainName}'`).get()
            }
            const titles: string[] = extgrps.map(item => item.GroupName?.Title);
            setExternalGroups(titles);
            //console.log("Ext", externalGroups)
            //console.log("external groups", titles)
        };
        void fetchExternalGroups();


    }, [DomainName, isSameDomain]);






    useEffect(() => {
        //console.log("active rowid", activerowId)
    }, [activerowId])

    //     const handlepdfupload = async (file: File)=>{
    //         const filearraybuffer= await file.arrayBuffer();
    //         // const pdfdoc= await pdf(filearraybuffer);

    //         // console.log("pdftext",pdfdoc.text)

    //         // for (const page of pages){
    //         //     const txtcontent= await page.getTextContent()
    //         // }
    //         const typedarray = new Uint8Array(filearraybuffer);

    //         const pdfdoc = await pdfjsLib.getDocument({data: typedarray}).promise
    //          let text:string="";
    //         for (let i=1; i<= pdfdoc.numPages; i++){
    //             const page = await pdfdoc.getPage(i);
    //             const textcontent = await page.getTextContent();
    //             const textitems = textcontent.items.map((item:any)=> item.str).join(' ');
    //             text+=textitems + '\n'
    //             // console.log("textitems",textitems)

    //         }
    //         console.log("pdftext",text)
    //         console.log("arraybuffer",filearraybuffer)
    //   }

    const handledropdownchange = async (Id: Number, val: string) => {
        if (Id && val) {

            try {
                await sp.web.lists.getByTitle("MyZoneProcessList").items.getById(Number(Id)).update({
                    Status: val,
                });
                setFiltereddata(prev => prev?.map((item) => {
                    return item.Id === Id ? { ...item, status: val } : item;
                }))  // after patching to sharepoint update it in frontend dropdown also
            }
            catch (err) {
                console.log("error in updating status", err)
            }

        }
    }

    const handleSearch = (searchTerm: string) => {
        debugger
        if (!searchTerm) {
            setFiltereddata(originaldata); // Reset to full data
            setmeetingitems(originalmeetingitems);
            return;
        }

        const filtered = filtereddata?.filter((item) =>

            item.status?.toLowerCase().includes(searchTerm.toLowerCase()) ||
            item.AssignedTo?.toLowerCase().includes(searchTerm.toLowerCase()) ||
            item.TaskId?.toLowerCase().includes(searchTerm.toLowerCase())
        );
        setFiltereddata(filtered);

        const filteredmeetings = meetingitems?.filter((item: any) =>

            item.MeetingID?.toLowerCase().includes(searchTerm.toLowerCase()) ||
            item.Attendees?.toLowerCase().includes(searchTerm.toLowerCase())
        )
        setmeetingitems(filteredmeetings)
    };

    const handleusertype = (UserAuth: string) => {
        //console.log("user retrieved", UserAuth)
        setcurrentusertype(UserAuth);
    }

    const handleUserDomain = (UserDomain: string) => {
        setIsSameDomain(UserDomain)

    };


    return (
        <div className="myzone" style={{ display: 'flex', flexDirection: 'column' }}>
            <TopNavBar onUserTypeRetrieved={handleusertype} onUserDomainRetrieved={handleUserDomain} />
            <div style={{ display: 'flex', backgroundColor: '#070D19' }}>
                <Sidebar activeMenu="My Zone" />
                <div className="myzonediv">
                    <div style={{ margin: '5px 5px 5px 5px' }}>
                        <section>
                            <div style={{ display: 'flex', alignItems: 'center', gap: '7px', margin: '20px 0px 0px 0px' }}>
                                <Button
                                    className={Togglestr === "To Do" ? 'myzonetglenabled tglenabled1' : 'myzonetgldisabled tgldisabled1'}
                                    onClick={() => { setTogglestr("To Do"); }}
                                >
                                    To Do
                                </Button>
                                <Button
                                    className={Togglestr === "My Actions" ? 'myzonetglenabled' : 'myzonetgldisabled'}
                                    onClick={() => { setTogglestr("My Actions"); }}
                                >
                                    My Actions
                                </Button>

                                <Button
                                    className={Togglestr === "My Meetings" ? 'myzonetglenabled' : 'myzonetgldisabled'}
                                    onClick={() => { setTogglestr("My Meetings"); }}
                                >
                                    My Meetings
                                </Button>
                                {currentusertype !== 'OnlyUser' && (
                                    <Button
                                        className={
                                            Togglestr === 'UnAssigned' ? 'myzonetglenabled' : 'myzonetgldisabled'
                                        }
                                        onClick={() => {
                                            setTogglestr('UnAssigned');
                                        }}
                                    >
                                        Un Assigned
                                    </Button>
                                )}
                                <div style={{ display: 'flex', marginLeft: 'auto' }}>
                                    {Togglestr === "To Do" ? (
                                        <Button
                                            className="addnew"
                                            onClick={() => {
                                                setisediting(false);
                                                setformdata({})
                                                setshowpopup(true)
                                            }}
                                        >Create Task</Button>) : ""}
                                    <div className="myzonesearchContainer">

                                        <input placeholder='Search...'
                                            className="myzonesearchBox"
                                            onChange={(e: React.ChangeEvent<HTMLInputElement>) => handleSearch(e.target.value)}
                                        />
                                        <img src={searchIcon} alt="Search Icon" className="myzonesearchIcon" />
                                    </div>
                                </div>
                            </div>
                        </section>
                        <div className='tabline'></div>
                        {Togglestr != "My Meetings" ?
                            <div className="tablediv">
                                {/* <table className="myzonetable">
                        <thead>
                            <tr>                              
                               <th>Task Id</th>
                               <th>Description</th>
                               <th>Status</th>
                               <th>Created On</th>
                              {Togglestr ==="My Actions"?<th>Assigned By</th>:<></>}
                               <th>Assigned To</th>
                               <th>Due Date</th>
                               <th></th>
 
                            </tr>
                        </thead>
                        <tbody>
                            {filtereddata?.map((item:tabledata,index)=>(
                                 <tr key={index}
                                 style={{color:'#B1B1B1',fontSize:'12px'}}
                                 onClick={()=>{
                                    if(Togglestr === "My Actions") return
                                    console.log("clickeditem",filtereddata[index])
                                    const selecteditem=filtereddata[index];
                                    setformdata((prev:any)=>({
                                        ...prev,
                                        Description: selecteditem.Description,
                                        AssignedTo: selecteditem.AssignedTo,
                                        AssignedToId: selecteditem.AssignedToId,
                                        AssignedToEMail: selecteditem.AssignedToEMail,
                                        DueDate: selecteditem.DueDate,
                                        Id: selecteditem.Id,
                                    }))
                                    setisediting(true);
                                    setshowpopup(true);
                                }}
                                 >
                                    <td>{item.TaskId}</td>
                                    <td>{item.Description}</td>
                                    <td
                                    onClick={(e)=>e.stopPropagation()}
                                    style={{minWidth:'95px'}}
                                    >{
                                        <select
                                        className="statusdropdown"
                                        onChange={async (e)=>{
                                            console.log("dropdownchanged",e.target.value)
                                            await handledropdownchange(item.Id,e.target.value)
 
                                        }}
                                        value={item.status }
                                        >
                                        {
                                            statusitems?.map((statusitem,index)=>{
                                                return <option key={index} value={statusitem}>{statusitem}</option>
                                            })
                                        }
                                        </select>
                                        }</td>
                                    <td>{item.createdon?.split("T")[0]}</td>
                                   { Togglestr==="My Actions"?<td>{item.createdby}</td>:<></>}
                                    <td>{item.AssignedTo}</td>
                                    <td>{item.DueDate?.split("T")[0]}</td>
                                    <td onClick={(e)=>e.stopPropagation()}>{<>                                      
                                        <MyZoneCommentBox isopen={item.Id === activerowId} ontoggle={()=>handlecommentstoggle(item.Id)} TaskId={item.TaskId} user={currentuser} 
                                        />
                                        </>}
                                    </td>
                                 </tr>
                            ))}
                       
                        </tbody>
                     </table> */}
                                {isloading ? (
                                    <LoadingSpinner />
                                ) : (<DataTable className="hover-table"

                                    value={filtereddata}
                                    responsiveLayout="scroll"
                                    paginator
                                    scrollable
                                    scrollHeight="500px"
                                    sortField='Modified'
                                    sortOrder={-1}
                                    paginatorTemplate="CurrentPageReport FirstPageLink PrevPageLink PageLinks NextPageLink LastPageLink RowsPerPageDropdown"
                                    currentPageReportTemplate="Showing {first} to {last} of {totalRecords}"
                                    rows={10}
                                    font-size='16px'
                                    rowsPerPageOptions={[10, 15, 20]}
                                    dataKey="ID"
                                    emptyMessage="No transactions found."
                                    onRowClick={(e) => {
                                        // if(Togglestr === "My Actions") return
                                        //console.log("clickeditem", e)
                                        const selecteditem = e.data;
                                        setformdata((prev: any) => ({
                                            ...prev,
                                            Description: selecteditem.Description,
                                            AssignedTo: selecteditem.AssignedTo,
                                            AssignedToId: selecteditem.AssignedToId,
                                            AssignedToEMail: selecteditem.AssignedToEMail,
                                            DueDate: selecteditem.DueDate,
                                            TaskId: selecteditem.TaskId,
                                            Id: selecteditem.Id,
                                        }))
                                        // if(Togglestr=== "My Actions"){
                                        //     setisediting(false)
                                        // }
                                        // else{
                                        //     setisediting(true);
                                        // }
                                        setisediting(true);
                                        setshowpopup(true);
                                    }}


                                >
                                    <Column field="TaskId" header="Task Id" sortable />
                                    <Column field="Description" header="Description" sortable className="datatablerow" />
                                    <Column field='Status' header="Status"
                                        style={{ minWidth: '95px' }}
                                        sortable
                                        body={(rowData) => (
                                            <select
                                                className="statusdropdown"
                                                onChange={async (e) => {
                                                    //console.log("dropdownchanged", e.target.value)
                                                    await handledropdownchange(rowData.Id, e.target.value)

                                                }}
                                                onClick={(e) => e.stopPropagation()}
                                                value={rowData.status}
                                            >
                                                {
                                                    statusitems?.map((statusitem, index) => {
                                                        return <option key={index} value={statusitem}>{statusitem}</option>
                                                    })
                                                }
                                            </select>
                                        )}
                                    />
                                    <Column field="createdon" header="Created On" sortable className="datatablerow"
                                        body={(rowData) => (
                                            rowData.createdon?.split("T")[0]
                                        )}
                                    />
                                    {Togglestr === "My Actions" ?
                                        <Column
                                            field="createdby"
                                            header="Assigned By"
                                            className="datatablerow"
                                            sortable
                                        /> : <></>}
                                    <Column
                                        field="AssignedTo"
                                        header="Assigned To"
                                        className="datatablerow"
                                        sortable
                                    />

                                    <Column
                                        field="DueDate"
                                        header="Due Date"
                                        className="datatablerow"
                                        sortable
                                        body={(rowData) => {
                                            return rowData.DueDate?.split("T")[0]
                                        }} />
                                    <Column
                                        field=""
                                        header="Comments"
                                        body={(rowData) => (
                                            <img
                                                src={commentsicon}
                                                className="commentsicon"
                                                onClick={(e) => {
                                                    e.stopPropagation();
                                                    const rect = (e.target as HTMLElement).getBoundingClientRect();
                                                    setPopupCoords({
                                                        top: rect.top + rect.height,
                                                        left: rect.left + rect.width - 350,
                                                    }); // track popup position
                                                    handlecommentstoggle(rowData.Id); // sets activeRowId
                                                    setpopupTaskId(rowData.TaskId); // store current taskId if needed
                                                }}
                                            />
                                        )}
                                    />
                                    {/* <Column
    field=""
    header="Comments"
    sortable
    body={(rowData) => {
        return <div
        onClick={(e) => {
            e.stopPropagation(); // Prevent row click from triggering
            handlecommentstoggle(rowData.Id);
        }}>                                      
            <MyZoneCommentBox isopen={rowData.Id === activerowId} ontoggle={()=>handlecommentstoggle(rowData.Id)} TaskId={rowData.TaskId} user={currentuser} 
            />
        </div>
    }} 
    /> */}
                                </DataTable>)}
                                {activerowId && (

                                    <div
                                        // style={{ top: popupCoords.top, left: popupCoords.left, position: 'absolute', zIndex: 999 }}
                                        onClick={(e) => e.stopPropagation()}
                                    >
                                        <MyZoneCommentBox
                                            isopen={true}
                                            top={`${popupCoords.top}px`}
                                            left={`${popupCoords.left}px`}
                                            ontoggle={() => handlecommentstoggle(null)}
                                            TaskId={popupTaskId}
                                            user={currentuser}
                                        />
                                    </div>
                                )}

                            </div> :
                            <div className="tablediv">

                                {/* <table className="myzonetable">
                     <thead>
                        <th>Meeting ID</th>
                        <th>Meeting Title</th>
                        <th>Organiser</th>
                        <th>Meeting Date</th>
                        <th>Time</th>
                        <th>Location</th>
                        <th>Attendees</th>                        
                        <th>MOM</th>
                        <th>Action Items</th>                        
                     </thead>
                     <tbody>
                       {
                        meetingitems?.map((item:any)=>{
                           return(
                            <tr
                            style={{color:'#B1B1B1',fontSize:'12px'}}
                            >
                                <td>{item.MeetingID}</td>
                                <td>{item.MeetingTitle}</td>
                                <td>{item.Organiser}</td>
                                <td>{item.MeetingDate}</td>
                                <td>{item.Time}</td>
                                <td>{item.Location}</td>
                                <td>{item.Attendees}</td>                                
                                <td>{                            
                                    item.MOMfileUrl ? 
                                    (
                                     
                                    <img 
                                        src={Attachicon}
                                        className="downloadicon"
                                        onClick={() => window.open(item.MOMfileUrl, '_blank')}
                                       
                                    />
                                   
                                    
                                ) :
                                    (<span>N/A</span>)
                                
                                }</td>  
                                <td>
                                {item.ActionfileUrl ? 
                                    (
                                    
                                    <img 
                                    src={Attachicon}
                                    className="downloadicon"
                                    onClick={() => window.open(item.ActionfileUrl, '_blank')}

                                   
                                     />
                                    ) :
                                    (<span>N/A</span>)
                               }
                                </td>                             
                            </tr>
                           )
                        })
                       }
                     </tbody>
                    </table>  */}
                                {isloading ? (
                                    <LoadingSpinner />
                                ) : (<DataTable className="hover-table"

                                    value={meetingitems}
                                    responsiveLayout="scroll"
                                    paginator
                                    scrollable
                                    scrollHeight="500px"
                                    sortField='Modified'
                                    sortOrder={-1}
                                    paginatorTemplate="CurrentPageReport FirstPageLink PrevPageLink PageLinks NextPageLink LastPageLink RowsPerPageDropdown"
                                    currentPageReportTemplate="Showing {first} to {last} of {totalRecords}"
                                    rows={10}
                                    font-size='16px'
                                    rowsPerPageOptions={[10, 15, 20]}
                                    dataKey="ID"
                                    emptyMessage="No transactions found."
                                    onRowClick={(e) => {
                                        setselectedmeeting(e.data);
                                        setmeetingpopup(true);
                                    }}
                                // rowClassName={(rowData, options) => {
                                //     const index = transactions.findIndex(item => item.ID === rowData.ID); // Find the row index
                                //     return index % 2 === 0 ? 'Table_tr_even' : 'Table_tr_odd'; // Apply CSS class
                                // }}

                                >
                                    <Column field="MeetingID" header="Meeting ID" sortable />
                                    <Column field="MeetingTitle" header="Meeting Title" sortable
                                        className="datatablerow"
                                        body={(rowData) => (
                                            <span className="datatablerow" title={rowData.MeetingTitle}>
                                                {rowData.MeetingTitle}
                                            </span>
                                        )}
                                    />
                                    <Column field='Organiser' header="Organiser" sortable
                                        className="datatablerow"
                                        body={(rowData) => (
                                            <span className="datatablerow" title={rowData.Organiser}>
                                                {rowData.Organiser}
                                            </span>
                                        )}
                                    />
                                    <Column field="MeetingDate" header="Meeting Date" sortable
                                        className="datatablerow"
                                        body={(rowData) => (
                                            <span className="datatablerow" title={rowData.MeetingDate}>
                                                {rowData.MeetingDate}
                                            </span>
                                        )}
                                    />
                                    <Column
                                        field="Time"
                                        header="Time"
                                        sortable
                                        className="datatablerow"
                                        body={(rowData) => (
                                            <span className="datatablerow" title={rowData.Time}>
                                                {rowData.Time}
                                            </span>
                                        )}
                                    />
                                    <Column
                                        field="Location"
                                        header="Location"
                                        sortable
                                        className="datatablerow"
                                        body={(rowData) => (
                                            <span className="datatablerow" title={rowData.Location}>
                                                {rowData.Location}
                                            </span>
                                        )}
                                    />

                                    <Column
                                        field="Attendees"
                                        header="Attendees"
                                        sortable
                                        className="datatablerow"
                                        body={(rowData) => (
                                            <span className="datatablerow" title={rowData.Attendees}>
                                                {rowData.Attendees}
                                            </span>
                                        )}
                                    />
                                    <Column
                                        field="MOMfileUrl"
                                        header="MOM"
                                        sortable
                                        body={(rowData) => {
                                            return rowData.MOMfileUrl ?
                                                (
                                                    <img
                                                        src={Attachicon}
                                                        className="downloadicon"
                                                        onClick={(e) => {
                                                            e.stopPropagation();
                                                            window.open(rowData.MOMfileUrl, '_blank')
                                                        }}
                                                    />
                                                ) :
                                                (<span>N/A</span>)
                                        }
                                        }
                                    />
                                    <Column
                                        field="ActionfileUrl"
                                        header="Action Items"
                                        sortable
                                        body={(rowData) => {
                                            return rowData.MOMfileUrl ?
                                                (
                                                    <img
                                                        src={Attachicon}
                                                        className="downloadicon"
                                                        onClick={(e) => {
                                                            e.stopPropagation();
                                                            window.open(rowData.ActionfileUrl, '_blank')
                                                        }}
                                                    />
                                                ) :
                                                (<span>N/A</span>)
                                        }
                                        }
                                    />
                                </DataTable>)}
                            </div>

                        }
                    </div>

                </div>
                {
                    meetingpopup && (
                        <div className="custom-dark-modal-overlay">
                            <div className="custom-dark-meeting-modal">
                                {/* <h2 className="newtasktitle">{selectedmeeting.MeetingID}</h2> */}
                                <div style={{ display: 'flex', flexDirection: 'column', gap: '5px', padding: '10px 10px 0px 15px' }}>

                                    <div style={{ display: 'flex' }}>
                                        <div className="meetingtitle" title={selectedmeeting.MeetingTitle}>{selectedmeeting.MeetingTitle}</div>
                                        <div className="meetingidlabel">{selectedmeeting.MeetingID}</div>
                                    </div>
                                    <div className="meetingitemfield">
                                        <div><img style={{ height: '20px' }} src={OrganiserIcon} /> </div>
                                        <div className="meetingpopuplabels" title={selectedmeeting.Organiser}>{selectedmeeting.Organiser}</div>
                                    </div>
                                    <div className="meetingitemfield">
                                        <div><img style={{ height: '20px' }} src={DateIcon} /></div>
                                        <div className="meetingpopuplabels" title={selectedmeeting.MeetingDate}>{selectedmeeting.MeetingDate}</div>
                                    </div>
                                    <div className="meetingitemfield">
                                        <div><img style={{ height: '20px' }} src={TimeIcon} /> </div>
                                        <div className="meetingpopuplabels" title={selectedmeeting.Time}>{selectedmeeting.Time}</div>
                                    </div>
                                    <div className="meetingitemfield">
                                        <div><img style={{ height: '20px' }} src={LocationIcon} /></div>
                                        <div className="meetingpopuplabels" title={selectedmeeting.Location}>{selectedmeeting.Location}</div>
                                    </div>
                                    <div className="meetingitemfield">
                                        <div><img style={{ height: '20px' }} src={AttendeesIcon} /></div>
                                        <div className="meetingpopuplabels" title={selectedmeeting.Attendees}>{selectedmeeting.Attendees}</div>
                                    </div>
                                </div>
                                <div className="buttonsdiv">
                                    <Button
                                        className="cancelbtn"
                                        style={{ marginLeft: 'auto' }}
                                        onClick={() => {
                                            setselectedmeeting({});
                                            setmeetingpopup(false);
                                        }}
                                    >
                                        Close
                                    </Button>
                                </div>
                            </div>
                        </div>
                    )
                }
                {showpopup && (
                    <div className="custom-dark-modal-overlay">
                        <div className="custom-dark-modal">
                            <h2 className="newtasktitle">{
                                isediting ? formdata.TaskId : "Create Task"
                            }</h2>
                            <div className="inputcontainer">
                                <label className="controllabel">Task Description<span style={{ color: 'red', paddingLeft: '3px' }}>*</span></label>
                                <textarea
                                    className="descriptiontextarea"
                                    placeholder="Enter Task Description"
                                    value={formdata.Description}
                                    disabled={Togglestr === "My Actions"}
                                    onChange={(e) => {
                                        handlechange("Description", e.target.value)
                                    }}
                                    maxLength={250}
                                />
                            </div>
                            <div className="inputcontainer">
                                <label className="controllabel">Due Date<span style={{ color: 'red', paddingLeft: '3px' }}>*</span></label>
                                <Calendar
                                    placeholder="Select Date"
                                    showIcon={true}
                                    disabled={Togglestr === "My Actions"}
                                    value={formdata.DueDate ? new Date(formdata.DueDate) : null}
                                    onChange={(e) => {
                                        if (e.target.value) {
                                            const selecteddate = new Date(e.target.value)
                                            selecteddate.setHours(23, 59, 59, 999);
                                            handlechange("DueDate", selecteddate.toISOString())
                                            //console.log("datechanged", e, typeof e.target.value, typeof selecteddate.toISOString())
                                        }
                                    }}
                                    className="calcontrol-myzone"
                                />
                            </div>
                            <div className="inputcontainer">
                                <label className="controllabel">Assign To<span style={{ color: 'red', paddingLeft: '3px' }}>*</span></label>
                                {/* {console.log("assignedemail", formdata?.AssignedToEMail)} */}
                                {
                                    // isediting ?
                                    <PeoplePicker
                                        context={getPeoplePickerContext()}
                                        groupName={""}
                                        disabled={Togglestr === "My Actions"}
                                        principalTypes={[PrincipalType.User, PrincipalType.SharePointGroup]}
                                        personSelectionLimit={1}
                                        defaultSelectedUsers={[formdata?.AssignedTo]}
                                        placeholder="select a person"
                                        onChange={async (e) => {
                                            if (e.length === 0) {
                                                //console.log("peoplepicker cleared")
                                                handlechange("AssignedTo", null)
                                                return
                                                // await handlepeoplepickerchange()
                                            }
                                            //console.log("people", e[0].text, e)
                                            await handlepeoplepickerchange(e);
                                            // handlechange("AssignedTo",e[0].text || ' ')
                                        }}
                                        allowUnvalidated={true}
                                        ensureUser={true}
                                        resultFilter={(results: any[]) =>
                                            results.filter(persona => {
                                                if (isSameDomain && isSameDomain === "NO") {
                                                    let email = persona?.loginName || '';
                                                    //console.log("persona in if ", persona)
                                                    if (!email.includes('#ext#')) {
                                                        const x = externalGroups.find(group => group === persona?.text);
                                                        return persona?.text.includes(x);
                                                    }
                                                    return email.includes(`${DomainName}`);
                                                } else if (isSameDomain && !persona?.imageUrl) {
                                                    //console.log("persona in else if ", persona)
                                                    return !externalGroups.includes(persona?.text);
                                                }
                                                else if (isSameDomain) {
                                                    //console.log("persona in else ", persona)
                                                    let email = persona?.loginName || '';
                                                    // if (persona?.imageUrl === null || persona?.imageUrl === undefined) {
                                                    //     return false;
                                                    // }
                                                    return !email.includes('#ext#');
                                                }
                                            })}

                                    />

                                }
                            </div>
                            <div className="buttonsdiv">
                                {isediting ? <></> :
                                    <Button
                                        className="createbtn"
                                        onClick={async () => {
                                            await handlecreate();
                                            // setshowpopup(false);
                                            await fetchprocessitems()
                                        }}
                                    >
                                        Create
                                    </Button>
                                }
                                {isediting ?
                                    <Button
                                        className="createbtn"
                                        disabled={Togglestr === "My Actions"}
                                        onClick={async () => {
                                            await handleupdate();
                                            await fetchprocessitems()
                                            // setshowpopup(false);
                                        }}
                                    >
                                        Update
                                    </Button> : <></>
                                }
                                <Button
                                    className="cancelbtn"
                                    onClick={() => {
                                        setformdata({});
                                        setshowpopup(false)
                                    }}
                                >
                                    Close
                                </Button>
                            </div>
                        </div>

                    </div>
                )}
            </div>
        </div>
    )
}
export default MyZone;