import React, { useEffect, useState } from "react";
import { sp } from '@pnp/sp/presets/all';
import { mySiteUrl } from "../ConfigURL/All_URLs";
import './MyZoneCommentBox.css';
// import commentsicon from '../../../assets/Images/CommentsIcons.png';
import { convertDateToLocal, convertUtcToLocalTime } from "../datetime";

interface commentboxprops {
    isopen: boolean,
    ontoggle: () => void,
    TaskId: string,
    user: string,
    top: string,
    left: string,
    // sendlatestcomment?: (TaskId:string,comment:string) => void,
}

const MyZoneCommentBox: React.FC<commentboxprops> = ({ isopen, ontoggle, TaskId, user,top,left}) => {
    // const [openbox,setopenbox]=useState<boolean>(false);
    // const [popupposition, setpopupposition] = useState({ top: 0, left: 0 });
    const [commentstr, setcommentstr] = useState<string>('');
    const [fetchedcomments, setfetchedcomments] = useState<any>()
    // const buttonref = useRef<HTMLButtonElement>(null);
    console.log("isopen",isopen,TaskId);
    sp.setup({
        sp: {
            baseUrl: mySiteUrl,
        },
    });
    const fetchcomments = async () => {
        try {
            const items = await sp.web.lists.getByTitle("MyZoneComments").items.select("Comment", "TaskID", "Author/Id", "Author/Title", "Created").expand("Author").filter(`TaskID eq '${TaskId}'`).getAll();
            console.log("fetchitems", items, TaskId)
            items.sort((a, b) => {
                const adate = new Date(a.Created);
                const bdate = new Date(b.Created);
                return adate > bdate ? -1 : 1
            })
            setfetchedcomments(items);

            // if(items?.length > 0 && sendlatestcomment){
            //     sendlatestcomment(TaskId,items[0].Comment);
            // }
            // else if(sendlatestcomment){
            //     sendlatestcomment(TaskId,'');
            // }
        }
        catch (err) {
            console.log(err)
        }
    }
    useEffect(() => {
        fetchcomments().catch((err) => console.log("error in fetching", err))
    }, [])
    useEffect(() => {
        console.log("commentfromuse", commentstr)
    }, [commentstr])
    const handlecomment = (comment: string) => {
        setcommentstr(comment)
    }
    const handlecommentsave = async () => {
        try {
            const item = await sp.web.lists.getByTitle("MyZoneComments").items.add(
                {
                    Comment: commentstr,
                    TaskID: TaskId,
                }
            )
            console.log("item", item)
            setfetchedcomments((prev: any) => [...prev, {
                Author: { Title: user },
                Comment: commentstr,
                Created: new Date().toLocaleTimeString(),
                TaskID: TaskId,
            }]);
            await fetchcomments();
        }
        catch (err) {
            console.log(err)
        }

    }
    return (
        <>
            {/* <button
                className="commentsicon-button"
                ref={buttonref}
                onClick={async (e) => {
                    e.stopPropagation();
                    ontoggle();
                    const rect = buttonref.current?.getBoundingClientRect();
                    console.log("rectobj",rect)
                    if (rect) {
                        setpopupposition({
                            top: rect.top + rect.height,
                            left: rect.left + rect.width - 350,
                        })
                    }
                    await fetchcomments()
                }}
            >
                <img
                    src={commentsicon}
                    className="commentsicon"

                />
            </button> */}

            {(isopen &&
                <div
                    className="commentboxoverlay"
                    onClick={() => ontoggle()}
                >
                    <div
                        className="myzonechatbox-container"
                        style={{ position: 'absolute', top: top, left: left, zIndex: 1000 }}
                        onClick={(e) => e.stopPropagation()}
                    >
{console.log("rectvariables",top,left)}
                        <div className="comments-header">Add Comments</div>
                        <div className="comments-box">

                            {fetchedcomments?.length > 0 ? (fetchedcomments?.map((item: any) => {
                                const commentdate = convertDateToLocal(item.Created.split("T")[0]);
                                const commenttime = convertUtcToLocalTime(item.Created);
                                return (
                                    <div className="message-item">
                                        <div className="comment-bubble">
                                            <h1>{item.Author?.Title?.split("|")[0]}</h1>
                                            <p className="comment-text">{item.Comment}</p>
                                            <div className="date-container">
                                                <p className="date">{commentdate}</p>
                                                <p className="date-time">{commenttime}</p>
                                            </div>
                                        </div>
                                    </div>

                                )
                            })) :
                                <p className="no-comments">No Comments Available</p>
                            }

                        </div>
        <div style={{display:'flex',marginTop:'auto',marginLeft:'5px',marginBottom:'5px',marginRight:'-10px'}}>
                            <textarea
                                className="commentinput"
                                style={{ width: '85%' }}
                                onChange={(e) => {
                                    console.log("commentchanged", e.target.value)
                                    handlecomment(e.target.value)
                                }}
                                placeholder="Type your comment..."
                                maxLength={250}
                            />
                            <button
                                onClick={async () => {
                                    if (commentstr != "") {
                                        console.log("clicked")
                                        await handlecommentsave();
                                    }
                                }}
                                className="addcomment"
                            >
                                Add
                            </button>
                        </div>
                    </div>
                </div>

            )}

        </>
    )
}
export default MyZoneCommentBox;