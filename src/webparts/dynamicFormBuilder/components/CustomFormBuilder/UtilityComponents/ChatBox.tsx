import React, { useEffect, useState, useRef } from "react";
import { sp } from "@pnp/sp/presets/all";
import { PeoplePicker, PrincipalType, IPeoplePickerContext } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { Toast } from "primereact/toast";
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
// import { faReply } from '@fortawesome/free-solid-svg-icons';
import { faXmark } from '@fortawesome/free-solid-svg-icons';
import { convertUtcToLocalTime, convertDateToLocal } from '../datetime';
import './CommunicationChannel.css';
import channelicon from '../../../assets/Images/channelicon.svg';
// import { fetchAllUsers } from "../ADService/ADService";
import { fetchTenantUser } from "../FetchTenantUser/fetchTenantUser";
interface ChatBoxProps {
  context: any;
  appSeqNo: string;
  userName: string;
  appcode: string;
  onChatResponse: (severity: string, message: string) => void;
  isSameDomain: string
}



const ChatBox: React.FC<ChatBoxProps> = ({ context, appSeqNo, userName, appcode, onChatResponse, isSameDomain }) => {
  const [chat, setChat] = useState("");
  const [communication, setCommunication] = useState<any[]>([]);
  const [isOpen, setIsOpen] = useState(false);
  const [chatBy, setChatBy] = useState<string | null>(null);
  const [pickerKey, setPickerKey] = useState(Date.now());
  const toast = useRef<Toast>(null);
  const [selectedPerson, setSelectedPerson] = useState('');
  const [domainfield, setDomainField] = useState<string>("");


  const textAreaRef = useRef<HTMLTextAreaElement | null>(null);

  const getPeoplePickerContext = (): IPeoplePickerContext => ({
    absoluteUrl: context.pageContext.web.absoluteUrl,
    msGraphClientFactory: context.msGraphClientFactory,
    spHttpClient: context.spHttpClient,
  });

  const fetchCommunicationLog = async () => {
    try {
      const items = await sp.web.lists
        .getByTitle("CommunicationLog")
        .items.select("Title", "Chat", "Author/Title", "Author/EMail", "Created", "MessageTo/Title")
        .expand("Author", "MessageTo")
        .filter(`Title eq '${appSeqNo}'`)
        .get();
      setCommunication(items);
    } catch (error) {
      console.error("Error fetching communication:", error);
    }
  };

  const getPersonData = async (items: any[]) => {
    try {
      const selectedUser = items[0]; // Only 1 user allowed
      const email = selectedUser?.secondaryText; // This is the email
      if (email) {
        setChatBy(email); // ✅ Set email, not ID
      }
    } catch (err) {
      console.log(err);
    }
  };




  useEffect(() => {
    if (selectedPerson) {
      void getPersonData([{ loginName: selectedPerson }]);
      setSelectedPerson("")
    }
  }, [selectedPerson]);

  const handleSendComment = async () => {
    if (!chat) {
      onChatResponse("warn", "Please enter a message");
      return;
    }

    try {
      const user = await sp.web.ensureUser(chatBy as string);
      await sp.web.lists.getByTitle("CommunicationLog").items.add({
        Chat: chat,
        Title: appSeqNo,
        MessageToId: user.data.Id,
        AppCode: appcode
      });


      setCommunication(prev => [...prev, {
        Author: { Title: userName },
        Chat: chat,
        Created: new Date().toLocaleTimeString(),
        MessageToId: user.data.Id
      }]);

      await fetchCommunicationLog();
      setChat("");
      setChatBy(null);
      // setSelectedUsers([]);
      setPickerKey(Date.now());

    } catch (error) {
      console.error("Error adding comment:", error);
      onChatResponse("warn", "Please tag a person");
    }
  };

  const fetchingTenantUsers = async () => {
    try {
      const domain = await fetchTenantUser()
      setDomainField(domain.TenantUsers)
    } catch (error) {
      console.error("Error fetching users:", error);
    }
  };



  const handleClose = () => setIsOpen(!isOpen);
  useEffect(() => {
    void fetchCommunicationLog();
    void fetchingTenantUsers();
  }, []);

  useEffect(()=>{
     void fetchCommunicationLog();
     void fetchingTenantUsers();

  },[appSeqNo]);

  return (
    <>
      <Toast ref={toast} />
      {isOpen && (
        <div className="chatbox-container">
          <div className="chatbox-header">
            <div className="header-content">
              <div className="chatbox-title">Communication Channel</div>
              <FontAwesomeIcon icon={faXmark} onClick={handleClose} className="close-icon" />
            </div>
          </div>

          <div className="chatbox-messages">
            {communication.length > 0 ? (
              <ul className="message-list">
                {communication.map((comment, index) => {
                  const loginUser = comment.Author?.Title || "";
                  const loginUserSplit = loginUser.split("|")[0];
                  const firstLetter = loginUserSplit
                    .split(" ")
                    .map((word: any) => word[0])
                    .join("");
                  // const isLoggedInUser = comment.Author?.EMail === userEmail;
                  const messageTime = convertUtcToLocalTime(comment.Created);
                  const oldDate = convertDateToLocal(comment.Created.split("T")[0]);

                  function stripHtmlTags(input: string): string {
                    const doc = new DOMParser().parseFromString(input, "text/html");
                    return doc.body.textContent || "";
                  }

                  return (
                    <li key={index} className="message-item">
                      <div className="avatar-container">
                        <div className="avatar" title={loginUserSplit}>
                          {firstLetter}
                        </div>
                      </div>

                      <div className="message-bubble">
                        <h1>@{comment.MessageTo?.Title.split("|")[0]}</h1>
                        <p className="message-text">{stripHtmlTags(comment.Chat)}</p>
                        <div className="datetime-container">
                          <p className="date">{oldDate}</p>
                          <p className="time">{messageTime}</p>
                        </div>
                      </div>
                    </li>
                  );
                })}
              </ul>
            ) : (
              <p className="no-messages">No communication available.</p>
            )}
          </div>

          <div className="chatbox-input">

            <PeoplePicker
              key={pickerKey}
              context={getPeoplePickerContext()}
              personSelectionLimit={1}
              required
              onChange={getPersonData}
              principalTypes={[PrincipalType.User]}
              placeholder="Send to..."
              defaultSelectedUsers={selectedPerson ? [selectedPerson] : []}
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


            <textarea
              ref={textAreaRef}
              value={chat}
              onChange={(e) => setChat(e.target.value)}
              placeholder="Type your message..."
              maxLength={250}
              className="message-input"
            />
            <button onClick={handleSendComment} className="send-button">Send</button>
          </div>
        </div>
      )}

      <button
        onClick={() => {
          setIsOpen(!isOpen);
          void fetchCommunicationLog();
        }}
        className="chat-icon-button"
        aria-label="Toggle Chat"
      >
        <img
          src={channelicon}  // <-- Replace with actual image path
          alt="Chat Icon"
          height="25"
          width="25"
          className="chat-icon"
        />
      </button>
    </>
  );
};

export default ChatBox;