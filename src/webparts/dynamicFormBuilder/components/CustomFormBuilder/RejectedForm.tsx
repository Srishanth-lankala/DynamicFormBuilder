import * as React from 'react';
import { useEffect, useState } from 'react';
import { useParams, useNavigate } from 'react-router-dom';
import { sp } from "@pnp/sp/presets/all";
import { ReactFormGenerator } from 'react-form-builder2';
import './EditTransactionForm.css';
import TopNavBar from './TopBar';
import Sidebar from './SideNavBar';
import { dataservice } from './encryptionutil';

const RejectedForm: React.FC = () => {
    const { id } = useParams<{ id: string }>(); // Retrieve the ID from the route
    const [formData, setFormData] = useState<any[]>([]);
    const [formJSON, setFormJSON] = useState<any>(null);
    const [loading, setLoading] = useState<boolean>(true);
    // const [isCurApprover, setIsCurApprover] = useState<boolean>(false); // State to track if the user is CurApprover
    const [curApproverEmail, setCurApproverEmail] = useState<string | null>(null);
    const dataserviceobj = new dataservice();
    const navigate = useNavigate();

    useEffect(() => {
        const fetchFormData = async () => {
            try {
                // Fetch the form data by ID from the WorkFlowProcessData list
                const transactionItem = await sp.web.lists.getByTitle("WorkFlowProcessData").items.getById(Number(id)).get();
                const parsedFormData = JSON.parse(transactionItem.FormData); // Parse FormData JSON
                setFormData(parsedFormData);

                const appCode = transactionItem.AppCode;
                const curApproverId = transactionItem.CurApproverId;

                setCurApproverEmail(curApproverId);

                console.log("CurAppover email from list ", curApproverEmail)
                // Fetch the corresponding FormJSON from the FormMaster list based on AppName
                const formMasterItems = await sp.web.lists.getByTitle("FormMaster").items.filter(`AppCode eq '${appCode}'`).get();
                if (formMasterItems.length > 0) {
                    const formMasterItem = formMasterItems[0];
                    const parsedFormJSON = JSON.parse(dataserviceobj.decryptjson(formMasterItem.FormJSON)); // Parse FormJSON JSON
                    setFormJSON(parsedFormJSON);
                }


                // const currentUser  = await sp.web.currentUser .get();
                // if (currentUser.Id === curApproverId) {
                //     setIsCurApprover(true); 
                // }
            } catch (error) {
                console.error("Error fetching form data:", error);
            } finally {
                setLoading(false);
            }
        };

        void fetchFormData();
    }, [id]);



    const handleClose = () => {

        navigate(`/RejectedView`);
    };

    return (
        <div style={{ display: 'flex', flexDirection: 'column' }}>
            <TopNavBar />
            <div style={{ display: 'flex' }}>
                <Sidebar />
                <div>
            <h2>Rejected Form</h2>
            {loading ? (
                <p>Loading...</p>
            ) : (
                <>
                    <ReactFormGenerator
                        data={formJSON} 
                        read_only={true}
                        form_action=""
                        form_method=""
                        answer_data={formData} 
                    />
                    {/* {isCurApprover && ( // Render buttons only if the user is CurApprover
                        <div>
                           
                            <button className='Submit-btn' onClick={handleCancel}>Close</button>
                        </div>
                    )} */}
                     <div>
                           
                           <button className='Submit-btn' onClick={handleClose}>Close</button>
                       </div>
                </>
            )}
        </div> 

            </div>

        </div>
    );
};

export default RejectedForm;



