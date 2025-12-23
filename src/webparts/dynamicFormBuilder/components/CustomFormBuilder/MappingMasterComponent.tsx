import * as React from 'react';
import { useState, useEffect } from 'react';
import { useParams, useLocation, useNavigate } from 'react-router-dom';
import { sp } from "@pnp/sp/presets/all";
import LoadingSpinner from './Loading';
import Sidebar from './SideNavBar';
import TopNavBar from './TopBar';
import "./MappingMasterComponent.css";
import { Timeline } from 'primereact/timeline';
import { mySiteUrl } from './ConfigURL/All_URLs';

interface MappingMasterData {
    Role: string;
    Level: string;
    Users_x002f_Groups: { Title: string };
    AppName: string;
    AppCode: string;
    Author: { Title: string };
}

const MappingMasterComponent: React.FC = () => {
    const { appCode } = useParams<{ appCode: string }>();
    const location = useLocation();
    const searchParams = new URLSearchParams(location.search);
    const appName = searchParams.get("appName") || "Default App Name";
    const authorName = searchParams.get("authorName") || "Default Author";
    const [mappingData, setMappingData] = useState<MappingMasterData[]>([]);
    const [error, setError] = useState<string | null>(null);
    const [isLoading, setIsLoading] = useState<boolean>(false);
    const [isSpinnerVisible, setIsSpinnerVisible] = useState<boolean>(false);
    const navigate = useNavigate()
    sp.setup({
        sp: {
            baseUrl: mySiteUrl,
        },
    });

    useEffect(() => {
        const fetchMappingData = async () => {
            setIsLoading(true);
            setError(null);
            setIsSpinnerVisible(true);
            const spinnerTimeout = setTimeout(() => {
                setIsSpinnerVisible(true);
            }, 10000);

            try {
                const data = await sp.web.lists
                    .getByTitle('MappingMaster')
                    .items.select('Role', 'Users_x002f_Groups/Title', 'Level', 'AppName', 'AppCode', 'Author/Title', 'Id')
                    .expand('Users_x002f_Groups', 'Author')
                    .filter(`AppCode eq '${appCode}'`)
                    .get();
                data.sort((a, b) => a.Level - b.Level);
                setMappingData(data);
            } catch (error) {
                console.error("Error fetching mapping data:", error);
                setError("An error occurred while fetching mapping data.");
            } finally {
                setIsLoading(false);
                setIsSpinnerVisible(false);
                clearTimeout(spinnerTimeout);
            }
        };

        void fetchMappingData();
    }, [appCode]);

    // Function to determine the background color based on role
    const getRoleBackgroundColor = (role: string) => {
        switch (role) {
            case 'Creator':
                return { backgroundColor: '#2196F3', color: 'white' }; // Green for Creator
            case 'Reviewer':
                return { backgroundColor: '#FF9800', color: 'white' }; // Orange for Reviewer
            case 'Approver':
                return { backgroundColor: '#4CAF50', color: 'white' }; // Blue for Approver
            default:
                return { backgroundColor: '#9E9E9E', color: 'white' }; // Default gray color for unknown roles
        }
    };

    return (
        <div style={{ display: 'flex', flexDirection: 'column' }}>
            <TopNavBar />
            <div style={{ display: 'flex' }}>
                <Sidebar />
                <div style={{ width: '100%', height: '100vh', display: 'flex', flexDirection: 'column' }}>
                    <div
                        style={{
                            position: 'sticky',
                            top: 0,
                            backgroundColor: '#fff',
                            zIndex: 10,
                            padding: '12px 25px',
                            display: 'flex',
                            justifyContent: 'space-between',
                            borderBottom: '1px solid #ddd'
                        }}
                    >
                        {/* <h6 style={{ marginBottom: '5px', fontWeight: 'bold', fontSize: '20px' }}>{appName}</h6> */}
                        <div style={{ display: 'flex', alignItems: 'center', gap: '7px' }}>
                            <img onClick={() => navigate(-1)} style={{ height: 18, width: 18, cursor: "pointer" }} src={require('../../assets/Images/previous.png')} alt="backicon" title='Back' />
                            <h6 style={{ marginBottom: '5px', fontWeight: 'bold', fontSize: '20px' }}>{appName}</h6>
                        </div>
                        <p style={{ fontStyle: 'italic', color: '#666' }}>Created By: {authorName.split('|')[0]}</p>
                    </div>

                    {/* Scrollable Content Section */}
                    <div style={{ flex: 1, overflowY: 'auto', padding: '20px' }}>
                        {error && <div style={{ color: "red", marginBottom: '20px' }}>{error}</div>}

                        {isSpinnerVisible && <LoadingSpinner />}

                        {isLoading ? (
                            <></>
                        ) : (
                            <div style={{ height: '100%' }}>
                                {mappingData.length > 0 ? (
                                    <div className='MappingGraph'>
                                        <Timeline
                                            value={mappingData}
                                            align="right"
                                            marker={(item) => (
                                                <div
                                                    className="timeline-marker"
                                                    title={item.Role}
                                                    style={{
                                                        ...getRoleBackgroundColor(item.Role),
                                                        padding: '10px',
                                                        borderRadius: '50%',
                                                    }}
                                                >
                                                    {item.Level}
                                                </div>
                                            )}
                                            content={(item) => (
                                                <div className="timeline-content">
                                                    <p>
                                                        {item.Users_x002f_Groups?.Title.split('|')[0] || 'Immediate Manager'}
                                                    </p>
                                                </div>
                                            )}
                                            className="w-full md:w-20rem"
                                        />
                                    </div>
                                ) : (
                                    <div>No workflow available.</div>
                                )}
                            </div>
                        )}
                    </div>
                </div>

            </div>
        </div>
    );
};

export default MappingMasterComponent;
