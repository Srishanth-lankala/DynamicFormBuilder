// import React from 'react';
// import { useNavigate } from 'react-router-dom';

// const NoAccess = () => {
//   const navigate = useNavigate();

//   return (
//     <div className="d-flex flex-column justify-content-center align-items-center vh-100 bg-light text-center px-3">
//       <h1 className="text-danger mb-3">Access Denied</h1>
//       <p className="text-muted mb-4">You do not have permission to view this page.</p>
//       <button className="btn btn-primary" onClick={() => navigate('/')}>
//         Go to Home
//       </button>
//     </div>
//   );
// };

// export default NoAccess;


// import React from 'react';
// import { useNavigate } from 'react-router-dom';

// const NoAccess = () => {
//     const navigate = useNavigate();

//     return (
//         <div className="d-flex flex-column justify-content-center align-items-center vh-100 bg-light text-center px-3">
//             <div className="border p-5 rounded shadow-sm bg-white">
//                 <h1 className="text-danger mb-3">🚫 Access Denied</h1>
//                 <p className="text-secondary mb-4">
//                     You do not have permission to view this page.
//                 </p>
//                 <button
//                     className="btn btn-outline-primary fw-semibold"
//                     onClick={() => navigate('/')}
//                 >
//                     ← Go to Home
//                 </button>
//             </div>
//         </div>
//     );
// };

// export default NoAccess;


import React from 'react';
import { useNavigate } from 'react-router-dom';

const NoAccess = () => {
    const navigate = useNavigate();

    return (
        <div className="d-flex flex-column justify-content-center align-items-center vh-100 text-white text-center px-3" style={{ backgroundColor: "#070d19" }}>
            <div className="border rounded p-5 bg-dark shadow-lg">
                <h1 className="text-danger mb-3">🚫 Access Denied</h1>
                <p className="text-light mb-4">You do not have permission to view this page.</p>
                <div className="d-flex gap-3 justify-content-center">
                    <button className="newlogocolorbtn" onClick={() => navigate('/')}>
                        Go to Home
                    </button>
                    <button className="newlogocolorbtn" onClick={() => navigate(-1)}>
                        Go Back
                    </button>
                </div>
            </div>
        </div >
    );
};

export default NoAccess;

