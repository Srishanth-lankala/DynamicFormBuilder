import * as React from 'react';
import './LoadingSpinner.css';

const LoadingSpinner = () => {
    return (
        <div className="Loading_background">
            <div className="loading">
                <img
                    src={require('../../assets/Images/loading.gif')}
                    alt="Loading..."
                    className="static-logo"
                />
            </div>
        </div>
    );
};

export default LoadingSpinner;
