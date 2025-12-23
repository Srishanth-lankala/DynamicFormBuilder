import React from "react";
import { useEffect, ReactNode } from "react";
import { useLocation } from "react-router-dom";

interface LayoutProps {
    children: ReactNode;
}

const Layout: React.FC<LayoutProps> = ({ children }) => {
    const location = useLocation();

    useEffect(() => {
        // console.log("Current path:", location.pathname); // Debugging
        if (location.pathname.includes("/viewform/")) {
            document.body.classList.add("viewform-mode");
        } else {
            document.body.classList.remove("viewform-mode");
        }
    }, [location.pathname]);

    return <>{children}</>;
};

export default Layout;
