import * as React from "react";
import { useState } from "react";
// import { Link } from "react-router-dom";
import "./SideNavBar.css";
import { Link } from "react-router-dom";
// useLocation
// import 'primereact/resources/themes/saga-blue/theme.css'; // Optional theme
// import 'primereact/resources/primereact.min.css';
// import 'primeicons/primeicons.css';

const Sidebar = () => {
  const [isToggled, setIsToggled] = useState(false);
  const [viewOptions, setViewOptions] = useState(1);
  const [isSidebarOpen, setIsSidebarOpen] = useState(false);
  const [isMobile, setIsMobile] = useState(false); // Determine if it's mobile view

  React.useEffect(() => {
    const handleResize = () => {
      setIsMobile(window.innerWidth <= 768); // Define mobile breakpoint
    };
    handleResize(); // Check on component mount
    window.addEventListener("resize", handleResize);
    return () => window.removeEventListener("resize", handleResize);
  }, []);

  let z: string | null;

  const toggleNav = () => {
    setIsToggled(!isToggled);
    setViewOptions(1);
    setIsSidebarOpen(!isSidebarOpen)
    
    z = localStorage.getItem("formName");

    console.log("zzzzzz", z);
    localStorage.removeItem("formName");
  };

//   const toggleSidebar = () => {
//     setIsSidebarOpen(!isSidebarOpen); // Toggles sidebar visibility
//   };

  const viewFunction = () => {
    setViewOptions(viewOptions + 1);

    if (viewOptions === 2) {
      setViewOptions(1);
    }
  };

  return (
    <div>
      {!isSidebarOpen &&<div className="burgerMenu" onClick={toggleNav}>
        {/* <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24">
          <path d="M3 6h18M3 12h18M3 18h18" />
        </svg> */}
         <img
          src={require("../../assets/Images/Toogle_4ac81a0a7dd064b5849a0695a2e5f1e3.svg")}
          alt=""
          className="navtoglcls"
        ></img>
      </div>
      }
      {/* <div className={`mainContainer ${isSidebarOpen ? "open" : ""}`}>
        <div className={`rightSideNavBar ${isSidebarOpen ? "open" : ""}`}></div> */}
      <div className={`mainContainer ${isSidebarOpen ? "open" : ""}`}>
      <div
        className={
          isMobile
            ? `rightSideNavBar ${isSidebarOpen ? "open" : ""}`
            : isToggled
            ? "rightSideNavBar"
            : "righttrans"
        }
      >
          <div className="navlist">
            <div className=" toggle_arrow">
              <Link to="/">
                <span className="link-text marketplace">Marketplace</span>
              </Link>
              <span className="icon " onClick={toggleNav}>
                {!isToggled ? (
                  <img
                    src={require("../../assets/Images/Toogle_4ac81a0a7dd064b5849a0695a2e5f1e3.svg")}
                    alt=""
                    className="navtoglcls"
                  ></img>
                ) : (
                  <img
                    src={require("../../assets/Images/Toogle_4ac81a0a7dd064b5849a0695a2e5f1e3.svg")}
                  />
                )}
              </span>
            </div>
            <Link
              style={{ display: "none" }}
              className={!isToggled ? "commonNavIcon" : "commonNavIcon2"}
              to="/App"
            >
              <span className="icon">
                <svg
                  xmlns="http://www.w3.org/2000/svg"
                  className="icon-tabler"
                  width="24"
                  height="24"
                  viewBox="0 0 24 24"
                  strokeWidth="1.5"
                  stroke="currentColor"
                  fill="none"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                >
                  <path stroke="none" d="M0 0h24v24H0z" fill="none" />
                  <path d="M12 5v14" />
                  <path d="M5 12h14" />
                </svg>
              </span>
              <span className="link-text">Create Form</span>
            </Link>
            <div className={!isToggled ? "commonNavIcon" : "commonNavIcon2"}>
              <div className="view_flex">
                <div
                  className={
                    isToggled ? "dropdown_removing_padding" : "commonNavIcon2"
                  }
                >
                  <span className="icon">
                    <svg
                      xmlns="http://www.w3.org/2000/svg"
                      height="24px"
                      viewBox="0 -960 960 960"
                      width="24px"
                      fill="white"
                      style={{ marginBottom: "3.5px" }}
                    >
                      <path d="M480-320q75 0 127.5-52.5T660-500q0-75-52.5-127.5T480-680q-75 0-127.5 52.5T300-500q0 75 52.5 127.5T480-320Zm0-72q-45 0-76.5-31.5T372-500q0-45 31.5-76.5T480-608q45 0 76.5 31.5T588-500q0 45-31.5 76.5T480-392Zm0 192q-146 0-266-81.5T40-500q54-137 174-218.5T480-800q146 0 266 81.5T920-500q-54 137-174 218.5T480-200Zm0-300Zm0 220q113 0 207.5-59.5T832-500q-50-101-144.5-160.5T480-720q-113 0-207.5 59.5T128-500q50 101 144.5 160.5T480-280Z" />
                    </svg>
                  </span>
                  <section className="link-text">
                    <span>Views</span>
                  </section>
                </div>
                <div className="dropdown_removing_padding">
                  {isToggled && (
                    <span onClick={viewFunction}>
                    {" "}
                    <svg
                      xmlns="http://www.w3.org/2000/svg"
                      height="24px"
                      viewBox="0 -960 960 960"
                      width="24px"
                      fill="white"
                    >
                      <path d="M480-344 240-584l56-56 184 184 184-184 56 56-240 240Z" />
                    </svg>{" "}
                  </span>
                  )}
                </div>
              </div>
            </div>
            <div>
              {viewOptions % 2 === 0 ? (
                <section>
                  <Link style={{ textDecoration: "none" }} to="/ViewByDate">
                    <section className="dropdown_option">
                      <svg
                        xmlns="http://www.w3.org/2000/svg"
                        height="24px"
                        viewBox="0 -960 960 960"
                        width="24px"
                        fill="white"
                      >
                        <path d="M400-280v-400l200 200-200 200Z" />
                      </svg>
                      By Date
                    </section>
                  </Link>
                  <Link style={{ textDecoration: "none" }} to="/ViewByStatus">
                    <section className="dropdown_option">
                      <svg
                        xmlns="http://www.w3.org/2000/svg"
                        height="24px"
                        viewBox="0 -960 960 960"
                        width="24px"
                        fill="white"
                      >
                        <path d="M400-280v-400l200 200-200 200Z" />
                      </svg>
                      By Status
                    </section>
                  </Link>
                </section>
              ) : (
                ""
              )}
            </div>
            <section>
              <Link
                style={{ display: "none" }}
                className={!isToggled ? "commonNavIcon" : "commonNavIcon2"}
                to="/LoadForm"
              >
                <span className="icon">
                  <svg
                    xmlns="http://www.w3.org/2000/svg"
                    className="icon-tabler"
                    width="24"
                    height="24"
                    viewBox="0 0 24 24"
                    strokeWidth="1.5"
                    stroke="currentColor"
                    fill="none"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                  >
                    <path stroke="none" d="M0 0h24v24H0z" fill="none" />
                    <path d="M9 12l2 2l4 -4" />
                    <path d="M12 7h.01" />
                    <path d="M17 17h.01" />
                    <path d="M17 7h.01" />
                    <path d="M7 17h.01" />
                    <path d="M7 7h.01" />
                    <path d="M3 12h1.99m16.02 0h2" />
                  </svg>
                </span>
                <span className="link-text">Load Form</span>
              </Link>

              <Link
                style={{ display: "none" }}
                className={!isToggled ? "commonNavIcon" : "commonNavIcon2"}
                to="/ConfigureWorkflow"
              >
                <span className="icon">
                  <svg
                    xmlns="http://www.w3.org/2000/svg"
                    className="icon icon-tabler icon-tabler-settings"
                    width="24"
                    height="24"
                    viewBox="0 0 24 24"
                    stroke-width="1.5"
                    stroke="currentColor"
                    fill="none"
                    stroke-linecap="round"
                    stroke-linejoin="round"
                  >
                    <path stroke="none" d="M0 0h24v24H0z" fill="none" />
                    <path d="M12 8a4 4 0 1 0 0 8a4 4 0 1 0 0 -8" />
                    <path d="M4.6 9l-.6 -2.3l2.3 -.6l1.4 2.2a7.8 7.8 0 0 0 -1 1.6l-2.1 -.9z" />
                    <path d="M19.4 9l.6 -2.3l-2.3 -.6l-1.4 2.2a7.8 7.8 0 0 0 1 1.6l2.1 -.9z" />
                    <path d="M12 4.5a7.5 7.5 0 0 1 3.8 1" />
                    <path d="M12 19.5a7.5 7.5 0 0 1 -3.8 -1" />
                    <path d="M9 12l-2 1" />
                    <path d="M15 12l2 -1" />
                  </svg>
                </span>
                <span
                  className="link-text Configure"
                  style={{ marginLeft: "-3px" }}
                >
                  Configure Workflow
                </span>
              </Link>
              <Link
                className={!isToggled ? "commonNavIcon" : "commonNavIcon2"}
                to={{
                  pathname: "/ReactDiagram",
                }}
              >
                <span className="icon">
                  <svg
                    xmlns="http://www.w3.org/2000/svg"
                    className="icon-workflow"
                    width="24"
                    height="24"
                    viewBox="0 0 24 24"
                    strokeWidth="1.5"
                    stroke="currentColor"
                    fill="none"
                    strokeLinecap="round"
                    strokeLinejoin="round"
                  >
                    <circle cx="5" cy="5" r="2" />
                    <circle cx="19" cy="5" r="2" />
                    <circle cx="12" cy="12" r="2" />
                    <circle cx="5" cy="19" r="2" />
                    <circle cx="19" cy="19" r="2" />
                    <path d="M5 5 L12 12" />
                    <path d="M19 5 L12 12" />
                    <path d="M12 12 L5 19" />
                    <path d="M12 12 L19 19" />
                  </svg>
                </span>
                <span
                  className="link-text Configure"
                  style={{ marginLeft: "-3px" }}
                >
                  Workflow
                </span>
              </Link>
            </section>
          </div>
        </div>
      </div>
    </div>
  );
};

export default Sidebar;
