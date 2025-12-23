import * as React from 'react';
import { sp } from '@pnp/sp';
import './TopBar.css';
import { Link } from 'react-router-dom';
import { mySiteUrl } from './ConfigURL/All_URLs';
import { myDomain } from './ConfigURL/All_URLs';
import { useEffect } from 'react';
import signout from '../../assets/Images/signout.svg';
import domain from '../../assets/Images/domain.svg';

interface IUser {
  title: string;
  picture: string;
}

interface ITopNavBarProps {
  onUserTypeRetrieved?: (userType: string) => void;
  onUserDomainRetrieved?: (domainType: string) => void
}

const TopNavBar: React.FunctionComponent<ITopNavBarProps> = ({ onUserTypeRetrieved, onUserDomainRetrieved }) => {
  const [user, setUser] = React.useState<IUser | null>(null);
  const [showSignOutMenu, setShowSignOutMenu] = React.useState(false); // State to manage sign-out menu visibility
  const [userDisplayDomain, setuserDisplayDomain] = React.useState("")
  const dropdownRef = React.useRef<HTMLDivElement | null>(null);
  sp.setup({
    sp: {
      baseUrl: mySiteUrl,
    },
  });
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (
        dropdownRef.current &&
        !dropdownRef.current.contains(event.target as Node)
      ) {
        setShowSignOutMenu(false); // Close if clicking outside
      }
    };

    document.addEventListener("mousedown", handleClickOutside);

    return () => {
      document.removeEventListener("mousedown", handleClickOutside);
    };
  }, []);

  const fetchUserInfo = async () => {
    try {
      const currentUser = await sp.web.currentUser();
      const currentLoggedInEmail = currentUser.Email;
      const currentLoggedId = currentUser.Id;
      const userDomain = currentLoggedInEmail.split("@")[1];
      const displaydomain = userDomain.split(".")[0]
      setuserDisplayDomain(displaydomain)



      let isSameDomain = "NO"
      isSameDomain = userDomain === myDomain ? "YES" : "NO";

      console.log("Is Same Domain:----------------------------------------------", isSameDomain);

      if (onUserDomainRetrieved) {
        onUserDomainRetrieved(isSameDomain)
      }


      const profileImageUrl = `${mySiteUrl}/_layouts/15/userphoto.aspx?size=L&accountName=${currentLoggedInEmail}&default=true`;

      const data = await sp.web.lists
        .getByTitle('AuthList')
        .items.select('Title', 'AuthName/Title', 'AuthName/Id', 'AdminUser ', 'NormalUser ')
        .expand('AuthName')
        .get();
      // Step 1: Find the user by currentLoggedId
      const currentUserData = data.find((item) => item.AuthName.Id === currentLoggedId);

      // Step 2: Determine the user type
      let userType = "OnlyUser"; // Default role
      if (currentUserData) {
        const { AdminUser, NormalUser } = currentUserData;
        if (AdminUser === 'Yes' && NormalUser === 'No') {
          userType = "OnlyAdmin";
        } else if (AdminUser === 'Yes' && NormalUser === 'Yes') {
          userType = "BothUser Admin";
        }
      }

      // console.log('User  Type:', userType);

      if (onUserTypeRetrieved) {
        onUserTypeRetrieved(userType);
      }

      // Set the user object
      setUser({
        title: currentUser.Title,
        picture: profileImageUrl,
      });

      return userType; // Return the determined user type
    } catch (error) {
      console.error('Error fetching user data', error);
    }
  };

  React.useEffect(() => {
    void fetchUserInfo();
  }, []); // Empty dependency array ensures it only runs once when the component mounts

  return (
    <header className="topNavbar">
      <Link to={'/LoadForm'}>
        <div>
          <img src={require('../../assets/Images/logo-smartoffice.svg')} alt="imageOflogo-smartoffice" style={{ width: 170, height: 30 }} />
        </div>
      </Link>

      <div className="userProfile">
        {user ? (
          <div className="profileWrapper" ref={dropdownRef}>
            <div className="profileContainer" onClick={() => setShowSignOutMenu(!showSignOutMenu)}>
              <div style={{ display: "flex", gap: "10px", alignItems: "center" }}>
                <div>{user.title.split('|')[0]}</div>
                <img src={user.picture} alt="User  Profile" className="profileImg" />
              </div>
              {showSignOutMenu && (
                <div className="signOutMenu">
                  <div className="menuItem">
                    <img src={user.picture} alt="User Icon" className="userIcon" />
                    <span>{user.title.split('|')[0]}</span>
                  </div>
                  <div className="menuItem">
                    <img src={domain} alt="Domain Icon" className="menuIcon" />
                    <span>{userDisplayDomain}</span>
                  </div>
                  <div className="signout" onClick={() => {
                            window.location.href = `${mySiteUrl}/_layouts/15/SignOut.aspx`;
                          }}
                    >
                  <a
                    href={`${mySiteUrl}/_layouts/15/SignOut.aspx`}
                    className="signOutButton"
                  >
                      <img src={signout} alt="Logout Icon" />
                    <span>Sign Out</span>
                  </a>
                  </div>
                </div>
              )}
            </div>
          </div>
        ) : (
          <div className="logo">
            <img
              src={`${mySiteUrl}/Shared%20Documents/Logo.png`}
              alt="Logo"
            />
          </div>
        )}
      </div>

    </header>
  );
};

export default TopNavBar;