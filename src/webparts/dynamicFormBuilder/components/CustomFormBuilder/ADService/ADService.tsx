import { PublicClientApplication, Configuration, AuthenticationResult } from "@azure/msal-browser";
import { auth } from "../ConfigURL/All_URLs";
import { sp } from '@pnp/sp';
// MSAL Configuration for authentication
const msalConfig: Configuration = { auth };

// Create MSAL instance
const msalInstance = new PublicClientApplication(msalConfig);

// Fetch access token
const getToken = async (): Promise<string | null> => {
  try {
    await msalInstance.initialize();
    const account = msalInstance.getAllAccounts()[0];
    let response: AuthenticationResult;

    if (account) {
      response = await msalInstance.acquireTokenSilent({
        scopes: ["User.Read", "User.Read.All"], // Added "User.Read.All" to access all users
        account: account
      });
    } else {
      response = await msalInstance.loginPopup({
        scopes: ["User.Read", "User.Read.All"]
      });
    }

    return response.accessToken;
  } catch (error) {
    console.error("Error acquiring token:", error);
    return null;
  }
};

// Fetch logged-in user's details from Microsoft Graph API
const fetchADUserDetails = async (accessToken: string) => {
  try {
    const response = await fetch("https://graph.microsoft.com/v1.0/me?$select=department,employeeId,jobTitle,mobilePhone", {
      method: "GET",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json"
      }
    });

    if (!response.ok) {
      throw new Error(`HTTP error! status: ${response.status}`);
    }

    const userDetails = await response.json();
    console.log("Current logged-in user details:", userDetails);
    return userDetails;
  } catch (error) {
    console.error("Error fetching current logged-in user details:", error);
    return null;
  }
};

// Fetch all users from AD
const fetchAllUsers = async () => {


  const token = await getToken();


  if (!token) {


    console.log("Failed to acquire token");


    return null;


  }

  try {


    const currentUser = await sp.web.currentUser();


    const loggedInEmail = currentUser.Email;


    const emailDomain = loggedInEmail.split("@")[1]; // Extract domain

    const baseUrl = "https://graph.microsoft.com/v1.0/users?$select=id,displayName,givenName,surname,mail,jobTitle,mobilePhone,officeLocation,userPrincipalName";


    let allUsers: any[] = [];


    let nextLink: string | null = baseUrl;

    while (nextLink) {


      const response = await fetch(nextLink, {


        method: "GET",


        headers: {


          Authorization: `Bearer ${token}`,


          "Content-Type": "application/json"


        }


      });

      if (!response.ok) {


        throw new Error(`HTTP error! status: ${response.status}`);


      }

      const data: any = await response.json();


      allUsers = allUsers.concat(data.value);

      // Check for more pages


      nextLink = data["@odata.nextLink"] || null;


    }

    // Optional: filter only users from your domain


    const filteredUsers = allUsers.filter(


      (user) => user.mail && user.mail.endsWith(`@${emailDomain}`)


    );

    console.log("All Filtered AD Users:", filteredUsers);


    return filteredUsers;


  } catch (error) {


    console.error("Error fetching all users:", error);


    return null;


  }


};




// Example usage of fetching user details
const fetchToken = async () => {
  const token = await getToken();
  if (token) {

    const userDetails = await fetchADUserDetails(token);
    console.log("User details:", userDetails);
    return userDetails;
  } else {
    console.log("Failed to acquire token");
    return null;
  }
};

export { fetchToken, fetchAllUsers };


