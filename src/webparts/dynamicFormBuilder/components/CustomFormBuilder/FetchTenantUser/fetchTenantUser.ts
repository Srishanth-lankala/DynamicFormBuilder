import { sp } from "@pnp/sp/presets/all"; // Ensure PnP library is imported

export const fetchTenantUser = async ():  Promise<{ CurrentUserId: number; TenantUsers: string }> => {
    try {
        const TenantUser = await sp.web.currentUser.get();
        const CurrentUserId=TenantUser.Id
        let TenantUsers = "";
        if (TenantUser.UserPrincipalName?.includes('#ext#')) {
            const userPrincipalName = TenantUser.UserPrincipalName;
            const parts = userPrincipalName.split('_');
            const afterUnderscore = parts[1];
            const finalPart = afterUnderscore.split('#')[0];
            TenantUsers = finalPart;
        } else if (TenantUser.Email) {
            const UserEmail = TenantUser.Email;
            const finalPart = UserEmail.split('@');
            TenantUsers = finalPart[1];
        } else {
            console.error("current logged in UserPrincipalName & Email are null or undefined");
        }
        return { CurrentUserId, TenantUsers };
    } catch (error) {
        console.error("Error fetching tenant user:", error);
        throw new Error("Failed to fetch tenant user.");
    }
};


