import { sp } from "@pnp/sp";
import { myDomain } from "../ConfigURL/All_URLs";

export const fetchWorkflowLimit = async (domain: string, appCode: string) => {
    try {
        if (myDomain === domain) {
            return { recordscount: 0, trnxRecLimitReached: false };
        }
        const items = await sp.web.lists
            .getByTitle("Workflowprocessdata")
            .items.filter(`Domain eq '${domain}' and substringof('${appCode}', AppCode) and Status ne  'Draft'`)
            .select("Id")
            .getAll();

        const domainlistItem = await sp.web.lists
            .getByTitle("DomainList")
            .items.select("Id", "LicenseType")
            .filter(`Tenant eq '${domain}'`).top(1).get();

        const licenseListItem = await sp.web.lists.getByTitle("LicenseList")
            .items.select("Id", "TransactionsPerForm")
            .filter(`LicenseType eq '${domainlistItem[0].LicenseType}'`)
            .top(1)
            .get();

        const recordslimit = licenseListItem?.[0]?.TransactionsPerForm ?? Infinity;

        return {
            recordscount: items.length,
            trnxRecLimitReached: items.length >= recordslimit,
        };

    } catch (err) {
        console.error("Error:", err);
        return { recordscount: 0, trnxRecLimitReached: false };
    }
};
export const fetchFormCount = async (domain: string) => {
    try {
        if (myDomain === domain) {
            return { formcount: 0, limitReached: false };
        }
        const domainItem = await sp.web.lists
            .getByTitle('DomainList')
            .items.select('Id', 'LicenseType')
            .filter(`Tenant eq '${domain}'`)
            .top(1)
            .get();

        const licenseItem = await sp.web.lists
            .getByTitle('LicenseList')
            .items.select('Id', 'FormCount')
            .filter(`LicenseType eq '${domainItem[0].LicenseType}'`)
            .top(1)
            .get();

        const formMasterItem = await sp.web.lists
            .getByTitle('FormMaster')
            .items.select('Id')
            .filter(`Domain eq '${domain}' and VisibilityFlag eq 1 and ParentAppCode eq null`)
            .getAll()

        const formlimit = licenseItem?.[0]?.FormCount ?? Infinity;

        return {
            formcount: formMasterItem.length,
            limitReached: formMasterItem.length >= formlimit,
        };

    }
    catch (err) {
        console.error("Error:", err);
        return { formcount: 0, limitReached: false };

    }
}
