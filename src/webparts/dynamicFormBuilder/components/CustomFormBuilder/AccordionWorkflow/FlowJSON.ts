import { sp } from "@pnp/sp/presets/all";

export interface NodeItem {
    id: string;
    data: {
        name: string;
        status: string;
        role: string;
        modified: string;
        color: string;
    };
}

export async function FlowData(transactionItem: any, levelsWithRoles: { Level: string; Role: string; User: string }[], endflow: boolean): Promise<NodeItem[]> {
    const FlowJSON: NodeItem[] = [];

    debugger;
    const appCode = transactionItem.AppCode;
    const seqNo = transactionItem.SeqNo;
    const currentLevel = Number(transactionItem.Level);

    //Checking if CurrentAppCode exists in the WorkFlowProcessData then use it else use AppCode column data
    const childAppCode = transactionItem?.CurrentAppCode ?? appCode;

    const currentLevelRole: string = levelsWithRoles.find(item => Number(item.Level) === currentLevel)?.Role || "";
    // 1. WorkFlowLog - levels less than current level

    const WFLogData = await sp.web.lists
        .getByTitle("WorkFlowLog")
        .items.select("ID", "Author/Title", "Status", "Role", "Modified")
        .expand("Author")
        .filter(`AppCode eq '${childAppCode}' and SeqNo eq '${seqNo}'`)
        .orderBy("Created", true) // Ascending order by Created (oldest first)
        .get();

    WFLogData.forEach((item: any) => {
        FlowJSON.push({
            id: item.ID.toString(),
            data: {
                name: item.Author?.Title || "",
                status: item.Status,
                role: item.Role,
                modified: new Date(item.Modified).toLocaleString().slice(0, 18).replace("T", " "),
                color: (item.Status === "Completed" || item.Status === "In-Progress") ? "#28a745" : item.Status === "Rejected" ? "red" : "gray",
            },
        });
        // console.log("case 1");
    });


    if (transactionItem.Status !== "Completed" && transactionItem.Status !== "Rejected" && transactionItem.Status !== "Cancelled") {
       
            // console.log("case 2 , case 3")
        // 2. TransactionItem itself
        // console.log("transactionItem in flowJson >>>>>>>>>>>>>>>>>>", transactionItem);
        FlowJSON.push({
            id: transactionItem.ID?.toString(),
            data: {
                name:transactionItem.Status==="Returned"? transactionItem.Author.Title: transactionItem?.CurApprover?.Title || "Immediate Manager",
                status: transactionItem.Status,
                role: currentLevelRole,
                modified: "  ",
                color: "#ffc107", // Yellow for current transaction item
            },
        });

        if (!endflow) {
            // 3. MappingMaster - levels greater than current level
            const MMData = await sp.web.lists
                .getByTitle("MappingMaster")
                .items.select("ID", "Users_x002f_Groups/Title", "Role", "Level", "Modified")
                .expand("Users_x002f_Groups")
                .filter(`AppCode eq '${childAppCode}' and Level gt ${currentLevel}`)
                .orderBy("Level", true)
                .get();

            MMData.forEach((item: any) => {
                FlowJSON.push({
                    id: item.ID.toString(),
                    data: {
                        name: item.Users_x002f_Groups?.Title,
                        status: "",
                        role: item.Role,
                        modified: "  ",
                        color: "grey"
                    },
                });
            });
        }
    }
    return FlowJSON;
}