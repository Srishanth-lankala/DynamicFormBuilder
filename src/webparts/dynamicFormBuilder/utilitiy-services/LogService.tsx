import { sp } from "@pnp/sp/presets/all";

export class LogService {
    private listName: string;

    constructor(listName: string = "WorkflowLogs") {
        this.listName = listName;
    }

    public async saveLog(logData: Record<string, any>): Promise<void> {
        try {
            await sp.web.lists.getByTitle(this.listName).items.add(logData);
            console.log("Log entry saved successfully.");
        } catch (error) {
            console.error("Error saving log entry:", error);
        }
    }

    public async getLogs(filters?: Record<string, any>): Promise<Record<string, any>[]> {
        try {
            let query = sp.web.lists.getByTitle(this.listName).items;

            // Apply filters dynamically if provided
            if (filters) {
                const filterStrings = Object.entries(filters)
                    .map(([key, value]) => `${key} eq '${value}'`);
                if (filterStrings.length > 0) {
                    query = query.filter(filterStrings.join(" and "));
                }
            }

            const logs = await query.get();
            return logs;
        } catch (error) {
            console.error("Error fetching logs:", error);
            return [];
        }
    }
}
