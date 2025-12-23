import { Web } from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/attachments";

// Constants
const TARGET_SITE_URL = "https://cloudangles.sharepoint.com/sites/Smartofficetst/SmartOfficeREDEV";
const LIST_NAME = "SMRT_PDF_Templates";

export interface TemplateRecord {
    id: string;
    name: string;
    fileData: string | string[]; // Base64 string(s) or PDF data
    mappings: any[];
    createdAt: string;
    hasAttachment?: boolean;
    AppName?: string;
    AppCode?: string;
}

export const SharePointService = {
    /**
     * Save a Template:
     * 1. Create List Item (Title, PDF_Template JSON, AppName, AppCode)
     */
    async saveTemplate(
        fileName: string,
        fileData: string | string[],
        mappings: any[],
        originalFile?: File,
        appName?: string,
        appCode?: string,
        templateName?: string
    ) {
        try {
            const targetWeb = Web(TARGET_SITE_URL);

            // Mapping data for the PDF_Template column
            const mappingData = {
                mappings: mappings,
                updatedAt: new Date().toISOString()
            };

            const mappingString = JSON.stringify(mappingData);

            // Item metadata
            const itemData: any = {
                Title: fileName,
                PDF_Template: mappingString,
                AppName: appName,
                AppCode: appCode,
                TemplateName: templateName
            };

            // If fileData is small enough or specifically requested, keep it in column. 
            // But we'll treat Attachment as primary.
            if (typeof fileData === 'string' && fileData.length < 32000) {
                itemData.FileData = fileData;
            }

            const itemAddResult = await targetWeb.lists.getByTitle(LIST_NAME).items.add(itemData);

            // Add Attachment
            try {
                if (originalFile) {
                    // Upload the actual file (PDF/Word)
                    await itemAddResult.item.attachmentFiles.add(fileName, originalFile);
                } else if (Array.isArray(fileData)) {
                    // Save all pages for multi-page image templates
                    await itemAddResult.item.update({ FileData: JSON.stringify(fileData) });

                    const firstPageMatch = fileData[0].match(/^data:(.*);base64,(.*)$/);
                    if (firstPageMatch) {
                        const content = firstPageMatch[2];
                        await itemAddResult.item.attachmentFiles.add(fileName, content);
                    }
                } else if (typeof fileData === 'string' && fileData.startsWith('data:')) {
                    const match = fileData.match(/^data:(.*);base64,(.*)$/);
                    if (match) {
                        const content = match[2];
                        await itemAddResult.item.attachmentFiles.add(fileName, content);
                    }
                }
            } catch (attachError) {
                console.warn("Failed to add attachment:", attachError);
            }

            return itemAddResult.item;
        } catch (error) {
            console.error("SharePointService.saveTemplate Error:", error);
            throw error;
        }
    },

    async getAllTemplates(): Promise<TemplateRecord[]> {
        try {
            debugger;
            const targetWeb = Web(TARGET_SITE_URL);
            const items = await targetWeb.lists.getByTitle(LIST_NAME)
                .items
                .select("Id", "Title", "PDF_Template", "Created", "AttachmentFiles", "AppName", "AppCode")
                .expand("AttachmentFiles")
                .orderBy("Created", false)
                .get();

            const templates: TemplateRecord[] = items.map((item: any) => {
                try {
                    let mappings: any[] = [];
                    if (item.PDF_Template) {
                        const parsedMapping = JSON.parse(item.PDF_Template);
                        mappings = parsedMapping.mappings || [];
                    }

                    let fileData: any = (item as any).FileData || '';
                    try {
                        if (fileData && (fileData.startsWith("[") || fileData.startsWith("{"))) {
                            fileData = JSON.parse(fileData);
                        }
                    } catch (e) { }

                    return {
                        id: `sp_${item.Id}`,
                        name: item.Title,
                        fileData: fileData || '',
                        mappings: mappings,
                        createdAt: item.Created,
                        hasAttachment: item.AttachmentFiles && item.AttachmentFiles.length > 0,
                        AppName: item.AppName,
                        AppCode: item.AppCode
                    } as TemplateRecord;
                } catch (e) {
                    console.warn("Failed to parse template for item", item.Id, e);
                    return null;
                }
            }).filter((t): t is TemplateRecord => t !== null);

            return templates;
        } catch (error) {
            console.error("SharePointService.getAllTemplates Error:", error);
            return [];
        }
    },

    async getTemplateById(templateId: string): Promise<TemplateRecord | null> {
        try {
            const idInt = parseInt(templateId.replace('sp_', ''));
            if (isNaN(idInt)) return null;

            const targetWeb = Web(TARGET_SITE_URL);
            const item = await targetWeb.lists.getByTitle(LIST_NAME)
                .items.getById(idInt)
                .select("Id", "Title", "PDF_Template", "Created", "AttachmentFiles", "AppName", "AppCode")
                .expand("AttachmentFiles")
                .get();

            if (!item) return null;

            let mappings: any[] = [];
            if (item.PDF_Template) {
                const parsedMapping = JSON.parse(item.PDF_Template);
                mappings = parsedMapping.mappings || [];
            }

            let fileData: any = (item as any).FileData || '';

            // If FileData column is empty, look into Attachments
            if (!fileData && item.AttachmentFiles && item.AttachmentFiles.length > 0) {
                console.log(`SharePointService: Fetching file from attachment: ${item.AttachmentFiles[0].FileName}`);
                const blob = await targetWeb.getFileByServerRelativeUrl(item.AttachmentFiles[0].ServerRelativeUrl).getBlob();
                fileData = await this.blobToBase64(blob);
            } else if (fileData) {
                try {
                    if (fileData.startsWith("[") || fileData.startsWith("{")) {
                        fileData = JSON.parse(fileData);
                    }
                } catch (e) { }
            }

            return {
                id: `sp_${item.Id}`,
                name: item.Title,
                fileData: fileData || '',
                mappings: mappings,
                createdAt: item.Created,
                AppName: item.AppName,
                AppCode: item.AppCode
            };
        } catch (error) {
            console.error("SharePointService.getTemplateById Error:", error);
            return null;
        }
    },

    blobToBase64(blob: Blob): Promise<string> {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onloadend = () => resolve(reader.result as string);
            reader.onerror = reject;
            reader.readAsDataURL(blob);
        });
    },

    async clearAllTemplates() {
        try {
            const targetWeb = Web(TARGET_SITE_URL);
            const list = targetWeb.lists.getByTitle(LIST_NAME);
            const items = await list.items.select("Id").get();

            // Batch delete (or loop)
            const batch = targetWeb.createBatch();
            items.forEach((item: any) => {
                void list.items.getById(item.Id).inBatch(batch).delete();
            });
            await batch.execute();
        } catch (error) {
            console.error("SharePointService.clearAllTemplates Error:", error);
            throw error;
        }
    }
};
