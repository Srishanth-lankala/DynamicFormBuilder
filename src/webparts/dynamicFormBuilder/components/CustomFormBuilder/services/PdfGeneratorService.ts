import { PDFDocument } from 'pdf-lib';
import { SharePointService } from './SharePointService';
import jsPDF from 'jspdf';

export interface PlacedField {
    id: string;
    type: string;
    value: string;
    x: number;
    y: number;
    width: number;
    height: number;
    fontSize: number;
    fieldId: string;
    pageIndex: number;
    name?: string;
    label?: string;
    originalId?: string;
    fontFamily?: string;
    tableData?: {
        columns: any[];
        rows: any[];
    };
    shape?: 'rectangle' | 'rounded' | 'pill';
}

export const PdfGeneratorService = {
    resolveReferenceValue(fieldId: string, value: string, formSchema: any[]) {
        if (!value || !formSchema) return value;
        const schemaField = formSchema.find(f => f.id === fieldId || f.field_name === fieldId || f.name === fieldId);
        if (!schemaField) return value;

        if (schemaField.options && Array.isArray(schemaField.options)) {
            const values = String(value).split(',').map(v => v.trim());
            const resolvedParts = values.map(val => {
                const option = schemaField.options.find((opt: any) =>
                    String(opt.key) === val || String(opt.value) === val || String(opt.text) === val
                );
                return option ? option.text : val;
            });
            return resolvedParts.join(', ');
        }
        return value;
    },

    async generatePDF(templateOrId: string | any, recordData: Record<string, any>, formSchema: any[]) {
        try {
            let template;
            if (typeof templateOrId === 'string') {
                template = await SharePointService.getTemplateById(templateOrId);
            } else {
                template = templateOrId;
                // CRITICAL: If fileData is missing (common with list-level fetch), 
                // re-fetch the full object using the ID to get the attachment blob.
                if (!template.fileData || (Array.isArray(template.fileData) && template.fileData.length === 0)) {
                    console.log("[PdfGenerator] File data missing, fetching full template from SharePoint...");
                    const fullTemplate = await SharePointService.getTemplateById(template.id);
                    if (fullTemplate) template = fullTemplate;
                }
            }
            if (!template) throw new Error('Template not found');

            console.log(`[PdfGenerator] Starting generation with ${Object.keys(recordData).length} data entries. Keys:`, Object.keys(recordData));

            const fields: PlacedField[] = template.mappings;
            const bgPages = Array.isArray(template.fileData) ? template.fileData : [template.fileData];
            const firstPageData = bgPages[0];

            let pdfDocInstance: any = null;
            const isPdf = typeof firstPageData === 'string' &&
                (firstPageData.includes('application/pdf') ||
                    firstPageData.startsWith('JVBERi') ||
                    firstPageData.includes('base64,JVBERi'));

            const isImage = typeof firstPageData === 'string' &&
                (firstPageData.startsWith('data:image/') ||
                    firstPageData.includes('base64,') && !isPdf);

            if (isPdf) {
                let base64Content = firstPageData.includes(',') ? firstPageData.split(',')[1] : firstPageData;
                base64Content = base64Content.replace(/\s/g, '');
                const binaryData = atob(base64Content);
                const pdfBytes = new Uint8Array(binaryData.length);
                for (let i = 0; i < binaryData.length; i++) pdfBytes[i] = binaryData.charCodeAt(i);
                pdfDocInstance = await PDFDocument.load(pdfBytes);
            }

            let pdfPointDims = { width: 595.28, height: 841.89 }; // A4 default
            if (pdfDocInstance) {
                const firstPage = pdfDocInstance.getPage(0);
                pdfPointDims = { width: firstPage.getWidth(), height: firstPage.getHeight() };
            } else if (isImage && bgPages.length > 0 && bgPages[0]) {
                const imgSrc = bgPages[0];
                console.log(`[PdfGenerator] Loading background image. Length: ${imgSrc.length}, Start: ${imgSrc.substring(0, 30)}...`);
                await new Promise<void>((resolve, reject) => {
                    const img = new Image();
                    img.onload = () => {
                        pdfPointDims = { width: img.width, height: img.height };
                        resolve();
                    };
                    img.onerror = () => reject(new Error(`Failed to load background image. Source length: ${imgSrc.length}`));
                    img.src = imgSrc;
                });
            } else {
                console.warn("[PdfGenerator] No background pages found in template mapping.");
            }

            // 1. Identify mode and initialize main document
            let mainDoc = pdfDocInstance;
            if (!mainDoc) {
                mainDoc = await PDFDocument.create();
            }

            const maxFieldIndex = fields.reduce((max, f) => Math.max(max, f.pageIndex || 0), 0);
            const templatePageCount = pdfDocInstance ? pdfDocInstance.getPageCount() : bgPages.length;
            const totalPages = Math.max(templatePageCount, maxFieldIndex + 1);

            console.log(`[PdfGenerator] Preparing PDF (${pdfDocInstance ? 'PDF' : 'Image'} Mode): ${totalPages} pages required (${templatePageCount} template pages).`);

            for (let i = 0; i < totalPages; i++) {
                let pageWidth = pdfPointDims.width;
                let pageHeight = pdfPointDims.height;

                let templatePage;
                if (pdfDocInstance) {
                    if (i < pdfDocInstance.getPageCount()) {
                        templatePage = pdfDocInstance.getPage(i);
                        pageWidth = templatePage.getWidth();
                        pageHeight = templatePage.getHeight();
                    } else {
                        console.log(`[PdfGenerator] Appending blank template page ${i + 1}`);
                        templatePage = pdfDocInstance.addPage([pageWidth, pageHeight]);
                    }
                } else {
                    // Image Mode: Add a new page for each iteration
                    templatePage = mainDoc.addPage([pageWidth, pageHeight]);
                }

                // 1. Create a FRESH jsPDF instance for JUST THIS PAGE
                const pageOverlay = new jsPDF({
                    orientation: pageWidth > pageHeight ? 'landscape' : 'portrait',
                    unit: 'pt',
                    format: [pageWidth, pageHeight]
                });

                // 2. Draw background image if applicable (image-based templates)
                if (!pdfDocInstance && i < bgPages.length && bgPages[i]) {
                    pageOverlay.addImage(bgPages[i], 'JPEG', 0, 0, pageWidth, pageHeight, undefined, 'FAST');
                }

                // 3. Render Fields for this page
                const pageFields = fields.filter(f => Number(f.pageIndex || 0) === i);
                console.log(`[PdfGenerator] Page ${i + 1}: Rendering ${pageFields.length} fields out of ${fields.length} total.`);

                pageFields.forEach(field => {
                    const x = Number(field.x);
                    const y = Number(field.y);
                    const w = Number(field.width);
                    const h = Number(field.height);
                    const fontSize = Number(field.fontSize || 12);

                    if (field.type === 'Table' || field.type === 'Timesheet') {
                        let dynamicData = recordData[field.fieldId];
                        if (dynamicData === undefined && field.originalId) dynamicData = recordData[field.originalId];
                        if (dynamicData === undefined && field.name) dynamicData = recordData[field.name];

                        // Fuzzy Lookup for Tables (more robust)
                        if (dynamicData === undefined) {
                            const possibleIds = [field.fieldId, field.originalId, field.name].filter(Boolean) as string[];
                            const matchingKey = Object.keys(recordData).find(key =>
                                possibleIds.some(id => key.endsWith(id) || id.endsWith(key))
                            );
                            if (matchingKey) dynamicData = recordData[matchingKey];
                        }

                        const dataToRender = dynamicData || field.tableData;

                        if (dataToRender) {
                            const cols = dataToRender.columns;
                            const rows = dataToRender.rows;
                            const colWidth = w / cols.length;
                            const rowHeight = fontSize + 6;

                            pageOverlay.setFillColor(243, 244, 246);
                            pageOverlay.rect(x, y, w, rowHeight, 'F');
                            pageOverlay.setDrawColor(0);
                            pageOverlay.rect(x, y, w, rowHeight, 'S');
                            pageOverlay.setFontSize(fontSize);
                            pageOverlay.setFont('helvetica', 'bold');
                            pageOverlay.setTextColor(0);

                            cols.forEach((col: any, idx: number) => {
                                let headerText = col.headername;
                                // Truncate header if it's too wide for the column
                                while (pageOverlay.getTextWidth(headerText + '...') > (colWidth - 4) && headerText.length > 0) {
                                    headerText = headerText.slice(0, -1);
                                }
                                if (headerText !== col.headername) headerText += '...';

                                pageOverlay.text(headerText, x + (idx * colWidth) + 2, y + (rowHeight / 2), { baseline: 'middle' });
                                if (idx > 0) pageOverlay.line(x + (idx * colWidth), y, x + (idx * colWidth), y + rowHeight);
                            });

                            pageOverlay.setFont('helvetica', 'normal');
                            let currentY = y + rowHeight;
                            for (const row of rows) {
                                let maxLineCount = 1;
                                const rowData = cols.map((col: any) => {
                                    const rawText = String(row[col.headername] || '');
                                    const wrappedText = pageOverlay.splitTextToSize(rawText, colWidth - 4);
                                    if (wrappedText.length > maxLineCount) maxLineCount = wrappedText.length;
                                    return wrappedText;
                                });

                                const requestedRowHeight = Math.max(rowHeight, (maxLineCount * fontSize) + 6);
                                const remainingHeight = (y + h) - currentY;

                                if (remainingHeight <= 0) break;

                                // Determine actual height (capped by available space)
                                const actualRowHeight = Math.min(requestedRowHeight, remainingHeight);

                                // Draw row box
                                pageOverlay.rect(x, currentY, w, actualRowHeight, 'S');

                                cols.forEach((_: any, c: number) => {
                                    const cellX = x + (c * colWidth);
                                    if (c > 0) pageOverlay.line(cellX, currentY, cellX, currentY + actualRowHeight);

                                    const originalLines = rowData[c];
                                    const cellLineHeight = fontSize * 1.15;
                                    const cellMaxLines = Math.floor((actualRowHeight - 4) / cellLineHeight);

                                    let linesToRender = originalLines;
                                    if (linesToRender.length > cellMaxLines && cellMaxLines > 0) {
                                        linesToRender = linesToRender.slice(0, cellMaxLines);
                                        let lastLine = linesToRender[cellMaxLines - 1];
                                        while (pageOverlay.getTextWidth(lastLine + '...') > (colWidth - 4) && lastLine.length > 0) {
                                            lastLine = lastLine.slice(0, -1);
                                        }
                                        linesToRender[cellMaxLines - 1] = lastLine + '...';
                                    }

                                    const textBlockHeight = linesToRender.length * cellLineHeight;
                                    const textY = currentY + (actualRowHeight / 2) - (textBlockHeight / 2);

                                    pageOverlay.text(linesToRender, cellX + 2, Math.max(currentY + 2, textY), {
                                        baseline: 'top',
                                        lineHeightFactor: 1.15
                                    });
                                });

                                currentY += actualRowHeight;
                                if (currentY >= y + h) break; // Reached bottom boundary
                            }
                        }
                    } else {
                        // 1. RESOLVE VALUE (Extreme Resilient Lookup)
                        let val: any = undefined;
                        const possibleKeys = [field.fieldId, field.originalId, field.name].filter(Boolean) as string[];

                        for (const key of possibleKeys) {
                            if (recordData[key] !== undefined) {
                                val = recordData[key];
                                break;
                            }
                        }

                        if (val === undefined) {
                            const recordKeys = Object.keys(recordData);
                            for (const rKey of recordKeys) {
                                if (possibleKeys.some(pk => rKey.endsWith(pk) || pk.endsWith(rKey) || rKey.includes(pk) || pk.includes(rKey))) {
                                    val = recordData[rKey];
                                    break;
                                }
                            }
                        }

                        if (val === undefined && field.label) {
                            const cleanLabel = field.label.replace(/<[^>]*>/g, '').trim();
                            const foundKey = Object.keys(recordData).find(k => k.toLowerCase() === cleanLabel.toLowerCase());
                            if (foundKey) val = recordData[foundKey];
                        }

                        const resolvedVal = (val !== undefined && val !== null && String(val).trim() !== "" && String(val) !== "[object Object]")
                            ? String(val)
                            : "";

                        let displayVal = resolvedVal;
                        if (displayVal && ["Radio", "Checkbox", "Dropdown", "RadioButtons", "Checkboxes"].includes(field.type)) {
                            displayVal = PdfGeneratorService.resolveReferenceValue(field.fieldId, displayVal, formSchema);
                        }

                        // Debug Log
                        console.log(`[PdfGenerator] Field: "${field.label || field.fieldId}"`, {
                            resolved: resolvedVal,
                            display: displayVal,
                            pos: { x, y, w, h },
                            page: i
                        });

                        // 2. Clear Shape/Border (Removed as requested)

                        // 3. Font
                        const fontMap: Record<string, string> = {
                            'Arial': 'helvetica', 'Times New Roman': 'times', 'Courier New': 'courier',
                            'Verdana': 'helvetica', 'Georgia': 'times'
                        };
                        const fontName = fontMap[field.fontFamily || 'Arial'] || 'helvetica';
                        pageOverlay.setFontSize(fontSize);
                        pageOverlay.setFont(fontName, 'normal');
                        pageOverlay.setTextColor(0, 0, 0);

                        // 4. Wrap and Truncate
                        const wrappedText = pageOverlay.splitTextToSize(displayVal, w - 4);
                        const cellLineHeight = fontSize * 1.15;
                        const maxLinesAvailable = Math.max(1, Math.floor((h - 2) / cellLineHeight));

                        let linesToRender: string[] = Array.isArray(wrappedText) ? wrappedText : [wrappedText];
                        if (linesToRender.length > maxLinesAvailable) {
                            linesToRender = linesToRender.slice(0, maxLinesAvailable);
                            let lastLine = linesToRender[maxLinesAvailable - 1] || "";
                            if (lastLine.length > 0) {
                                while (pageOverlay.getTextWidth(lastLine + '...') > (w - 4) && lastLine.length > 1) {
                                    lastLine = lastLine.slice(0, -1);
                                }
                                linesToRender[maxLinesAvailable - 1] = lastLine + '...';
                            }
                        }

                        // 5. Render Centered Text
                        const textBlockHeight = linesToRender.length * cellLineHeight;
                        const textY = y + (h / 2) - (textBlockHeight / 2);

                        pageOverlay.text(linesToRender, x + 2, Math.max(y + 2, textY), {
                            baseline: 'top',
                            lineHeightFactor: 1.15
                        });
                    }
                });

                // 4. Embed this page's overlay into the main document
                const pageBytes = pageOverlay.output('arraybuffer');
                const singlePagePdf = await PDFDocument.load(pageBytes);
                const [embeddedPage] = await mainDoc.embedPdf(singlePagePdf);

                templatePage.drawPage(embeddedPage, {
                    x: 0,
                    y: 0,
                    width: templatePage.getWidth(),
                    height: templatePage.getHeight()
                });
            }

            // --- Generate Final Download ---
            const finalPdfBytes = await mainDoc.save();
            const blob = new Blob([finalPdfBytes as any], { type: 'application/pdf' });
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = `${template.name}_Export.pdf`;
            link.click();

            return true;
        } catch (error) {
            console.error('PDF Generation Service Error:', error);
            throw error;
        }
    }
};
