import React, { useState, useRef, useEffect } from 'react';
import { Upload, X, ZoomIn, ZoomOut } from 'lucide-react';
// import { Database } from 'lucide-react';
import FieldSidebar from './fieldsidebar';
import { type MappedField, mapSchemaToFields } from './dataMapping';
// import formSchema from './formjson.json';
import { SharePointService } from '../services/SharePointService';
import './InvoiceGenerator.css';
import * as pdfjsLib from 'pdfjs-dist/legacy/build/pdf';
import { Toast } from 'primereact/toast';
import pdfWorker from 'pdfjs-dist/legacy/build/pdf.worker.entry';
// import { jsPDF } from 'jspdf';
import { dataservice } from '../encryptionutil';
import { mySiteUrl } from '../ConfigURL/All_URLs';
import { Web } from '@pnp/sp/webs';
import TopNavBar from '../TopBar';



// Type definitions
interface Field {
    id: string;
    name?: string;
    type: string;
    value: string;
    label?: string; // Added from MappedField compatibility
    originalId?: string; // Added from MappedField compatibility
    tableData?: {
        columns: any[];
        rows: any[];
    };
    shape?: 'rectangle' | 'rounded' | 'pill';
    fontFamily?: string;
    widthChar?: number;
}

interface PlacedField extends Field {
    x: number;
    y: number;
    width: number;
    height: number;
    fontSize: number;
    fieldId: string;
    pageIndex: number;
    fontFamily: string;
    widthChar: number;
}

interface PdfPage {
    pageNum: number;
    dataUrl: string;
    width: number;
    height: number;
}





// Set worker source
pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker;

// Extend Window interface for jsPDF if needed locally, though we prefer module imports
declare global {
    interface Window {
        jspdf: any;
        // pdfjsLib: any; // validation removed as we use module import
    }
}

export default function TemplateMapper() {
    const toast = React.useRef<Toast>(null);
    const [formmasteritems, setFormmasteritems] = useState<any[]>([]);
    const [allFormMasterItems, setAllFormMasterItems] = useState<any[]>([]);
    const [uploadedFile, setUploadedFile] = useState<File | null>(null);
    const [fileContent, setFileContent] = useState<string | ArrayBuffer>('');
    const [fileType, setFileType] = useState<string>('');
    const [placedFields, setPlacedFields] = useState<PlacedField[]>([]);
    const [draggedField, setDraggedField] = useState<Field | null>(null);
    const [isDragging, setIsDragging] = useState<boolean>(false);
    const [zoom, setZoom] = useState<number>(100);
    const [pdfPages, setPdfPages] = useState<PdfPage[]>([]);
    const [selectedField, setSelectedField] = useState<string | null>(null);
    const [dataserviceobj] = useState<dataservice>(new dataservice());
    // Unified Sidebar State
    const [sidebarFields, setSidebarFields] = useState<MappedField[]>([]);
    const [selectedApp, setSelectedApp] = useState<string>('');
    const [showNameDialog, setShowNameDialog] = useState(false);
    const [templateName, setTemplateName] = useState('');
    const [, setisSameDomain] = useState<boolean>(false);
    const handleUserDomain = (UserDomain: string) => {
        setisSameDomain(true)
        //console.log("Logged-in USER DOMAIN IS   ----------------:", UserDomain);

    };
    // Dragging state refs
    const dragState = useRef<{
        isMoving: boolean;
        fieldId: string | null;
        startX: number;
        startY: number;
        initialFieldX: number;
        initialFieldY: number;
        pageIndex: number;
    }>({
        isMoving: false,
        fieldId: null,
        startX: 0,
        startY: 0,
        initialFieldX: 0,
        initialFieldY: 0,
        pageIndex: -1 // Added pageIndex
    });

    // Resizing state refs
    const resizeState = useRef<{
        isResizing: boolean;
        fieldId: string | null;
        direction: 'n' | 's' | 'e' | 'w' | 'ne' | 'nw' | 'se' | 'sw' | null;
        startX: number;
        startY: number;
        initialX: number;
        initialY: number;
        initialWidth: number;
        initialHeight: number;
        pageIndex: number; // Added pageIndex
    }>({
        isResizing: false,
        fieldId: null,
        direction: null,
        startX: 0,
        startY: 0,
        initialX: 0,
        initialY: 0,
        initialWidth: 0,
        initialHeight: 0,
        pageIndex: -1
    });

    const fileInputRef = useRef<HTMLInputElement>(null);
    const fetchFormMasterItems = async () => {
        const targetWeb = Web(mySiteUrl);
        const RAWitems = await targetWeb.lists.getByTitle('FormMaster')
            .items
            .select("*", "Id", "IsParentForm", "FormJSON", "ParentAppCode", "VisibilityFlag")
            .get();

        const decryptedItems = RAWitems?.map((item: any) => {
            return {
                ...item,
                FormJSON: item.FormJSON ? dataserviceobj.decryptjson(item.FormJSON) : '[]'
            };
        }) || [];

        console.log('Decrypted Form Master Items', decryptedItems);
        setAllFormMasterItems(decryptedItems);

        const filteredItems = decryptedItems.filter((item: any) => !item.ParentAppCode && item.VisibilityFlag === true);
        setFormmasteritems(filteredItems);
        return filteredItems;
    }

    const fetchdata = (items: any[], appCode: string) => {
        if (!appCode || !items.length) return;

        const selectedItem = items.find(item => item.AppCode === appCode);
        if (!selectedItem) return;

        let mergedFields: any[] = [];

        if (selectedItem.IsParentForm) {
            // Find all forms whose ParentAppCode is the selected AppCode
            const children = allFormMasterItems.filter(item => item.ParentAppCode === appCode);

            console.log(`Merging ${children.length} child forms for parent ${appCode}`);

            children.forEach(child => {
                try {
                    const schema = typeof child.FormJSON === 'string' ? JSON.parse(child.FormJSON) : child.FormJSON;
                    if (Array.isArray(schema)) {
                        mergedFields = [...mergedFields, ...schema];
                    }
                } catch (error) {
                    console.error("Error parsing child form schema:", error);
                    // Fallback if it's already an array
                    if (Array.isArray(child.FormJSON)) {
                        mergedFields = [...mergedFields, ...child.FormJSON];
                    }
                }
            });

            // Also check if parent itself has fields
            try {
                const parentSchema = typeof selectedItem.FormJSON === 'string' ? JSON.parse(selectedItem.FormJSON) : selectedItem.FormJSON;
                if (Array.isArray(parentSchema)) {
                    mergedFields = [...mergedFields, ...parentSchema];
                }
            } catch (e) { }

        } else {
            // Normal form logic
            let formSchema1 = selectedItem.FormJSON;
            if (formSchema1) {
                try {
                    mergedFields = typeof formSchema1 === 'string' ? JSON.parse(formSchema1) : formSchema1;
                } catch (error) {
                    console.error("Error parsing form schema:", error);
                    if (Array.isArray(formSchema1)) {
                        mergedFields = formSchema1;
                    }
                }
            }
        }

        if (mergedFields && (Array.isArray(mergedFields) ? mergedFields.length > 0 : true)) {
            const sidebarfeilds = mapSchemaToFields(mergedFields as any);
            setSidebarFields(sidebarfeilds);
        }
    };
    useEffect(() => {
        const init = async () => {
            try {
                await fetchFormMasterItems();
                // Removed auto-selection of first item to force user to choose
            } catch (error) {
                console.log(error);
            }
        };
        void init();
    }, []);

    useEffect(() => {
        if (selectedApp && formmasteritems.length > 0) {
            fetchdata(formmasteritems, selectedApp);
        }
    }, [selectedApp, formmasteritems, allFormMasterItems]);



    const renderPDF = async (fileData: ArrayBuffer) => {
        try {
            const loadingTask = pdfjsLib.getDocument({ data: fileData });
            const pdf = await loadingTask.promise;

            const pages: PdfPage[] = [];
            for (let i = 1; i <= pdf.numPages; i++) {
                const page = await pdf.getPage(i);
                const viewport = page.getViewport({ scale: 1.0 });

                const canvas = document.createElement('canvas');
                const context = canvas.getContext('2d');
                canvas.height = viewport.height;
                canvas.width = viewport.width;

                if (context) {
                    await page.render({
                        canvasContext: context,
                        viewport: viewport
                    }).promise;

                    pages.push({
                        pageNum: i,
                        dataUrl: canvas.toDataURL('image/jpeg', 0.7),
                        width: viewport.width,
                        height: viewport.height
                    });
                }
            }
            setPdfPages(pages);
        } catch (error) {
            console.error('Error rendering PDF:', error);
            alert('Error rendering PDF. Please try another file.');
        }
    };

    const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
        const file = e.target.files?.[0];
        if (!file) return;

        setUploadedFile(file);
        // Load fields from dynamic form schema instead of default JSON
        setPlacedFields([]);
        setPdfPages([]);

        const fileExt = file.name.split('.').pop()?.toLowerCase() || '';
        setFileType(fileExt);

        const reader = new FileReader();

        if (fileExt === 'pdf') {
            reader.onload = async (e) => {
                const arrayBuffer = e.target?.result as ArrayBuffer;
                setFileContent(arrayBuffer);
                await renderPDF(arrayBuffer);
            };
            reader.readAsArrayBuffer(file);
        } else if (['jpg', 'jpeg', 'png', 'gif', 'bmp', 'webp'].includes(fileExt)) {
            reader.onload = (e) => {
                const img = new Image();
                img.onload = () => {
                    setPdfPages([{
                        pageNum: 1,
                        dataUrl: e.target?.result as string,
                        width: img.width,
                        height: img.height
                    }]);
                };
                img.src = e.target?.result as string;
                setFileContent(e.target?.result as string);
            };
            reader.readAsDataURL(file);
        } else if (['txt', 'html', 'htm', 'css', 'js', 'json', 'xml'].includes(fileExt)) {
            reader.onload = (e) => {
                setFileContent(e.target?.result as string);
            };
            reader.readAsText(file);
        } else if (['doc', 'docx'].includes(fileExt)) {
            setFileContent('Word documents can be used as templates. Please map your fields relative to the document layout.');
            // For now, using a placeholder for Word documents in the visual mapper
            setPdfPages([{
                pageNum: 1,
                dataUrl: '', // No image for word unless we convert it
                width: 800,
                height: 1100
            }]);
        } else {
            alert('Unsupported file format. Please upload a PDF or Word document.');
            return;
        }
    };

    const handleDragStart = (field: Field) => {
        setDraggedField(field);
        setIsDragging(true);
    };

    const handleDragEnd = () => {
        setDraggedField(null);
        setIsDragging(false);
    };

    const handleDrop = (e: React.DragEvent<HTMLDivElement>, pageIndex: number) => {
        e.preventDefault();
        if (!draggedField || !uploadedFile) return;
        setIsDragging(false);

        const page = pdfPages[pageIndex];
        if (!page) return;

        const imgElement = e.currentTarget.querySelector('img');
        if (!imgElement) return;
        const rect = imgElement.getBoundingClientRect();

        const scaleX = page.width / rect.width;
        const scaleY = page.height / rect.height;

        let width = draggedField.type === 'Table' || draggedField.type === 'Timesheet' ? 300 : 150;
        let height = 22;

        // Calculate PDF-space mouse coordinates
        let mouseX = (e.clientX - rect.left) * scaleX;
        let mouseY = (e.clientY - rect.top) * scaleY;

        // Center the field on the mouse cursor
        let x = mouseX - (width / 2);
        let y = mouseY - (height / 2);

        // Constrain to page bounds
        x = Math.max(0, Math.min(x, page.width - width));
        y = Math.max(0, Math.min(y, page.height - height));

        const newField: PlacedField = {
            ...draggedField,
            x,
            y,
            width,
            height,
            fontSize: 12,
            fontFamily: 'Arial',
            widthChar: 50, // Default character limit
            fieldId: draggedField.id,
            pageIndex,
        };

        setPlacedFields(prev => [...prev, newField]);
        setDraggedField(null);
        setSelectedField(newField.id);
    };

    const handleDragOver = (e: React.DragEvent<HTMLDivElement>) => {
        e.preventDefault();
    };

    const removeField = (id: string) => {
        setPlacedFields(prev => prev.filter(f => f.id !== id));
        if (selectedField === id) setSelectedField(null);
    };

    const updateField = (id: string, updates: Partial<PlacedField>) => {
        setPlacedFields(prev => prev.map(f =>
            f.id === id ? { ...f, ...updates } : f
        ));
    };

    const startMoving = (e: React.MouseEvent, fieldId: string, field: PlacedField) => {
        e.stopPropagation();
        setSelectedField(fieldId);

        document.body.style.cursor = 'move';

        dragState.current = {
            isMoving: true,
            fieldId,
            startX: e.clientX,
            startY: e.clientY,
            initialFieldX: field.x,
            initialFieldY: field.y,
            pageIndex: field.pageIndex
        };
        document.addEventListener('mousemove', handleMouseMove);
        document.addEventListener('mouseup', handleMouseUp);
    };

    const startResizing = (e: React.MouseEvent, fieldId: string, field: PlacedField, direction: 'n' | 's' | 'e' | 'w' | 'ne' | 'nw' | 'se' | 'sw') => {
        e.stopPropagation();
        e.preventDefault();
        setSelectedField(fieldId);

        const cursorMap: Record<string, string> = {
            'n': 'ns-resize', 's': 'ns-resize',
            'e': 'ew-resize', 'w': 'ew-resize',
            'ne': 'nesw-resize', 'sw': 'nesw-resize',
            'nw': 'nwse-resize', 'se': 'nwse-resize'
        };
        document.body.style.cursor = cursorMap[direction] || 'default';

        resizeState.current = {
            isResizing: true,
            fieldId,
            direction,
            startX: e.clientX,
            startY: e.clientY,
            initialX: field.x,
            initialY: field.y,
            initialWidth: field.width,
            initialHeight: field.height,
            pageIndex: field.pageIndex
        };
        document.addEventListener('mousemove', handleMouseMove);
        document.addEventListener('mouseup', handleMouseUp);
    };

    function handleMouseMove(e: MouseEvent) {
        if (dragState.current.isMoving) {
            if (dragState.current.fieldId) {
                // Find which page container we are over
                const allPageContainers = document.querySelectorAll('[id^="page-container-"]');
                let targetPageIdx = dragState.current.pageIndex;
                let targetRect: DOMRect | null = null;

                for (let i = 0; i < allPageContainers.length; i++) {
                    const r = allPageContainers[i].getBoundingClientRect();
                    if (e.clientX >= r.left && e.clientX <= r.right && e.clientY >= r.top && e.clientY <= r.bottom) {
                        targetPageIdx = i;
                        targetRect = allPageContainers[i].querySelector('img')?.getBoundingClientRect() || r;
                        break;
                    }
                }

                if (targetRect) {
                    const page = pdfPages[targetPageIdx];
                    const scaleX = page.width / targetRect.width;
                    const scaleY = page.height / targetRect.height;

                    const localX = (e.clientX - targetRect.left) * scaleX;
                    const localY = (e.clientY - targetRect.top) * scaleY;

                    const field = placedFields.find(f => f.id === dragState.current.fieldId);
                    if (field) {
                        const newX = Math.max(0, Math.min(localX - (field.width / 2), page.width - field.width));
                        const newY = Math.max(0, Math.min(localY - (field.height / 2), page.height - field.height));
                        updateField(dragState.current.fieldId, {
                            x: newX,
                            y: newY,
                            pageIndex: targetPageIdx
                        });
                    }
                }
            }
        } else if (resizeState.current.isResizing) {
            const page = pdfPages[resizeState.current.pageIndex];
            if (!page) return;

            const imgElement = document.querySelector(`#page-container-${resizeState.current.pageIndex} img`);
            if (!imgElement) return;
            const rect = imgElement.getBoundingClientRect();

            const scaleX = page.width / rect.width;
            const scaleY = page.height / rect.height;

            const deltaX = (e.clientX - resizeState.current.startX) * scaleX;
            const deltaY = (e.clientY - resizeState.current.startY) * scaleY;

            let newX = resizeState.current.initialX;
            let newY = resizeState.current.initialY;
            let newWidth = resizeState.current.initialWidth;
            let newHeight = resizeState.current.initialHeight;

            switch (resizeState.current.direction) {
                case 'n':
                    newY = resizeState.current.initialY + deltaY;
                    newHeight = resizeState.current.initialHeight - deltaY;
                    break;
                case 's':
                    newHeight = resizeState.current.initialHeight + deltaY;
                    break;
                case 'e':
                    newWidth = resizeState.current.initialWidth + deltaX;
                    break;
                case 'w':
                    newX = resizeState.current.initialX + deltaX;
                    newWidth = resizeState.current.initialWidth - deltaX;
                    break;
                case 'ne':
                    newY = resizeState.current.initialY + deltaY;
                    newHeight = resizeState.current.initialHeight - deltaY;
                    newWidth = resizeState.current.initialWidth + deltaX;
                    break;
                case 'nw':
                    newY = resizeState.current.initialY + deltaY;
                    newHeight = resizeState.current.initialHeight - deltaY;
                    newX = resizeState.current.initialX + deltaX;
                    newWidth = resizeState.current.initialWidth - deltaX;
                    break;
                case 'se':
                    newHeight = resizeState.current.initialHeight + deltaY;
                    newWidth = resizeState.current.initialWidth + deltaX;
                    break;
                case 'sw':
                    newHeight = resizeState.current.initialHeight + deltaY;
                    newX = resizeState.current.initialX + deltaX;
                    newWidth = resizeState.current.initialWidth - deltaX;
                    break;
            }

            // Simple constraint to min size
            if (newWidth < 20) newWidth = 20;
            if (newHeight < 20) newHeight = 20;

            // Simple constraint to page bounds
            newX = Math.max(0, Math.min(newX, page.width - newWidth));
            newY = Math.max(0, Math.min(newY, page.height - newHeight));
            newWidth = Math.min(newWidth, page.width - newX);
            newHeight = Math.min(newHeight, page.height - newY);

            if (resizeState.current.fieldId) {
                updateField(resizeState.current.fieldId, { x: newX, y: newY, width: newWidth, height: newHeight });
            }
        }
    }

    function handleMouseUp() {
        if (dragState.current.isMoving) {
            dragState.current.isMoving = false;
            dragState.current.fieldId = null;
        }
        if (resizeState.current.isResizing) {
            resizeState.current.isResizing = false;
            resizeState.current.fieldId = null;
            resizeState.current.direction = null;
        }
        document.body.style.cursor = 'default';
    }

    useEffect(() => {
        window.addEventListener('mousemove', handleMouseMove);
        window.addEventListener('mouseup', handleMouseUp);
        return () => {
            window.removeEventListener('mousemove', handleMouseMove);
            window.removeEventListener('mouseup', handleMouseUp);
        };
    }, []);

    const clearTemplate = () => {
        setUploadedFile(null);
        setFileContent('');
        setFileType('');
        setPlacedFields([]);
        setPdfPages([]);
        if (fileInputRef.current) {
            fileInputRef.current.value = '';
        }
    };

    const showSuccessNotification = (message: string) => {
        const notification = document.createElement('div');
        notification.className = 'fixed top-20 right-6 bg-green-500 text-white px-6 py-3 rounded-lg shadow-lg z-50';
        notification.textContent = message;
        document.body.appendChild(notification);
        setTimeout(() => notification.remove(), 2500);
    };

    const saveToSystem = async () => {
        if (!uploadedFile) {
            alert('Please upload a template first');
            return;
        }
        setShowNameDialog(true);
    };

    const handleConfirmSave = async () => {
        if (!templateName.trim()) {
            alert('Please enter a template name');
            return;
        }

        const selectedAppItem = formmasteritems.find(item => item.AppCode === selectedApp);
        const appPrefix = selectedAppItem ? `${selectedAppItem.AppName}_${selectedAppItem.AppCode}_` : '';
        const finalFileName = `${appPrefix}${templateName}${fileType ? `.${fileType}` : ''}`;

        try {
            let fileData: string | string[] = '';

            if (pdfPages.length > 0) {
                fileData = pdfPages.map(page => page.dataUrl);
            } else if (typeof fileContent === 'string') {
                fileData = fileContent;
            } else {
                const blob = new Blob([fileContent]);
                const reader = new FileReader();
                fileData = await new Promise<string>((resolve) => {
                    reader.onloadend = () => resolve(reader.result as string);
                    reader.readAsDataURL(blob);
                });
            }

            console.log("[TemplateMapper] Fields to save:", JSON.stringify(placedFields, null, 2));

            await SharePointService.saveTemplate(
                finalFileName,
                fileData,
                placedFields,
                uploadedFile || undefined,
                selectedAppItem?.AppName,
                selectedApp,
                templateName
            );

            setShowNameDialog(false);
            setTemplateName('');
            showSuccessNotification(`Template Saved as: ${finalFileName}`);
        } catch (error) {
            console.error('Error saving to system:', error);
            alert('Failed to save template to system');
        }
    };

    const renderFileContent = () => {
        if (pdfPages.length > 0) {
            return (
                <div className="overflow-visible" >
                    {pdfPages.map((page, index) => (
                        <div
                            key={index}
                            id={`page-container-${index}`}
                            className={`relative overflow-visible shadow-lg bg-white mb-8 mx-auto ${isDragging ? 'ring-4 ring-blue-400 ring-opacity-50' : ''}`}
                            style={{
                                width: `${(page.width * zoom) / 100}px`,
                                height: `${(page.height * zoom) / 100}px`,
                                padding: 0,
                                lineHeight: 0,
                                fontSize: 0,
                                position: 'relative' // Ensure fields anchor to this container
                            }}
                            onDragOver={handleDragOver}
                            onDrop={(e) => handleDrop(e, index)}
                        >
                            <img
                                src={page.dataUrl}
                                alt={`Page ${page.pageNum}`}
                                style={{
                                    width: '100%',
                                    height: '100%',
                                    display: 'block',
                                    verticalAlign: 'top',
                                    margin: 0,
                                    padding: 0
                                }}
                                draggable={false}
                            />

                            {/* Fields for this page */}
                            {placedFields.filter(f => f.pageIndex === index).map(field => (
                                <div
                                    key={field.id}
                                    className={`placed-field group ${selectedField === field.id ? 'selected' : ''}`}
                                    style={{
                                        position: 'absolute',
                                        left: `${(field.x * zoom) / 100}px`,
                                        top: `${(field.y * zoom) / 100}px`,
                                        width: `${(field.width * zoom) / 100}px`,
                                        height: `${(field.height * zoom) / 100}px`
                                    }}
                                    onMouseDown={(e) => startMoving(e, field.id, field)}
                                    onClick={(e) => { e.stopPropagation(); setSelectedField(field.id); }}
                                >
                                    <div className={`placed-field-content ${selectedField === field.id ? 'selected' : ''} 
                    ${field.shape === 'pill' ? 'rounded-full' : field.shape === 'rectangle' ? 'rounded-none' : 'rounded-lg'}
                  `} style={{ overflow: 'hidden', height: '100%', boxSizing: 'border-box' }}>
                                        <div className="placed-field-value" style={{ height: (field.type === 'Table' || field.type === 'Timesheet') ? '100%' : '22px', display: 'flex', alignItems: (field.type === 'Table' || field.type === 'Timesheet') ? 'flex-start' : 'center', overflow: 'hidden', boxSizing: 'border-box' }}>
                                            {(field.type === 'Table' || field.type === 'Timesheet') && field.tableData ? (
                                                <div style={{ width: '100%', height: '100%', overflow: 'hidden', backgroundColor: 'rgba(255, 255, 255, 0.9)', padding: '2px', boxSizing: 'border-box' }}>
                                                    <table style={{ width: '100%', borderCollapse: 'collapse', border: '1px solid #999', tableLayout: 'fixed', boxSizing: 'border-box' }}>
                                                        <thead>
                                                            <tr style={{ backgroundColor: '#f1f5f9' }}>
                                                                {field.tableData.columns.map((col: any, i: number) => (
                                                                    <th key={i} style={{ border: '1px solid #cbd5e1', padding: '2px 4px', textAlign: 'left', fontWeight: 'bold', fontSize: `${(field.fontSize * 0.8 * zoom) / 100}px`, width: `${100 / (field.tableData?.columns.length || 1)}%`, boxSizing: 'border-box' }}>
                                                                        <div style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{col.headername}</div>
                                                                    </th>
                                                                ))}
                                                            </tr>
                                                        </thead>
                                                        <tbody>
                                                            {field.tableData.rows.map((row: any, i: number) => (
                                                                <tr key={i}>
                                                                    {field.tableData!.columns.map((col: any, j: number) => (
                                                                        <td key={j} style={{ border: '1px solid #cbd5e1', padding: '2px 4px', fontSize: `${(field.fontSize * 0.8 * zoom) / 100}px`, boxSizing: 'border-box' }}>
                                                                            <div style={{
                                                                                display: '-webkit-box',
                                                                                WebkitBoxOrient: 'vertical',
                                                                                WebkitLineClamp: 2,
                                                                                overflow: 'hidden',
                                                                                lineHeight: '1.2'
                                                                            }}>
                                                                                {row[col.headername] || ''}
                                                                            </div>
                                                                        </td>
                                                                    ))}
                                                                </tr>
                                                            ))}
                                                        </tbody>
                                                    </table>
                                                </div>
                                            ) : (
                                                <div className="leading-tight break-words p-1 h-full w-full overflow-hidden"
                                                    style={{
                                                        fontSize: `${(field.fontSize * zoom) / 100}px`,
                                                        fontFamily: field.fontFamily || 'Arial',
                                                        fontWeight: 'normal',
                                                        fontStyle: 'normal',
                                                        color: '#000000',
                                                        display: '-webkit-box',
                                                        WebkitBoxOrient: 'vertical',
                                                        WebkitLineClamp: Math.max(1, Math.floor(field.height / ((field.fontSize || 12) * 1.15))),
                                                        lineHeight: '1.15'
                                                    }}>
                                                    {field.value || field.label || "Sample Text"}
                                                </div>
                                            )}
                                        </div>
                                    </div>

                                    {/* Controls: Visible on Selection OR Hover */}
                                    <div className={`absolute pointer-events-none ${selectedField === field.id ? 'opacity-100' : 'opacity-0 group-hover:opacity-100'} transition-opacity duration-200`} style={{ top: 0, left: 0, width: '100%', height: '100%' }}>
                                        {/* Resize Handles (8-point) */}
                                        {/* SE (Bottom-Right) */}
                                        <div style={{ position: 'absolute', bottom: '-6px', right: '-6px', width: '12px', height: '12px', zIndex: 1001 }}
                                            className="bg-transparent pointer-events-auto cursor-nwse-resize flex items-center justify-center"
                                            title="Resize"
                                            onMouseDown={(e) => startResizing(e, field.id, field, 'se')}>
                                            <div style={{
                                                width: '10px',
                                                height: '10px',

                                                backgroundColor: '#000',
                                                border: '2px solid #fff',
                                                borderRadius: '50%',
                                                boxShadow: '0 0 3px rgba(0,0,0,0.5)'
                                            }} />
                                        </div>
                                        {/* Delete Button */}
                                        <button
                                            onClick={(e) => { e.stopPropagation(); removeField(field.id); }}
                                            style={{ position: 'absolute', top: '-18px', right: '1px', zIndex: 1000, height: '12px', width: '10px', background: '#ff0000', border: '0px' }}
                                            className="w-7 h-7 bg-red-500 text-white rounded-full flex items-center justify-center hover:bg-red-600 shadow-lg pointer-events-auto"
                                            title="Remove Field"
                                        >
                                            <X style={{ marginLeft: '-12px', paddingBottom: '12px' }} />
                                        </button>

                                    </div>

                                </div>
                            ))
                            }
                        </div >
                    ))
                    }
                </div >
            );
        } else if (fileContent && typeof fileContent === 'string') {
            return (
                <div className="bg-white p-12 shadow-lg max-w-4xl mx-auto min-h-[800px] relative">
                    <pre className="whitespace-pre-wrap text-sm text-gray-700 font-mono leading-relaxed">
                        {fileContent}
                    </pre>
                    {/* Fields for text document */}
                    {placedFields.map(field => (
                        <div
                            key={field.id}
                            className="absolute bg-yellow-200 border border-yellow-500 px-2 rounded opacity-80"
                            style={{ left: field.x, top: field.y }}
                        >
                            {field.value}
                        </div>
                    ))}
                </div>
            );
        }

        return (
            <div className="p-12 text-center text-gray-500">
                <p>Please upload a PDF or Word document to use the visual mapper.</p>
            </div>
        );
    };

    return (
        <div style={{ display: 'flex', flexDirection: 'column', height: '100vh', width: '100%', overflow: 'hidden' }}>
            <TopNavBar onUserDomainRetrieved={handleUserDomain} />
            <div style={{ display: 'flex', flex: 1, minHeight: 0 }}>
                <div style={{ display: 'flex', flex: 1, minHeight: 0 }}>
                    {/* <SideBar activeMenu" /> */}
                    <div className="Table" style={{ display: 'flex', flexDirection: 'column', flex: 1, minHeight: 0, width: '100%' }}>
                        <div className='d-flex justify-content-between align-items-center' style={{ margin: '0', flexShrink: 0, width: '100%', background: '#0b1120', color: 'white', borderBottom: '1px solid #1e293b' }}>
                            {/* Left Section: Title and Main Controls */}
                            <div style={{ flex: 1, display: 'flex', alignItems: 'center', padding: '12px 20px', gap: '20px' }}>
                                <h1 className="white-title" style={{ margin: 0, fontSize: '1.25rem', fontWeight: 500, letterSpacing: '0.5px' }}>
                                    Template Mapper
                                </h1>
                                <select
                                    className="child-dropdown"
                                    style={{ backgroundColor: '#1e293b', color: 'white', border: '1px solid #334155', borderRadius: '6px', padding: '6px 12px', fontSize: '14px' }}
                                    value={selectedApp}
                                    onChange={(e) => { console.log(e.target.value); setSelectedApp(e.target.value) }}>
                                    <option value="" disabled>Select a form</option>
                                    {formmasteritems.map((data: { AppName: any; AppCode: any }, index: number) => (
                                        <option key={index} value={data.AppCode} style={{ background: '#0b1120' }}>
                                            {data.AppName}
                                        </option>
                                    ))}
                                </select>
                                {uploadedFile && (
                                    <div className="zoom-controls" style={{ display: 'flex', alignItems: 'center', gap: '12px', background: '#1e293b', padding: '4px 12px', borderRadius: '6px' }}>
                                        <button onClick={() => setZoom(z => Math.max(50, z - 10))} className="zoom-btn" style={{ color: '#94a3b8', border: 'none', background: 'transparent', cursor: 'pointer' }}><ZoomOut size={18} /></button>
                                        <span className="zoom-value" style={{ fontSize: '14px', fontWeight: 400, minWidth: '45px', textAlign: 'center' }}>{zoom}%</span>
                                        <button onClick={() => setZoom(z => Math.min(200, z + 10))} className="zoom-btn" style={{ color: '#94a3b8' }}><ZoomIn size={18} /></button>
                                    </div>
                                )}

                                <div style={{ flex: 1 }}></div>

                                <div className="template-header-actions" style={{ display: 'flex', alignItems: 'center', gap: '12px' }}>
                                    {placedFields.length > 0 && (
                                        <>
                                            <button onClick={saveToSystem} className="newblackcolorbtn" style={{ background: '#1e293b', border: '1px solid #334155', padding: '8px 16px', borderRadius: '6px', fontSize: '14px' }}>
                                                Save Template
                                            </button>
                                            <button onClick={clearTemplate} className="newblackcolorbtn" style={{ background: '#1e293b', border: '1px solid #334155', padding: '8px 16px', borderRadius: '6px', fontSize: '14px' }}>
                                                Clear
                                            </button>
                                        </>
                                    )}
                                    {/* {selectedApp && (
                                        <label className="upload-label" style={{ margin: 0, background: '#0ea5e9', color: 'white', padding: '8px 16px', borderRadius: '6px', fontSize: '14px', fontWeight: 500, cursor: 'pointer' }}>
                                            {uploadedFile ? 'Change File' : 'Upload Template'}
                                            <input ref={fileInputRef} type="file" onChange={handleFileUpload} accept=".pdf,.doc,.docx" className="hidden" />
                                        </label>

                                    )} */}
                                    {!selectedApp && (
                                        <Toast ref={toast} />)}
                                </div>
                            </div>

                            {/* Right Section: Form Fields Title (Matches Sidebar Width) */}
                            <div style={{ width: '300px', display: 'flex', alignItems: 'center', justifyContent: 'center', background: '#0b1120', borderLeft: '1px solid #1e293b', alignSelf: 'stretch' }}>
                                <h2 style={{ margin: 0, fontSize: '1.25rem', color: 'white', fontWeight: 500, letterSpacing: '0.5px' }}>Form Fields</h2>
                            </div>
                        </div>

                        <div style={{ display: 'flex', flex: 1, minHeight: 0, overflow: 'hidden' }}>
                            {/* Canvas Area */}
                            <div style={{ display: 'flex', flex: 1, minHeight: 0 }}>
                                <div className="template-canvas" style={{ flex: 1, overflow: 'auto', position: 'relative', backgroundColor: 'black' }}>
                                    {selectedField && (
                                        (() => {
                                            const field = placedFields.find(f => f.id === selectedField);
                                            if (!field) return null;
                                            return (
                                                <div className='d-flex justify-content-between ' style={{ display: 'flex', flexDirection: 'row', width: '100%', backgroundColor: 'white', position: 'sticky', top: 0, zIndex: 1100 }}>
                                                    {/* Positioned relative to viewport at bottom center */}
                                                    <div className="flex flex-col gap-1" style={{ width: '30%' }}>
                                                        <label className="text-xs font-medium text-gray-500 uppercase">Font Size</label>
                                                        <input
                                                            style={{ width: '30%' }}
                                                            type="number"
                                                            value={field.fontSize}
                                                            onChange={(e) => updateField(field.id, { fontSize: parseInt(e.target.value) })}
                                                            className="w-16 border rounded px-2 py-1 text-sm text-center"
                                                        />
                                                    </div>

                                                    <div className="flex flex-col gap-1" style={{ width: '30%' }} >
                                                        <label className="text-xs font-medium text-gray-500 uppercase">Font Family</label>
                                                        <select
                                                            value={field.fontFamily}
                                                            style={{ width: '45%', backgroundColor: 'black', color: 'white' }}
                                                            onChange={(e) => updateField(field.id, { fontFamily: e.target.value })}
                                                            className="border rounded px-2 py-1 text-sm w-32"
                                                        >
                                                            <option value="Arial">Arial</option>
                                                            <option value="Times New Roman">Times New Roman</option>
                                                            <option value="Courier New">Courier New</option>
                                                            <option value="Verdana">Verdana</option>
                                                            <option value="Georgia">Georgia</option>
                                                        </select>
                                                    </div>

                                                    <div className="flex flex-col gap-1" style={{ width: '30%' }}>
                                                        <label className="text-xs font-medium text-gray-500 uppercase">Width (px)</label>
                                                        <input
                                                            style={{ width: '30%' }}
                                                            type="number"
                                                            value={Math.round(field.width)}
                                                            onChange={(e) => updateField(field.id, { width: parseInt(e.target.value) })}
                                                            className="w-20 border rounded px-2 py-1 text-sm text-center"
                                                        />
                                                    </div>

                                                    <div className="flex flex-col gap-1" style={{ width: '30%' }}>
                                                        <label className="text-xs font-medium text-gray-500 uppercase">Height (px)</label>
                                                        <input
                                                            style={{ width: '30%' }}
                                                            type="number"
                                                            value={Math.round(field.height)}
                                                            onChange={(e) => updateField(field.id, { height: parseInt(e.target.value) })}
                                                            className="w-20 border rounded px-2 py-1 text-sm text-center"
                                                        />
                                                    </div>

                                                    <div className="h-8 w-px bg-gray-300 mx-2"></div>

                                                    <button
                                                        onClick={() => setSelectedField(null)}
                                                        className="newblackcolorbtn"
                                                    >
                                                        Done
                                                    </button>
                                                </div>
                                            );
                                        })()
                                    )}
                                    {!uploadedFile ? (
                                        <div className="upload-placeholder">
                                            <div className="upload-placeholder-content"
                                                onClick={() =>
                                                    selectedApp ? (fileInputRef.current?.click()) :
                                                        (toast.current?.show({ severity: 'info', summary: 'Select a Form First', detail: 'Please select a form from the dropdown to enable template upload.', life: 3000 }))
                                                }
                                            >
                                                <Upload style={{ color: '#0ea5e9', cursor: 'pointer' }} />
                                                <h3>Select an App</h3>
                                                <input ref={fileInputRef} type="file" onChange={handleFileUpload} accept=".pdf" className="hidden" style={{ width: '0%', height: '0%' }} />
                                                <h3>Upload Template</h3>
                                                <p>PDF document is only supported</p>
                                            </div>
                                        </div>
                                    ) : (
                                        <div className="template-content" style={{ padding: 0, margin: 0 }}>
                                            <div className="template-wrapper" style={{ padding: 0, margin: 0 }}>
                                                {renderFileContent()}
                                            </div>
                                        </div>
                                    )}
                                </div>


                            </div>
                            {/* Right Panel (FieldSidebar) */}
                            <div className="fields-panel" style={{ width: '300px', display: 'flex', flexDirection: 'column', borderLeft: '1px solid #1e293b', backgroundColor: '#0b1120' }}>
                                <div className="flex-1 overflow-y-auto" style={{ backgroundColor: '#0b1120' }}>
                                    {selectedApp && (
                                        <FieldSidebar
                                            fields={sidebarFields}
                                            onDragStart={handleDragStart}
                                            onDragEnd={handleDragEnd}
                                        />
                                    )}
                                </div>
                            </div>

                        </div>
                    </div>
                </div>

                {/* Field Properties Panel - Floating when selected */}

                {showNameDialog && (
                    <div className="custom-dark-modal-overlay">
                        <div className="custom-dark-modal">
                            <h2 className='addsteptitle' style={{ fontSize: '1rem' }}>
                                Save the Template, for "{formmasteritems.find(item => item.AppCode === selectedApp)?.AppName || 'Unknown'}":
                            </h2>
                            <div className="form-group" style={{ padding: '0px 20px' }}>
                                <div className="form-control-group">
                                    <p className="text-sm text-gray-600 mb-4">Give your template a recognizable name.</p>
                                    <input
                                        type="text"
                                        className="form-control"
                                        style={{ backgroundColor: 'black', color: 'white' }}
                                        placeholder="e.g., Q3 Invoice Template"
                                        value={templateName}
                                        onChange={(e) => setTemplateName(e.target.value)}
                                        autoFocus
                                    />
                                </div>
                            </div>
                            <div className="modal-buttons" style={{ padding: '10px 20px 20px 20px' }}>
                                <div style={{ width: '50%' }} className='modal-buttons'>
                                    <button
                                        onClick={handleConfirmSave}
                                        className="newlogocolorbtn">
                                        Save</button>
                                    <button
                                        onClick={() => setShowNameDialog(false)}
                                        className="newblackcolorbtn">
                                        Cancel
                                    </button>
                                </div>
                            </div>
                        </div>
                    </div>
                )}
            </div>
        </div>
    );
}
