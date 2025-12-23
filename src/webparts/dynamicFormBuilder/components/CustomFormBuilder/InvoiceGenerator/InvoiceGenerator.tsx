import * as React from 'react';
import { useState, useEffect } from 'react';
import { FileText, Loader, CheckCircle } from 'lucide-react';
import { SharePointService, type TemplateRecord } from '../services/SharePointService';
import { PdfGeneratorService } from '../services/PdfGeneratorService';
import { MOCK_RECORD } from './mockData';
import * as pdfjsLib from 'pdfjs-dist/legacy/build/pdf';
import pdfWorker from 'pdfjs-dist/legacy/build/pdf.worker.entry';

// Configure PDF.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc = pdfWorker;

// Props interface
interface InvoiceGeneratorProps {
    answerData: any[];
    tableData: any[];
    timesheetData: any[];
    formSchema: any[];
}

export default function InvoiceGenerator({ answerData, tableData, timesheetData, formSchema }: InvoiceGeneratorProps) {
    // Create REAL_RECORD from props
    const REAL_RECORD = {
        id: 'r_real_data',
        name: 'Real Data (answerjson.json)',
        data: {
            ...answerData.reduce((acc, curr) => {
                const updated = { ...acc };
                if (curr.id) updated[curr.id] = curr.value;
                if (curr.name) updated[curr.name] = curr.value;
                if (curr.custom_name) updated[curr.custom_name] = curr.value;
                return updated;
            }, {} as Record<string, any>),
            ...tableData.reduce((acc, curr) => {
                const updated = { ...acc };
                if (curr.id) updated[curr.id] = curr;
                if (curr.name) updated[curr.name] = curr;
                if (curr.custom_name) updated[curr.custom_name] = curr;
                return updated;
            }, {} as Record<string, any>),
            ...timesheetData.reduce((acc, curr) => {
                const updated = { ...acc };
                if (curr.id) updated[curr.id] = curr;
                if (curr.name) updated[curr.name] = curr;
                if (curr.custom_name) updated[curr.custom_name] = curr;
                return updated;
            }, {} as Record<string, any>)
        }
    };

    const RECORDS = [MOCK_RECORD, REAL_RECORD];
    const [templates, setTemplates] = useState<TemplateRecord[]>([]);
    const [selectedTemplate, setSelectedTemplate] = useState<string>('');
    const [selectedRecord, setSelectedRecord] = useState<string>('r_real_data');
    const [loading, setLoading] = useState(true);
    const [generating, setGenerating] = useState(false);
    const [error, setError] = useState<string | null>(null);

    const loadTemplates = async () => {
        setLoading(true);
        setError(null);
        try {
            console.log('InvoiceGenerator: Fetching templates from SharePoint...');
            const data = await SharePointService.getAllTemplates();
            console.log('InvoiceGenerator: Fetched templates:', data);

            setTemplates(data);

            if (data.length > 0) {
                // Keep existing selection if it still exists, otherwise pick first
                setSelectedTemplate(prev => {
                    if (prev && data.some(t => t.id === prev)) return prev;
                    return data[0].id;
                });
            } else {
                setSelectedTemplate('');
            }
        } catch (err) {
            console.error('InvoiceGenerator: Error loading templates:', err);
            setError('Failed to load templates from SharePoint. Please check your connection and the SMRT_PDF_Templates list.');
        } finally {
            setLoading(false);
        }
    };

    // Load available templates on mount
    useEffect(() => {
        void loadTemplates();
    }, []);

    const handleGenerate = async () => {
        if (!selectedTemplate || !selectedRecord) {
            alert('Please select both a template and a record.');
            return;
        }
        setGenerating(true);

        try {
            const record = RECORDS.find(r => r.id === selectedRecord) || MOCK_RECORD;
            const recordData = record.data as Record<string, any>;

            await PdfGeneratorService.generatePDF(
                selectedTemplate,
                recordData,
                formSchema
            );

            alert('Invoice Generated (High Quality) Successfully!');
        } catch (error: any) {
            console.error('Generation Error:', error);
            alert(`Failed to generate invoice: ${error.message}`);
        } finally {
            setGenerating(false);
        }
    };

    return (
        <div className="generator-container">
            <h1 className="generator-title">
                <FileText className="w-8 h-8" />
                End User Invoice Generator
            </h1>

            {error && (
                <div className="bg-red-50 border border-red-200 text-red-700 px-4 py-3 rounded-lg text-sm mb-6 flex justify-between items-center">
                    <span>{error}</span>
                    <button onClick={loadTemplates} className="underline font-medium hover:text-red-800">Retry</button>
                </div>
            )}

            {loading ? (
                <div className="loading-spinner">
                    <Loader className="w-6 h-6 animate-spin" /> Loading Templates...
                </div>
            ) : (
                <div className="generator-form">
                    <div className="generator-btn-group">
                        <div className="flex-1">
                            <label className="generator-label">Select Template (From SharePoint)</label>
                            <div className="flex gap-2">
                                <select
                                    value={selectedTemplate}
                                    onChange={(e) => setSelectedTemplate(e.target.value)}
                                    className="generator-select"
                                >
                                    {templates.map(t => (
                                        <option key={t.id} value={t.id}>{t.name} (Created: {new Date(t.createdAt).toLocaleDateString()})</option>
                                    ))}
                                    {templates.length === 0 && <option value="">No templates found</option>}
                                </select>
                                <button
                                    onClick={loadTemplates}
                                    className="p-2 border border-gray-300 rounded-lg hover:bg-gray-50 flex items-center justify-center"
                                    title="Refresh Templates"
                                >
                                    <Loader className={`w-4 h-4 ${loading ? 'animate-spin' : ''}`} />
                                </button>
                            </div>
                        </div>
                        <button
                            onClick={async () => {
                                if (confirm('Are you sure you want to delete ALL templates? This cannot be undone.')) {
                                    await SharePointService.clearAllTemplates();
                                    await loadTemplates();
                                    alert('Templates cleared.');
                                }
                            }}
                            className="generator-btn-reset"
                        >
                            Reset
                        </button>
                    </div>

                    <div className="generator-field-group">
                        <label className="generator-label">Select Record</label>
                        <select
                            value={selectedRecord}
                            onChange={(e) => setSelectedRecord(e.target.value)}
                            className="generator-select"
                        >
                            {RECORDS.map(r => (
                                <option key={r.id} value={r.id}>{r.name}</option>
                            ))}
                        </select>
                    </div>

                    <button
                        onClick={handleGenerate}
                        disabled={generating || !selectedTemplate}
                        className="generator-btn-main"
                    >
                        {generating ? <Loader className="w-5 h-5 animate-spin" /> : <CheckCircle className="w-5 h-5" />}
                        {generating ? 'Generating PDF...' : 'Generate Invoice'}
                    </button>

                    <div className="generator-info-box">
                        <strong>Note:</strong> This simulates the End User view. The user does not see the template editor, only this form to pick a template and data.
                    </div>
                </div>
            )}
        </div>
    );
}
