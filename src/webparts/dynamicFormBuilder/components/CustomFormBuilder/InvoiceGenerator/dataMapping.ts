interface SchemaOption {
    value: string;
    text: string;
    key?: string;
}

interface SchemaItem {
    id: string;
    field_name?: string;
    label?: string;
    text?: string;
    name?: string;
    element: string;
    options?: SchemaOption[];
    content?: string;
    columns?: any[];
}

interface AnswerItem {
    id: string;
    value?: string | string[];
}

export interface MappedField {
    id: string;
    originalId: string;
    label: string;
    value: string;
    type: string;
    tableData?: {
        columns: any[];
        rows: any[];
    };
}

export const mapSchemaToFields = (schema: SchemaItem[]): MappedField[] => {
    debugger;
    return schema
        .filter(item => {
            const excludedElements = ['Header', 'LineBreak', 'ThreeColumnRow', 'TwoColumnRow', 'FileUpload'];

            // Allow Table (Label with content='Table') and Timesheet (Download with content='Timesheet')
            if (item.element === 'Label' && item.content === 'Table') return true;
            if (item.element === 'Download' && item.content === 'Timesheet') return true;

            // Also excluding layout containers that don't have direct values usually
            if (item.element === 'Label' || item.element === 'Download') return false;

            return !excludedElements.includes(item.element);
        })
        .map(item => {
            let type = item.element;
            if (item.element === 'Label' && item.content === 'Table') type = 'Table';
            if (item.element === 'Download' && item.content === 'Timesheet') type = 'Timesheet';

            const rawLabel = item.label || item.text || item.name || 'Untitled Field';
            // Strip HTML tags from label
            const cleanLabel = rawLabel.replace(/<[^>]*>/g, '');

            return {
                id: item.field_name || item.id,
                originalId: item.id,
                label: cleanLabel,
                value: '', // Initial empty value
                type: type
            };
        });
};

export const linkFormAndAnswers = (
    schema: SchemaItem[],
    answers: AnswerItem[],
    tableAnswers: any[] = [],
    timesheetAnswers: any[] = []
): MappedField[] => {
    const fields = mapSchemaToFields(schema);

    return fields.map(field => {
        const schemaItem = schema.find(item => item.id === field.originalId);

        // Handle Tables
        if (field.type === 'Table') {
            const tableData = tableAnswers.find(t => t.id === field.originalId);
            return {
                ...field,
                value: 'Table Data',
                tableData: {
                    columns: schemaItem?.['columns'] || [],
                    rows: tableData?.rows || []
                }
            };
        }

        // Handle Timesheets
        if (field.type === 'Timesheet') {
            const sheetData = timesheetAnswers.find(t => t.id === field.originalId);
            return {
                ...field,
                value: 'Timesheet Data',
                tableData: {
                    columns: schemaItem?.['columns'] || [],
                    rows: sheetData?.rows || []
                }
            };
        }

        const answer = answers.find(a => a.id === field.originalId);

        let displayValue = '';
        if (answer) {
            if (field.type === 'RadioButtons' || field.type === 'Checkboxes') {
                // For Radio/Checkboxes, answer.value contains option keys (IDs) or values
                // We need to match these with schemaItem.options
                const answerValues = Array.isArray(answer.value) ? answer.value : [answer.value];

                if (schemaItem?.options) {
                    const selectedOptions = schemaItem.options.filter(opt =>
                        answerValues.includes(opt.value) || answerValues.includes(opt.key || '')
                    );
                    displayValue = selectedOptions.map(opt => (opt.text || '').replace(/<[^>]*>/g, '')).join(', ');
                } else {
                    displayValue = String(answer.value || '');
                }

            } else {
                if (Array.isArray(answer.value)) {
                    displayValue = answer.value.join(', ');
                } else {
                    displayValue = String(answer.value || '');
                }
            }
        }

        return {
            ...field,
            value: displayValue
        };
    });
};
