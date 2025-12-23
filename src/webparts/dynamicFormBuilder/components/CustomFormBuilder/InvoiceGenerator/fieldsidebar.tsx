import React from 'react';
import { Move } from 'lucide-react';
import type { MappedField } from './dataMapping';

interface FieldSidebarProps {
    fields: MappedField[];
    onDragStart: (field: MappedField) => void;
    onDragEnd: () => void;
}

export default function FieldSidebar({ fields, onDragStart, onDragEnd }: FieldSidebarProps) {
    return (
        <div className="fields-list" style={{ backgroundColor: 'linear-gradient(to bottom, #1c02aaff, #030e7fff)' }}>
            {fields.map(field => (
                <div
                    key={field.id}
                    draggable
                    onDragStart={() => onDragStart(field)}
                    onDragEnd={onDragEnd}
                    className="field-card"
                >
                    <div className="field-card-content">
                        <div className="field-card-info">
                            <div className="field-card-name-row">
                                <Move className="w-4 h-4 text-gray-400 group-hover:text-blue-500 flex-shrink-0" />
                                <span className="field-card-name" title={field.label}>
                                    {field.label}
                                </span>
                            </div>
                            {field.value ? (
                                <div className="field-card-value">
                                    {field.value}
                                </div>
                            ) : (
                                <div className="text-xs italic text-gray-400 mt-1">
                                    {`{{${field.label}}}`}
                                </div>
                            )}
                        </div>
                        <span className="field-type-badge">{field.type}</span>
                    </div>
                </div>
            ))}
        </div>
    );
}
