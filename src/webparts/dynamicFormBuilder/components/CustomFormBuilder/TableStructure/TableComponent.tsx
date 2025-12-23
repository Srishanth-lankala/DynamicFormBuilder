import * as React from 'react';
import { useState, useRef, useEffect } from 'react';
import { IconButton } from '@fluentui/react';
import { initializeFileTypeIcons, getFileTypeIconProps } from '@fluentui/react-file-type-icons';
import { Icon } from '@fluentui/react/lib/Icon';
import { sp } from "@pnp/sp/presets/all";
import { Callout, DirectionalHint } from '@fluentui/react/lib/Callout';
// import { TooltipHost } from '@fluentui/react/lib/Tooltip';
import './Table.css';
import { fileFetchUrl, mySiteUrl } from '../ConfigURL/All_URLs';

initializeFileTypeIcons();

// Define interfaces for restrictions
interface DropdownCondition {
  value: string;
  min?: number;
  max?: number;
}

interface Restrictions {
  restrictionType?: 'number' | 'conditional';
  min?: number;
  max?: number;
  dropdownColumnIndex?: number;
  dropdownConditions?: DropdownCondition[];
}

interface Option {
  id: string;
  value: string;
}

interface Header {
  headername: string;
  columntype: 'Text' | 'Number' | 'Attachment' | 'Dropdown';
  options?: Option[];
  restrictions?: Restrictions;
}

export interface TableData {
  id: string;
  content: string;
  text: string;
  numberOfColumns: number;
  columns: Header[];
  rows?: any[];
}

interface TableComponentProps {
  tableConfig: TableData;
  onTableUpdate: (updatedTable: TableData) => void;
  onRestrictionBreachUpdate?: (hasBreaches: boolean) => void;
  readonly?: boolean;
  seqno?: string;
}

const TableComponent: React.FC<TableComponentProps> = ({ tableConfig, onTableUpdate, onRestrictionBreachUpdate, readonly = false, seqno }) => {
  const [tableData, setTableData] = useState<TableData>(tableConfig);
  const [showSum, setShowSum] = useState<boolean[]>(new Array(tableConfig.columns?.length).fill(false));
  const [, setHasRestrictionBreaches] = useState<boolean>(false);
  const [, setCalloutTarget] = useState<HTMLElement | null>(null);
  const [, setCalloutContent] = useState<string>('');
  const [calloutCellId, setCalloutCellId] = useState<string | null>(null);
  const fileInputRefs = useRef<{ [key: string]: HTMLInputElement | null }>({});
  const [showTableInfoTooltip, setShowTableInfoTooltip] = useState(false);
  const [tableInfoTooltipTarget, setTableInfoTooltipTarget] = useState<HTMLElement | null>(null);





  const isInputBreached = (column: Header, row: any): boolean => {
    if (column.columntype !== 'Number') return false;
    const inputValue = parseFloat(row[column.headername] as string);
    if (isNaN(inputValue)) return false;

    const restrictions = column.restrictions || {};
    if (restrictions.restrictionType === 'number') {
      const { min, max } = restrictions;
      return (min !== undefined && inputValue < min) || (max !== undefined && inputValue > max);
    } else if (restrictions.restrictionType === 'conditional' && restrictions.dropdownColumnIndex !== undefined) {
      const dropdownColumn = tableData.columns[restrictions.dropdownColumnIndex];
      const dropdownValue = row[dropdownColumn.headername] as string;
      if (!dropdownValue) return false;
      const condition = restrictions.dropdownConditions?.find(cond => cond.value === dropdownValue);
      if (!condition) return false;
      return (condition.min !== undefined && inputValue < condition.min) || (condition.max !== undefined && inputValue > condition.max);
    }
    return false;
  };

  const updateRestrictionBreaches = () => {
    let breaches = false;
    tableData.rows?.forEach(row => {
      tableData.columns.forEach(column => {
        if (isInputBreached(column, row)) {
          breaches = true;
        }
      });
    });
    setHasRestrictionBreaches(breaches);
    if (onRestrictionBreachUpdate) {
      onRestrictionBreachUpdate(breaches);
    }
  };

  useEffect(() => {
    if (!tableData.rows || tableData.rows.length === 0 && !readonly) {
      const newRow: any = { id: `row-${Date.now()}` };
      tableData.columns.forEach(header => {
        newRow[header.headername] = header.columntype === 'Number' ? 0 : '';
      });

      setTableData(prev => {
        const updated = {
          ...prev,
          rows: [...(prev.rows || []), newRow]
        };
        onTableUpdate(updated);
        return updated;
      });
    }
  }, [readonly]);

  sp.setup({
    sp: {
      baseUrl: mySiteUrl,
    },
  });

  const addNewRow = (e: React.MouseEvent<HTMLButtonElement>) => {
    e.preventDefault();
    const newRow: any = { id: `row-${Date.now()}` };
    tableData.columns.forEach(header => {
      newRow[header.headername] = header.columntype === 'Number' ? 0 : '';
    });

    setTableData(prev => {
      const updated = {
        ...prev,
        rows: [...(prev.rows || []), newRow]
      };
      onTableUpdate(updated);
      return updated;
    });
  };

  const deleteRow = (rowId: string) => {
    setTableData(prev => {
      const updated = {
        ...prev,
        rows: prev.rows?.filter(row => row.id !== rowId) || []
      };
      onTableUpdate(updated);
      return updated;
    });
    updateRestrictionBreaches();
  };

  const handleCellChange = (rowId: string, headerName: string, value: any) => {
    if (readonly) return;

    setTableData(prev => {
      const updatedRows = prev.rows?.map(row => {
        if (row.id === rowId) {
          return { ...row, [headerName]: value };
        }
        return row;
      }) || [];

      const updated = {
        ...prev,
        rows: updatedRows
      };
      onTableUpdate(updated);
      return updated;
    });
    updateRestrictionBreaches();
  };

  const handleFileUpload = (rowId: string, headerName: string, e: React.ChangeEvent<HTMLInputElement>) => {
    if (readonly) return;
    const file = e.target.files?.[0];
    if (file) {
      setTableData(prev => {
        const updatedRows = prev.rows?.map(row => {
          if (row.id === rowId) {
            return { ...row, [headerName]: { name: file.name, fileObject: file } };
          }
          return row;
        }) || [];

        const updated = {
          ...prev,
          rows: updatedRows
        };
        onTableUpdate(updated);
        return updated;
      });
    }
  };

  const toggleSum = (index: number) => {
    const newShowSum = [...showSum];
    newShowSum[index] = !newShowSum[index];
    setShowSum(newShowSum);
  };

  const calculateSum = (headerName: string): number => {
    return tableData.rows?.reduce((sum, row) => {
      const value = parseFloat(row[headerName]) || 0;
      return sum + value;
    }, 0) || 0;
  };

  const handleColumnResize = (e: React.MouseEvent, headerName: string) => {
    const startX = e.clientX;
    const startWidth = (e.target as HTMLElement).parentElement?.offsetWidth || 0;

    const handleMouseMove = (moveEvent: MouseEvent) => {
      const newWidth = startWidth + (moveEvent.clientX - startX);
      document.documentElement.style.setProperty(`--col-${headerName}-width`, `${newWidth}px`);
    };

    const handleMouseUp = () => {
      document.removeEventListener('mousemove', handleMouseMove);
      document.removeEventListener('mouseup', handleMouseUp);
    };

    document.addEventListener('mousemove', handleMouseMove);
    document.addEventListener('mouseup', handleMouseUp);
  };

  const getRestrictionTooltip = (column: Header, row: any): string => {
    const restrictions = column.restrictions || {};
    if (restrictions.restrictionType === 'number') {
      const { min, max } = restrictions;
      return `Min: ${min !== undefined ? min : 'None'}, Max: ${max !== undefined ? max : 'None'}`;
    } else if (restrictions.restrictionType === 'conditional' && restrictions.dropdownColumnIndex !== undefined) {
      const dropdownColumn = tableData.columns[restrictions.dropdownColumnIndex];
      const dropdownValue = row[dropdownColumn.headername] as string;
      if (!dropdownValue) {
        return 'Select a dropdown value to see restrictions';
      }
      const condition = restrictions.dropdownConditions?.find(cond => cond.value === dropdownValue);
      if (!condition) return 'No restrictions for this dropdown value';
      return `For "${dropdownValue}" - Min: ${condition.min !== undefined ? condition.min : 'None'}, Max: ${condition.max !== undefined ? condition.max : 'None'}`;
    }
    return '';
  };

  const getDetailedRestrictionContent = (column: Header, row: any): JSX.Element => {
    const restrictions = column.restrictions || {};
    if (restrictions.restrictionType === 'number') {
      const { min, max } = restrictions;
      return (
        <div>
          <p><strong>Number Restrictions:</strong></p>
          <ul>
            <li>Min: {min !== undefined ? min : 'None'}</li>
            <li>Max: {max !== undefined ? max : 'None'}</li>
          </ul>
        </div>
      );
    } else if (restrictions.restrictionType === 'conditional' && restrictions.dropdownColumnIndex !== undefined) {
      const dropdownColumn = tableData.columns[restrictions.dropdownColumnIndex];
      const dropdownValue = row[dropdownColumn.headername] as string;
      if (!dropdownValue) {
        return <p>Select a dropdown value to see restrictions</p>;
      }
      const condition = restrictions.dropdownConditions?.find(cond => cond.value === dropdownValue);
      if (!condition) {
        return <p>No restrictions for this dropdown value</p>;
      }
      return (
        <div>
          <p><strong>Conditional Restrictions for "{dropdownValue}":</strong></p>
          <ul>
            <li>Min: {condition.min !== undefined ? condition.min : 'None'}</li>
            <li>Max: {condition.max !== undefined ? condition.max : 'None'}</li>
          </ul>
        </div>
      );
    }
    return <p>No restrictions available</p>;
  };

  const handleInfoClick = (event: React.MouseEvent<HTMLElement>, column: Header, row: any, cellId: string) => {
    setCalloutTarget(event.currentTarget as HTMLElement);
    setCalloutContent(getRestrictionTooltip(column, row));
    setCalloutCellId(cellId);
  };

  // const dismissCallout = () => {
  //   setCalloutTarget(null);
  //   setCalloutCellId(null);
  // };

  useEffect(() => {
    updateRestrictionBreaches();
  }, [tableData]);

  return (
    <div className="dynamic-table-container">
      <div className='d-flex align-items-center justify-content-between mb-1'>
        <h6 style={{ marginBottom: 0  , fontSize: '14px', fontWeight: 400 }}>{tableData.content}</h6>
        {!readonly && (
          <button className='btn btn-primary addrow' onClick={addNewRow}>+ Add Row</button>
        )}
      </div>

      <div className="table-responsive">
        <table className={`dynamic-table ${readonly ? 'readonly-mode' : ''}`}>
          <thead>
            <tr>
              {tableData.columns.map((header, index) => (
                <th
                  key={header.headername}
                  style={{
                    width: `var(--col-${header.headername}-width, auto)`,
                    position: 'relative', backgroundColor: '#161d2f'
                  }}
                >
                  <div style={{ display: 'flex', alignItems: 'center' }}>
                    {header.headername}
                    {header.columntype === 'Number' && (
                      <IconButton
                        // iconProps={{ iconName: showSum[index] ? 'CalculatorSubtract' : 'CalculatorAddition' }}
                        title="Toggle Sum"
                        onClick={() => toggleSum(index)}
                        styles={{
                          root: { marginLeft: '5px', height: '20px' },
                          rootHovered: {
                            backgroundColor: '#212b46',
                            color: 'white',
                          }
                          ,
                          rootPressed: {
                            backgroundColor: '#212b46', // Custom click color
                            color: 'white',
                          }
                        }}
                        onRenderIcon={() => <span style={{ color: 'white', fontSize: '16px' }}>{showSum[index] ? '-' : '∑'}</span>}
                      />
                    )}
                  </div>
                  <div
                    className="column-resize-handle" style={{ background: '#070D19' }}
                    onMouseDown={(e) => handleColumnResize(e, header.headername)}
                  />
                </th>
              ))}
              {!readonly && <th style={{ textAlign: 'center', backgroundColor: '#161d2f' }}>Actions</th>}
            </tr>
          </thead>
          <tbody>
            {tableData.rows?.map((row) => (
              <tr key={row.id}>
                {tableData.columns.map((header) => {
                  const cellValue = row[header.headername] || '';
                  const isBreached = isInputBreached(header, row);
                  const cellId = `${row.id}-${header.headername}`;

                  return (
                    <td
                      key={cellId}
                      style={{ backgroundColor: isBreached ? '' : '#0D1322' }}
                      className={isBreached ? 'error-cell' : ''}
                    >
                      {header.columntype === 'Dropdown' ? (
                        <select
                          value={cellValue}
                          onChange={(e) => handleCellChange(row.id, header.headername, e.target.value)}
                          className="borderless-input"
                          disabled={readonly}
                          style={{ paddingLeft: 7 }}
                        >
                          <option value="" disabled selected hidden></option>
                          <option value="" disabled style={{ backgroundColor: 'lightgray' }}>Select</option>
                          {header.options?.map(opt => (
                            <option key={opt.id} value={opt.value}>
                              {opt.value}
                            </option>
                          ))}
                        </select>
                      ) : header.columntype === 'Attachment' ? (
                        <div style={{ display: 'flex', flexDirection: 'column' }}>
                          <div style={{ display: 'flex', alignItems: 'center' }}>
                            {cellValue?.name ? (
                              <>
                                <Icon style={{ marginBottom: "5px" }} {...getFileTypeIconProps({ extension: cellValue.name.split('.').pop() || '', size: 16 })} />
                                {
                                  <a
                                    href={`${fileFetchUrl}/TableAttachments/${seqno}/${tableData.id}/${cellValue.name}`}
                                    target="_blank"
                                    rel="noopener noreferrer"
                                    onClick={(e) => e.stopPropagation()}
                                    style={{ marginLeft: '5px' }}
                                  >
                                    {cellValue.name}
                                  </a>
                                }
                              </>
                            ) : null}
                            {!readonly && (
                              <>
                                <input
                                  type="file"
                                  ref={el => (fileInputRefs.current[`${row.id}-${header.headername}`] = el)}
                                  style={{ display: 'none' }}
                                  onChange={(e) => handleFileUpload(row.id, header.headername, e)}
                                />
                                <button
                                  type='button'
                                  onClick={() => fileInputRefs.current[`${row.id}-${header.headername}`]?.click()}
                                  style={{ marginLeft: '10px' }}
                                  className=' btn btn-primary addrow'
                                >
                                  {cellValue ? 'Change' : 'Attach'}
                                </button>
                                {console.log("row", row)}
                                {console.log("cellvalue", cellValue.fileObject?.size)}
                              </>
                            )}
                          </div>
                          {!readonly && cellValue?.fileObject?.size > 1024 * 1024 && (
                            <div style={{ color: 'red', marginTop: '4px' }}>
                              File exceeds 1MB
                            </div>
                          )}
                        </div>
                      ) : (
                        <div style={{ display: 'flex', alignItems: 'center' }}>
                          <input
                            type={header.columntype === 'Number' ? 'number' : 'text'}
                            value={cellValue}
                            onChange={(e) => {
                              if (header.columntype === 'Number') {
                                if (e.target.value === '' || !isNaN(Number(e.target.value))) {
                                  handleCellChange(row.id, header.headername, e.target.value);
                                }
                              } else {
                                handleCellChange(row.id, header.headername, e.target.value);
                              }
                            }}
                            className="borderless-input no-arrows"
                            disabled={readonly}
                            style={{ paddingLeft: 7 }}
                          />
                          {header.columntype === 'Number' && header.restrictions && Object.keys(header.restrictions).length > 0 && (
                            // <TooltipHost
                            //   content={getRestrictionTooltip(header, row)}
                            //   directionalHint={DirectionalHint.topCenter}
                            // >
                            <IconButton
                              iconProps={{ iconName: 'Info' }}
                              styles={{
                                root: { marginLeft: '5px', height: '20px', color: 'white' },
                                rootHovered: {
                                  backgroundColor: '#212b46',
                                  color: 'white',
                                },
                                rootPressed: {
                                  backgroundColor: '#212b46', // Custom click color
                                  color: 'white',
                                }
                              }}
                              onMouseEnter={(e) => {
                                setTableInfoTooltipTarget(e.currentTarget as HTMLElement);
                                handleInfoClick(e as React.MouseEvent<HTMLElement>, header, row, cellId);
                                setShowTableInfoTooltip(true);
                              }}
                              onMouseLeave={() => setShowTableInfoTooltip(false)}
                            />
                            // </TooltipHost>
                          )}
                        </div>
                      )}
                    </td>
                  );
                })}
                {!readonly && (
                  <td style={{ textAlign: 'center' }}>
                    <IconButton
                      iconProps={{ iconName: 'Delete' }}
                      title="Delete row"
                      onClick={() => deleteRow(row.id)}
                      styles={{
                        root: { color: 'white' },
                        rootHovered: {
                          backgroundColor: '#212b46',
                          color: 'white',
                        },
                        rootPressed: {
                          backgroundColor: '#212b46', // Custom click color
                          color: 'white',
                        }
                      }}
                    />
                  </td>
                )}
              </tr>
            ))}
            {tableData.columns.some((_, i) => showSum[i]) && (
              <tr className="sum-row">
                {tableData.columns.map((header, index) => (
                  <td key={`sum-${header.headername}`}>
                    {showSum[index] && header.columntype === 'Number' ? (
                      <strong>Sum: {calculateSum(header.headername).toFixed(2)}</strong>
                    ) : null}
                  </td>
                ))}
                {!readonly && <td></td>}
              </tr>
            )}
          </tbody>
        </table>
      </div>
      {showTableInfoTooltip && tableInfoTooltipTarget && (
        <Callout
          target={tableInfoTooltipTarget}
          onDismiss={() => setShowTableInfoTooltip(false)}
          directionalHint={DirectionalHint.topCenter}
          styles={{
            root: {
              backgroundColor: '#fff',
              border: '1px solid #dee2e6',
              borderRadius: '4px',
              boxShadow: '0 2px 8px rgba(0,0,0,0.15)',
              padding: '10px',
              minWidth: '200px'
            }
          }}
        >
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '8px' }}>
            <span><strong>Restrictions</strong></span>
            {/* <IconButton
              iconProps={{ iconName: 'Cancel' }}
              onClick={dismissCallout}
              styles={{ root: { height: '20px' } }}
            /> */}
          </div>
          {tableData.rows?.map(row =>
            tableData.columns.map(header => {
              const cellId = `${row.id}-${header.headername}`;
              if (cellId === calloutCellId) {
                return (
                  <div key={cellId}>
                    {getDetailedRestrictionContent(header, row)}
                  </div>
                );
              }
              return null;
            })
          )}
        </Callout>
      )}
      {/* <div className="alert alert-info">
        Has Restriction Breaches: {hasRestrictionBreaches ? 'Yes' : 'No'}
      </div> */}
    </div>
  );
};

export default TableComponent;