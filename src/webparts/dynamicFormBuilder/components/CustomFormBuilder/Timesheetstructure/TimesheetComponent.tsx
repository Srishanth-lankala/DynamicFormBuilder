
import * as React from 'react';
import { useState, useRef, useEffect } from 'react';
import { IconButton } from '@fluentui/react';
import { initializeFileTypeIcons, getFileTypeIconProps } from '@fluentui/react-file-type-icons';
import { Icon } from '@fluentui/react/lib/Icon';
import { sp } from "@pnp/sp/presets/all";
import { Callout, DirectionalHint } from '@fluentui/react/lib/Callout';
// import { TooltipHost } from '@fluentui/react/lib/Tooltip';
import './Timesheet.css';
import { fileFetchUrl, mySiteUrl } from '../ConfigURL/All_URLs';
import { Calendar } from 'primereact/calendar';

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
  isdynamic: boolean;
  weekday: string;
}

export interface TimeSheetData {
  id: string;
  content: string;
  text: string;
  numberOfColumns: number;
  columns: Header[];
  weekends?: string[];
  rows?: any[];
  fromdate?: string;
  todate?: string;
  minHoursPerCell: number;
  maxHoursPerCell: number;
}

interface TableComponentProps {
  tableConfig: TimeSheetData;
  onTableUpdate: (updatedTable: TimeSheetData) => void;
  onRestrictionBreachUpdate?: (hasBreaches: boolean) => void;
  readonly?: boolean;
  seqno?: string;
}

const TimesheetComponent: React.FC<TableComponentProps> = ({ tableConfig, onTableUpdate, onRestrictionBreachUpdate, readonly = false, seqno }) => {
  const [tableData, setTableData] = useState<TimeSheetData>(tableConfig);
  const [, setHasRestrictionBreaches] = useState<boolean>(false);
  const [calloutTarget, setCalloutTarget] = useState<HTMLElement | null>(null);
  const [, setCalloutContent] = useState<string>('');
  const [calloutCellId, setCalloutCellId] = useState<string | null>(null);
  const [fromdate, setFromDate] = useState<Date | null>(tableConfig?.fromdate ? new Date(tableConfig.fromdate) : null);
  const [todate, setToDate] = useState<Date | null>(tableConfig?.todate ? new Date(tableConfig.todate) : null);
  const [staticcolumns, setStaticcolumns] = useState<any>();
  const [showTableInfoTooltip, setShowTableInfoTooltip] = useState(false);
  const [tableInfoTooltipTarget, setTableInfoTooltipTarget] = useState<HTMLElement | null>(null);

  const fileInputRefs = useRef<{ [key: string]: HTMLInputElement | null }>({});
  const fetchfiledetails = async () => {
    const files = await sp.web.lists
      .getByTitle("TableAttachments")
      .items.select("FileLeafRef", "FileRef", "FileDirRef")
      .filter(`FileDirRef eq '${fileFetchUrl}/TableAttachments/${seqno}/3138EBE7-F08B-4C2B-86BE-7A5135290154'`)
      .get();
    console.log('files>>>>>>>>>>>>>>>>>>>>>>>', files);
  };


  const calculateSum = (headerName: string): number => {
    return tableData.rows?.reduce((sum, row) => {
      const value = parseFloat(row[headerName]) || 0;
      return sum + value;
    }, 0) || 0;
  };
  const totalsum =()=> {
    return tableData.columns?.reduce((colsum, header) => {
    return (header.columntype === 'Number' && header.isdynamic) ?
      colsum + tableData.rows?.reduce((rowsum, row) => {
        return rowsum + parseFloat(row[header.headername]) || 0;
      }, 0) || 0 : 0
  }, 0)
}
  console.log("totalsum", totalsum());


  const isInputBreached = (column: Header, row: any): boolean => {
    if (column.columntype !== 'Number') return false;
    const inputValue = parseFloat(row[column.headername] as string);
    if (isNaN(inputValue)) return false;
    // Only apply min/max hours restriction for dynamic columns
    if (column.isdynamic) {
      if (inputValue < tableData.minHoursPerCell || inputValue > tableData.maxHoursPerCell) {
        return true;
      }
    }
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
    if(!breaches){
      tableData.columns?.map((header, index) => {
        if(header.isdynamic === true){
        const sum=calculateSum(header.headername).toFixed(2);
        if(Number(sum)>=24){
          console.log("breach",Number(sum)>=24,header.headername )
          breaches = true;
        }
      }
        return 
         })
    }
    setHasRestrictionBreaches(breaches);
    if (onRestrictionBreachUpdate) {
      onRestrictionBreachUpdate(breaches);
    }
  };

  useEffect(() => {
    setStaticcolumns(tableData.columns?.filter(item => item.isdynamic === false));
    console.log("st", tableData.columns)
    console.log("static", staticcolumns)
  }, [])

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

    void fetchfiledetails();
  }, [readonly]);

  useEffect(() => {
    console.log("useffecttriggered", fromdate, todate, staticcolumns)
    if (fromdate && todate && (fromdate.getTime() < todate.getTime()) && staticcolumns) {
      fromdate.setHours(0, 0, 0, 0);
      todate.setHours(23, 59, 59, 999);
      const diffmillis = todate.getTime() - fromdate.getTime();
      console.log("millis", diffmillis)
      const datediffdays = Math.ceil(diffmillis / (1000 * 60 * 60 * 24));
      console.log("daysdif", datediffdays)
      const newfromdate = new Date(fromdate);
      newfromdate.setDate(fromdate.getDate() + 1);
      const newcols: Header[] = Array(datediffdays).fill(1).map((_, i) => {
        const headerdate = new Date(fromdate);
        headerdate.setDate(fromdate.getDate() + i)
        const formatheaderdate = headerdate.toLocaleDateString('en-GB', {
          day: '2-digit',
          month: 'short',
        })
        const weekday = headerdate.toLocaleDateString('en-GB', { weekday: 'long' })
        return {
          headername: `${formatheaderdate}`,
          columntype: 'Number',
          options: [],
          restrictions: {},
          isdynamic: true,
          weekday: `${weekday}`
        }
      })
      // sort so that weekends always appear at the end
      const sortednewcols = newcols.sort((a, b) => {
        const aIsWeekend = tableConfig.weekends?.includes(a.weekday);
        const bIsWeekend = tableConfig.weekends?.includes(b.weekday);

        if (aIsWeekend && !bIsWeekend) return 1;
        if (!aIsWeekend && bIsWeekend) return -1;
        return 0;
      });
      console.log("headerdates", sortednewcols)
      const totalcols = Array.from(new Set([...staticcolumns, ...newcols]))
      console.log("totalcols", totalcols);
      const totaluniquecols = Array.from(new Map(totalcols?.map(col => [col.headername, col])).values());
      console.log("totaluniquecols", totaluniquecols);
      const allowedKeys = new Set(totaluniquecols?.map(col => col.headername));
      allowedKeys.add('id'); // Keep the "id" field 

      // Cleaned rows: remove extra fields in each row that are not in columns field
      const cleanedRows = tableData.rows?.map(row => {
        const cleaned: Record<string, any> = {};
        for (const key in row) {
          if (allowedKeys.has(key)) {
            cleaned[key] = row[key];
          }
        }
        return cleaned;
      });
      setTableData(prev => {
        console.log("paaaaa", newfromdate.toISOString().split('T')[0])
        return {
          ...prev,
          // columns:[...(prev.columns), ...newcols]
          columns: [...totaluniquecols],
          rows: cleanedRows,
          fromdate: newfromdate.toISOString().split('T')[0],
          todate: todate.toISOString().split('T')[0]
        }
      })

    }

  }, [fromdate, todate, staticcolumns])

  useEffect(() => {
    console.log("timedata", tableData)
  }, [tableData])

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


  // const calculateSum = (headerName: string): number => {
  //   return tableData.rows?.reduce((sum, row) => {
  //     const value = parseFloat(row[headerName]) || 0;
  //     return sum + value;
  //   }, 0) || 0;
  // };
  // const totalsum = tableData.columns.reduce((colsum, header) => {
  //   return (header.columntype === 'Number' && header.isdynamic) ?
  //     colsum + tableData.rows?.reduce((rowsum, row) => {
  //       return rowsum + parseFloat(row[header.headername]) || 0;
  //     }, 0) || 0 : 0
  // }, 0)
  // console.log("totalsum", totalsum);

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

  const dismissCallout = () => {
    setCalloutTarget(null);
    setCalloutCellId(null);
  };

  useEffect(() => {
    updateRestrictionBreaches();
  }, [tableData]);

  const today = new Date();
  let maxToDate = null;
  if (fromdate) {
    const sevenDaysAgo = new Date(today);
    sevenDaysAgo.setDate(today.getDate() - 7);
    if (fromdate >= sevenDaysAgo) {
      maxToDate = today;
    } else {
      maxToDate = new Date(fromdate.getTime() + 6 * 24 * 60 * 60 * 1000);
    }
  }
  const thirtyDaysAgo = new Date();
  thirtyDaysAgo.setDate(today.getDate() - 31);
  return (
    <div className="dynamic-table-container">
      <div className='d-flex align-items-center justify-content-between mb-1'>
        <h6 style={{ marginBottom: 0, fontSize: '14px', fontWeight: 400 }}>{tableData.content}</h6>
        <IconButton
          iconProps={{ iconName: 'Info' }}
          styles={{
            root: {
              marginLeft: '5px',
              height: '20px',
              color: 'white',
              backgroundColor: 'transparent'
            },
            rootHovered: {
              backgroundColor: '#212b46',
              color: 'white',
            },
            rootPressed: {
              backgroundColor: '#212b46',
              color: 'white',
            }
          }}
          onMouseEnter={(e) => {
            setTableInfoTooltipTarget(e.currentTarget as HTMLElement);
            setShowTableInfoTooltip(true);
          }}
          onMouseLeave={() => setShowTableInfoTooltip(false)}
        />
        <div style={{ marginLeft: 'auto', display: 'flex', gap: '10px', paddingRight: 10 }}>
          {(fromdate && todate && !readonly &&
            <IconButton
              title="Toggle Week"
              onClick={() => {
                const newfrom = new Date(fromdate.setDate(fromdate.getDate() - 7));
                const newto = new Date(todate.setDate(todate.getDate() - 7));
                setFromDate(newfrom);
                setToDate(newto);
              }}
              disabled={(fromdate.getTime() - thirtyDaysAgo.getTime() <= 7 * 24 * 60 * 60 * 1000)}
              styles={{
                root: { height: '29px', width: '29px', color: 'white' },
                rootHovered: {
                  backgroundColor: '#212b46',
                  color: 'white',
                }
                ,
                rootPressed: {
                  backgroundColor: '#212b46', // Custom click color
                  color: 'white',
                },
                rootDisabled: {
                  backgroundColor: 'transparent', // Change to your preferred disabled background
                  color: '#6A6E79',           // Change to your preferred disabled icon/text color
                },
              }}
              iconProps={{ iconName: 'ChevronLeft' }} 
            />
          )}
          <Calendar
            value={fromdate ?? undefined}
            className="calcontrol-timesheet"
            showButtonBar
            showIcon
            // disabledDays={[0,2,3,4,5,6]}
            maxDate={today}
            minDate={thirtyDaysAgo}
            disabled={readonly}
            placeholder="From Date"
            dateFormat="dd-mm-yy"
            onClearButtonClick={() => { setFromDate(null); console.log("clear button clicked", fromdate) }}
            onChange={(e) => {
              if (e.target.value) {
                const fromDateObj = new Date(e.target.value);
                setTableData(prev => {
                  const updated = {
                    ...prev,
                    fromdate: fromDateObj.toISOString().split('T')[0],
                    todate: '', // reset todate when fromdate changes
                  };
                  onTableUpdate(updated);
                  return updated;
                })
                // tableConfig.fromdate = fromDateObj.toISOString().split('T')[0];
                setFromDate(fromDateObj);
                setToDate(null); // reset toDate when fromDate changes
                // tableConfig.todate = ''; // clear the stored todate
              }
            }}
          />

          <Calendar
            className="calcontrol-timesheet"
            value={todate ?? undefined}
            showIcon
            showButtonBar
            placeholder="To Date"
            disabled={!fromdate || readonly}
            minDate={fromdate ?? undefined}
            maxDate={maxToDate ?? undefined}
            onClearButtonClick={() => { setToDate(null); console.log("clear button clicked", todate) }}
            onChange={(e) => {
              if (e.target.value) {
                const toDateObj = new Date(e.target.value);
                setTableData(prev => {
                  const updated = {
                    ...prev,
                    todate: toDateObj.toISOString().split('T')[0], // reset todate when fromdate changes
                  };
                  onTableUpdate(updated);
                  return updated;
                })
                setToDate(toDateObj);
              }
            }}
            dateFormat="dd-mm-yy"
          />

          {(fromdate && todate && !readonly &&
            <IconButton
              title="Toggle Week"
              onClick={() => {
                const newfrom = new Date(fromdate.setDate(fromdate.getDate() + 7));
                const newto = new Date(todate.setDate(todate.getDate() + 7));
                setFromDate(newfrom);
                setToDate(newto);
                console.log("setfromdate", fromdate.setDate(fromdate.getDate() + 7))
              }}
              disabled={today.getTime() - todate.getTime() <= 7 * 24 * 60 * 60 * 1000}
              styles={{
                root: { height: '29px', width: '29px', color: 'white' },
                rootHovered: {
                  backgroundColor: '#212b46',
                  color: 'white',
                }
                ,
                rootPressed: {
                  backgroundColor: '#212b46', // Custom click color
                  color: 'white',
                },
                rootDisabled: {
                  backgroundColor: 'transparent', // Change to your preferred disabled background
                  color: '#6A6E79',           // Change to your preferred disabled icon/text color
                },
              }}
             iconProps={{ iconName: 'ChevronRight' }}
            />
          )}
        </div>
        {!readonly && (
          <button className='btn btn-primary addrow' onClick={addNewRow}>+ Add Row</button>
        )}
      </div>

      <div className="table-responsive">
        <table className={`dynamic-table ${readonly ? 'readonly-mode' : ''}`}>
          <thead>
            <tr>
              {tableData.columns?.map((header, index) => (
                <th
                  key={header.headername}
                  style={{
                    width: `var(--col-${header.headername}-width, auto)`,
                    position: 'relative', backgroundColor: (header.isdynamic ? (tableConfig?.weekends?.includes(header.weekday) ? '#4A4A4A' : '#161d2f') : '#161d2f')
                  }}
                >
                  <div style={{ display: 'flex', alignItems: 'center' }}>
                    {header.headername}

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
                {tableData.columns?.map((header) => {
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
                          {/* <option value="" disabled >Select</option> */}
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
                                {!readonly ? (
                                  <span style={{ marginLeft: '5px' }}>{cellValue.name}</span>
                                ) : (
                                  <a
                                    href={`${fileFetchUrl}/TableAttachments/${seqno}/${tableData.id}/${cellValue.name}`}
                                    target="_blank"
                                    rel="noopener noreferrer"
                                    onClick={(e) => e.stopPropagation()}
                                    style={{ marginLeft: '5px' }}
                                  >
                                    {cellValue.name}
                                  </a>
                                )}
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
                            min={0}
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
                              onClick={(e) => handleInfoClick(e as React.MouseEvent<HTMLElement>, header, row, cellId)}
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
            {(
              <tr className="sum-row">
                {tableData.columns?.map((header, index) => {
               const sum=calculateSum(header.headername).toFixed(2);
               return <td 
                key={`sum-${header.headername}`}                       
                className={ header.isdynamic && Number(sum) >= 24 ? 'error-cell' : ''}
                >
                    {(header.columntype === 'Number' && header.isdynamic) ? (
                      <strong>{calculateSum(header.headername).toFixed(2)} Hrs</strong>
                    ) : null}
                  </td>
                })}
                {!readonly && <td></td>}
              </tr>
            )}
          </tbody>
        </table>
      </div>
      <div style={{ display: 'flex', justifyContent: 'flex-end' }}>
        <strong>Total: {totalsum()} Hrs</strong>
      </div>
      {calloutTarget && calloutCellId && (
        <Callout
          target={calloutTarget}
          onDismiss={dismissCallout}
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
            <IconButton
              iconProps={{ iconName: 'Cancel' }}
              onClick={dismissCallout}
              styles={{ root: { height: '20px' } }}
            />
          </div>
          {tableData.rows?.map(row =>
            tableData.columns?.map(header => {
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
      {showTableInfoTooltip && tableInfoTooltipTarget && (
        <Callout
          target={tableInfoTooltipTarget}
          onDismiss={() => setShowTableInfoTooltip(false)}
          directionalHint={DirectionalHint.topCenter}
          styles={{
            root: {
              backgroundColor: '#212b46',
              color: 'black',
              border: '1px solid #3a455c',
              borderRadius: '4px',
              padding: '8px 12px',
              boxShadow: '0 2px 4px rgba(0,0,0,0.2)'
            }
          }}
        >
          <div style={{ display: 'flex', flexDirection: 'column' }}>

            <p><strong>Hours Restrictions Per Cell</strong></p>
            <ul>
              <li>Minimum: {tableData.minHoursPerCell}</li>
              <li>Maximum: {tableData.maxHoursPerCell}</li>
            </ul>

          </div>
        </Callout>
      )}

      {/* <div className="alert alert-info">
        Has Restriction Breaches: {hasRestrictionBreaches ? 'Yes' : 'No'}
      </div> */}
    </div>
  );
};

export default TimesheetComponent;