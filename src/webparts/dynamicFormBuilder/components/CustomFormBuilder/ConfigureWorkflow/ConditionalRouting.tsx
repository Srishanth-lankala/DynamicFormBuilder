import React, { useState, useEffect, useCallback } from 'react';
import { Trash2 } from 'lucide-react';
import '../tree.css';
import DeleteCondition from '../../../assets/Images/ConfigureWorkflowIcons/DeleteCondition.svg';
import 'bootstrap/dist/css/bootstrap.min.css';
// import { sp } from '@pnp/sp';
// import '@pnp/sp/webs';
// import '@pnp/sp/lists';
// import '@pnp/sp/items'; 
// import { BaseWebPartContext } from '@microsoft/sp-webpart-base';

// Types for condition and group 
interface SingleConditionType {
  id: string;
  type: 'single';
  field: string;
  operator: string;
  value: string;
  elementType: string;
}
interface IConditionBuilder {
  filteredjson: [],
  onDataSend: any,
  databool: boolean,
  initialConditions?: ConditionGroupType | null
}

interface IOperatorsJson {
  elementType: string,
  Operators: string[]
}

interface ConditionGroupType {
  id: string;
  type: 'group';
  operator: string;
  conditions: Condition[];
}

function stripHtmlTags(label: any) { return label.replace(/<\/?[^>]+(>|$)/g, "").trim() || ""; }

type Condition = SingleConditionType | ConditionGroupType;

// Helper function to count total conditions
const countTotalConditions = (group: ConditionGroupType): number => {
  return group.conditions.reduce((total, condition) => {
    if (condition.type === 'single') {
      return total + 1;
    }
    return total + countTotalConditions(condition as ConditionGroupType);
  }, 0);
};



interface OperatorSelectProps {
  value: string;
  onChange: (value: string) => void;
}

// const OperatorSelect: React.FC<OperatorSelectProps> = ({ value, onChange }) => {
//   const [isOpen, setIsOpen] = useState(false);

//   return (
//     <div className="position-relative">
//       <button
//         onClick={() => setIsOpen(!isOpen)}
//         style={{backgroundColor:'var(--color-navigation)',height:'25px',fontSize:'13px',fontWeight:'400'}}
//         type="button"
//         className="btn btn-dark d-flex align-items-center"
//       >
//         {value}
//         <ChevronDown size={16} />
//       </button>

//       {isOpen && (
//         <>
//           <div
//             className="dropdown position-fixed top-0 start-0 w-100 h-100"
//             onClick={() => setIsOpen(false)}
//           />
//           <div className="dropdown-menu show position-absolute" style={{padding:2,backgroundColor:'var(--color-navigation)'}}>

//             {['And', 'Or'].map((op) => (
//               <button
//                 key={op}
//                 className="dropdown-item"
//                 style={{color:'var(--color-white)'}}
//                 onClick={() => {
//                   onChange(op);
//                   setIsOpen(false);
//                 }}
//               >
//                 {op}
//               </button>
//             ))}
//           </div>
//         </>
//       )}
//     </div>
//   );
// };

// Props for SingleCondition

const OperatorSelect: React.FC<OperatorSelectProps> = ({ value, onChange }) => {
  return (
    <select
      value={value}
      onChange={(e) => onChange(e.target.value)}
      className="form-select operator-select"
    >
      {['And', 'Or'].map((op) => (
        <option key={op} value={op}>
          {op}
        </option>
      ))}
    </select>
  );
};

interface SingleConditionProps {
  condition: SingleConditionType;
  onDelete: () => void;
  onChange: (updatedCondition: SingleConditionType) => void;
  FormFields: []
}

const SingleCondition: React.FC<SingleConditionProps> = ({
  condition,
  onDelete,
  onChange,
  FormFields,
}) => {
  const [operatorOptions, setOperatorOptions] = useState<string[]>();

  const OperatordDD: IOperatorsJson[] = [
    {
      elementType: "TextInput",
      Operators: ["EqualsTo", "Contains", "StartsWith"],
    },
    {
      elementType: "NumberInput",
      Operators: [
        "EqualsTo",
        "Not EqualsTo",
        "GreaterThan",
        "LessThan",
        "GreaterThan or EqualsTo",
        "LessThan or EqualsTo",
      ],
    },
    {
      elementType: "Dropdown",
      Operators: [
        "EqualsTo",
        "Not EqualsTo"
      ]
    }
  ];

  // Function to update operator options based on field selection
  const operatorSelection = (fieldType: string) => {
    const json: any = FormFields.find((ele: any) => ele.label.includes(fieldType)) ///////json has single json data about a field in the form
    console.log("dfgjklhfgjkhxcjkcv", json)

    const selectedOperator = OperatordDD.find(
      (item) => item.elementType === json.element
    );
    console.log("selectedopppppp", selectedOperator)
    console.log("fieldtypee", fieldType)
    if (selectedOperator) {
      setOperatorOptions(selectedOperator.Operators);
    }
  };

  useEffect(() => {
    if (condition.field) {
      operatorSelection(condition.field);
    }
  }, [condition.field, FormFields]);

  return (
    <div className="d-flex align-items-center gap-2 position-relative" >
      <select
        className="form-select newDDSelect"
        value={condition.field}
        onChange={(e) => {
          const json: any = FormFields.find((ele: any) => ele.label.includes(e.target.value))
          const selectedOperator = OperatordDD.find(
            (item) => item.elementType === json.element
          );
          onChange({ ...condition, field: e.target.value, elementType: selectedOperator?.elementType || '' });
          console.log("dddddddd", e)
          operatorSelection(e.target.value); // Update operator options based on field
        }}
      >
        <option value="">Choose a value</option>
        {FormFields.map((ele: any) => (
          <option key={ele.id} value={ele.type}>
            {stripHtmlTags(ele.label)}
          </option>
        ))}
      </select>

      <select
        // style={{
        //   lineHeight: 1.3, border: '1px solid #6A6E79', backgroundColor: '#161D2F', color: 'white'
        // }}
        className="form-select newDDSelect"
        value={condition.operator}
        onChange={(e) => onChange({ ...condition, operator: e.target.value })}
      >
        <option value="">Choose an operator</option>
        {operatorOptions && operatorOptions?.map((ele: string) => (
          <option key={ele} value={ele}>
            {ele}
          </option>
        ))}
      </select>

      <input
      style={{marginTop:0}}
        type="text"
        className="form-control  newDDSelect"
        placeholder="Choose a value"
        value={condition.value}
        onChange={(e) => onChange({ ...condition, value: e.target.value })}
      />

      <button className="btn p-2 d-flex align-items-center" onClick={onDelete}>
        <img className="buttonicon" style={{ height: 16, width: 16 }} src={DeleteCondition} alt='Delete'></img>
      </button>
    </div>
  );
};


interface ConditionGroupProps {
  group: ConditionGroupType;
  onUpdate: (updatedGroup: ConditionGroupType) => void;
  onDelete?: () => void;
  level?: number;
  FormFieldsCon: []
}

const ConditionGroup: React.FC<ConditionGroupProps> = ({
  group,
  onUpdate,
  onDelete,
  level = 0,
  FormFieldsCon
}) => {
  const totalConditions = countTotalConditions(group);
  const limit = 4;
  const isLimitReached = totalConditions >= limit;

  const addCondition = () => {
    const newCondition: SingleConditionType = {
      id: Math.random().toString(36).substr(2, 9),
      type: 'single',
      field: '',
      operator: '',
      value: '',
      elementType: ''
    };
    onUpdate({
      ...group,
      conditions: [...group.conditions, newCondition],
    });
  };

  //   const addNestedGroup = () => {
  //     const newGroup: ConditionGroupType = {
  //       id: Math.random().toString(36).substr(2, 9),
  //       type: 'group',
  //       operator: 'And',
  //       conditions: [],
  //     };
  //     onUpdate({
  //       ...group,
  //       conditions: [...group.conditions, newGroup],
  //     });
  //   };

  const deleteCondition = (conditionId: string) => {
    const newConditions = group.conditions.filter((c) => c.id !== conditionId);
    onUpdate({ ...group, conditions: newConditions });
  };

  const updateCondition = (conditionId: string, updatedCondition: Condition) => {
    const newConditions = group.conditions.map((c) =>
      c.id === conditionId ? updatedCondition : c
    );
    onUpdate({ ...group, conditions: newConditions });
  };

  return (
    <div className={`ms-${level > 0 ? 3 : 0} mt-3`}>
      {group.conditions.length > 1 && (
        <div className="d-flex align-items-center gap-2 mb-3">
          <OperatorSelect
            value={group.operator}
            onChange={(newOperator) => onUpdate({ ...group, operator: newOperator })}
          />
          {level > 0 && (
            <button
              className="btn p-2 d-flex align-items-center"
              onClick={onDelete}
            >
              <Trash2 size={16} />
            </button>
          )}
        </div>)}
      {/* single condition */}
      <div className="d-flex flex-column gap-3">
        {group.conditions.map((condition) => (
          <div key={condition.id}>
            {condition.type === 'single' ? (
              <SingleCondition
                condition={condition as SingleConditionType}
                onDelete={() => deleteCondition(condition.id)}
                onChange={(updatedCondition) =>
                  updateCondition(condition.id, updatedCondition as SingleConditionType)

                }
                FormFields={FormFieldsCon}
              />
            ) : (
              <ConditionGroup
                group={condition as ConditionGroupType}
                onUpdate={(updatedGroup) =>
                  updateCondition(condition.id, updatedGroup as ConditionGroupType)
                }
                onDelete={() => deleteCondition(condition.id)}
                level={level + 1}
                FormFieldsCon={FormFieldsCon}
              />
            )}
          </div>
        ))}
      </div>

      <div className=" modal-buttons d-flex gap-2 mt-3">
        <button
          // style={{ height: '32px', fontSize: '14px', fontWeight: '400' }}
          onClick={addCondition}
          disabled={isLimitReached}
          className={`${isLimitReached ? 'btn-secondary' : 'newlogocolorbtn'}`}
        >
          + Add Condition
        </button>
        {isLimitReached && (
          <span className="text-danger small" style={{ marginTop: '6px' }}>
            Maximum limit of {limit} conditions reached
          </span>
        )}
      </div>
    </div>
  );
};

const ConditionBuilder: React.FC<IConditionBuilder> = ({ filteredjson, onDataSend, databool, initialConditions }) => {
  const [rootGroup, setRootGroup] = useState<ConditionGroupType>(() => {
    if (initialConditions) {
      return initialConditions;
    }
    return {
      id: 'root',
      type: 'group',
      operator: 'And',
      conditions: [],
    };
  });
  useEffect(() => {
    if (initialConditions) {
      setRootGroup(initialConditions);
    }
  }, []);
  const getJsonOutput = useCallback(() => {
    return JSON.stringify(rootGroup, null, 2);
  }, [rootGroup]);
  const sendDatatoparent = () => {
    console.log("____rootGroup-----", rootGroup)
    onDataSend(rootGroup)
  }
  if (databool) {
    sendDatatoparent()
  }
  getJsonOutput()

  return (
    <div className="container mt-1">
      <div className="card mb-4" style={{ backgroundColor: 'var(--color-other-bg)' }}>
        <div className="card-body " style={{ padding: '0px 0px 0px 7px' }}>
          <ConditionGroup group={rootGroup} onUpdate={setRootGroup} FormFieldsCon={filteredjson} />
        </div>
      </div>
      {/* <div className="card bg-light">
        <div className="card-body">
          <h5 className="card-title">JSON Output:</h5>
          <pre className="bg-dark text-light p-3 rounded">
            {getJsonOutput()}
          </pre>
        </div>
      </div> */}
    </div>
  );
};
export default ConditionBuilder;