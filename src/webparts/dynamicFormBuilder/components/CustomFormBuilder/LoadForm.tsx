import * as React from 'react';
import { ILoadFormProps } from './ILoadFormProps'; // Import your props interface

const LoadForm: React.FC<ILoadFormProps> = (props) => {
  const { formData } = props;

  const renderField = (field: any) => {
    switch (field.element) {
      case "TextInput":
        return (
          <div key={field.id}>
            <label htmlFor={field.field_name}>{field.label}</label>
            <input
              type="text"
              id={field.field_name}
              name={field.field_name}
              required={field.required}
            />
          </div>
        );

      case "TextArea":
        return (
          <div key={field.id}>
            <label htmlFor={field.field_name}>{field.label}</label>
            <textarea
              id={field.field_name}
              name={field.field_name}
              required={field.required}
            />
          </div>
        );

        case "Dropdown":
          return (
            <div key={field.id}>
              <label htmlFor={field.field_name}>{field.label}</label>
              <select id={field.field_name} name={field.field_name} required={field.required}>
                {field.options.map((option: any) => (
                  <option key={option.key} value={option.value}>
                    {option.text}
                  </option>
                ))}
              </select>
            </div>
          );
        

          case "Checkboxes":
            return (
              <div key={field.id}>
                <label>{field.label}</label>
                {field.options.map((option: any) => (
                  <div key={option.key}>
                    <label>
                      <input
                        type="checkbox"
                        id={`${field.field_name}_${option.value}`}
                        name={field.field_name}
                        value={option.value}
                        required={field.required}
                      />
                      {option.text}
                    </label>
                  </div>
                ))}
              </div>
            );
          
          case "RadioButtons":
            return (
              <div key={field.id}>
                <label>{field.label}</label>
                {field.options.map((option: any) => (
                  <div key={option.key}>
                    <label>
                      <input
                        type="radio"
                        name={field.field_name}
                        value={option.value}
                        required={field.required}
                      />
                      {option.text}
                    </label>
                  </div>
                ))}
              </div>
            );
          

      

      default:
        return null; // Return null for unrecognized elements
    }
  };

  const renderColumnFields = (fields: any[]) => {
    return fields.map((subField, index) => (
      <div key={subField.id} style={{ flex: 1, marginRight: index < fields.length - 1 ? '10px' : '0' }}>
        {renderField(subField)}
      </div>
    ));
  };

  return (
    <>
      {/* {console.log(formData)} */}
      <form>
        {/* {title && <h2>{title}</h2>} */}
        {formData.map((field) => {
          switch (field.element) {
            case "TextInput":
            case "TextArea":
            case "Dropdown":
            case "Checkboxes":
            case "RadioButtons":
              return renderField(field);

            case "TwoColumn":
              return (
                <div key={field.id} style={{ display: 'flex' }}>
                  {renderColumnFields(field.fields)}
                </div>
              );

            case "ThreeColumn":
              return (
                <div key={field.id} style={{ display: 'flex' }}>
                  {renderColumnFields(field.fields)}
                </div>
              );

            case "FourColumn":
              return (
                <div key={field.id} style={{ display: 'flex' }}>
                  {renderColumnFields(field.fields)}
                </div>
              );

            default:
              return null; // Return null for unrecognized elements
          }
        })}
        {/* <button type="submit">Submit</button> */}
      </form>
    </>
  );
};

export default LoadForm;