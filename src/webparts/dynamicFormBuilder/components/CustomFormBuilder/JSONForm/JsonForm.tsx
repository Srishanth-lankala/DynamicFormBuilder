import React, { useState, useEffect } from "react";
import { sp } from "@pnp/sp";
import { useGlobalState } from "../GlobalVariable/GlobalStateContext";
import { dataservice } from "../encryptionutil";

interface FormField {
  id: string;
  element: string;
  text: string;
  group_name: string;
  static: boolean;
  required: boolean;
  canHaveAnswer: boolean;
  content: string;
  canHavePageBreakBefore: boolean;
  canHaveAlternateForm: boolean;
  canHaveDisplayHorizontal: boolean;
  canHaveOptionCorrect: boolean;
  canHaveOptionValue: boolean;
  canPopulateFromApi: boolean;
  field_name: string;
  label: string;
  dirty: boolean;
}

const FormRenderer: React.FC<{ json: FormField[] }> = ({ json }) => {
  const { globalVariable } = useGlobalState();
  const parsedGlobalVariable = globalVariable ? JSON.parse(globalVariable) : null;
  const FormCode = parsedGlobalVariable.appcode;
  const FormName = parsedGlobalVariable.appname;

  return (
    <div>
      <h2>{FormCode}</h2>
      <h2>{FormName}</h2>
      <form>
        {json.map((field) => {
          switch (field.element) {
            case "TextInput":
              return (
                <div key={field.id} style={{ marginBottom: "10px" }}>
                  <label htmlFor={field.field_name}>{field.label}</label>
                  <input
                    type="text"
                    id={field.field_name}
                    name={field.field_name}
                    placeholder={field.content}
                  />
                </div>
              );
            case "NumberInput":
              return (
                <div key={field.id} style={{ marginBottom: "10px" }}>
                  <label htmlFor={field.field_name}>{field.label}</label>
                  <input
                    type="number"
                    id={field.field_name}
                    name={field.field_name}
                    placeholder={field.content}
                  />
                </div>
              );
            default:
              return null;
          }
        })}
      </form>
    </div>
  );
};

const fetchJsonData = async (formCode: string): Promise<FormField[]> => {
 
  const dataserviceobj = new dataservice();

  try {
    const RAWitems = await sp.web.lists
      .getByTitle("FormMaster")
      .items.filter("AppCode eq 'EMP-77HT0I9'") // Use dynamic filter
      .select("FormJSON")
      .get();
      const items = RAWitems?.map((item:any)=>{
        return { ...item, FormJSON: dataserviceobj.decryptjson(item.FormJSON)};
    })
    if (items.length > 0 && items[0].FormJSON) {
      const json = JSON.parse(items[0].FormJSON);
      return json.map((item: any) => ({
        id: item.id,
        element: item.element,
        text: item.text,
        content: item.content,
        field_name: item.field_name,
        label: item.label,
        group_name: "",
        static: false,
        required: false,
        canHaveAnswer: true,
        canHavePageBreakBefore: true,
        canHaveAlternateForm: true,
        canHaveDisplayHorizontal: true,
        canHaveOptionCorrect: true,
        canHaveOptionValue: true,
        canPopulateFromApi: true,
        dirty: false,
      }));
    }

    return [];
  } catch (error) {
    console.error("Error fetching data from SharePoint:", error);
    return [];
  }
};

const App: React.FC = () => {
  const [formData, setFormData] = useState<FormField[]>([]);
  const { globalVariable } = useGlobalState();
  const parsedGlobalVariable = globalVariable ? JSON.parse(globalVariable) : null;
  const FormCode = parsedGlobalVariable?.appcode || "";

  useEffect(() => {
    const loadData = async () => {
      try {
        if (FormCode) {
          const data = await fetchJsonData(FormCode); // Pass FormCode here
          setFormData(data);
        }
      } catch (error) {
        console.error("Error loading form data:", error);
      }
    };

    void loadData();
  }, [FormCode]);

  return (
    <div>
      <FormRenderer json={formData} />
    </div>
  );
};

export default App;
