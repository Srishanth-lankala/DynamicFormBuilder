import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs"
import "@pnp/sp/lists"
import "@pnp/sp/items"
import ListItem from "./ListItem";
import * as React from "react";
import { IDynamicFormBuilderProps } from "../IDynamicFormBuilderProps";
// import { ReactFormGenerator } from 'react-form-builder2';
// export interface IListProps{
//     Title:string
// }
export interface IListstate{
    allItems:any[]
}

export default class List extends React.Component<IDynamicFormBuilderProps,IListstate>{
    constructor(props:IDynamicFormBuilderProps){
        super(props)
        this.state={
            allItems:[]
        }
        this.getAllListitems=this.getAllListitems.bind(this)
    }
 
    componentDidMount(): void {
        sp.setup({
            spfxContext:this.props.context as any
        })
        this.getAllListitems()
    }

    getAllListitems=async()=>{
        const all=await sp.web.lists.getByTitle("Forms").items.get()
        console.log(all)
        this.setState({
            allItems:all 
        })
        console.log(this.state.allItems)
        const x=this.state.allItems
        console.log(x)
       return x
        
    }

    

    render(){
        return(
            <>
                <h1>List of Forms</h1>
                <button className="btn btn-primary" onClick={async()=>{
                    const x=await this.getAllListitems();
                    console.log("inside onclick",x)
                }}>Dummy</button>
                <h1>this is form</h1>
                <div>
                    
                    {
                    this.state.allItems.map((element,index)=>{
                        return <ListItem key={index} Element={element} />
                        
                    })
                }
                </div>
                {/* <div>
          {this.state.allItems.length > 0 ? (
            this.state.allItems.map((element, index) => (
              <ListItem key={index} Title={element.Title} />
            ))
          ) : (
            <p>No items to display</p>
          )}
        </div> */}


            </>
        )
    }
}

// import { sp } from "@pnp/sp/presets/all";
// import "@pnp/sp/webs";
// import "@pnp/sp/lists";
// import "@pnp/sp/items";
// import * as React from "react";
// import { IDynamicFormBuilderProps } from "../IDynamicFormBuilderProps";
// import { ReactFormGenerator } from 'react-form-builder2';

// interface IListState {
//   formJson: any[];
// }

// export default class List extends React.Component<IDynamicFormBuilderProps, IListState> {
//   constructor(props: IDynamicFormBuilderProps) {
//     super(props);
//     this.state = {
//       formJson: [], // Initialize with an empty array
//     };
//   }

//   componentDidMount(): void {
//     sp.setup({
//       spfxContext: this.props.context as any,
//     });

//     this.getAllListItems(); // Fetch items when component mounts
//   }

//   // Fetch the JSON from SharePoint list and set state
//   getAllListItems = async () => {
//     try {
//       const allItems = await sp.web.lists.getByTitle("Forms").items.get();
//       if (allItems.length > 0) {
//         // Assuming 'form_data' is the field containing your form JSON
//         const formData = allItems[0].FormJSON;
//         this.setState({ formJson: JSON.parse(formData) });
//       }
//     } catch (error) {
//       console.error("Error fetching list items", error);
//     }
//   };

//   render() {
//     const { formJson } = this.state;

//     return (
//       <>
//         <h1>List of Forms</h1>
//         <button className="btn btn-primary" onClick={this.getAllListItems}>
//           Fetch Form Data
//         </button>

//         <h2>Generated Form</h2>

//         {formJson.length > 0 ? (
//           <ReactFormGenerator
//             data={formJson}
//             form_action="/submit"
//             form_method="POST"
//             hide_actions={true} // Hides the submit/reset buttons
//           />
//         ) : (
//           <p>No form data available</p>
//         )}
//       </>
//     );
//   }
// }
