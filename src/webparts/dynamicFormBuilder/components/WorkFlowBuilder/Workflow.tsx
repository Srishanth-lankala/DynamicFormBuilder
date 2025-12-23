import * as React from "react";
import "react-form-builder2/dist/app.css";
import type { IDynamicFormBuilderProps } from "../IDynamicFormBuilderProps";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { sp } from "@pnp/sp/presets/all";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface RRJson {
  Id: number;
  Actions: string;
  Status: string;
  CurrentQueue: string | null;
  DestinationQueue: string | null;
  ID: number;
}
 export interface MMJson{
  Role:string;
  Level:number;
  AppCode:string;
  AppName:string;
  ID:number;
 }
export interface IWorkflowProps {
  RR: RRJson[];
 MM: MMJson[];
  context: WebPartContext;
}

export default class Workflow extends React.Component<IWorkflowProps> {
  private iframeRef = React.createRef<HTMLIFrameElement>();

  constructor(props: IWorkflowProps) {
    super(props);
  }

  componentDidMount() {
    SPComponentLoader.loadCss(
      "https://use.fontawesome.com/releases/v5.13.0/css/all.css"
    );
    SPComponentLoader.loadCss(
      "https://stackpath.bootstrapcdn.com/bootstrap/4.4.1/css/bootstrap.min.css"
    );

    sp.setup({
      spfxContext: this.props.context as any,
    });
  }

  sendDataToIframe = () => {
    const sampleData = {
      type: "FROM_SPFX",
      MM:this.props.MM,
      payload: this.props.RR
    };

    if (this.iframeRef.current && this.iframeRef.current.contentWindow) {
      this.iframeRef.current.contentWindow.postMessage(sampleData, "*");
      console.log("Data sent toiframe:", sampleData);
    } 
  };   

  public render(): React.ReactElement<IDynamicFormBuilderProps> {
    return ( 
      <section>
        <div>
          <iframe
            ref={this.iframeRef}
            src="https://smartofficenxtbuilder.cloudangles.com/workflow/"
            style={{
              border: "none",
              display: "block",
              width: "100%",
              height: "100vh",
            }}
            title="Embedded Application"
            onLoad={this.sendDataToIframe} // Send data once iframe loads
          ></iframe>
        </div>
      </section>
    );
  }
}