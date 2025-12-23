import * as React from "react";

export interface IListItemProps{
    Element:any

}

export default class ListItem extends React.Component<IListItemProps>{
    render(){
        return(
            <>
                <p>{this.props.Element.Title}</p> 
            </>
        )
    }
}