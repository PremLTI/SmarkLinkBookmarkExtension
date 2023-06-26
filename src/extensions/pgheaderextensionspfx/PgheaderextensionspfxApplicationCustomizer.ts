import { Log } from '@microsoft/sp-core-library';
import {
  PlaceholderContent,  
  PlaceholderName,
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
// import { Dialog } from '@microsoft/sp-dialog';
import * as strings from 'PgheaderextensionspfxApplicationCustomizerStrings';
import * as React from "react";  
import * as ReactDOM from "react-dom";  
import ReactHeader, { IReactHeaderProps } from "./ReactHeader"
// import ReactFooter, { IReactFooterProps } from "./ReactFooter";  
const LOG_SOURCE: string = 'PgheaderextensionspfxApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IPgheaderextensionspfxApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
   Bottom: string; 
  Top:string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class PgheaderextensionspfxApplicationCustomizer
  extends BaseApplicationCustomizer<IPgheaderextensionspfxApplicationCustomizerProperties> {

private _topPlaceholder: PlaceholderContent | undefined;

  public onInit(): Promise<void> {
  
     Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);  
  
    // Added to handle possible changes on the existence of placeholders.  
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);  
      
    // Call render method for generating the HTML elements.  
    this._renderPlaceHolders();  
  
    return Promise.resolve();
  }


private _renderPlaceHolders(): void {  
  console.log('Available placeholders: ',  
  this.context.placeholderProvider.placeholderNames.map(name => PlaceholderName[name]).join(', '));  
     

 // Handling the top placeholder
 if (!this._topPlaceholder) {
  this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
    PlaceholderName.Top,
    { onDispose: this._onDispose }
  );

  // The extension should not assume that the expected placeholder is available.
  if (!this._topPlaceholder) {
    console.error("The expected placeholder (Top) was not found.");
    return;
  }

 if (this.properties) {
    let topString: string = this.properties.Top;
    if (!topString) {
      topString = "(Top property was not defined.)";
    }
    // creating element of ReactHeader on Extension
    if (this._topPlaceholder.domElement){
      try {
       {
        const elem: React.ReactElement<IReactHeaderProps> = React.createElement(ReactHeader,
          {
            context: this.context
          });  
        ReactDOM.render(elem, this._topPlaceholder.domElement);  
        }
      } 
      catch (error) {
  
      }
      
    }

  } 

}
  // Handling the bottom placeholder  
  // if (!this._bottomPlaceholder) {  
  //   this._bottomPlaceholder =  
  //     this.context.placeholderProvider.tryCreateContent(  
  //       PlaceholderName.Bottom,  
  //       { onDispose: this._onDispose });  
    
  //   // The extension should not assume that the expected placeholder is available.  
  //   if (!this._bottomPlaceholder) {  
  //     console.error('The expected placeholder (Bottom) was not found.');  
  //     return;  
  //   }  

  //   const elem: React.ReactElement<IReactFooterProps> = React.createElement(ReactFooter);  
  //   ReactDOM.render(elem, this._bottomPlaceholder.domElement);      
  // }  
}

private _onDispose(): void {  
    console.log('[ReactHeaderFooterApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.');  
}  
}
