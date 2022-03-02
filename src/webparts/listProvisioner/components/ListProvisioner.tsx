import * as React from 'react';

import { SPHttpClient, 
  SPHttpClientResponse,
  ISPHttpClientOptions
} from '@microsoft/sp-http';

import { 
  WebPartContext 
} from '@microsoft/sp-webpart-base';

import styles from './ListProvisioner.module.scss';
import { IListProvisionerProps } from './IListProvisionerProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class ListProvisioner extends React.Component<IListProvisionerProps, {}> {

  private provisionList() : void{
    const getListUrl: string =  this.props.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('My List')";

    this.props.context.spHttpClient.get(getListUrl, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
      if (response.status === 200) {
        alert("List already exists.");
        return; // list already exists
      }
      if (response.status === 404) {
        const url: string = this.props.context.pageContext.web.absoluteUrl + "/_api/web/lists";

        const listDefinition : any = {
          "Title": "My List",
          "Description": "My description",
          "AllowContentTypes": true,
          "BaseTemplate": 100,
          "ContentTypesEnabled": true
        };

        const spHttpClientOptions: ISPHttpClientOptions = {
          "body": JSON.stringify(listDefinition)
        };

        this.props.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
        .then((response: SPHttpClientResponse) => {
          if (response.status === 201) {
            alert("List created successfully");
          } 
          else {
            alert("Response status "+response.status+" - "+response.statusText);
          }
        });
      } else {
        alert("Something went wrong. "+response.status+" "+response.statusText);
      }
    });

    console.log(getListUrl);
  }

  public render(): React.ReactElement<IListProvisionerProps> {
    return (
      <div className={ styles.listProvisioner }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Click the button to check if a list exists, if it doesnt, its created.</span>
              <p className={ styles.description }>{escape(this.props.context.pageContext.web.absoluteUrl)}</p>
              <br />
              <button onClick={() => this.provisionList()}>Click me!</button>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
