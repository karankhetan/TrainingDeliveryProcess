import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  WebPartContext
} from '@microsoft/sp-webpart-base';

import * as strings from 'TraningManagerWebPartStrings';
import TraningManager from './components/TraningManager';
import { ITraningManagerProps } from './components/ITraningManagerProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface ITraningManagerWebPartProps {
  description: string;
}

export default class TraningManagerWebPart extends BaseClientSideWebPart<ITraningManagerWebPartProps> {

  public render(): void {

    const element: React.ReactElement<ITraningManagerProps> = React.createElement(
      TraningManager,
      {
        description: this.properties.description,
        getlistItem: this.getTheListItems.bind(this),
        DeleteListItem: this.DeleteListItem.bind(this),
        context:this.context
      }
    );
    ReactDom.render(element, this.domElement);



  }

  public DeleteListItem(id: any): Promise<any> {
    return new Promise<any>((resolve) => {

      this.context.spHttpClient.post(`${this.context.pageContext.site.absoluteUrl}/_api/web/lists/getbytitle('TraningEventList')/items(${id})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=verbose',
            'odata-version': '',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'DELETE'
          }
        })
        .then((response: SPHttpClientResponse): void => {
          resolve("success");
        }, (error: any): void => {

        });




    });

  }

  public getTheListItems(): Promise<any[]> {

    return new Promise<any[]>((resolve, reject) => {
      this.context.spHttpClient.get(`${this.context.pageContext.site.absoluteUrl}/_api/web/lists/GetByTitle('TraningEventList')/items?$select=OData__ModerationStatus,Title,DateOfTraining,Id`, SPHttpClient.configurations.v1).then(
        (spHttpClientResponse: SPHttpClientResponse) => {
          spHttpClientResponse.json().then(
            (jsonresponse: any) => {
              resolve(jsonresponse.value)
            }
          );
        }
      );
    });
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
