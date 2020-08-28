import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './GraphDemoWebPart.module.scss';
import * as strings from 'GraphDemoWebPartStrings';

import { AadHttpClient } from "@microsoft/sp-http";

export interface IGraphDemoWebPartProps {
  description: string;
}

export default class GraphDemoWebPart extends BaseClientSideWebPart<IGraphDemoWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.graphDemo }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Graph API Output!</span>
              <p class="${ styles.subTitle }">Information returned via Graph API.</p>
              <p class="${ styles.description }" id="userinfo">
                Loading...
              </p>
            </div>
          </div>
        </div>
      </div>`;

      this.getUserInfo();
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

  private getUserInfo() {
    this.context.aadHttpClientFactory.getClient("https://graph.microsoft.com")
      .then((client: AadHttpClient) => {
        client.get("https://graph.microsoft.com/v1.0/me",AadHttpClient.configurations.v1)
          .then(resp =>{
              return resp.json();
          })
          .then(data => {
              let html : string = `<div>
                Name: ${ data.displayName } <br/>
                Email: ${ data.email }
              </div>`;

            this.domElement.querySelector("#userinfo").innerHTML = html;
          })
          .catch(err=> {
            this.domElement.querySelector("#userinfo").innerHTML = `<div>Error getting user info: ${err}</div>`;
          });
      })
  }
}
