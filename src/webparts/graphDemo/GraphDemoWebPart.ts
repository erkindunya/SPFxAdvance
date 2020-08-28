import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './GraphDemoWebPart.module.scss';
import * as strings from 'GraphDemoWebPartStrings';

import { AadHttpClient, MSGraphClient } from "@microsoft/sp-http";

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
              <h3>Current User</h3>
              <p class="${ styles.description }" id="userinfo">
                Loading...
              </p>
              <h3>List of Users</h3>
              <p class="${ styles.description }" id="users">
                Loading...
              </p>  
            </div>
          </div>
        </div>
      </div>`;

      this.getUserInfo(); // Uses AADHttpClient

      this.getAllUsers(); // Uses MS Graph Client
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
                Email: ${ data.mail }
              </div>`;

            this.domElement.querySelector("#userinfo").innerHTML = html;
          })
          .catch(err=> {
            this.domElement.querySelector("#userinfo").innerHTML = `<div>Error getting user info: ${err}</div>`;
          });
      })
  }

  private getAllUsers() {
    this.context.msGraphClientFactory.getClient()
      .then((client: MSGraphClient) => {
          client.api("users")
            .version("v1.0")
            .select("displayName,mail,userPrincipalName")
            .get((err,res) => {
              if(err) {
                console.log("Error fetching users :"  +err);
                return;
              }

              let html = "";

              for(let u of res.value) {
                html += `<div>
                  ${ u.displayName } <br/>
                  ${ u.mail } <br/>
                  ${ u.userPrincipalName }
                </div>`;

                this.domElement.querySelector("#users").innerHTML = html;
              }
            });
      })
  }
}