import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import { AadHttpClient, HttpClientResponse } from "@microsoft/sp-http";

import styles from './AzureFuncDemoWebPart.module.scss';
import * as strings from 'AzureFuncDemoWebPartStrings';

export interface IAzureFuncDemoWebPartProps {
  description: string;
}

interface IProduct {
  id: number;
  name: string;
  qty: number;
  price: number;
}

export default class AzureFuncDemoWebPart extends BaseClientSideWebPart<IAzureFuncDemoWebPartProps> {
  private aadClient : AadHttpClient;
  
  protected onInit() : Promise<void> {
    return this.context.aadHttpClientFactory.getClient('3d55f89c-0686-4523-b857-e0da5570c37b')
      .then((client : AadHttpClient) => {
        this.aadClient = client;
        return Promise.resolve();
      })
      .catch(err=> {
        console.log("Error getting AadHttpClient : " +err);
      });
  }

  private getProducts() : Promise<IProduct[]> {
    if(!this.aadClient) {
      throw new Error('AadHttpClient not initialized!');
    }
    return this.aadClient.get('https://erkindunya.azurewebsites.net/api/GetProducts',AadHttpClient.configurations.v1)
      .then((res: HttpClientResponse) =>{
        return res.json();
      })
      .then((json: any) => {
        return json as IProduct[];
      })
      .catch(err=> {
        console.log("Error fetching products:" + err);

        return [];
      });
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.azureFuncDemo }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Azure Function Demo!</span>
              <p class="${ styles.subTitle }">SPFx AadHttpClient call to Azure Function API</p>
              <div id="output">
                Loading...
              </div>
            </div>
          </div>
        </div>
      </div>`;

      try {
        this.getProducts()
          .then((items: IProduct[]) =>
          {
              let html = "";

              for(let p of items) {
                console.log(JSON.stringify(p));

                html += `<div class="${ styles.product }">
                            ${ p.id } <br/>
                            ${ p.name } <br/>
                            ${ p.price } <br/>
                            ${ p.qty }
                         </div>`;
              }

              this.domElement.querySelector("#output").innerHTML = html;
          })
          .catch(err => {
            console.log("Error fetching products:" + err);
          })
      } catch(err) {
        console.log("Error occurred :" + err);
      }
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